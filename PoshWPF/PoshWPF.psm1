Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase

Function New-WPFWindow {
    Param(
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage="XAML code of UI")]
        [ValidateNotNullOrEmpty()]
        [xml]$xaml
    )
    $FormattedXAML = Format-WPFXAML -xaml $xaml
    $Global:PoshWPFHashTable = [HashTable]::Synchronized(@{})
    $Global:PoshWPFHashTable.Host = $Host
    $Global:PoshWPFHashTable.xaml = $FormattedXAML
    $Global:PoshWPFHashTable.Actions = New-Object System.Collections.ArrayList
    $Global:PoshWPFHashTable.ActionsMutex = New-Object System.Threading.Mutex($false, 'ActionsMutex')
    $Global:PoshWPFHashTable.WindowShown = $false
    $Global:PoshWPFHashTable.WindowControls = @{}
    $Global:PoshWPFHashTable.ScriptDirectory = $PSScriptRoot
    $Runspace = [RunspaceFactory]::CreateRunspace()
    $Runspace.ApartmentState = 'STA'
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable('PoshWPFHashTable', $Global:PoshWPFHashTable)
    $PS = [PowerShell]::Create()
    $PS.Runspace = $Runspace
    $null = $PS.AddScript({
        $ScriptDirectory = $PoshWPFHashTable.ScriptDirectory
        . "$ScriptDirectory\PoshWPF-UI-Code.ps1"
        [xml]$xaml = $PoshWPFHashTable.xaml
        Show-WPFWindow -xaml $xaml
    })
    $Global:PoshWPFHashTable.Handle = $PS.BeginInvoke()
    $Global:PoshWPFHashTable.Runspace = $Runspace
    while(!$Global:PoshWPFHashTable.WindowShown) {
        Start-Sleep -Milliseconds 10
    }
}

Function Format-WPFXAML {
    param(
        [xml]$xaml
    )
    if($xaml.Window) {
        $Attributes = $xaml.Window.Attributes
        $AttributesToRemove = @()
        foreach($Attribute in $Attributes) {
            Switch($Attribute.LocalName) {
                'Class' {
                    $AttributesToRemove += @($Attribute.Name)
                }
                'Local' {
                    $AttributesToRemove += @($Attribute.Name)
                }
                'Ignorable' {
                    $AttributesToRemove += @($Attribute.Name)
                }
            }
        }
        foreach($Attribute in $AttributesToRemove){
            $xaml.Window.RemoveAttribute($Attribute)
        }
        $xaml
    }
    else {
        Throw 'No window object!'
    }
}

Function Add-WPFAction {
    param(
        [scriptblock]$ScriptBlock
    )
    $null = $Global:PoshWPFHashTable.ActionsMutex.WaitOne()
    $null = $Global:PoshWPFHashTable.Actions.Add($ScriptBlock)
    $null = $Global:PoshWPFHashTable.ActionsMutex.ReleaseMutex()
}

Function Invoke-WPFAction {
    param(
        [ScriptBlock]$Action
    )
    $Global:PoshWPFHashTable.WindowControls.Window.Dispatcher.Invoke([action]$Action, 'Normal')
}

Function Get-WPFControl {
    Param(
        [string]$ControlName,
        [string]$PropertyName
    )
    if($ControlName -ne 'Window') { $ControlName = "Window_$($ControlName)" }
    $Control = $Global:PoshWPFHashTable.WindowControls[$ControlName]
    if($null -ne $Control) {
        if([string]::IsNullOrEmpty($PropertyName)) {
            $PropertyHash = @{}
            $ControlProperties = ($Control | Get-Member -MemberType Property).Name
            foreach($ControlProperty in $ControlProperties) {
                $PropertyHash[$ControlProperty] = $Control."$($ControlProperty)"
            }
            $PropertyHash
        }
        else {
            $Control."$PropertyName"
        }
    }
    else {
        Throw 'Control not found!'
    }
}

Function Set-WPFControl {
    Param(
        [string]$ControlName,
        [string]$PropertyName,
        [object]$Value
    )
    if($ControlName -ne 'Window') { $ControlName = "Window_$ControlName" }
    $Guid = (New-Guid).Guid
    $Global:PoshWPFHashTable[$guid] = $Value
    $strScriptBlock = "`$PoshWPFHashTable.WindowControls['$($ControlName)'].$($PropertyName) = `$PoshWPFHashTable['$guid'];" + `
                      "`$null = `$PoshWPFHashTable.Remove('$guid')"
    $ScriptBlock = [ScriptBlock]::Create($strScriptBlock)
    Invoke-WPFAction -Action $ScriptBlock
}

Function New-WPFEvent {
    Param(
        [string]$ControlName,
        [string]$EventName,
        [scriptblock]$Action
    )
    if($ControlName -ne 'Window') { $ControlName = "Window_$ControlName" }
    $WinControls = $Global:PoshWPFHashTable.WindowControls
    $null = Register-ObjectEvent -InputObject $WinControls[$ControlName] -EventName 'Click' -Action $Action
}