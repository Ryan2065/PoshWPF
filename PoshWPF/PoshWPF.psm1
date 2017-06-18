Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase

Function New-WPFWindow {
    <#
        .SYNOPSIS
        Creates new WPF window in background thread
        
        .DESCRIPTION
        Creates new WPF window in background thread to allow you to keep using the PowerShell console
        
        .PARAMETER xaml
        XAML of window
        
        .EXAMPLE
        New-WPFWindow -XAML $XAML
        
        .NOTES
            .Author Ryan Ephgrave
    #>
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
    $Global:PoshWPFHashTable.WaitEvent = $true
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
    $Global:PoshWPFHashTable.PowerShell = $PS
    while(!$Global:PoshWPFHashTable.WindowShown) {
        Start-Sleep -Milliseconds 10
    }

    $null = New-WPFEvent -ControlName 'Window' -EventName 'Closing' -Action {
        $null = $Global:PoshWPFHashTable.PowerShell.EndInvoke($Global:PoshWPFHashTable.Handle)
        $null = $Global:PoshWPFHashTable.PowerShell.Dispose()
        $null = $Global:PoshWPFHashTable.Runspace.Close()
        $null = $Global:PoshWPFHashTable.Runspace.Dispose()
        $Global:PoshWPFHashTable.WaitEvent = $false
    }
}

Function Format-WPFXAML {
    <#
        .SYNOPSIS
        Removes Visual Studio specific XAML properties
        
        .DESCRIPTION
        Removes the properties Visual Studio adds to XAML which causes crashing outside of VS
        
        .PARAMETER xaml
        XAMl of the window
        
        .EXAMPLE
        Format-WPFXAML -XAML $xaml
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
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

Function Invoke-WPFAction {
    <#
        .SYNOPSIS
        Runs an action in the UI thread
        
        .DESCRIPTION
        Runs a scriptblock in the UI thread
        
        .PARAMETER Action
        Scriptblock to run in the UI thread
        
        .EXAMPLE
        Invoke-WPFAction -Action $Scriptblock
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
    param(
        [ScriptBlock]$Action
    )
    $Global:PoshWPFHashTable.WindowControls.Window.Dispatcher.Invoke([action]$Action, 'Normal')
}

Function Get-WPFControl {
    <#
        .SYNOPSIS
        Returns a hash of properties of the WPF control
        
        .DESCRIPTION
        Returns a hash because if you try to interact with the objects in the HashTable you'll get errors
        
        .PARAMETER ControlName
        Name of WPF control
        
        .PARAMETER PropertyName
        Name of the property you want
        
        .EXAMPLE
        Get-WPFControl -ControlName 'Window' -PropertyName 'Title'
        Only returns Title from Window

        .EXAMPLE
        Get-WPFControl -ControlName 'Window'
        Returns all properties from Window
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
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
    <#
        .SYNOPSIS
        Updates a property on a WPF control
        
        .DESCRIPTION
        Will update the property by running Invoke-WPFAction
        
        .PARAMETER ControlName
        Name of the control
        
        .PARAMETER PropertyName
        Name of the property
        
        .PARAMETER Value
        Object with the new value
        
        .EXAMPLE
        Set-WPFControl -ControlName 'Window' -PropertyName 'Title' -Value 'My new title!'
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
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
    <#
        .SYNOPSIS
        Creates an event to run in the main thread when a UI action is run in the UI thread
        
        .DESCRIPTION
        Creates an event to run in the main thread when a UI action is run in the UI thread
        
        .PARAMETER ControlName
        Name of the control
        
        .PARAMETER EventName
        Name of the event on the control
        
        .PARAMETER Action
        Action to run
        
        .EXAMPLE
        New-WPFEvent -ControlName 'Button' -EventName 'Click' -Action { Write-Host 'Button clicked!' }
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
    Param(
        [string]$ControlName,
        [string]$EventName,
        [scriptblock]$Action
    )
    if($ControlName -ne 'Window') { $ControlName = "Window_$ControlName" }
    $WinControls = $Global:PoshWPFHashTable.WindowControls
    $null = Register-ObjectEvent -InputObject $WinControls[$ControlName] -EventName $EventName -Action $Action
}

Function Start-WPFSleep {
    <#
        .SYNOPSIS
        Waits for action to be done
        
        .DESCRIPTION
        When running the UI in a separate thread, you may want to pause the main thread
        until an action is done in the UI. This is very necessary if you run the script
        without the -NoExit switch. The PowerShell session will simply close!
        
        .EXAMPLE
        Start-WPFSleep
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
    while($Global:PoshWPFHashTable.WaitEvent){
        Wait-Event -Timeout 2
    }
}

