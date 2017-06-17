Function Show-WPFWindow {
    Param(
        [xml]$xaml
    )
    try {
        Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
        $Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
        $Global:PoshWPFHashTable.WindowControls['Window'] = $Window
        $xaml.SelectNodes("//*[@Name]") | Foreach-Object { 
            $Global:PoshWPFHashTable.WindowControls["Window_$($_.Name)"] = $Window.FindName($_.Name)
        }
        #$Timer = New-Object System.Windows.Threading.DispatcherTimer
        #$Timer.Interval = [timespan]"0:0:0.50"
        #$Timer.Add_Tick( { New-WPFTick } )
        #$Timer.Start()
        $Global:PoshWPFHashTable.WindowShown = $true
        $null = $Window.ShowDialog()
    }
    catch {
        $Global:PoshWPFHashTable.WindowShown = $true
        Write-WPFError -Exc $_
    }
}

Function Write-WPFError {
    Param($Exc)
    if($PoshWPFHashTable.ErrorList -eq $null) {
        $PoshWPFHashTable.ErrorList = New-Object System.Collections.ArrayList
    }
    $null = $PoshWPFHashTable.ErrorList.Add($Exc)
}

Function New-WPFTick {
    $null = $PoshWPFHashTable.ActionsMutex.WaitOne()
    $RunActions = $false
    $ActionsToRun = @()
    if($PoshWPFHashTable.Actions.Count -gt 0) {
        $RunActions = $true
        foreach($action in $PoshWPFHashTable.Actions) {
            $ActionsToRun += @($action)
        }
        $null = $PoshWPFHashTable.Actions.Clear()
    }
    $null = $PoshWPFHashTable.ActionsMutex.ReleaseMutex()
    if($RunActions) {
        foreach($instance in $ActionsToRun) {
            try {
                Invoke-Command -ScriptBlock $instance
            }
            catch {
                Write-WPFError -Exc $_
            }
        }
    }
}
