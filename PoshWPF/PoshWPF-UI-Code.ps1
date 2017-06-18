Function Show-WPFWindow {
    <#
        .SYNOPSIS
        Shows a window
        
        .DESCRIPTION
        Takes XAML and turns it into a window
        
        .PARAMETER xaml
        XAML of the window
        
        .EXAMPLE
        Show-WPFWindow -xaml $xaml
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
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
        $Global:PoshWPFHashTable.WindowShown = $true
        $null = $Window.ShowDialog()
    }
    catch {
        $Global:PoshWPFHashTable.WindowShown = $false
        Write-WPFError -Exc $_
    }
    $Global:PoshWPFHashTable.WindowShown = $false
}

Function Write-WPFError {
    <#
        .SYNOPSIS
        Adds errors from the WPF window into the Synchronized hashtable for easy troubleshooting
        
        .DESCRIPTION
        Adds errors from the WPF window into the sync hash
        
        .PARAMETER Exc
        Exception
        
        .EXAMPLE
        Write-WPFError -Exc $Exception
        
        .NOTES
        .Author: Ryan Ephgrave
    #>
    Param($Exc)
    if($PoshWPFHashTable.ErrorList -eq $null) {
        $PoshWPFHashTable.ErrorList = New-Object System.Collections.ArrayList
    }
    $null = $PoshWPFHashTable.ErrorList.Add($Exc)
}
