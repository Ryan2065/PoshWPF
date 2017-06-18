# Simple XAML with button and ListBox created in Visual Studio
Import-Module PoshWPF
$xaml = @'
<Window x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="SMA Administration" Height="350" Width="525">
    <StackPanel>
        <Button Name="Button" Width="75" Height="30" Content="Button"/>
        <ListBox Name="List" Height="200"/>
    </StackPanel>
</Window>

'@

#Show XAML Window
New-WPFWindow -xaml $xaml

#Add event to update ListBox when clicked
New-WPFEvent -ControlName 'Button' -EventName 'Click' -Action {
    #Get existing list items
    $ListItems = Get-WPFControl -ControlName 'List' -PropertyName 'Items'
    
    #Add list items to new array
    $NewArray = @()
    $NewArray += $ListItems
    
    #Add new entry to list
    $NewEntry = Get-Random
    $NewArray += @($NewEntry)

    #Set new list to control List
    Set-WPFControl -ControlName 'List' -PropertyName 'ItemsSource' -Value $NewArray

}

#Start sleep loop to end when window closed. Necessary if PowerShell isn't opened with -noexit
Start-WPFSleep