function Show-DebugWindow
{    
    Param
    (
        [parameter(Mandatory=$True)]
        $ItemsSource
    )

    #
    # Setup Window
    #
    $debugWindow = @{}
    $debugWindow.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $script:xamldebugWindow))
    $debugWindow.Window.Title = "Debug view"
    $style = ($debugWindow.Window.FindResource("iconColor")).Color = $script:Settings.IconColor
    foreach($guiObject in $xamldebugWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        #Write-Log -LogString ("Adding " + $guiObject.Name + " to $debugWindow") -Severity "Debug"
        $debugWindow.$($guiObject.Name) = $debugWindow.Window.FindName($guiObject.Name)
    }

    $debugWindow.dg.ItemsSource = $ItemsSource
   
    $debugWindow.Window.ShowDialog()

}
