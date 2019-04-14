###############################################################################
#
#.SYNOPSIS
#    Shows a computer password in an independent Window
#
#.PARAMETER Password
#    Password to show
#
#.PARAMETER Hostname
#    Name of computer the password belongs to
#
###############################################################################
function Show-LapsPassword
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$Password,
        [Parameter(Mandatory=$true)]
        [string]$Hostname
    )

    #
    # Setup Window
    #
    $lapsWindow = @{}
    $lapsWindow.Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $script:xamllapsWindow))
    $lapsWindow.Window.Title = "LAPS Password"
    $style = ($lapsWindow.Window.FindResource("iconColor")).Color = $script:Settings.IconColor
    foreach($guiObject in $xamllapsWindow.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]"))
    {
        Write-Log -LogString ("Adding " + $guiObject.Name + " to $lapsWindow") -Severity "Debug"
        $lapsWindow.$($guiObject.Name) = $lapsWindow.Window.FindName($guiObject.Name)
    }

    $lapsWindow.btnClose.add_Click({
        $lapsWindow.Window.Close()
    })

    $lapsWindow.btnShowPassword.add_Click({
        if($lapsWindow.txtPasswd.Text -ne $Password)
        {
            $lapsWindow.txtPasswd.Text = $Password
        }
        else
        {
            $lapsWindow.txtPasswd.Text = "********"       
        }
        
    })
    
    $lapsWindow.txtHostname.Text = $Hostname
    $lapsWindow.txtPasswd.Text = "********"     
    $lapsWindow.Window.ShowDialog()

}