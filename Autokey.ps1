################################
# Startup Script
# Editted by Chris Gergler
################################
# How to Adjust
# Screen starts in top left corner (0,0)
# And goes to bottom right corner (1920,1080)
# Or highest resolution.
################################


#Set the average time to your system load and open an iexplorer page
Start-Sleep -Seconds 3

[system.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null

# Set the exactly position of cursor in some iexplore hyperlink between the (open parenthesis) below: 

function Click-MouseButton
{
    $signature=@' 
      [DllImport("user32.dll",CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
      public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@ 

    $SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru 

        $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0);
}
    

function TargetClick
{
    param (
    $xLoc,
    $yLoc
    )
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($xLoc,$yLoc)


}

function SendKey
{
    param(
    $key,
    $window
    )

    $getWindow = New-Object -ComObject wscript.shell;
    IF ($window) {$getWindow.AppActivate($window)}
    Sleep 1
    IF ($key){$getWindow.SendKeys($key)}
}


TargetClick -xLoc 1280 -yLoc 480
Click-MouseButton