# Load necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create GUI Form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Mouse Jiggler"
$Form.Size = New-Object System.Drawing.Size(300, 200)
$Form.StartPosition = "CenterScreen"

# Status Label
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Status: Stopped"
$StatusLabel.Location = New-Object System.Drawing.Point(100, 20)
$StatusLabel.Size = New-Object System.Drawing.Size(200, 20)
$Form.Controls.Add($StatusLabel)

# Interval Label
$IntervalLabel = New-Object System.Windows.Forms.Label
$IntervalLabel.Text = "Interval (sec):"
$IntervalLabel.Location = New-Object System.Drawing.Point(20, 60)
$IntervalLabel.Size = New-Object System.Drawing.Size(100, 20)
$Form.Controls.Add($IntervalLabel)

# Interval TextBox
$IntervalTextBox = New-Object System.Windows.Forms.TextBox
$IntervalTextBox.Text = "30"
$IntervalTextBox.Location = New-Object System.Drawing.Point(120, 60)
$IntervalTextBox.Size = New-Object System.Drawing.Size(50, 20)
$Form.Controls.Add($IntervalTextBox)

# Start Button
$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Text = "Start"
$StartButton.Location = New-Object System.Drawing.Point(50, 100)
$StartButton.Size = New-Object System.Drawing.Size(75, 30)
$Form.Controls.Add($StartButton)

# Stop Button
$StopButton = New-Object System.Windows.Forms.Button
$StopButton.Text = "Stop"
$StopButton.Location = New-Object System.Drawing.Point(150, 100)
$StopButton.Size = New-Object System.Drawing.Size(75, 30)
$StopButton.Enabled = $false
$Form.Controls.Add($StopButton)

# Import Mouse Movement Function
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class MouseJiggler {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    public struct POINT {
        public int X;
        public int Y;
    }
}
"@ -Language CSharp

# Global variable to control the loop
$Script:Running = $false

# Function to get current mouse position
function Get-MousePosition {
    $Point = New-Object MouseJiggler+POINT
    [MouseJiggler]::GetCursorPos([ref]$Point)
    return $Point
}

# Start Button Event
$StartButton.Add_Click({
    $Script:Running = $true
    $StatusLabel.Text = "Status: Running"
    $StartButton.Enabled = $false
    $StopButton.Enabled = $true
    $Interval = [int]$IntervalTextBox.Text

    # Get current mouse position
    $StartPos = Get-MousePosition
    $StartX = $StartPos.X
    $StartY = $StartPos.Y

    # Start a separate thread for movement
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("Interval", $Interval)
    $Runspace.SessionStateProxy.SetVariable("Running", [ref]$Script:Running)
    $Runspace.SessionStateProxy.SetVariable("StartX", $StartX)
    $Runspace.SessionStateProxy.SetVariable("StartY", $StartY)
    $PowerShell = [powershell]::Create()
    $PowerShell.Runspace = $Runspace

    # Mouse Jiggler Script Block
    $PowerShell.AddScript({
        $x = $StartX
        $y = $StartY
        while ($Running.Value) {
            [MouseJiggler]::SetCursorPos($x, $y)
            $x = $x + 1
            $y = $y + 1
            Start-Sleep -Seconds $Interval
            [MouseJiggler]::SetCursorPos($x, $y)
            $x = $x - 1
            $y = $y - 1
            Start-Sleep -Seconds $Interval
        }
    })

    $PowerShell.BeginInvoke()
    $Script:Runspace = $Runspace
    $Script:PowerShell = $PowerShell
})

# Stop Button Event
$StopButton.Add_Click({
    $Script:Running = $false
    $StatusLabel.Text = "Status: Stopped"
    $StartButton.Enabled = $true
    $StopButton.Enabled = $false

    # Close the separate thread
    if ($Script:Runspace) {
        $Script:PowerShell.Dispose()
        $Script:Runspace.Close()
        $Script:Runspace.Dispose()
        $Script:Runspace = $null
        $Script:PowerShell = $null
    }
})

# Show Form
$Form.ShowDialog()
