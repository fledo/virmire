<#
    .SYNOPSIS
    Open files and folders with global hotkeys. 

    .DESCRIPTION
    Opens a GUI where system wide hotkeys can be configured. 
    This requires a background listener process to be started. 

    .LINK
    https://github.com/owlnical/virmire

    .LICENSE
    The MIT License (MIT)

    Copyright (c) 2014-2018 Teddy Wong, Fred Uggla

    Permission is hereby granted, free of charge, to any person obtaining a copy of
    this software and associated documentation files (the "Software"), to deal in
    the Software without restriction, including without limitation the rights to
    use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
    of the Software, and to permit persons to whom the Software is furnished to do
    so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>

# Path/files
$Appdata = "$env:APPDATA\Virmire"
$PsEventingFile = "$Appdata\Nivot.PowerShell.Eventing.Extensions.dll"
$SettingsFile = "$Appdata\settings.csv"
$PidFile = "$Appdata\pid.txt"
$ListenerFile = "$Appdata\listener.ps1"

<#
    .SYNOPSIS
    Extract icon from DLL

    .DESCRIPTION
    Opens a new dialog. User choice will be saved in the parameter Button
    object.Target. The hotkey will be enabledand the Button GUI image updated.
    
    .LINK
    http://stackoverflow.com/questions/6872957/how-can-i-use-the-images-within-shell32-dll-in-my-c-sharp-project
    https://social.technet.microsoft.com/Forums/scriptcenter/en-US/16444c7a-ad61-44a7-8c6f-b8d619381a27

    .EXAMPLE
    $Form = New-Object System.Windows.Forms.Form
    $Form.Icon = [System.IconExtractor]::Extract("shell32.dll", 4, $true)
#>
$FolderIcon = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System {
    public class IconExtractor {
        public static Icon Extract(string file, int number, bool largeIcon) {
	       IntPtr large;
	       IntPtr small;
	       ExtractIconEx(file, number, out large, out small, 1);
	       try {
	           return Icon.FromHandle(largeIcon ? large : small);
	       } 
           catch {
	           return null;
           }
	   }
	   [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	   private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
	}
}
"@

<#
    .SYNOPSIS
    Update button caption and image

    .DESCRIPTION
    Updates button if its status is enabled. Set image to icon associated
    with target or folder icon from shell32.dll. Set text to target name.
    
    .PARAMETER Button
    Object representing a hotkey. 

    .EXAMPLE
    Choose-Folder -Button $Button
#>
Function Update-ButtonContent {
    param (
        $Button
    )
    if ($Button.Target -ne "no-target") {
        if (Test-Path -Path $Button.target -PathType Leaf) {
            $Button.Image = [System.Drawing.Icon]::ExtractAssociatedIcon($Button.Target)
        } else {
            $Button.Image = [System.IconExtractor]::Extract("shell32.dll", 4, $true)
        }
        $Button.Text = "$($Button.name)`n`n`n`n`n`n$([io.path]::GetFileNameWithoutExtension($Button.Target))"
    }
}

<#
    .SYNOPSIS
    Show file browser dialog, save choice as hotkey target

    .DESCRIPTION
    Opens a new dialog. User choice will be saved in the parameter Button
    object.Target. The hotkey will be enabledand the Button GUI image updated.
    
    .PARAMETER Button
    Object representing a hotkey.

    .EXAMPLE
    Choose-File -Button $this # $this = button clicked in GUI in click event
#>
Function Choose-File {
    param (
        $Button
    )
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    if ($OpenFileDialog.ShowDialog() -eq "OK") {
        $Button.Target = $OpenFileDialog.FileName
        Update-ButtonContent -Button $Button
    } else {
        $Button.Target = "no-target"
        $Button.Image = $null
    }
} 

<#
    .SYNOPSIS
    Show folder browser dialog, save choice as hotkey target

    .DESCRIPTION
    Opens a new dialog. User choice will be saved in the parameter Button 
    object.Target. The hotkey will be enabled and the Button GUI image updated.
    
    .PARAMETER Button
    Object representing a hotkey. 

    .EXAMPLE
    Choose-Folder -Button $this # Used from within button click event
#>
Function Choose-Folder {
    param (
        $Button
    )
    $OpenFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($OpenFolderDialog.ShowDialog() -eq "OK") {
        $Button.Target = $OpenFolderDialog.SelectedPath
        Update-ButtonContent -Button $Button
    } else {
        $Button.Target = "no-target"
        $Button.Image = $null
    }
}

<#
    .SYNOPSIS
    Save script settings.

    .DESCRIPTION
    Saves the current settings to $SettingsFile and reloads the listener.

    .EXAMPLE
    Save-Settings
#>
Function Save-Settings {
    $Output = @()
    Foreach ($object in $Buttons) {
        $Output += New-Object -TypeName PSObject -Property @{Key=$object.Key; Target=$object.Target}
    }
    $Output | Export-Csv $SettingsFile -Force
    Start-Listener
}

<#
    .SYNOPSIS
    Show the Graphical User Interface

    .DESCRIPTION
    Renders entire GUI and register click events for buttons

    .EXAMPLE
    Show-GUI
#>
Function Show-GUI {

    # Form Base 
    $Form = New-Object system.Windows.Forms.Form
    $Form.Width = 1125
    $Form.Height = 420
    $Form.StartPosition = "CenterScreen"
    $Form.BackColor = [System.Drawing.Color]::Ivory
    $Form.Text = "Virmire"

    # Label
    $Font = New-Object System.Drawing.Font("Arial",18)
    $title = New-Object System.Windows.Forms.Label
    $title.Text = "Choose a hotkey for CTRL + ALT"
    $title.Width = 450
    $title.Height = 50
    $title.Font = $Font
    $title.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $Form.Controls.Add($title)

    # Keyboard buttons
    $left = 10
    $top = 50
    foreach ($object in $Data) {
        if ($object.Key -eq "A") {
        $left = 30
        $top  += 110
    } elseif ($object.Key -eq "Z") {
        $left = 70
        $top  += 110
    }
        $button = New-Object System.Windows.Forms.Button
        $button | Add-Member -NotePropertyName Target -NotePropertyValue $object.Target
        $button | Add-Member -NotePropertyName Key -NotePropertyValue $object.Key
        $button.name = $object.Key
        $button.add_MouseDown({
            if ($this.Target -eq "no-target") {
                if ($_.button -eq "Left") {
                    Choose-File -Button $this
                } elseif ($_.button -eq "Right") {
                    Choose-Folder -Button $this
                }
            } else {
                $this.Target = "no-target"
                $this.Image = $null
                $this.Text = $this.name
            }
        })
        $button.Text = $object.Key
        $button.Width = 100
        $button.Height = 100
        $button.Top = $top
        $button.Left = $left
        $button.Padding = 1
        $button.BackColor = [System.Drawing.Color]::LightGray
        $button.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
        $button.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        Update-ButtonContent -Button $button
        $Form.Controls.Add($button)
        $global:Buttons += $button
        $left += 110
    }
    
    # Menu buttons
    $titles = @("Help", "Reset keys", "Start Listener", "Stop Listener", "Save", "Exit")
    $menu = @()
    $left = 450
    For ($i=0; $i -lt $titles.Length; $i++) {
        $menu += New-Object System.Windows.Forms.Button
        $menu[$i].Text = $titles[$i]
        $menu[$i].Width = 100
        $menu[$i].Height = 25
        $menu[$i].Top = 10
        $menu[$i].Left = $left
        $menu[$i].BackColor = [System.Drawing.Color]::LightGray
        $left += 110
    }

    # Help button
    $menu[0].Add_Click({
        $popup = new-object -comobject wscript.shell
        $popup.popup("Left click button to choose target file.`n 
Right click to choose target folder.`n
Left/Right click to clear target.`n
When done, click save.
Start listener to enable hotkeys (this requires the powershell module PSEventing Plus which is included with Virmire. It's also available for download on pseventing.codeplex.com/)`n
You can autostart the listener by running the listener file '$Appdata\listener.ps1' using Task Scheduler or by placing a shortcut in 'Program Files\Startup'.`n
Further info at github.com/owlnical/virmire and in the README.md file.")
    })

    # Reset button
    $menu[1].Add_Click({
        $YesNo = new-object -comobject wscript.shell
        $intAnswer = $YesNo.popup("Do you want to remove all configured hotkeys?", 0, "Remove hotkeys", 4)
        If ($intAnswer -eq 6) {
          Set-Defaults
          $Form.Close()
          $YesNo.popup("Configured keys removed. This program will now close.")
        }
    })

    # Start listener button
    $menu[2].ForeColor = [System.Drawing.Color]::DarkGreen
    $menu[2].Add_Click({
        Start-Listener
    })

    # Stop listener button
    $menu[3].ForeColor = [System.Drawing.Color]::DarkRed
    $menu[3].Add_Click({
        Stop-Listener
    })

    # Save button
    $menu[4].Add_Click({
        Save-Settings
        [System.Windows.Forms.MessageBox]::Show("Data saved to $SettingsFile")
    })

    # Exit button
    $menu[5].Add_Click({
        $Form.Close()
    })

    For ($i=0; $i -lt $menu.Length; $i++) {
        $Form.Controls.Add($menu[$i])
    }

    $Form.ShowDialog() > $null
}

<#
    .SYNOPSIS
    Kill background listener

    .DESCRIPTION
    Stop the process with ID from $Appdata\pid

    .EXAMPLE
    Stop-Listener
#>
Function Stop-Listener {
    sleep -Milliseconds 1000
    $id = Get-Content $PidFile -ErrorAction SilentlyContinue
    if ($id) {
        $Process = Get-Process -id $id -ErrorAction SilentlyContinue
        if (($Process.name -eq "powershell") -and ($id -eq $Process.Id)) {
            Stop-Process -id $id
        }
    }
}

<#
    .SYNOPSIS
    Restore default settings

    .DESCRIPTION
    Deletes current settings and creates a CSV in $Appdata with default settings

    .EXAMPLE
    Set-Defaults
#>
function Set-Defaults {
    Stop-Listener
    Remove-Item $Appdata -Force -Recurse -ErrorAction SilentlyContinue
    New-Item $Appdata -type directory -Force > $null
    $Data = @()
    $Keys = ("Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "A", "S", "D", "F", "G", "H", "J", "K", "L", "Z", "X", "C", "V", "B", "N", "M")
    foreach ($key in $Keys) {
        $Data += New-Object -TypeName PSObject -Property @{Key=$key; Target="no-target"}
    }
    $Data | Export-Csv $SettingsFile -Force
}

<#
    .SYNOPSIS
    (Re)load the background process listening for hotkeys

    .DESCRIPTION
    Write a PS script which registers a new hotkey event for each enabled key.
    The script is then started in a hidden window with pid saved in $appdata.

    .EXAMPLE
    Start-Listener
#>
function Start-Listener {
    Stop-Listener
    $Data = Import-Csv $SettingsFile
    "import-module -Name $Appdata\Nivot.PowerShell.Eventing.Extensions.dll" | Out-File -FilePath $ListenerFile -Force -Encoding UTF8
    Foreach ($object in $Data) {
        if ($object.Target -ne "no-target"){
            Add-Content -Value "register-hotkeyevent 'ctrl+alt+$($object.Key)' -action  { start '$($object.Target)' } -global" -Path $ListenerFile 
        }
    }
    Add-Content -Value "`$pid | Out-File -FilePath '$PidFile' -Force" -Path $ListenerFile
    Add-Content -Value "write-host `"Keys registered. Do not close this window.``nPID: `$pid. Saved to $PidFile`"" -Path $ListenerFile
    
    # Start the Listener. Keep alive and hidden.
    start-process powershell.exe -WindowStyle Hidden -argument "-NoExit -nologo -noprofile -executionpolicy bypass -command . '$ListenerFile'" 
}

# Check for version 3+ of powershell, required by "Add-Member -NotePropertyName" 
if ($PSVersionTable.PSVersion.Major -le 3) {
    Write-error "Detected version $($PSVersionTable.PSVersion.Major) of PowerShell. Virmire requires version 3."
    exit
}

Add-Type -TypeDefinition $FolderIcon -ReferencedAssemblies System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Load settings, make sure we have the pseventingmodule and show GUI
$global:Buttons = @() # Why is this global!?
if (-not (Test-Path $Appdata)) {
    Set-Defaults
}
if (-not (Test-Path $PsEventingFile)) {
    Copy-Item "Nivot.PowerShell.Eventing.Extensions.dll" -Destination $PsEventingFile
}
$Data = Import-Csv $SettingsFile
Show-GUI
