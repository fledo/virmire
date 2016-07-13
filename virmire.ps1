# Path/files
$Appdata = "$env:APPDATA\Virmire"
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
    if ($Button.Status -eq $True) {
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
        $Button.Status = $True
        Update-ButtonContent -Button $Button
    } else {
        $Button.Target = "null"
        $Button.Status = $False
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
        $Button.Status = $True
        Update-ButtonContent -Button $Button
    } else {
        $Button.Target = "null"
        $Button.Status = $False
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
        $Output += New-Object -TypeName PSObject -Property @{Key=$object.Key; Target=$object.Target; Status = $object.Status}
    }
    $Output | Export-Csv $SettingsFile -Force
    Load-Listener
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
    $Form.FormBorderStyle = 'None'
    $Form.Width = 1110
    $Form.Height = 380
    $Form.StartPosition = "CenterScreen"
    $Form.BackColor = [System.Drawing.Color]::DarkGray

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
        $button | Add-Member -NotePropertyName Status -NotePropertyValue $object.Status
        $button | Add-Member -NotePropertyName Key -NotePropertyValue $object.Key
        $button.name = $object.Key
        $button.add_MouseDown({
            if ($this.Status -eq $False) {
                if ($_.button -eq "Left") {
                    Choose-File -Button $this
                } elseif ($_.button -eq "Right") {
                    Choose-Folder -Button $this
                }
            } else {
                $this.Target = "null"
                $this.Status = $False
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
    
    # Help button
    $help = New-Object System.Windows.Forms.Button
    $help.BackColor = [System.Drawing.Color]::LightGray
    #$help.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
    $help.Text = "Help"
    $help.Width = 100
    $help.Height = 25
    $help.Top = 10
    $help.Left = 450
    $help.Add_Click({
        $popup = new-object -comobject wscript.shell
        $popup.popup("hej")
    })
    $Form.Controls.Add($help)
    
    # Reset button
    $remove = New-Object System.Windows.Forms.Button
    $remove.BackColor = [System.Drawing.Color]::LightGray
    $remove.Text = "Reset keys"
    $remove.Width = 100
    $remove.Height = 25
    $remove.Top = 10
    $remove.Left = 560
    $remove.Add_Click({
        $YesNo = new-object -comobject wscript.shell
        $intAnswer = $YesNo.popup("Do you want to remove all configured hotkeys?", 0, "Remove hotkeys", 4)
        If ($intAnswer -eq 6) {
          Set-Defaults
          $Form.Close()
          $YesNo.popup("Configured keys removed. This program will now close.")
        }
    })
    $Form.Controls.Add($remove)
    
    # Start listener button
    $start = New-Object System.Windows.Forms.Button
    $start.BackColor = [System.Drawing.Color]::LightGray
    $start.ForeColor = [System.Drawing.Color]::DarkGreen
    $start.Text = "Start Listener"
    $start.Width = 100
    $start.Height = 25
    $start.Top = 10
    $start.Left = 670
    $start.Add_Click({
        Load-Listener
    })
    $Form.Controls.Add($start)
    
    # Stop listener button
    $stop = New-Object System.Windows.Forms.Button
    $stop.BackColor = [System.Drawing.Color]::LightGray
    $stop.ForeColor = [System.Drawing.Color]::DarkRed
    $stop.Text = "Stop Listener"
    $stop.Width = 100
    $stop.Height = 25
    $stop.Top = 10
    $stop.Left = 780
    $stop.Add_Click({
        Stop-Listener
    })
    $Form.Controls.Add($stop)
    
    # Save button
    $save = New-Object System.Windows.Forms.Button
    $save.BackColor = [System.Drawing.Color]::LightGray
    $save.Text = "Save"
    $save.Width = 100
    $save.Height = 25
    $save.Top = 10
    $save.Left = 890
    $save.Add_Click({
        Save-Settings
        [System.Windows.Forms.MessageBox]::Show("Data saved to $SettingsFile")
    })
    $Form.Controls.Add($save)
    
    # Exit button
    $exit = New-Object System.Windows.Forms.Button
    $exit.BackColor = [System.Drawing.Color]::LightGray
    $exit.Text = "Exit"
    $exit.Width = 100
    $exit.Height = 25
    $exit.Top = 10
    $exit.Left = 1000
    $exit.Add_Click({
        $Form.Close()
    })
    $Form.Controls.Add($exit)
    
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
        $Data += New-Object -TypeName PSObject -Property @{Key=$key; Target="null"; Status = $False}
    }
    $Data | Export-Csv $SettingsFile -Force
}

<#
    .SYNOPSIS
    Tries to load a module.

    .DESCRIPTION
    Loads the module from the parameter or displays an
    error message with a link where it can be downloaded.

    .PARAMETER Name
    Specifies the module to be loaded.

    .PARAMETER URL
    Specifies where the module can be downloaded.
   
    .PARAMETER RequiredBy
    Specifies what requires the module

    .EXAMPLE
    Load-Module -Name ShowUI -Url showui.codeplex.com -RequiredBy $MyInvocation.MyCommand.Name
#>
function Load-Module {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$URL,
        [string]$RequiredBy = "This script"
    )
    if (-not (Get-Module $Module))
    {
        try {
            Import-Module $Module -Force
        } catch {
            write-error "$RequiredBy requires the module '$Name'. Please download it from $URL`n"
        }
    }
}

<#
    .SYNOPSIS
    (Re)load the background process listening for hotkeys

    .DESCRIPTION
    Write a PS script which registers a new hotkey event for each enabled key.
    The script is then started in a hidden window with pid saved in $appdata.

    .EXAMPLE
    Load-Listener
#>
function Load-Listener {
    # Make sure we have the required module and that the previous Listener is dead
    Load-Module -Name pseventingplus -Url http://pseventing.codeplex.com/releases/view/66587 -RequiredBy "The background listener process"
    Stop-Listener
    $Data = Import-Csv $SettingsFile
    "import-module pseventingplus" | Out-File -FilePath $ListenerFile -Force -Encoding UTF8
    Foreach ($object in $Data) {
        if ($object.Status -eq $True){
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
    Write-error "Detected version $($PSVersionTable.PSVersion.Major) of PowerShell. This script requires version 3."
    exit
}

Add-Type -TypeDefinition $FolderIcon -ReferencedAssemblies System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Load settings and show GUI
$global:Buttons = @() # Why is this global!?
if (-not (Test-Path $Appdata)) {
    Set-Defaults
}
$Data = Import-Csv $SettingsFile
Show-GUI
