# ==============================================================================
#                               INITIATE WINDOWS FORM
# ==============================================================================

# Import the Windows Forms library for creating the graphical interface
Add-Type -AssemblyName System.Windows.Forms

# Import the Drawing library for handling fonts, colors, and graphics
Add-Type -AssemblyName System.Drawing 



# ==============================================================================
#                               GRAPHIC AND DPI Scaling Logic
# The $Global:DpiScale calculation ensures your toolkit looks sharp on 4K monitors or 150% zoom settings
# It sets the AutoScaleMode to Dpi to prevent text from looking blurry
# ==============================================================================

# Create a temporary graphics object to detect the system's screen resolution (DPI)
$Graphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)

# Calculate scaling factor (e.g., 1.5 for 150% zoom) to keep the UI sharp on 4K screens
$Global:DpiScale = $Graphics.DpiX / 96 

# Release the memory used by the temporary graphics object
$Graphics.Dispose()



# ==============================================================================
#                               Reflective Helper
# The Enable-DoubleBuffering function uses reflection to force double-buffering on child controls like panels,
# further smoothing out the user interface
# ==============================================================================

# Helper function to enable smooth rendering (no flicker) for specific UI elements
function Enable-DoubleBuffering ($Control) {
    # Search for the hidden 'DoubleBuffered' property in the .NET object
    $PropertyInfo = $Control.GetType().GetProperty("DoubleBuffered", 
        [System.Reflection.BindingFlags]::Instance -bor 
        [System.Reflection.BindingFlags]::NonPublic)
    
    # If the property exists, set it to True to prevent graphical flashing
    if ($PropertyInfo) { $PropertyInfo.SetValue($Control, $true, $null) }
}



# ==============================================================================
#                               Double Buffering
# By defining a custom BufferedForm class and using WS_EX_COMPOSITED,
# the toolkit renders all controls to a memory buffer before showing them on screen
# This eliminates the "flashing" effect when you switch tabs or resize the window
# ==============================================================================


# Define a custom Form class that uses 'Composite' rendering to eliminate window flicker
if (-not ([System.Type]::GetType("BufferedForm"))) {
    Add-Type -TypeDefinition @"
    using System.Windows.Forms;
    public class BufferedForm : Form {
        protected override CreateParams CreateParams {
            get {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000; // WS_EX_COMPOSITED
                return cp;
            }
        }
    }
"@ -ReferencedAssemblies System.Windows.Forms
}



# ==============================================================================
#                               MAIN FORM
# ==============================================================================


# Create a new instance of our flicker-free Main Window
$Form = New-Object BufferedForm

# Set the title displayed at the top of the application window
$Form.Text = "Windows Automation Toolkit"

# Set initial window size (600x500 pixels) before scaling
$Form.Size = New-Object System.Drawing.Size(600, 500)

# Ensure the application starts filled across the entire screen
$Form.WindowState = "Maximized"

# Center the window on the screen when it is opened 
$Form.StartPosition = "CenterScreen"

# Set the window background color to pure white
$Form.BackColor = "#FFFFFF"

# Configure the form to scale its size and controls based on the monitor's DPI settings
$Form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

# Set the default application font to Segoe UI at 9-point size
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Load the Visual Basic library to use easy popup tools like 'InputBox' 
Add-Type -AssemblyName Microsoft.VisualBasic

$Form.Controls.Add($TabControl)



# ==============================================================================
#                 MAIN FORM --> Floating Action AI Button (FAB) 
# ==============================================================================


# 2. BUTTON CODE
$Btn_AiAssistant = New-Object System.Windows.Forms.Button
$Btn_AiAssistant.Text = "AI Assistant"
$Btn_AiAssistant.Size = New-Object System.Drawing.Size(110, 50)

# Position: Bottom Right with Anchoring 
$Btn_AiAssistant.Location = New-Object System.Drawing.Point(($Form.Width - 150), ($Form.Height - 130))
$Btn_AiAssistant.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

# Styling to look like a FAB 
$Btn_AiAssistant.BackColor = "#0078d4" # Microsoft Blue
$Btn_AiAssistant.ForeColor = "White"
$Btn_AiAssistant.FlatStyle = "Flat"
$Btn_AiAssistant.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

# Event Logic: Trigger the WebView2 Sidebar

$Btn_AiAssistant.Add_Click({
    # Launch Microsoft Edge with the specific URL
    Start-Process "msedge.exe" "https://copilot.microsoft.com"
})

$Form.Controls.Add($Btn_AiAssistant)




# ==============================================================================
#                               HEADER
# This code block constructs a professional, responsive header for toolkit by utilizing Z-ordering and Docking
# Logic Breakdown:
#                *Right Panel ($PanelRight): Docks to the right edge and hosts the "Admin Mode" button. It includes a security check to determine if the toolkit is running with elevated privileges.
#                *Clock Label ($LblClock): Docks to the left edge. A timer updates this label every 1,000ms to provide a real-time clock.
#                *Title Label ($LblTitle): Uses Dock = Fill. In WinForms, the control that "Fills" must have the lowest priority in the Z-order to occupy the remaining space between the left and right docked elements
#                *Z-Order Priority: The Controls.Clear() and subsequent Add() methods manually define the stack. By adding $LblTitle first, it becomes the "bottom" layer, allowing the clock and admin button to "claim" their edge space first
# ==============================================================================


# Create the main container for the top of the window
$HeaderPanel = New-Object System.Windows.Forms.Panel
# Attach the panel firmly to the top edge of the window
$HeaderPanel.Dock = "Top"
# Set a fixed height of 40 pixels for the header bar
$HeaderPanel.Height = 40
# Set the background color to a professional blue shade
$HeaderPanel.BackColor = "#007ACC"

# --- 1. RIGHT SIDE CONTAINER (Holds the Admin Button) ---
# Create a sub-panel to group items on the right side
$PanelRight = New-Object System.Windows.Forms.Panel
# Attach this sub-panel to the right edge of the header
$PanelRight.Dock = "Right"
# Set a fixed width for the right-side section
$PanelRight.Width = 150
# Allow the header color to show through the background
$PanelRight.BackColor = "Transparent" 

# Initialize the button used for administrative elevation
$Btn_HeaderAdmin = New-Object System.Windows.Forms.Button
# Set the button's physical dimensions
$Btn_HeaderAdmin.Size = New-Object System.Drawing.Size(120, 26)
# Manually position the button within the 150px right panel
$Btn_HeaderAdmin.Location = New-Object System.Drawing.Point(15, 7)
# Use a flat visual style for a modern look
$Btn_HeaderAdmin.FlatStyle = "Flat"
# Remove the default border for a cleaner appearance
$Btn_HeaderAdmin.FlatAppearance.BorderSize = 0
# Set bold font for high visibility
$Btn_HeaderAdmin.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
# Set text color to white for contrast against blue/green/red
$Btn_HeaderAdmin.ForeColor = "White"
# Change the mouse pointer to a hand icon when hovering
$Btn_HeaderAdmin.Cursor = [System.Windows.Forms.Cursors]::Hand

# SECURITY CHECK: Determine if current user is an Administrator
$Identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$Principal = New-Object Security.Principal.WindowsPrincipal($Identity)

# Logic for when the tool is already running with Admin rights
if ($Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $Btn_HeaderAdmin.Text = "Admin Mode"
    $Btn_HeaderAdmin.BackColor = "#28a745" # Set color to Green for success
    $Btn_HeaderAdmin.Enabled = $false      # Disable button as no action is needed
} else {
    # Logic for when the tool needs elevation
    $Btn_HeaderAdmin.Text = "Run as Admin"
    $Btn_HeaderAdmin.BackColor = "#ff4d4d" # Set color to Red for warning
    $Btn_HeaderAdmin.Add_Click({
        # Relaunch the script with the 'RunAs' verb to trigger UAC prompt
        Start-Process powershell.exe -ArgumentList "-NoExit -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        # Close the non-admin window
        $Form.Close()
    })
}

# Add the button into the Right-side panel
$PanelRight.Controls.Add($Btn_HeaderAdmin)
# Add the Right-side panel into the main Header
$HeaderPanel.Controls.Add($PanelRight)

# --- 2. LEFT SIDE CONTAINER (Holds the Clock) ---
# Create a label to display the current date and time
$LblClock = New-Object System.Windows.Forms.Label
# Initial text format: Day-Month-Year : Seconds
$LblClock.Text = $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')
# Use a monospace font (Consolas) to prevent text jumping during updates
$LblClock.Font = New-Object System.Drawing.Font("Consolas", 13, [System.Drawing.FontStyle]::Bold)
# Set clock text color to white
$LblClock.ForeColor = "White"
# Attach the clock firmly to the left edge of the header
$LblClock.Dock = "Left"
# Vertically center the text within the 40px height
$LblClock.TextAlign = "MiddleLeft"
# Allow the label to grow/shrink based on text length
$LblClock.AutoSize = $true
# Add 10 pixels of space from the left edge of the window
$LblClock.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
# Add the clock into the main Header
$HeaderPanel.Controls.Add($LblClock)

# --- 3. CENTER TITLE (Fills the middle) ---
# Create the main application title label
$LblTitle = New-Object System.Windows.Forms.Label
# Application Name
$LblTitle.Text = "IT Support Automation Suite"
# Set bold, larger font for the title
$LblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
# Set title text color to white
$LblTitle.ForeColor = "White"
# Set the title to fill all available space between Left and Right components
$LblTitle.Dock = "Fill"
# Precisely center the title text in the middle of the header
$LblTitle.TextAlign = "MiddleCenter"
# Must be false for 'Dock = Fill' to function correctly
$LblTitle.AutoSize = $false
# Add the title into the main Header
$HeaderPanel.Controls.Add($LblTitle)

# --- 4. CONTROL ORDER (Crucial for Docking) ---
# Reset controls to manually define Z-Order (Dock priority)
$HeaderPanel.Controls.Clear()
# The 'Fill' control must be added first (it sits at the back)
$HeaderPanel.Controls.Add($LblTitle)
# The 'Right' control claimed its edge
$HeaderPanel.Controls.Add($PanelRight)
# The 'Left' control claimed its edge
$HeaderPanel.Controls.Add($LblClock)

# Add the completed header bar to the main application form
$Form.Controls.Add($HeaderPanel)

# --- 5. CLEAN TIMER (Replaces any old timers) ---
# Create a timer to keep the clock current
$TimerHeader = New-Object System.Windows.Forms.Timer
# Set update interval to 1 second (1000 milliseconds)
$TimerHeader.Interval = 1000
# Define what happens every 1 second
$TimerHeader.Add_Tick({
    # Update only the text; docking maintains the position
    $LblClock.Text = $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')
})
# Start the timer immediately
$TimerHeader.Start()



# ==============================================================================
#                               BUTTONS
# This function provides a sophisticated, modular way to style toolkit buttons using GDI+ custom drawing
# Logic Breakdown:
#                *High DPI Scaling: Uses $Global:DpiScale to multiply base dimensions, ensuring buttons remain usable on high-resolution displays
#                *Visual Styling: Implements a modern "Flat" design with semi-transparent gray backgrounds and bold Segoe UI typography
#                *Text Rendering Hack: Standard WinForms text doesn't support outlines. This code moves the button's text to the .Tag property and clears .Text to hide the default rendering
#                *The Paint Event: Uses System.Drawing.Drawing2D.GraphicsPath to "trace" the text shape. It draws a 3-pixel white outline first, then fills the inside with black. This makes the text readable against any background color or image.
#                *Dynamic Updates: The Add_TextChanged event ensures that if a script updates the button text (e.g., from "Start" to "Scanning..."), the custom outline drawing updates automatically
# ==============================================================================


# Function to apply high-quality styling and custom text effects to WinForms buttons
function Format-Button ($Button) {

    # Define the base dimensions for the button at 100% scaling
    $BaseWidth = 230
    $BaseHeight = 60
    # Apply global DPI scaling to ensure buttons aren't tiny on high-res screens 
    $Button.Width = $BaseWidth * $Global:DpiScale
    $Button.Height = $BaseHeight * $Global:DpiScale
    
    # Set background color with partial transparency (Alpha=128)
    $Button.BackColor = [System.Drawing.Color]::FromArgb(128, 200, 200, 200)
    # Use 'Flat' style to remove the old-school 3D button look
    $Button.FlatStyle = "Flat"
    # Set the width of the button's outer border
    $Button.FlatAppearance.BorderSize = 1
    # Set the color of the button border to black
    $Button.FlatAppearance.BorderColor = "Black"
    # Define a bold Segoe UI font at 11pt for clear readability
    $Button.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    # Set external spacing around the button (Top/Bottom = 25px)
    $Button.Margin = New-Object System.Windows.Forms.Padding(10, 25, 10, 25)
    # Set internal spacing inside the button edges
    $Button.Padding = New-Object System.Windows.Forms.Padding(10)

    # 2. Initial Text Hiding (Move Text to Tag)
    # WinForms can't draw outlines natively, so we move text to 'Tag' to hide it 
    if ($Button.Text -ne "") {
        $Button.Tag = $Button.Text # Store the actual name in a hidden property
        $Button.Text = ""          # Clear visible text so it doesn't overlap the custom drawing
    }

    # 3. AUTO-FIX: Watch for text changes (e.g. "Scanning...")
    # If a script updates the button name, this moves the new text back to the hidden 'Tag'
    $Button.Add_TextChanged({
        if ($this.Text -ne "") {
            $this.Tag = $this.Text # Capture the new text
            $this.Text = ""        # Hide it again immediately
            $this.Invalidate()     # Force the button to redraw with the new name
        }
    })

    # 4. Custom Paint Logic (Draws the Outline Text from 'Tag')
    # This event triggers every time the button needs to be drawn on screen 
    $Button.Add_Paint({
        param($sender, $e)
        if ($sender.Tag) {
            $g = $e.Graphics # Get the drawing surface
            $g.SmoothingMode = "AntiAlias"        # Smooth out edges for high quality
            $g.TextRenderingHint = "AntiAlias"    # Smooth out text edges

            # Create a path object to define the shape of the text characters
            $Path = New-Object System.Drawing.Drawing2D.GraphicsPath
            $Rect = $sender.ClientRectangle # Get the button's boundaries
            $Format = New-Object System.Drawing.StringFormat
            $Format.Alignment = "Center"      # Center text horizontally
            $Format.LineAlignment = "Center" # Center text vertically
            $EmSize = $sender.Font.Size * 1.3 # Scale font size for path drawing

            # Create the actual geometric shape of the text based on the hidden 'Tag' 
            $Path.AddString($sender.Tag, $sender.Font.FontFamily, [int]$sender.Font.Style, $EmSize, $Rect, $Format)

            # Draw the white outline first (3 pixels wide)
            $Pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 3)
            $g.DrawPath($Pen, $Path)

            # Fill the center of the text with black
            $Brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)
            $g.FillPath($Brush, $Path)
        }
    })
}



# ==============================================================================
#                               TAB CONTROL
# This code block creates a responsive TabControl that serves as the main navigation hub for toolkit. 
# By utilizing the previously defined $Global:DpiScale, it ensures the interface remains usable and clear across different display resolutions
# Logic Breakdown:
#                *Full-Size Layout: Using $TabControl.Dock = "Fill" ensures the tab system automatically expands to occupy the entire available space within the form
#                *DPI-Aware Typography: The font size is dynamically multiplied by your DPI scale factor. This prevents the tab headers from appearing tiny or unreadable on high-resolution monitors
#                *Proportional Padding: The Padding property (expressed as a Point) controls the "breathing room" around the text inside the tab headers. Scaling these values ensures that the clickable area of the tab remains large enough for easy mouse interaction
#                *Z-Order Management: $TabControl.BringToFront() ensures the tab system stays visible on top of any background panels or other controls
# ==============================================================================


# Create the main TabControl object to act as the primary navigation hub
$TabControl = New-Object System.Windows.Forms.TabControl

# Set the tab control to fill all available space within its parent container 
$TabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

# DYNAMIC FONT: Calculate bold Segoe UI font size based on current DPI scale
# (e.g., a base of 12pt becomes 18pt on a 150% scaled display) 
$TabFontSize = 12 * $Global:DpiScale

# Apply the scaled bold font to the TabControl headers
$TabControl.Font = New-Object System.Drawing.Font("Segoe UI", $TabFontSize, [System.Drawing.FontStyle]::Bold)

# DYNAMIC PADDING: Calculate horizontal (30px) and vertical (5px) breathing room
$PadX = [int](30 * $Global:DpiScale) # Scaled horizontal padding for clickable area
$PadY = [int](5 * $Global:DpiScale)  # Scaled vertical padding for header height
# Apply scaled padding to ensure tab headers are easy to click on high-res screens 
$TabControl.Padding = New-Object System.Drawing.Point($PadX, $PadY)


# Ensure the tab control is layered at the very top of the visual stack 
$TabControl.BringToFront()


# *****************************************************************************************************************************************************************************************************
# ==============================================================================
#                              DIAGNOSTIC TAB
# The Diagnostic Tab serves as the central hub for system maintenance within your toolkit
# it includes several high-value features:
#                *Maintenance Automation: Features buttons for clearing temporary files, resetting the print spooler, and flushing system caches to resolve common Tier 1 helpdesk issues
#                *Deep System Repair: Includes the "Troubleshoot Slowness" tool, which launches a specialized batch environment for running intensive SFC and DISM repairs without freezing the main GU
#                *Audit Capabilities: Provides one-click access to installed software lists, system uptime, and BitLocker encryption status reports
# This tab is designed to be the first stop for IT analysts when troubleshooting a malfunctioning workstation
# ==============================================================================


# Initialize a new TabPage object for the Diagnostic section
$TabDiagnostic = New-Object System.Windows.Forms.TabPage
# Set the text label that appears on the tab header
$TabDiagnostic.Text = "Diagnostic"
# Set a dark gray background color for the tab page
$TabDiagnostic.BackColor = "#2d2d30" 

# Create a FlowLayoutPanel to automatically organize buttons in a grid-like fashion
$FlowDiagnostic = New-Object System.Windows.Forms.FlowLayoutPanel
# Set the panel to expand and fill the entire area of the TabPage 
$FlowDiagnostic.Dock = "Fill"
# Enable scrollbars so users can access buttons if they exceed the window height
$FlowDiagnostic.AutoScroll = $true
# Allow the panel to resize itself based on the buttons added
$FlowDiagnostic.AutoSize = $true
# Configure the panel to grow or shrink to fit its contents exactly
$FlowDiagnostic.AutoSizeMode = "GrowAndShrink"

# Apply Double Buffering to the FlowLayoutPanel to prevent flickering during scrolling
Enable-DoubleBuffering $FlowDiagnostic 
# Apply Double Buffering to the main TabControl for smooth tab switching
Enable-DoubleBuffering $TabControl

# Background Image
# Define the path for the background texture image
$ImgPath = "$PSScriptRoot\Images\WhiteBG.png"
# Check if the image file exists before attempting to apply it 
if (Test-Path $ImgPath) {
    # Load and set the background image for the Diagnostic tab
    $TabDiagnostic.BackgroundImage = [System.Drawing.Image]::FromFile($ImgPath)
    # Stretch the image to cover the entire tab area
    $TabDiagnostic.BackgroundImageLayout = "Stretch"
}

# Set the layout panel background to transparent to show the TabPage's background image
$FlowDiagnostic.BackColor = [System.Drawing.Color]::Transparent



# ==============================================================================
#                         BUTTON --> Cleanup Temp Files
# ==============================================================================

# Create the button object for the Cleanup tool
$Btn_Cleanup = New-Object System.Windows.Forms.Button
# Set the visible text on the button
$Btn_Cleanup.Text = "Cleanup Temp Files"
# Apply your custom formatting function for scaling and styles
Format-Button $Btn_Cleanup

# Define what happens when the button is clicked
$Btn_Cleanup.Add_Click({
    # Display a warning message to the user to confirm the deletion process 
    $Confirm = [System.Windows.Forms.MessageBox]::Show("Delete all temp files and empty Recycle Bin?", "Confirm Cleanup", "YesNo", "Warning")
    
    # If the user clicks 'Yes', proceed with the cleanup logic
    if ($Confirm -eq "Yes") {
        try {
            # 1. Clean User Temp: Deletes files in the current user's temporary folder 
            Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
            
            # 2. Clean System Temp: Deletes global Windows temporary files (requires Admin rights) 
            Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
            
            # 3. Empty Recycle Bin: Permanently removes all items from the bin for all drives 
            Clear-RecycleBin -Force -ErrorAction SilentlyContinue
            
            # 4. Clean Windows Update Cache: Stops the update service to release file locks 
            Stop-Service -Name "wuauserv" -Force -ErrorAction SilentlyContinue
            # Removes downloaded update files that have already been installed
            Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
            # Restarts the Windows Update service
            Start-Service -Name "wuauserv" -ErrorAction SilentlyContinue

            # Show a success message once all tasks are finished
            [System.Windows.Forms.MessageBox]::Show("✅ Cleanup Complete!", "Success")
        } catch {
            # Capture and display any errors (like permission issues) in a popup
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed")
        }
    }
})

# Add the button to the Diagnostic flow layout panel
$FlowDiagnostic.Controls.Add($Btn_Cleanup)



# ==============================================================================
#                         BUTTON --> System Slowness Troubleshooter
# ==============================================================================

# Create the button object for the Slowness Troubleshooter
$Btn_Slow = New-Object System.Windows.Forms.Button
# Set the visible text on the button
$Btn_Slow.Text = "Troubleshoot Slowness"
# Apply custom formatting for DPI scaling and visual style
Format-Button $Btn_Slow

# Define the action to take when the button is clicked
$Btn_Slow.Add_Click({
    # Prepare a detailed warning message about execution time and reboots 
    $Msg = "This process runs deep diagnostics (SFC, DISM, Defrag, CHKDSK).`nIt can take 30-60 minutes and may schedule a reboot.`n`nLaunch maintenance window?"
    # Show confirmation dialog; execution only continues if user selects 'Yes'
    $Confirm = [System.Windows.Forms.MessageBox]::Show($Msg, "System Maintenance", "YesNo", "Warning")

    if ($Confirm -eq "Yes") {
        # 1. Define the Batch Script Content (Embedded heredoc)
        # We use a batch file to handle legacy commands like SFC and CHKDSK easily
        $BatchScript = @"
@echo off
Title System Slowness Diagnostics Tool
color 1f
cls
echo ===========================================
echo    Starting System Slowness Diagnostics
echo ===========================================
echo.

REM --- 1. METRICS: Gather current hardware load data ---
echo [1/11] CPU Load: & wmic cpu get loadpercentage /value
echo [2/11] Free Memory: & wmic os get freephysicalmemory /value
echo [3/11] Disk Usage: & wmic logicaldisk get size,freespace,caption
echo.

REM --- 2. REPAIRS: Fix system corruption and clear clutter ---
echo [6/11] Running System File Checker (SFC)...
sfc /scannow
echo.
echo [7/11] Running DISM Image Repair...
Dism /online /cleanup-image /restorehealth
echo.
echo [8/11] Disk Cleanup...
cleanmgr /d C: /sagerun:1
echo.
echo [9/11] Clearing Temp Files...
rmdir /s /q %temp%
echo.

REM --- 3. MAINTENANCE: Optimize disk performance ---
echo [10/11] Defragging Drive C...
defrag C: /U /V
echo.
echo [11/11] Scheduling Disk Check...
echo You must reboot for this to run.
chkdsk C: /f /r
echo.
echo ===========================================
echo    Diagnostics Complete.
echo ===========================================
pause
"@ 
        # 2. Save the batch content to a temporary file in the user's temp directory 
        $TempBatch = "$env:TEMP\Run-Diagnostics.bat"
        # Force overwrite the file and ensure ASCII encoding for CMD compatibility
        $BatchScript | Out-File $TempBatch -Encoding ASCII -Force

        # 3. Execute in a new visible CMD window with Admin privileges
        try {
            # Verb 'RunAs' triggers the UAC prompt for elevation 
            Start-Process cmd.exe -ArgumentList "/c `"$TempBatch`"" -Verb RunAs
        } catch {
            # Catch errors if the user denies the UAC prompt
            [System.Windows.Forms.MessageBox]::Show("Failed to launch. User cancelled UAC.", "Error")
        }
    }
})

# Add the button to the Diagnostic tab's layout panel (Updated from FlowSec)
$FlowDiagnostic.Controls.Add($Btn_Slow)



# ==============================================================================
#                         BUTTON --> SYSTEM HEALTH CHECK
# ==============================================================================

# BUTTON: SYSTEM HEALTH CHECK
# Initialize a new button object for launching a comprehensive system checkup
$Btn_Health = New-Object System.Windows.Forms.Button
# Set the visible text displayed on the button
$Btn_Health.Text = "System Health Check"
# Apply the toolkit's standard visual styling and DPI scaling
Format-Button $Btn_Health

# Define the script logic to execute when the button is clicked
$Btn_Health.Add_Click({
    # Prepare a descriptive message warning the user about reboots and deep repairs 
    $Msg = "This runs SFC, DISM, CHKDSK, and cleanup. A REBOOT is required at the end. Continue?"
    
    # Show a Yes/No confirmation box before proceeding with heavy operations
    if ([System.Windows.Forms.MessageBox]::Show($Msg, "Confirm", "YesNo") -eq "Yes") {
        
        # Change the mouse pointer to a waiting spinner for visual feedback
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Launch a new elevated PowerShell window to run repair commands as Administrator
        Start-Process powershell -Verb RunAs -ArgumentList "-NoExit", "-Command", "
            Write-Host '--- Starting System Health Check ---' -ForegroundColor Yellow;
            
            # REPAIR: Fix Windows image corruption using Deployment Image Servicing and Management 
            DISM /Online /Cleanup-Image /RestoreHealth; 
            
            # SCAN: Verify and repair protected system files using System File Checker
            sfc /scannow; 
            
            # OPTIMIZE: Defragment and optimize drive C for better performance
            defrag C: /O /V;
            
            # CLEANUP: Remove user temporary files to free up disk space 
            Remove-Item '$env:TEMP\*' -Recurse -Force -ErrorAction SilentlyContinue;
            
            Write-Host 'Checkup complete. System will restart in 60 seconds.' -ForegroundColor Green;
            
            # REBOOT: Trigger a system restart to finalize all repairs 
            shutdown /r /t 60;
            
            # SCHEDULE: Queue a disk check to run on the next startup
            chkdsk C: /f /r; 
        "
        
        # Restore the mouse cursor to default once the process has launched
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the button into the Diagnostic Tab's layout container (Updated to Diagnostic panel) [[2](https://github.com/lazywinadmin/WinFormPS)]
$FlowDiagnostic.Controls.Add($Btn_Health)



# ==============================================================================
#                         BUTTON --> AUDIO TROUBLESHOOTER
# ==============================================================================

# BUTTON: AUDIO TROUBLESHOOTER
# Create a new button object for the Audio diagnostic tool
$Btn_AudioTrouble = New-Object System.Windows.Forms.Button
# Set the visible text displayed on the button
$Btn_AudioTrouble.Text = "Sound/Microphone Troubleshooter"
# Apply the toolkit's standard visual styling and DPI scaling
Format-Button $Btn_AudioTrouble

# Define the logic to execute when the button is clicked
$Btn_AudioTrouble.Add_Click({
    # Change the mouse pointer to a waiting spinner for visual feedback
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # 1. Launch Built-in Troubleshooters: Invoke Microsoft Support Diagnostic Tool for playback and recording
        Start-Process msdt.exe -ArgumentList "/id AudioPlaybackDiagnostic"
        Start-Process msdt.exe -ArgumentList "/id AudioRecordingDiagnostic"

        # 2. Restart Audio Services: Force a refresh of the core Windows Audio service
        Restart-Service -Name Audiosrv -Force -ErrorAction SilentlyContinue
        
        # 3. Reset Audio Devices: Identify hardware devices reporting an 'Error' status in the Media class
        $ErrorDevices = Get-PnpDevice | Where-Object { $_.Status -eq 'Error' -and $_.Class -eq 'Media' }
        foreach ($Device in $ErrorDevices) {
            # Cycle the device state (Disable then Enable) to clear temporary hardware glitches
            Disable-PnpDevice -InstanceId $Device.InstanceId -Confirm:$false
            Enable-PnpDevice -InstanceId $Device.InstanceId -Confirm:$false
        }

        # Notify the user that maintenance tasks have been triggered successfully
        [System.Windows.Forms.MessageBox]::Show("✅ Audio services restarted and troubleshooters launched.", "Audio Diagnostic")
    } catch {
        # Catch and display any errors, such as missing permissions
        [System.Windows.Forms.MessageBox]::Show("❌ Error: $($_.Exception.Message)", "Troubleshooter Failed")
    } finally {
        # Restore the standard mouse pointer
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
# Add the button to the Diagnostic Tab's layout container
$FlowDiagnostic.Controls.Add($Btn_AudioTrouble)



# ==============================================================================
#                         BUTTON --> Teams Cache Delete
# ==============================================================================

# BUTTON: TEAMS CACHE BY USER
# Initialize a new button object for clearing the Teams cache of a specific user
$Btn_TeamsUser = New-Object System.Windows.Forms.Button
# Set the visible text displayed on the button
$Btn_TeamsUser.Text = "Teams Cache Delete"
# Apply the toolkit's standard visual styling and DPI scaling
Format-Button $Btn_TeamsUser

# Define the script logic to execute when the button is clicked
$Btn_TeamsUser.Add_Click({
    # 1. Get all local user folders: Scan the C:\Users directory, excluding system and public profiles
    $UserFolders = Get-ChildItem -Path C:\Users -Directory | Where-Object { $_.Name -notmatch "Public|Default|All Users" }
    
    # 2. Show selection menu: Use Out-GridView to present a searchable list of users for the analyst to choose from
    $SelectedUser = $UserFolders | Select-Object @{N='Username';E={$_.Name}}, @{N='Profile Path';E={$_.FullName}} | 
                    Out-GridView -Title "Select User Profile to Clear Teams Cache" -OutputMode Single

    # Proceed only if a user was selected from the list
    if ($SelectedUser) {
        # Display a warning to confirm that Teams should be closed and data deleted
        $Confirm = [System.Windows.Forms.MessageBox]::Show("Close Teams and clear cache for $($SelectedUser.Username)?", "Confirm", "YesNo", "Warning")
        
        if ($Confirm -eq "Yes") {
            # Change the mouse pointer to a waiting spinner
            $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            # 3. Close Teams: Forcefully terminate any running Teams processes to release file locks
            Get-Process ms-teams, Teams -ErrorAction SilentlyContinue | Stop-Process -Force
            # Brief pause to ensure the application has fully shut down
            Start-Sleep -Seconds 2

            # 4. Target path for selected user: Define the specific AppData path where Teams stores its cache
            $CachePath = "$($SelectedUser.'Profile Path')\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe"
            
            # Check if the cache directory actually exists before attempting deletion
            if (Test-Path $CachePath) {
                # Recursively delete all files and subfolders within the Teams cache directory
                Remove-Item -Path "$CachePath\*" -Recurse -Force -ErrorAction SilentlyContinue 
                # Notify the analyst of a successful operation
                [System.Windows.Forms.MessageBox]::Show("✅ Cache cleared for $($SelectedUser.Username).", "Success")
            } else {
                # Inform the analyst if the folder path could not be found
                [System.Windows.Forms.MessageBox]::Show("❌ Teams cache folder not found for this user.", "Not Found")
            }
            # Restore the standard mouse pointer
            $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})

# Add the button into the Diagnostic Tab's layout container (FlowSec)
$FlowDiagnostic.Controls.Add($Btn_TeamsUser)



# ==============================================================================
#                         BUTTON --> OUTLOOK REPAIR
# ==============================================================================

# BUTTON: OUTLOOK REPAIR
$Btn_OutlookTrouble = New-Object System.Windows.Forms.Button
$Btn_OutlookTrouble.Text = "Outlook Troubleshooting"
Format-Button $Btn_OutlookTrouble

$Btn_OutlookTrouble.Add_Click({
    # 1. Create the specialized Sub-Window
    $OutlookForm = New-Object System.Windows.Forms.Form
    $OutlookForm.Text = "Outlook Repair Center"; $OutlookForm.Size = "450,550"; $OutlookForm.StartPosition = "CenterParent"
    $OutlookForm.BackColor = "#f0f0f0" # Light gray background for contrast
    
    # 2. Main Layout Container
    $OutFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $OutFlow.Dock = "Fill"; $OutFlow.AutoScroll = $true; $OutFlow.Padding = New-Object System.Windows.Forms.Padding(15)
    $OutFlow.FlowDirection = "TopDown"; $OutFlow.WrapContents = $false # Forces vertical stacking
    $OutlookForm.Controls.Add($OutFlow)

    # --- HELPER: Browser-Style Button Formatter ---
    function Format-OutSubButton ($Btn) {
        $Btn.Width = 400; $Btn.Height = 50
        $Btn.BackColor = "White"; $Btn.FlatStyle = "Flat"
        $Btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $Btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
        $Btn.TextAlign = "MiddleLeft" # Aligns text to the left like the browser menu
        $Btn.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
    }

    # 3. Define Repair Actions
    $Actions = @{
        "1. Start in Safe Mode"      = @("/safe", "Starts Outlook without add-ins (fixes startup crashes)")
        "2. Reset Navigation Pane"   = @("/resetnavpane", "Regenerates the Folder Pane for the current profile")
        "3. Clean Interface Views"   = @("/cleanviews", "Restores default views. WARNING: Custom views will be lost.")
        "4. Restore Missing Folders" = @("/resetfolders", "Restores missing system folders for default delivery location.")
        "5. View Performance Logs"   = @("/PerfLog", "Opens Application Event Logs to check for Outlook issues")
    }

    # 4. Generate the Buttons
    $Actions.Keys | Sort-Object | ForEach-Object {
        $Key = $_; $B = New-Object System.Windows.Forms.Button
        $B.Text = $Key; Format-OutSubButton $B
        
        $ActionData = $Actions[$Key]
        $B.Add_Click({
            if ([System.Windows.Forms.MessageBox]::Show($ActionData[1], "Confirm Outlook Action", "YesNo", "Question") -eq "Yes") {
                Start-Process "outlook.exe" -ArgumentList $ActionData[0]
            }
        }.GetNewClosure())
        
        $OutFlow.Controls.Add($B)
    }

    $OutlookForm.ShowDialog()
})

$FlowDiagnostic.Controls.Add($Btn_OutlookTrouble)



# ==============================================================================
#                         BUTTON --> Browser Repair Center
# The Browser Repair Center implements a highly modular troubleshooting system that allows IT analysts to resolve application-specific issues for Edge, Chrome, and Firefox.
# Logic Highlights:
#                  *Dynamic Hierarchy: The tool uses a selection form ($SelForm) that dynamically launches browser-specific sub-menus using the Show-BrowserMenu helper function.
#                  *Safety Procedures: Every high-impact action, such as clearing cache or resetting profiles, triggers a standard Windows "Yes/No" dialog to ensure analyst confirmation.
#                  *Path Resolution: The script automatically resolves localized environment variables (like $env:LOCALAPPDATA) and performs pattern matching for dynamic Mozilla Firefox profile names.
#                  *Layout Update: Per your request, the button is now correctly added to the $FlowDiagnostic container to maintain the toolkit's updated tab structure.
# ==============================================================================

# BUTTON: BROWSER REPAIR CENTER (Modular)
# Create a new button object for the Browser Repair toolkit
$Btn_BrowserFix = New-Object System.Windows.Forms.Button
# Set the visible label for the main toolkit button
$Btn_BrowserFix.Text = "Browser Repair Center"
# Apply global formatting and high-DPI scaling via your custom function
Format-Button $Btn_BrowserFix

# Define the script logic to execute when the main button is clicked
$Btn_BrowserFix.Add_Click({
    # 1. Main Selection Window Initialization
    $SelForm = New-Object System.Windows.Forms.Form
    # Set title, dimensions, and ensure it pops up centered over the toolkit
    $SelForm.Text = "Select Browser to Repair"
    $SelForm.Size = New-Object System.Drawing.Size(350, 400)
    $SelForm.StartPosition = "CenterParent"
    $SelForm.BackColor = "#f0f0f0"

    # Create a layout panel to organize the primary browser selection buttons
    $FlowSel = New-Object System.Windows.Forms.FlowLayoutPanel
    $FlowSel.Dock = "Fill"
    $FlowSel.Padding = New-Object System.Windows.Forms.Padding(20)
    $FlowSel.FlowDirection = "TopDown" # Stack buttons vertically
    $SelForm.Controls.Add($FlowSel)

    # --- HELPER FUNCTION: Create Browser Menu ---
    # This reusable function builds a unique repair sub-window for any browser
    function Show-BrowserMenu ($BrowserName, $ProcessName, $ProfilePath, $CachePaths) {
        $BForm = New-Object System.Windows.Forms.Form
        $BForm.Text = "$BrowserName Troubleshooting"
        $BForm.Size = New-Object System.Drawing.Size(400, 450)
        $BForm.StartPosition = "CenterParent"
        
        # Internal panel for specific repair action buttons
        $BFlow = New-Object System.Windows.Forms.FlowLayoutPanel
        $BFlow.Dock = "Fill"
        $BFlow.FlowDirection = "TopDown"
        $BFlow.Padding = New-Object System.Windows.Forms.Padding(10)
        $BForm.Controls.Add($BFlow)

        # Style Helper for sub-buttons inside the browser repair window
        function Add-ActionBtn ($Text, $ScriptBlock) {
            $Btn = New-Object System.Windows.Forms.Button
            $Btn.Text = $Text
            $Btn.Width = 360
            $Btn.Height = 45
            $Btn.FlatStyle = "Flat"
            $Btn.BackColor = "White"
            $Btn.Margin = New-Object System.Windows.Forms.Padding(0,0,0,10)
            $Btn.Add_Click($ScriptBlock) # Attach the specific repair logic
            $BFlow.Controls.Add($Btn)
        }

        # === ACTION 1: KILL PROCESS ===
        # Forcefully ends the browser tasks to resolve freezing or locked files
        Add-ActionBtn "1. Force Close $BrowserName" {
            Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
            [System.Windows.Forms.MessageBox]::Show("✅ $BrowserName has been terminated.", "Done")
        }

        # === ACTION 2: CLEAR CACHE ===
        # Deletes temp files from known paths while preserving user bookmarks
        Add-ActionBtn "2. Clear Cache (Keep Bookmarks)" {
            if ([System.Windows.Forms.MessageBox]::Show("Close $BrowserName and clear temporary cache files?", "Confirm", "YesNo") -eq "Yes") {
                Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
                
                $Deleted = 0
                foreach ($P in $CachePaths) {
                    if (Test-Path $P) { 
                        Remove-Item -Path "$P\*" -Recurse -Force -ErrorAction SilentlyContinue 
                        $Deleted++
                    }
                }
                [System.Windows.Forms.MessageBox]::Show("✅ Cache cleared from $Deleted folder locations.", "Success")
            }
        }

        # === ACTION 3: SOFT PROFILE RESET ===
        # Renames the user profile to .OLD to fix deep corruption without permanent loss
        Add-ActionBtn "3. Soft Reset (Rename Profile)" {
            $Msg = "This will rename your '$BrowserName' profile to '.OLD'.`n`n• You will lose extensions and settings.`n• Bookmarks can be recovered from the .OLD folder.`n`nContinue?"
            
            if ([System.Windows.Forms.MessageBox]::Show($Msg, "Warning", "YesNo", "Warning") -eq "Yes") {
                try {
                    Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2
                    
                    if (Test-Path $ProfilePath) {
                        # Create a timestamped backup folder name
                        $NewName = "$ProfilePath.OLD_$(Get-Date -Format 'yyyyMMdd_HHmm')"
                        Rename-Item -Path $ProfilePath -NewName $NewName
                        [System.Windows.Forms.MessageBox]::Show("✅ Profile Reset!`n`nOld Profile: $NewName`n`n$BrowserName will now restart clean.", "Success")
                        Start-Process "$ProcessName.exe"
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("❌ Default profile folder not found.`nPath: $ProfilePath", "Error")
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed")
                }
            }
        }

        # === ACTION 4: EXTENSION CHECK ===
        # Directly opens the directory where browser extensions are stored for audit
        Add-ActionBtn "4. Open Extensions Folder" {
            $ExtPath = Join-Path (Split-Path $ProfilePath) "Extensions"
            if (-not (Test-Path $ExtPath)) { $ExtPath = $ProfilePath } # Fallback to profile root
            
            if (Test-Path $ExtPath) { Invoke-Item $ExtPath } 
            else { [System.Windows.Forms.MessageBox]::Show("Extension folder not found.", "Info") }
        }

        $BForm.ShowDialog() # Display the browser-specific sub-menu
    }

    # === MAIN MENU BUTTONS: Selection Logic ===

    # 1. MICROSOFT EDGE Configuration
    $Btn_Edge = New-Object System.Windows.Forms.Button
    $Btn_Edge.Text = "Microsoft Edge"
    $Btn_Edge.Width = 280; $Btn_Edge.Height = 60
    $Btn_Edge.BackColor = "#0078D7"; $Btn_Edge.ForeColor = "White"
    $Btn_Edge.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $Btn_Edge.Margin = New-Object System.Windows.Forms.Padding(0,0,0,15)
    
    $Btn_Edge.Add_Click({
        Show-BrowserMenu "Edge" "msedge" "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default" @(
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache",
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache"
        )
    })
    $FlowSel.Controls.Add($Btn_Edge)

    # 2. GOOGLE CHROME Configuration
    $Btn_Chrome = New-Object System.Windows.Forms.Button
    $Btn_Chrome.Text = "Google Chrome"
    $Btn_Chrome.Width = 280; $Btn_Chrome.Height = 60
    $Btn_Chrome.BackColor = "#DB4437"; $Btn_Chrome.ForeColor = "White"
    $Btn_Chrome.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $Btn_Chrome.Margin = New-Object System.Windows.Forms.Padding(0,0,0,15)

    $Btn_Chrome.Add_Click({
        Show-BrowserMenu "Chrome" "chrome" "$env:LOCALAPPDATA\Google\Chrome\User Data\Default" @(
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Code Cache",
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\GPUCache"
        )
    })
    $FlowSel.Controls.Add($Btn_Chrome)

    # 3. MOZILLA FIREFOX Configuration
    $Btn_Firefox = New-Object System.Windows.Forms.Button
    $Btn_Firefox.Text = "Mozilla Firefox"
    $Btn_Firefox.Width = 280; $Btn_Firefox.Height = 60
    $Btn_Firefox.BackColor = "#FF9400"; $Btn_Firefox.ForeColor = "White"
    $Btn_Firefox.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    
    $Btn_Firefox.Add_Click({
        # Find the dynamic Firefox Profile string ending in 'default-release'
        $FFRoot = "$env:APPDATA\Mozilla\Firefox\Profiles"
        if (Test-Path $FFRoot) {
            $Profile = Get-ChildItem -Path $FFRoot | Where-Object { $_.Name -like "*default-release" } | Select-Object -First 1
            
            if ($Profile) {
                Show-BrowserMenu "Firefox" "firefox" $Profile.FullName @(
                    "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\$($Profile.Name)\cache2",
                    "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\$($Profile.Name)\startupCache"
                )
            } else {
                [System.Windows.Forms.MessageBox]::Show("Could not find a standard 'default-release' Firefox profile.", "Error")
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Firefox AppData folder not found.`nIs Firefox installed?", "Error")
        }
    })
    $FlowSel.Controls.Add($Btn_Firefox)

    $SelForm.ShowDialog() # Display the initial browser picker
})

# Finalize placement: Add the tool to the toolkit's Diagnostic Hub
# FIXED: Updated $FlowSec to $FlowDiagnostic for tab consistency
$FlowDiagnostic.Controls.Add($Btn_BrowserFix)



# ==============================================================================
#                         BUTTON --> Intune Device Manager
# The Intune Device Manager provides a specialized GUI for managing Microsoft Intune and Entra ID compliance.
# It solves common "stuck" synchronization issues through three primary mechanisms:
#                  *Direct Compliance Validation: It attempts to read the Company Portal Local Cache to find the exact compliance state without waiting for cloud updates.
#                  *Forced Synchronization: It utilizes specialized protocol triggers (intunemanagementextension://synccompliance) to force the Intune Management Extension agent to check in with the server immediately.
#                  *Service Remediation: It includes a one-click button to restart the Intune Management Extension service, which is a standard troubleshooting step when application deployments are hanging or failing.
# This tool simplifies complex command-line tasks like dsregcmd /status and agent manipulation into a simple interface for IT analysts.
# ==============================================================================

# [BUTTON] INTUNE DEVICE MANAGER
# Create the main button object for launching the Intune Manager sub-window
$Btn_Intune = New-Object System.Windows.Forms.Button
# Set the visible text displayed on the main toolkit button
$Btn_Intune.Text = "Intune Device Manager"
# Apply the global button formatting and DPI scaling logic
Format-Button $Btn_Intune

# Define the logic to execute when the Intune Manager button is clicked
$Btn_Intune.Add_Click({
    # 1. Create the specialized Sub-Window for Intune Management
    $IntuneForm = New-Object System.Windows.Forms.Form
    # Set the title bar text for the new window
    $IntuneForm.Text = "Intune & Compliance Manager"
    # Define window dimensions with a height of 500 pixels
    $IntuneForm.Size = New-Object System.Drawing.Size(450, 500)
    # Ensure the sub-form pops up in the center of the main toolkit window
    $IntuneForm.StartPosition = "CenterParent"
    # Set a light gray background color for the form
    $IntuneForm.BackColor = "#f0f0f0"
    
    # Initialize a FlowLayoutPanel to organize sub-buttons vertically
    $IntFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    # Dock the panel to fill the entire form area
    $IntFlow.Dock = "Fill"
    # Enable a scrollbar if the buttons exceed the window height
    $IntFlow.AutoScroll = $true
    # Add a 10-pixel margin around the inside of the panel
    $IntFlow.Padding = New-Object System.Windows.Forms.Padding(10)
    # Arrange buttons in a vertical "Top-Down" column
    $IntFlow.FlowDirection = "TopDown"
    # Prevent buttons from wrapping into a second column
    $IntFlow.WrapContents = $false
    # Add the layout panel to the Intune form's control collection
    $IntuneForm.Controls.Add($IntFlow)

    # --- HELPER: Local Function to style sub-buttons consistently ---
    function Format-SubButton ($Btn) {
        # Set a fixed width for buttons to match the layout
        $Btn.Width = 400
        # Set a fixed height for a professional dashboard look
        $Btn.Height = 50
        # Set button background to white for contrast
        #$Btn.BackColor = "White"
        # Use a modern flat border style
        $Btn.FlatStyle = "Flat"
        # Set bold Segoe UI font for the button labels
        $Btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        # Add bottom spacing between buttons
        $Btn.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
        # Align text to the middle-left for readability
        $Btn.TextAlign = "MiddleLeft"
        # Add left padding so text doesn't touch the button edge
        $Btn.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
        # Border color Black
        $Btn.FlatAppearance.BorderColor = "Black"
    }

    # === OPTION 1: DEEP COMPLIANCE CHECK (App Cache + Cloud) ===
    # This tool validates device health using the Company Portal's local data
    $Btn_CompCheck = New-Object System.Windows.Forms.Button
    $Btn_CompCheck.Text = "1. Check Compliance (App Cache Method)"
    $Btn_CompCheck.BackColor = "#0078D7"; $Btn_CompCheck.ForeColor = "White"
    # Apply the sub-button styling
    Format-SubButton $Btn_CompCheck
    $Btn_CompCheck.Add_Click({
        # Show a loading spinner cursor during the check
        $IntuneForm.Cursor = "WaitCursor"
        $StatusFound = $false
        
        # METHOD A: Identify the local JSON cache path for the Company Portal App
        $PkgPath = "$env:LOCALAPPDATA\Packages\Microsoft.CompanyPortal_8wekyb3d8bbwe\TempState\ApplicationCache"
        
        # Check if the cache directory exists on the system
        if (Test-Path $PkgPath) {
            try {
                # Find the most recently updated JSON status file in the cache
                $CacheFile = Get-ChildItem -Path $PkgPath -Filter *.json -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                
                if ($CacheFile) {
                    # Read the raw JSON content from the file
                    $JsonContent = Get-Content $CacheFile.FullName -Raw | ConvertFrom-Json
                    # Extract the nested 'ComplianceState' value from the portal data
                    $InnerData = $JsonContent.data | ConvertFrom-Json
                    $AppStatus = $InnerData.ComplianceState
                    
                    if ($AppStatus) {
                        $StatusFound = $true
                        # Display a success message if the device is compliant
                        if ($AppStatus -eq "Compliant") {
                            [System.Windows.Forms.MessageBox]::Show("✅ APP CONFIRMED: Device is COMPLIANT.`n`nSource: Company Portal App Cache`nLast Update: $($CacheFile.LastWriteTime)", "Compliance Status")
                        } else {
                             # Warn the user if the app reports a non-compliant state
                             [System.Windows.Forms.MessageBox]::Show("⚠️ APP STATUS: $AppStatus`n`nThe Company Portal app is reporting this device as Non-Compliant.", "Compliance Status")
                        }
                    }
                }
            } catch {
                # Silently catch errors if the cache file is locked or unreadable
            }
        }

        # METHOD B: Fallback to the Entra ID status CLI if app data is unavailable
        if (-not $StatusFound) {
            # Execute dsregcmd to get the registration status and redirect output to a temp file
            $Proc = Start-Process dsregcmd.exe -ArgumentList "/status" -NoNewWindow -PassThru -RedirectStandardOutput "$env:TEMP\dsreg.txt"
            $Proc.WaitForExit()
            $Content = Get-Content "$env:TEMP\dsreg.txt" | Out-String
            
            # Check for the presence of a valid Primary Refresh Token (AzureAdPrt)
            if ($Content -match "AzureAdPrt\s+:\s+YES") {
                $Msg = "✅ HYBRID JOINED & CONNECTED.`n`nThis device has a valid Primary Refresh Token (PRT).`n`nNote: 'IsDeviceCompliant' is not visible in CLI for Hybrid devices.`nPlease open the Company Portal app to confirm visual status."
                $Icon = "Information"
            } else {
                $Msg = "❌ ERROR: No PRT Token found.`nDevice is not communicating with Azure AD properly."
                $Icon = "Error"
            }
            
            # Offer to open the Company Portal app to resolve status issues
            if ([System.Windows.Forms.MessageBox]::Show("$Msg`n`nDo you want to open the Company Portal App now?", "Status Result", "YesNo", $Icon) -eq "Yes") {
                Start-Process "companyportal:"
            }
        }
        # Reset the cursor to default
        $IntuneForm.Cursor = "Default"
    })
    # Add the compliance button to the layout panel
    $IntFlow.Controls.Add($Btn_CompCheck)

    # === OPTION 2: FORCE AGENT SYNC (The New Command) ===
    # This tool triggers the Intune Management Extension agent directly
    $Btn_AgentSync = New-Object System.Windows.Forms.Button
    $Btn_AgentSync.Text = "2. Force Agent Compliance Sync"
    $Btn_AgentSync.BackColor = "#ed0e0e"; $Btn_AgentSync.ForeColor = "White"
    Format-SubButton $Btn_AgentSync
    $Btn_AgentSync.Add_Click({
        # Define the path to the Intune background agent executable
        $AgentPath = "C:\Program Files (x86)\Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe"
        
        if (Test-Path $AgentPath) {
            # Execute a specific URI trigger to force the agent into a compliance check
            Start-Process -FilePath $AgentPath -ArgumentList "intunemanagementextension://synccompliance"
            [System.Windows.Forms.MessageBox]::Show("✅ Sync command sent to Agent.", "Sync Initiated")
        } else {
            # Error if the device is likely not managed by Intune
            [System.Windows.Forms.MessageBox]::Show("❌ Agent not found. Is this device enrolled?", "Error")
        }
    })
    $IntFlow.Controls.Add($Btn_AgentSync)

    # === OPTION 3: FORCE BACKGROUND TASK SYNC ===
    # This tool runs the built-in Windows scheduled sync task
    $Btn_Sync = New-Object System.Windows.Forms.Button
    $Btn_Sync.Text = "3. Run Scheduled Sync Task"
    $Btn_Sync.BackColor = "#FF9400"; $Btn_Sync.ForeColor = "White"
    Format-SubButton $Btn_Sync
    $Btn_Sync.Add_Click({
        $IntuneForm.Cursor = "WaitCursor"
        # Locate the specific synchronization task created by the MDM client
        $Task = Get-ScheduledTask | Where-Object { $_.TaskName -like "Schedule #3 created by client*" } 
        
        if ($Task) {
            # Start the task to trigger a policy check-in with the cloud
            Start-ScheduledTask -TaskName $Task.TaskName -TaskPath $Task.TaskPath
            [System.Windows.Forms.MessageBox]::Show("✅ Scheduled Task Started.`n`nWindows is syncing policies in the background.", "Sync Started")
        } else {
            # Fallback to the DeviceEnroller tool if the scheduled task is missing
            Start-Process DeviceEnroller.exe -ArgumentList "/c /AutoEnrollMDM" -NoNewWindow -Wait
            [System.Windows.Forms.MessageBox]::Show("✅ Fallback: Sync command sent to DeviceEnroller.", "Sync Started")
        }
        $IntuneForm.Cursor = "Default"
    })
    $IntFlow.Controls.Add($Btn_Sync)

    # === OPTION 4: RESTART INTUNE SERVICE ===
    # This tool restarts the IME service to unfreeze app installations
    $Btn_RestartIME = New-Object System.Windows.Forms.Button
    $Btn_RestartIME.Text = "4. Restart Intune Service (Fix Stuck Apps)"
    $Btn_RestartIME.BackColor = "#9200d1"; $Btn_RestartIME.ForeColor = "White"
    Format-SubButton $Btn_RestartIME
    $Btn_RestartIME.Add_Click({
        # Ensure the script is running with elevated privileges
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            [System.Windows.Forms.MessageBox]::Show("Run as Administrator to restart services.", "Access Denied"); return
        }

        $IntuneForm.Cursor = "WaitCursor"
        try {
            # Forcefully restart the Intune Management Extension service
            Restart-Service -Name "IntuneManagementExtension" -Force -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show("✅ Intune Service Restarted.", "Success")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Service not found or Access Denied.", "Error")
        }
        $IntuneForm.Cursor = "Default"
    })
    $IntFlow.Controls.Add($Btn_RestartIME)

    # === OPTION 5: AZURE AD STATUS ===
    # This tool displays the full Entra ID registration log
    $Btn_Dsreg = New-Object System.Windows.Forms.Button
    $Btn_Dsreg.Text = "5. View Azure AD Status Log"
    $Btn_Dsreg.BackColor = "#d1009d"; $Btn_Dsreg.ForeColor = "White"
    Format-SubButton $Btn_Dsreg
    $Btn_Dsreg.Add_Click({
        $IntuneForm.Cursor = "WaitCursor"
        # Export full registration details to a text file
        $Proc = Start-Process dsregcmd.exe -ArgumentList "/status" -NoNewWindow -PassThru -RedirectStandardOutput "$env:TEMP\dsreg_full.txt"
        $Proc.WaitForExit()
        
        # Create a new viewer form to show the log content
        $ViewForm = New-Object System.Windows.Forms.Form
        $ViewForm.Text = "Azure AD Status"; $ViewForm.Size = "600,600"; $ViewForm.StartPosition = "CenterParent"
        $Txt = New-Object System.Windows.Forms.TextBox
        # Configure the text box to show multiline log data with a scrollbar
        $Txt.Multiline = $true; $Txt.ScrollBars = "Vertical"; $Txt.Dock = "Fill"; $Txt.Font = "Consolas, 10"
        $Txt.Text = Get-Content "$env:TEMP\dsreg_full.txt" -Raw
        $ViewForm.Controls.Add($Txt)
        $ViewForm.ShowDialog()
        $IntuneForm.Cursor = "Default"
    })
    $IntFlow.Controls.Add($Btn_Dsreg)

    # Display the final Intune Manager sub-form
    $IntuneForm.ShowDialog()
})

# Add the completed Intune button to the main Diagnostic panel collection
$FlowDiagnostic.Controls.Add($Btn_Intune)



# ==============================================================================
#                         BUTTON --> DOMAIN ACCOUNT MANAGEMENT (Modular)
# The Active Directory Dashboard provides a centralized interface for performing high-frequency domain management tasks.
# This module leverages the legacy but highly reliable net user command suite to interact with Active Directory without requiring the full RSAT (Remote Server Administration Tools) module to be loaded.
# Key Logic Highlights:
#                      *Independent Execution: Each command launches in a new, persistent PowerShell window (-NoExit). This prevents the main toolkit from freezing during long domain queries and allows analysts to verify the output.
#                      *Modular Interaction: Uses the Visual Basic InputBox for streamlined user data entry, ensuring the GUI remains clean and focused.
#                      *Visual Status Coding: Buttons are color-coded (Blue for Info, Green for Unlock, Orange for Resets) to help analysts identify tasks at a glance.
#                      *No RSAT Required: By using standard command-line tools, this dashboard remains functional on any domain-joined machine, even those without advanced management tools installed.
# ==============================================================================

# BUTTON: DOMAIN ACCOUNT MANAGEMENT (Modular)
# Create the main button object for Active Directory tasks
#$Btn_ADManage = New-Object System.Windows.Forms.Button
# Set the visible label for the button
#$Btn_ADManage.Text = "Domain Account Management"
# Apply standardized global formatting and DPI scaling to the button
#Format-Button $Btn_ADManage

# Define the script logic to execute when the button is clicked
#$Btn_ADManage.Add_Click({
    # 1. Initialize a modal sub-window for the Active Directory Dashboard
#    $ADForm = New-Object System.Windows.Forms.Form
    # Set title, dimensions, and ensure it pops up centered over the toolkit
#    $ADForm.Text = "Active Directory Dashboard"; $ADForm.Size = New-Object System.Drawing.Size(350, 400)
#    $ADForm.StartPosition = "CenterParent"; $ADForm.BackColor = "#f0f0f0"

    # Create a layout panel to organize AD action buttons vertically
#    $FlowAD = New-Object System.Windows.Forms.FlowLayoutPanel
    # Configure the panel to fill the window and stack buttons top-to-bottom
#    $FlowAD.Dock = "Fill"; $FlowAD.Padding = New-Object System.Windows.Forms.Padding(20); $FlowAD.FlowDirection = "TopDown"
#    $ADForm.Controls.Add($FlowAD)

    # 1. HELPER FUNCTION: Captures success/failure with interactive pop-ups
#    function Execute-ADCmd ($Cmd) {
        # Inform the user via the console what command is being triggered
#        Write-Host "Running: $Cmd"
        
        # Launch a separate PowerShell window to execute the AD command and show the output
        # Use -NoExit so analysts can review the results before closing the window
#        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "Write-Host 'EXECUTING: $Cmd' -ForegroundColor Cyan; $Cmd; Write-Host '`nProcess Complete. You can close this window.' -ForegroundColor Green"
#    }

    # Helper function to style and add buttons to the AD dashboard
#    function Add-DashAction ($Text, $Color, $Script) {
#        $Btn = New-Object System.Windows.Forms.Button
        # Set button properties: text, size, and vibrant color coding
#        $Btn.Text = $Text; $Btn.Width = 280; $Btn.Height = 60
#        $Btn.BackColor = $Color; $Btn.ForeColor = "White"; $Btn.FlatStyle = "Flat"
        # Use bold font for high visibility on the dashboard
#        $Btn.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        # Attach the provided script logic to the button's click event
#        $Btn.Add_Click($Script); $FlowAD.Controls.Add($Btn)
#    }

    # 2. Add individual Action Buttons using the styling helper
    
    # ACTION: Retrieve Domain User Properties
#    Add-DashAction "ACCOUNT PROPERTIES" "#007ACC" { 
        # Prompt for target username and execute the 'net user' query
#        $u = [Microsoft.VisualBasic.Interaction]::InputBox("Username?", "AD Query"); if($u){ Execute-ADCmd "net user $u /domain" } 
#    }
    
    # ACTION: Unlock a Locked Domain Account
#    Add-DashAction "UNLOCK ACCOUNT" "#28A745" { 
        # Prompt for username and trigger the unlock command via 'net user'
#        $u = [Microsoft.VisualBasic.Interaction]::InputBox("Username?", "AD Unlock"); if($u){ Execute-ADCmd "net user $u /active:yes /domain" } 
#    }
    
    # ACTION: Force a Domain Password Change
#    Add-DashAction "PASSWORD CHANGE" "#FF9400" { 
        # Collect both username and new password to perform the update
#        $u = [Microsoft.VisualBasic.Interaction]::InputBox("User?", "AD Reset"); $p = [Microsoft.VisualBasic.Interaction]::InputBox("Pass?", "AD Reset")
#        if($u -and $p){ Execute-ADCmd "net user $u $p /domain" }
#    }

    # Display the completed dashboard as a modal dialog
#    $ADForm.ShowDialog() 
#})

# Finalize placement: Add the button to the main Diagnostic Hub (FlowDiagnostic)
#$FlowDiagnostic.Controls.Add($Btn_ADManage)



# ==============================================================================
#                         BUTTON --> WINDOWS UPDATE MANAGER
# The Windows Update Manager block provides a GUI-driven way to manage system patches using the native Windows Update Agent (WUA) COM object.
# Key Logic Stages:
#                  *Scanning: It utilizes the Microsoft.Update.Session to perform an online search for uninstalled and non-hidden patches.
#                  *Validation: Analysts can review pending updates in a searchable grid before committing, ensuring control over system changes.
#                  *Automation Fixes: The script includes an automated EULA-acceptance loop, preventing installation failures caused by manual agreement prompts.
#                  *Reporting: It monitors the installation result in real-time and provides a status update, specifically flagging when a reboot is mandatory.
#                  *UI Feedback: Uses [System.Windows.Forms.Application]::DoEvents() to update button text dynamically (Scanning, Downloading, Installing) so the tool doesn't appear "frozen" during long operations.
# ==============================================================================

# BUTTON: WINDOWS UPDATE MANAGER
# Initialize the primary button object for Windows Update tasks
$Btn_Updates = New-Object System.Windows.Forms.Button
# Set the visible text displayed on the button
$Btn_Updates.Text = "Windows Update"
# Apply standardized global formatting and DPI scaling
Format-Button $Btn_Updates

# Define the script logic to execute when the button is clicked
$Btn_Updates.Add_Click({
    # Disable the button to prevent multiple scans simultaneously
    $Btn_Updates.Enabled = $false
    $Btn_Updates.Text = "Scanning..."
    # Force the UI to refresh and show the updated button text
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Initialize the COM object to interact with the Windows Update Agent
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Searcher = $Session.CreateUpdateSearcher()
        # Query for updates that are not yet installed and not hidden
        $Search = $Searcher.Search("IsInstalled=0 and IsHidden=0")
        $Updates = $Search.Updates

        if ($Updates.Count -gt 0) {
            # 1. Review: Gather titles and KB IDs for analyst inspection
            $Results = foreach ($Update in $Updates) {
                [PSCustomObject]@{ Title = $Update.Title; KB = $($Update.KBArticleIDs) -join ',' }
            }
            # Display findings in a searchable grid window for review
            $Results | Out-GridView -Title "Review Pending Updates ($($Updates.Count) found)" -Wait

            # 2. Confirm: Prompt the user before initiating download and installation
            $Confirm = [System.Windows.Forms.MessageBox]::Show("Found $($Updates.Count) updates. Install now?", "Confirm", "YesNo", "Question")
            
            if ($Confirm -eq "Yes") {
                # --- CRITICAL FIX: ACCEPT EULA ---
                # Some updates require explicit EULA acceptance before downloading
                $Btn_Updates.Text = "Preparing..."
                [System.Windows.Forms.Application]::DoEvents()
                foreach ($Update in $Updates) {
                    if (-not $Update.EulaAccepted) { $Update.AcceptEula() }
                }

                # 3. Download: Initiate the background download of confirmed updates
                $Btn_Updates.Text = "Downloading..."
                [System.Windows.Forms.Application]::DoEvents()
                $Downloader = $Session.CreateUpdateDownloader()
                $Downloader.Updates = $Updates
                $Downloader.Download() 

                # 4. Install: Finalize by launching the installer engine
                $Btn_Updates.Text = "Installing..."
                [System.Windows.Forms.Application]::DoEvents()
                $Installer = $Session.CreateUpdateInstaller()
                $Installer.Updates = $Updates
                
                # Execute the installation process
                $InstallResult = $Installer.Install()
                
                # 5. Audit: Determine if the machine requires a reboot to complete
                $StatusMsg = "Installation Finished.`n"
                if ($InstallResult.RebootRequired) { 
                    $StatusMsg += "⚠️ REBOOT REQUIRED to complete updates." 
                } else {
                    $StatusMsg += "✅ Updates installed successfully."
                }

                [System.Windows.Forms.MessageBox]::Show($StatusMsg, "Update Result")
            }
        } else {
            # Notify the analyst if no new updates are pending
            [System.Windows.Forms.MessageBox]::Show("✅ System is up to date.", "No Updates")
        }
    } catch {
        # Catch and display any engine errors (e.g., service not running)
        [System.Windows.Forms.MessageBox]::Show("❌ Error: $($_.Exception.Message)", "Failed")
    } finally {
        # Restore the button to its original state for future use
        $Btn_Updates.Enabled = $true
        $Btn_Updates.Text = "Windows Update"
    }
})

# Add the completed button to the Diagnostic panel
$FlowDiagnostic.Controls.Add($Btn_Updates)



# ==============================================================================
#                         BUTTON --> Check Disk Space
# he Check Disk Space tool provides a one-click audit for local storage health. By utilizing CIM instances
# it retrieves high-performance disk metrics and includes an automated safety check against a 15GB threshold.
# ==============================================================================

# BUTTON: Check Disk Space (C: Drive)
# Create a new Button object for disk space monitoring
$Btn_Disk = New-Object System.Windows.Forms.Button
# Set the text label for the button
$Btn_Disk.Text = "Check Disk Space (C:)"
# Apply the global button styling defined in your script
Format-Button $Btn_Disk

# Define what happens when the button is clicked
$Btn_Disk.Add_Click({
    # Define settings for the target drive and warning threshold
    $Drive = "C:"
    $ThresholdGB = 15

    try {
        # Query the Win32_LogicalDisk class for the specific drive
        $Disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$Drive'"
        # Calculate free space in GB and round to two decimal places
        $FreeSpaceGB = [math]::Round($Disk.FreeSpace / 1GB, 2)

        # Check if free space is below the defined threshold
        if ($FreeSpaceGB -lt $ThresholdGB) {
            # Display a warning message if space is low
            [System.Windows.Forms.MessageBox]::Show("⚠️ LOW DISK SPACE!`nDrive $Drive has only $FreeSpaceGB GB remaining.", "Warning", "OK", "Warning")
        } else {
            # Display a healthy status message
            [System.Windows.Forms.MessageBox]::Show("Healthy: Drive $Drive has $FreeSpaceGB GB free.", "Disk Status", "OK", "Information")
        }
    } catch {
        # Display an error message if the query fails
        [System.Windows.Forms.MessageBox]::Show("Error checking disk: $($_.Exception.Message)", "Error")
    }
})

# Add the button to the Diagnostic panel collection
$FlowDiagnostic.Controls.Add($Btn_Disk)



# ==============================================================================
#                         BUTTON --> RESET PRINT SPOOLER
# The Reset Print Spooler tool automates the standard IT fix for stuck print queues.
# Logic Overview:
#                *Service Management: The script uses Stop-Service and Start-Service to cycle the background spooling engine.
#                *Queue Purging: It targets the protected System32\spool\PRINTERS directory to remove corrupted .shd and .spl files that cause jobs to hang.
#                *Administrator Requirements: Since manipulating system services requires elevated rights, the script includes specific error handling to alert the user if they are not running in Admin mode.
# ==============================================================================

# BUTTON: Reset Print Spooler
# Initialize a new Button control for the Print Spooler troubleshooting tool.
$Btn_Spooler = New-Object System.Windows.Forms.Button
# Set the visible label text for the button.
$Btn_Spooler.Text = "Reset Print Spooler"
# Apply the toolkit's standard visual styling and high-DPI scaling.
Format-Button $Btn_Spooler

# Define the script logic to execute when the button is clicked.
$Btn_Spooler.Add_Click({
    # Prompt the user for confirmation before stopping services and deleting print jobs.
    $Confirm = [System.Windows.Forms.MessageBox]::Show("This will stop the Print Spooler, clear pending jobs, and restart the service. Continue?", "Confirm Reset", "YesNo", "Warning")
    
    if ($Confirm -eq "Yes") {
        # 1. Visual Feedback: Show the wait cursor and disable the button to prevent multiple clicks.
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $Btn_Spooler.Enabled = $false
        # Force the UI to refresh immediately to reflect button changes.
        [System.Windows.Forms.Application]::DoEvents()

        try {
            # 2. Stop Service: Forcefully stop the 'Spooler' service to release file locks on print jobs.
            Stop-Service -Name Spooler -Force -ErrorAction Stop
            
            # 3. Clear Cache: Recursively delete all pending print job files from the system spool directory.
            $SpoolPath = "$env:SystemRoot\System32\spool\PRINTERS\*"
            Remove-Item -Path $SpoolPath -Recurse -Force -ErrorAction SilentlyContinue
            
            # 4. Restart Service: Re-initialize the 'Spooler' service to resume normal printing operations.
            Start-Service -Name Spooler -ErrorAction Stop
            
            # Notify the user of a successful reset operation.
            [System.Windows.Forms.MessageBox]::Show("Print Spooler reset successfully!", "Success")
        }
        catch {
            # Catch and display errors, often related to missing Administrative permissions.
            [System.Windows.Forms.MessageBox]::Show("❌ Error: $($_.Exception.Message)`nEnsure you are running as Administrator.", "Failed")
        }
        finally {
            # 5. Reset UI: Restore the standard cursor and re-enable the button.
            $Form.Cursor = [System.Windows.Forms.Cursors]::Default
            $Btn_Spooler.Enabled = $true
        }
    }
})

# Add the button into the Diagnostic Tab's layout container (Updated for consistency).
$FlowDiagnostic.Controls.Add($Btn_Spooler)



# ==============================================================================
#                         BUTTON --> CHECK SYSTEM UP TIME
# The Check System Uptime tool provides a critical diagnostic metric for identifying "ghost" issues caused by lack of restarts.
# ==============================================================================

# [BUTTON: Check System Uptime
# Initialize the Button object for monitoring system uptime
$Btn_Uptime = New-Object System.Windows.Forms.Button
# Set the visible label for the button
$Btn_Uptime.Text = "Check System Uptime"
# Apply standard toolkit styling and scaling via your global helper function
Format-Button $Btn_Uptime

# Define the script logic to execute when the button is clicked
$Btn_Uptime.Add_Click({
    try {
        # 1. Get OS Information: Retrieve system boot data using the CIM instance
        # Get-CimInstance is the recommended method as it is fast and reliable
        $OS = Get-CimInstance -ClassName Win32_OperatingSystem
        # Identify the exact timestamp the system last completed its boot sequence
        $LastBoot = $OS.LastBootUpTime
        $CurrentTime = Get-Date
        
        # 2. Calculate Difference: Subtract boot time from current time to find total uptime
        $Uptime = $CurrentTime - $LastBoot
        
        # 3. Format the Message: Construct a readable breakdown of days, hours, and minutes
        $Msg = "Your system has been running for:`n`n" +
               "📅 Days:    $($Uptime.Days)`n" +
               "⌚ Hours:   $($Uptime.Hours)`n" +
               "⏱️ Minutes: $($Uptime.Minutes)`n`n" +
               "Last Boot Time: $($LastBoot.ToString('yyyy-MM-dd HH:mm:ss'))"

        # 4. Show Popup: Display the final uptime report to the IT analyst
        [System.Windows.Forms.MessageBox]::Show($Msg, "System Uptime Status", "OK", "Information")
        
    } catch {
        # Display an error message if the CIM query or calculation fails
        [System.Windows.Forms.MessageBox]::Show("Error retrieving uptime: $($_.Exception.Message)", "Error")
    }
})

# Finalize placement by adding the button to the Diagnostic panel
$FlowDiagnostic.Controls.Add($Btn_Uptime)



# ==============================================================================
#                         BUTTON --> LIST INSTALLED SOFTWARE
# ==============================================================================

# BUTTON: List Installed Software
# Initialize the button for listing and uninstalling software
$Btn_Software = New-Object System.Windows.Forms.Button
# Set the visible label for the button
$Btn_Software.Text = "Installed Software"
# Apply standardized global formatting and scaling
Format-Button $Btn_Software

# Define the script logic for the button click event
$Btn_Software.Add_Click({
    # Disable the button and update text to provide visual feedback during loading
    $Btn_Software.Enabled = $false
    $Btn_Software.Text = "Loading List..."
    # Force the UI to refresh to show the new button state immediately
    [System.Windows.Forms.Application]::DoEvents()

    # 1. Gather all installed software including the 'UninstallString' from the registry
    # This queries 64-bit, 32-bit (WOW6432Node), and current user registry paths
    $SoftwareList = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*, 
                                     HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*,
                                     HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | 
                    Where-Object { $_.DisplayName -ne $null } |
                    Select-Object DisplayName, DisplayVersion, Publisher, UninstallString | 
                    Sort-Object DisplayName

    # 2. Open GridView and wait for the user to select one row for uninstallation
    # The -PassThru parameter allows the selection to be captured as an object
    $SelectedApp = $SoftwareList | Out-GridView -Title "Select an Application to Uninstall" -PassThru

    if ($SelectedApp) {
        # 3. Confirmation Dialog to prevent accidental uninstalls
        $Confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to uninstall: `n`n$($SelectedApp.DisplayName)?", "Confirm Uninstall", "YesNo", "Warning")
        
        if ($Confirm -eq "Yes") {
            try {
                # Update button text to show uninstallation is in progress
                $Btn_Software.Text = "Uninstalling..."
                [System.Windows.Forms.Application]::DoEvents()

                # 4. Execute the Uninstall String retrieved from the registry
                if ($SelectedApp.UninstallString) {
                    # Handle MsiExec strings specifically for more reliable execution
                    if ($SelectedApp.UninstallString -match "MsiExec.exe") {
                        # Strip the executable name and add a no-restart flag
                        $Args = $SelectedApp.UninstallString -replace "MsiExec.exe", ""
                        $Args += " /norestart" 
                        Start-Process "MsiExec.exe" -ArgumentList $Args -Wait
                    } else {
                        # For standard EXE uninstallers, run the command string via cmd.exe
                        cmd.exe /c $SelectedApp.UninstallString
                    }
                    [System.Windows.Forms.MessageBox]::Show("Uninstaller launched for $($SelectedApp.DisplayName). Please check for any background windows.", "Task Started")
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Could not find a valid uninstall command for this app.", "Error")
                }
            } catch {
                # Display any errors encountered during the process
                [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed")
            }
        }
    }

    # Restore the button state and text after the process completes
    $Btn_Software.Enabled = $true
    $Btn_Software.Text = "Installed Software"
})

# Add the button to the Diagnostic panel collection
$FlowDiagnostic.Controls.Add($Btn_Software)



# ==============================================================================
#                         BUTTON --> WINDOWS CREDENTIAL MANAGER
# ==============================================================================

# BUTTON: WINDOWS CREDENTIAL MANAGER
$Btn_CredManager = New-Object System.Windows.Forms.Button
$Btn_CredManager.Text = "Credential Manager"
Format-Button $Btn_CredManager

$Btn_CredManager.Add_Click({
    # 1. Create Sub-Window
    $CredForm = New-Object System.Windows.Forms.Form
    $CredForm.Text = "Credential Manager Tools"
    $CredForm.Size = New-Object System.Drawing.Size(450, 400)
    $CredForm.StartPosition = "CenterParent"
    $CredForm.BackColor = "#f0f0f0"

    $CredFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $CredFlow.Dock = "Fill"; $CredFlow.Padding = New-Object System.Windows.Forms.Padding(10)
    $CredFlow.FlowDirection = "TopDown"; $CredFlow.WrapContents = $false
    $CredForm.Controls.Add($CredFlow)


    # 2. Option 1: Clear System Cached Passwords
    # Set Modern Style at TOP of script
[System.Windows.Forms.Application]::EnableVisualStyles()

$Btn_ClearCreds = New-Object System.Windows.Forms.Button
$Btn_ClearCreds.Text = "1. Clear System Cached Passwords"
$Btn_ClearCreds.Size = New-Object System.Drawing.Size(400, 50) # Fixed Height
$Btn_ClearCreds.BackColor = "#0078D7"; $Btn_ClearCreds.ForeColor = "White"
$Btn_ClearCreds.FlatStyle = "Flat" # Modern Look
    $Btn_ClearCreds.Add_Click({
        $Confirm = [System.Windows.Forms.MessageBox]::Show("Clear all generic Windows credentials?", "Confirm", "YesNo", "Warning")
        if ($Confirm -eq "Yes") {
            # Logic to enumerate and delete generic credentials using cmdkey
            cmdkey /list | ForEach-Object { if($_ -match "target=(.*)") { cmdkey /delete:$($matches[1]) } }
            [System.Windows.Forms.MessageBox]::Show("Generic credentials cleared.", "Success")
        }
    })
    $CredFlow.Controls.Add($Btn_ClearCreds)
    $CredForm.ShowDialog()
})
$FlowDiagnostic.Controls.Add($Btn_CredManager)



# ==============================================================================
#                         BUTTON --> SYSTEM INFO
# ==============================================================================

# BUTTON: SYSTEM INFO
# Create a new button object for the System Information tool 
$Btn_SysInfo = New-Object System.Windows.Forms.Button
# Set the button label text
$Btn_SysInfo.Text = "System Information"
# Apply global styling and scaling via your custom function
Format-Button $Btn_SysInfo

# Define the action to take when the button is clicked 
$Btn_SysInfo.Add_Click({
    # 1. Admin Notice: Inform the user that some data requires elevated rights
    $Msg = "To retrieve complete information (TPM, Bitlocker, etc.), please ensure this script is running with Administrator privileges."
    [System.Windows.Forms.MessageBox]::Show($Msg, "Privilege Notice", "OK", "Information")

    # Change the mouse pointer to a loading spinner during data collection
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # 2. Precise OS Detection: Pull raw version data from the registry
        $WinVer = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $OS = Get-ComputerInfo
        $OSName = $WinVer.ProductName
        $OSDisplay = $WinVer.DisplayVersion

        # Logic fix: Ensure Windows 11 is not incorrectly reported as Windows 10
        if ($OSName -match "Windows 10" -and [environment]::OSVersion.Version.Build -ge 22000) {
            $OSName = "Windows 11 Enterprise"
        }

        # 3. Gather Hardware Data: Query system and security components
        $CS = Get-CimInstance -ClassName Win32_ComputerSystem
        $BIOS = Get-CimInstance -ClassName Win32_BIOS
        $TPM = Get-CimInstance -Namespace Root\CIMV2\Security\MicrosoftTpm -ClassName Win32_Tpm -ErrorAction SilentlyContinue
        $Bitlocker = Get-BitLockerVolume -MountPoint C: -ErrorAction SilentlyContinue
        $Disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"

        # 4. Build Data Object: Map collected data to readable properties
        $SysData = [PSCustomObject]@{
            'Hostname'         = $env:COMPUTERNAME
            'Make/Model'       = "$($CS.Manufacturer) / $($CS.Model)"
            'Serial Number'    = $BIOS.SerialNumber
            'Domain Joined'    = if($CS.PartOfDomain){"Yes"}else{"No"}
            'Azure AD Joined'  = if((dsregcmd /status | Select-String "AzureAdJoined").ToString() -match "YES"){"Yes"}else{"No"}
            'OS / Version'     = "$OSName / $OSDisplay"
            'OS Install Date'     = $OS.WindowsInstallDateFromRegistry
            'Bitlocker (C:)'   = if($Bitlocker.VolumeStatus -eq 'FullyEncrypted'){"Yes"}else{"No"}
            'TPM Version'      = if($TPM){$TPM.SpecVersion}else{"Not Found"}
            'BIOS Ver / Date'  = "$($BIOS.SMBIOSBIOSVersion) / $($BIOS.ReleaseDate)"
            'Processor'        = (Get-CimInstance Win32_Processor).Name
            'RAM (GB)'         = [math]::Round($CS.TotalPhysicalMemory / 1GB, 2)
            'Storage Info'     = ($Disks | ForEach-Object { "$($_.DeviceID) Free: $([math]::Round($_.FreeSpace / 1GB, 2))GB" }) -join " | "
        }

        # 5. Transform and Show: Display results in a searchable GridView window [[5](https://bonguides.com/winforms-creating-guis-in-windows-powershell-with-winforms/)]
        $SysData.psobject.Properties | Select-Object @{Name='System Setting';Expression={$_.Name}}, @{Name='Details';Expression={$_.Value}} | Out-GridView -Title "Full System Audit Report"
    } finally {
        # Restore the standard mouse pointer when complete
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the completed button to the Diagnostic panel (Updated from FlowSec)
$FlowDiagnostic.Controls.Add($Btn_SysInfo)



# ==============================================================================
#                         BUTTON --> BITLOCKER KEY STATUS
# ==============================================================================

# [BUTTON: Check BitLocker Key Status
# Initialize a new Button control for BitLocker security auditing.
$Btn_BitLocker = New-Object System.Windows.Forms.Button
# Set the visible label for the button.
$Btn_BitLocker.Text = "BitLocker Status"
# Apply the toolkit's standard visual styling and scaling via your global helper function.
Format-Button $Btn_BitLocker

# Define the script logic to execute when the button is clicked.
$Btn_BitLocker.Add_Click({
    # 1. Admin Privilege Check: BitLocker management requires elevated rights for security.
    $Principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        [System.Windows.Forms.MessageBox]::Show("⚠️ ACCESS DENIED`n`nBitLocker keys are sensitive security data.`nYou must run this toolkit as Administrator to view them.", "Admin Required", "OK", "Error")
        return
    }

    # 2. Run the Command: Change the cursor to wait mode during data retrieval.
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Retrieve detailed information about all BitLocker-capable volumes.
        $Vols = Get-BitLockerVolume -ErrorAction Stop
        
        # Format the BitLocker data into a readable custom object collection.
        $Results = foreach ($Vol in $Vols) {
            # Extract and join the types of active key protectors (e.g., Tpm, RecoveryPassword).
            $KeyTypes = ($Vol.KeyProtector).KeyProtectorType -join ", " 
            
            [PSCustomObject]@{
                'Drive'          = $Vol.MountPoint
                'Status'         = $Vol.VolumeStatus
                'Encryption'     = "$($Vol.EncryptionPercentage)%"
                'Key Protectors' = $KeyTypes
                'Protection'     = $Vol.ProtectionStatus
            }
        }
        
        # Display the formatted security report in an interactive GridView window.
        $Results | Out-GridView -Title "BitLocker Security Report"
    } catch {
        # Catch and display any engine errors during the retrieval process.
        [System.Windows.Forms.MessageBox]::Show("❌ Error retrieving BitLocker info:`n$($_.Exception.Message)", "Failed")
    } finally {
        # Restore the standard cursor once the operation is complete.
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the button into the Diagnostic Tab's layout container.
$FlowDiagnostic.Controls.Add($Btn_BitLocker)



# ==============================================================================
#                         BUTTON --> Monitor CPU/Memory Peaks
# The Monitor CPU/Memory Peaks tool is designed for real-time performance auditing.
# It captures system behavior over a 60-second window, which is ideal for identifying intermittent "freezes" or resource spikes.
# ==============================================================================

# [BUTTON: Monitor CPU/Memory Peaks
# Initialize a new Button control for performance monitoring.
$Btn_Monitor = New-Object System.Windows.Forms.Button
# Set the visible label for the button.
$Btn_Monitor.Text = "Monitor CPU/Memory Peaks (60s)"
# Apply the toolkit's standard visual styling and scaling via your global helper function.
Format-Button $Btn_Monitor

# Define the script logic to execute when the button is clicked.
$Btn_Monitor.Add_Click({
    # Define the log path on the current user's desktop.
    $LogPath = "$env:USERPROFILE\Desktop\PeakMonitor_Log.csv"
    # Disable the button to prevent multiple simultaneous monitoring sessions.
    $Btn_Monitor.Enabled = $false
    # Initialize the CSV file with headers.
    "Timestamp, CPU_Usage(%), Memory_Usage(%)" | Out-File $LogPath -Encoding utf8
    
    # Start a 60-second monitoring loop.
    for ($i = 1; $i -le 60; $i++) {
        # Update the PowerShell progress bar in the console.
        Write-Progress -Activity "Monitoring Performance" -Status "Collecting Data ($i/60s)" -PercentComplete (($i/60)*100)
        
        # 1. Retrieve CPU Usage: Gets the 'CookedValue' for total processor time.
        $CPU = [math]::Round((Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue, 2)
        
        # 2. Retrieve Memory Usage: Gets the percentage of committed bytes currently in use.
        $Mem = [math]::Round((Get-Counter '\Memory\% Committed Bytes In Use').CounterSamples.CookedValue, 2)
        
        # Log the current time and metrics to the CSV file.
        "$((Get-Date -Format T)), $CPU, $Mem" | Out-File $LogPath -Append
        
        # Prevent the WinForms UI from freezing during the loop.
        [System.Windows.Forms.Application]::DoEvents() 
        # Wait 1 second before the next sample.
        Start-Sleep -Seconds 1
    }
    
    # Close the progress bar once finished.
    Write-Progress -Activity "Monitoring Performance" -Completed
    
    # Prompt the user to open the generated log file.
    $Ans = [System.Windows.Forms.MessageBox]::Show("Finished! Log saved to: $LogPath`n`nOpen log now?", "Task Complete", "YesNo")
    if ($Ans -eq "Yes") { Start-Process "notepad.exe" $LogPath }
    
    # Re-enable the button for future use.
    $Btn_Monitor.Enabled = $true
})

# Add the button to the Diagnostic Tab's layout container.
$FlowDiagnostic.Controls.Add($Btn_Monitor)



# ==============================================================================
#                         BUTTON --> WINDOWS ACTIVATION STATUS
# ==============================================================================

# BUTTON: WINDOWS ACTIVATION STATUS
# Create a new button object to check license validity
$Btn_Activation = New-Object System.Windows.Forms.Button
# Set the button label text
$Btn_Activation.Text = "Windows Activation Status"
# Apply standard toolkit formatting
Format-Button $Btn_Activation

# Define the action to take when clicked
$Btn_Activation.Add_Click({
    # Change the mouse pointer to a loading spinner
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Querying the license status using CIM: Filter for products with partial keys to find the OS license 
        $License = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL" | 
                   Where-Object { $_.Name -match "Windows" }
        
        # LicenseStatus 1 indicates the system is fully Licensed/Activated 
        if ($License.LicenseStatus -eq 1) {
            $StatusMsg = "✅ Windows is ACTIVATED"
        } else {
            # Provide a warning if the system is in a grace period or unactivated state 
            $StatusMsg = "❌ Windows is NOT ACTIVATED"
        }
        
        # Display the final activation status in a popup
        [System.Windows.Forms.MessageBox]::Show($StatusMsg, "Activation Status")
    } catch {
        # Inform the user if the WMI/CIM query fails (common on non-admin sessions) 
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Check Failed")
    } finally {
        # Reset the mouse cursor to default
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
# Add the button to the layout panel
$FlowDiagnostic.Controls.Add($Btn_Activation)



# ==============================================================================
#                         BUTTON --> CHECK PENDING REBOOT
#The Check Pending Reboot tool is an essential diagnostic for IT analysts to determine if a system requires a restart before further troubleshooting or software installations.
# ==============================================================================

# BUTTON: Check Pending Reboot
# Initialize the Button object to check for pending system restarts
$Btn_Reboot = New-Object System.Windows.Forms.Button
# Set the visible label for the button
$Btn_Reboot.Text = "Pending Reboot?"
# Apply standard toolkit styling and DPI scaling via your global helper function
Format-Button $Btn_Reboot

# Define the script logic to execute when the button is clicked
$Btn_Reboot.Add_Click({
    # Define registry paths that Windows uses to flag a pending reboot
    $Paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending", 
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    )
    
    # Set the initial status to healthy/no reboot
    $Status = "No Reboot Required"
    
    # Iterate through the registry paths to see if any of them exist
    foreach($P in $Paths){ 
        # If any path is found, it confirms a reboot is waiting
        if(Test-Path $P){ 
            $Status = "⚠️ REBOOT REQUIRED" 
        } 
    }
    
    # Display the final status result to the IT analyst in a popup box
    [System.Windows.Forms.MessageBox]::Show($Status, "Reboot Check")
})

# Add the button to the Diagnostic Tab's layout container
$FlowDiagnostic.Controls.Add($Btn_Reboot)



# ==============================================================================
#                         BUTTON --> Local Admin Account Status
# The Local Admin Account Status tool provides a security audit by identifying which local users have elevated privileges on the machine.
# ==============================================================================

# [BUTTON 14] Local Admin Account Status
# Initialize the Button object for auditing local administrator accounts
$Btn_Admins = New-Object System.Windows.Forms.Button
# Set the visible label for the button
$Btn_Admins.Text = "Local Admin Status"
# Apply the toolkit's standard visual styling and high-DPI scaling
Format-Button $Btn_Admins

# Define the script logic to execute when the button is clicked
$Btn_Admins.Add_Click({
    # Change the mouse pointer to a waiting spinner during data retrieval
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Fetch members of the 'Administrators' group, ignoring errors from orphaned SIDs [[1](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.localaccounts/get-localgroupmember?view=powershell-5.1)]
        $Admins = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue 
        
        # Iterate through the group members to filter and extract user-specific details
        $Results = foreach($A in $Admins){
            # Only process members that are local user accounts; skip groups or domain accounts [[6](https://techcommunity.microsoft.com/blog/itopstalkblog/how-to-manage-local-users-and-groups-using-powershell/733544)]
            if ($A.PrincipalSource -eq "Local" -and $A.ObjectClass -eq "User") {
                # Extract the username by removing the machine prefix (e.g., 'PC-NAME\Admin')
                $UserName = $A.Name.Split('\')[-1]
                # Retrieve full local user details to check status and lockout info
                $User = Get-LocalUser -Name $UserName -ErrorAction SilentlyContinue
                
                if ($User) {
                    # Create a custom object with security audit properties
                    [PSCustomObject]@{ 
                        Name            = $User.Name
                        Active          = $User.Enabled
                        PasswordCreated = $User.PasswordLastSet 
                        Locked          = $User.AccountLockout
                    }
                }
            }
        }
        
        # If local admins are found, display them in an interactive grid for review [[4](https://powershellfaqs.com/list-local-administrators-using-powershell/)]
        if ($Results) {
            $Results | Out-GridView -Title "Local Admin Security Audit"
        } else {
            # Notify if no accounts matching the 'local user' criteria exist in the group
            [System.Windows.Forms.MessageBox]::Show("No local admin users found.", "Info")
        }
    } finally {
        # Restore the standard mouse pointer once the audit is complete
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the button to the Diagnostic Tab's layout container
$FlowDiagnostic.Controls.Add($Btn_Admins)



# ==============================================================================
#                         BUTTON --> DEFENDER FULL SCAN
# ==============================================================================

# BUTTON: DEFENDER FULL SCAN
# Create the button object for initiating a comprehensive system scan
$Btn_DefScan = New-Object System.Windows.Forms.Button
# Set the visible text label on the button
$Btn_DefScan.Text = "Defender Full Scan"
# Apply standardized global formatting and DPI scaling
Format-Button $Btn_DefScan

# Define the script logic to execute when the button is clicked
$Btn_DefScan.Add_Click({
    # 1. Define the warning message for the user regarding scan duration
    $MsgText = "A Full Scan checks all files and running programs on your hard drive. This process can take over 1 hour depending on your system size.`n`nDo you want to proceed?"
    
    # 2. Show the Yes/No confirmation box to prevent accidental long-running tasks
    $Response = [System.Windows.Forms.MessageBox]::Show($MsgText, "Defender Full Scan Warning", "YesNo", "Warning")

    # 3. Only execute the scan command if the user explicitly clicks 'Yes'
    if ($Response -eq "Yes") {
        # Launch a separate elevated PowerShell window to run the long-duration scan
        # This keeps the main toolkit responsive during the process
        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "Write-Host 'Starting Full Scan...'; Start-MpScan -ScanType FullScan"
    }
})

# Add the button to the Diagnostic panel layout container
$FlowDiagnostic.Controls.Add($Btn_DefScan)



# ==============================================================================
#                         BUTTON --> OPEN DEFENDER LOGS
# ==============================================================================

# BUTTON: OPEN DEFENDER LOGS
# Create the button object for opening the Defender support logs directory
$Btn_DefLogs = New-Object System.Windows.Forms.Button
# Set the visible text label for the button
$Btn_DefLogs.Text = "Defender Logs Location"
# Apply standardized global formatting
Format-Button $Btn_DefLogs

# Define the logic to open the specific system folder when clicked
$Btn_DefLogs.Add_Click({
    # Use Windows Explorer to open the hidden ProgramData path where Defender logs are stored
    Start-Process explorer.exe "C:\ProgramData\Microsoft\Windows Defender\Support"
})

# Add the logs button to the Diagnostic panel layout container
$FlowDiagnostic.Controls.Add($Btn_DefLogs)









# 1. Add the layout panel to the Diagnostic Tab
$TabDiagnostic.Controls.Add($FlowDiagnostic) 

# 2. Add the Diagnostic Tab to the Main Tab Control
$TabControl.TabPages.Add($TabDiagnostic) 

# *****************************************************************************************************************************************************************************************************


# *****************************************************************************************************************************************************************************************************
# ==============================================================================
#                              HARDWARE HEALTH TAB
# The Hardware Health Tab provides a specialized monitoring suite for IT professionals. It leverages the following key features: 
#                *Reliability Tracking: Uses Get-StorageReliabilityCounter to estimate the remaining life of SSDs and powercfg to analyze battery degradation
#                *Thermal Diagnostics: Implements high-frequency sampling to track CPU temperatures, helping identify cooling issues before hardware failure occurs
#                *Responsive Design: Automatically scales button sizes and fonts using the $Global:DpiScale variable, ensuring visibility on both standard and 4K displays
#                *Z-Order Management: $TabControl.BringToFront() ensures the tab system stays visible on top of any background panels or other controls
# This tab centralizes vital physical metrics, reducing the need for multiple third-party diagnostic tools
# ==============================================================================


# --- TAB 4: Hardware Health ---
# Create the container for hardware monitoring tools
$TabHealth = New-Object System.Windows.Forms.TabPage
$TabHealth.Text = "Hardware Health"
$TabHealth.BackColor = "#2d2d30"

# Setup a FlowLayoutPanel to organize hardware diagnostic buttons
$FlowHealth = New-Object System.Windows.Forms.FlowLayoutPanel
$FlowHealth.Dock = "Fill" 
$FlowHealth.AutoScroll = $true

# Apply flicker reduction to the hardware panel
Enable-DoubleBuffering $FlowHealth

# Background Image integration for visual consistency
$ImgPath = "$PSScriptRoot\Images\WhiteBG.png"
if (Test-Path $ImgPath) {
    $TabHealth.BackgroundImage = [System.Drawing.Image]::FromFile($ImgPath)
    $TabHealth.BackgroundImageLayout = "Stretch"
}

# Set transparency to show the underlying background texture
$FlowHealth.BackColor = [System.Drawing.Color]::Transparent



# ==============================================================================
#                         BUTTON --> SSD Health Status
# ==============================================================================
# SSD HEALTH: Estimates drive life using storage reliability counters
$Btn_SSDHealth = New-Object System.Windows.Forms.Button
$Btn_SSDHealth.Text = "SSD Health Status"
Format-Button $Btn_SSDHealth

$Btn_SSDHealth.Add_Click({
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Query physical SSDs and calculate remaining life based on wear levels
    $SSDReport = Get-PhysicalDisk | Where-Object MediaType -eq 'SSD' | ForEach-Object {
        $reliability = Get-StorageReliabilityCounter -PhysicalDisk $_
        $wear = if ($reliability.Wear -ne $null) { "$($reliability.Wear)%" } else { "N/A" }

        [PSCustomObject]@{
            Disk      = $_.FriendlyName
            Health    = if ($reliability.Wear -ne $null) { "$(100 - $reliability.Wear)%" } else { "Unknown" }
            Temp      = "$($reliability.Temperature) °C"
            WearLevel = $wear
            Status    = $_.OperationalStatus
        }
    }

    if ($SSDReport) {
        $SSDReport | Out-GridView -Title "SSD Hardware Health Estimation"
    } else {
        [System.Windows.Forms.MessageBox]::Show("No SSD devices found.", "Hardware Info")
    }
    
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})
$FlowHealth.Controls.Add($Btn_SSDHealth)



# ==============================================================================
#                         BUTTON --> Battery Health Status
# ==============================================================================
# BATTERY HEALTH: Analyzes wear by comparing design vs. current capacity
$Btn_Battery = New-Object System.Windows.Forms.Button
$Btn_Battery.Text = "Battery Health Status"
Format-Button $Btn_Battery

$Btn_Battery.Add_Click({
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Generate the official Windows battery report XML [[5](https://windowsforum.com/threads/how-to-check-windows-laptop-battery-health-with-powercfg-battery-report.387953/)]
    $ReportPath = "$env:TEMP\battery_report.xml"
    powercfg /batteryreport /output $ReportPath /xml | Out-Null
    
    if (Test-Path $ReportPath) {
        [xml]$xml = Get-Content $ReportPath
        $Design = $xml.BatteryReport.Batteries.Battery.DesignCapacity
        $Full = $xml.BatteryReport.Batteries.Battery.FullChargeCapacity
        # Calculate percentage of original health remaining [[5](https://windowsforum.com/threads/how-to-check-windows-laptop-battery-health-with-powercfg-battery-report.387953/)]
        $Health = [math]::Round(($Full / $Design) * 100, 2)
        
        $Msg = "Design Capacity: $Design mWh`nFull Charge: $Full mWh`nHealth: $Health %"
        [System.Windows.Forms.MessageBox]::Show($Msg, "Battery Health Status")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Battery data not available (Desktop PC?).", "Info")
    }
    
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})
$FlowHealth.Controls.Add($Btn_Battery)



# ==============================================================================
#                         BUTTON --> CPU Temperature
# ==============================================================================
# CPU TEMPERATURE: Retrieves real-time thermal data via WMI [[3](https://www.itsupportguides.com/knowledge-base/windows-10/powershell-script-display-battery-information-on-windows-11/)]
$Btn_CPUTemp = New-Object System.Windows.Forms.Button
$Btn_CPUTemp.Text = "CPU Temperature"
Format-Button $Btn_CPUTemp

$Btn_CPUTemp.Add_Click({
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Query specialized thermal zone counters for live temperature data
        $tempData = Get-WmiObject -Query "SELECT * FROM Win32_PerfFormattedData_Counters_ThermalZoneInformation" -Namespace "root/CIMV2"
        $maxTemp = $null

        foreach ($zone in $tempData) {
            # Convert Kelvin units to Celsius
            $celsius = $zone.Temperature - 273.15
            if ($null -eq $maxTemp -or $celsius -gt $maxTemp) { $maxTemp = $celsius }
        }

        if ($maxTemp) {
            $formattedTemp = [math]::Round($maxTemp, 2)
            [System.Windows.Forms.MessageBox]::Show("Current CPU Temperature: $formattedTemp °C", "Thermal Diagnostic")
        } else {
            throw "No temperature data found."
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Hardware sensors not reporting via WMI. Try running as Admin.", "Error")
    }
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})
$FlowHealth.Controls.Add($Btn_CPUTemp)



# ==============================================================================
#                         BUTTON --> Thermal Log (60s)
# ==============================================================================
# CPU THERMAL LOG: Tracks heat fluctuations over a 60-second window
$Btn_ThermalLog = New-Object System.Windows.Forms.Button
$Btn_ThermalLog.Text = "Thermal Log (60s)"
Format-Button $Btn_ThermalLog

$Btn_ThermalLog.Add_Click({
    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $SaveDialog.FileName = "CPU_Thermal_Log_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    
    if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $LogPath = $SaveDialog.FileName
        $Results = New-Object System.Collections.Generic.List[PSObject]
        
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.MessageBox]::Show("Logging started. Capturing data every second for 60 seconds.", "Thermal Logging")

        for ($i = 1; $i -le 60; $i++) {
            $tempData = Get-WmiObject -Query "SELECT * FROM Win32_PerfFormattedData_Counters_ThermalZoneInformation" -Namespace "root/CIMV2"
            $maxKelvin = ($tempData | Measure-Object -Property Temperature -Maximum).Maximum
            $celsius = [math]::Round($maxKelvin - 273.15, 2)
            
            $Results.Add([PSCustomObject]@{
                Second      = $i
                Timestamp   = Get-Date -Format "HH:mm:ss"
                TempCelsius = $celsius
            })
            Start-Sleep -Seconds 1 # Pause execution for 1 second between captures
        }

        # Export findings to CSV and offer to open the file
        $Results | Export-Csv -Path $LogPath -NoTypeInformation
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        
        $Msg = "Log saved successfully at: $LogPath`n`nDo you want to open the log file now?"
        if ([System.Windows.Forms.MessageBox]::Show($Msg, "Success", "YesNo", "Information") -eq "Yes") {
            Invoke-Item $LogPath
        }
    }
})
$FlowHealth.Controls.Add($Btn_ThermalLog)

# Finalize the Hardware tab by nesting panels and adding to the main Hub
$TabHealth.Controls.Add($FlowHealth)
$TabControl.TabPages.Add($TabHealth)



# *****************************************************************************************************************************************************************************************************

# *****************************************************************************************************************************************************************************************************
# ==============================================================================
#                              NETWORK TAB
# 
# ==============================================================================

# --- TAB 3: Network ---
# Initialize a new TabPage object for the Network section
$TabNet = New-Object System.Windows.Forms.TabPage 

# Set the title text displayed on the tab
$TabNet.Text = "Network" 

# Define the background color using a hex code string
$TabNet.BackColor = "#2d2d30" 

# FlowLayout Setup
# Create a FlowLayoutPanel to arrange child controls automatically
$FlowNet = New-Object System.Windows.Forms.FlowLayoutPanel 

# Set the panel to fill the entire area of the parent tab
$FlowNet.Dock = "Fill" 

# Enable scrollbars if the content exceeds the panel size
$FlowNet.AutoScroll = $true 

# Double Buffering
# Call custom function to reduce flickering during UI redraws
Enable-DoubleBuffering $FlowNet 

# Background Image
# Define the local path to the background image file
$ImgPath = "$PSScriptRoot\Images\WhiteBG.png" 

# Check if the image file exists before attempting to load it
if (Test-Path $ImgPath) {
    # Load the image and assign it as the tab's background
    $TabNet.BackgroundImage = [System.Drawing.Image]::FromFile($ImgPath)
    
    # Set the image to stretch and fit the tab dimensions
    $TabNet.BackgroundImageLayout = "Stretch" 
}

# Transparency
# Set the layout panel background to transparent to show the tab's image
$FlowNet.BackColor = [System.Drawing.Color]::Transparent



# ==============================================================================
#                         BUTTON --> Test Remote Connectivity
# ==============================================================================

# BUTTON: Test Remote Connectivity
# Create a new Button object for the remote connectivity tool
$Btn_RemoteTest = New-Object System.Windows.Forms.Button 

# Set the text displayed on the button surface
$Btn_RemoteTest.Text = "Test Remote Connectivity" 

# Apply standardized formatting using your custom function
Format-Button $Btn_RemoteTest 

# Define the action to take when the button is clicked
$Btn_RemoteTest.Add_Click({ 
    # Load the Visual Basic assembly to enable InputBox functionality
    Add-Type -AssemblyName Microsoft.VisualBasic 
    
    # Prompt the user to enter a hostname or IP address with a default value
    $Target = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Hostname or IP:", "Remote Test", "8.8.8.8")

    # Proceed only if the user provided a non-empty target address
    if (-not [string]::IsNullOrWhiteSpace($Target)) {
        
        # 1. Create a Temporary Script File
        # Define the file path for the temporary diagnostic script
        $TempScript = "$env:TEMP\RemoteTest.ps1" 
        
        # Build the script content using a Here-String for easier formatting
        $ScriptContent = @"
            # Update the console window title to show current target
            `$Host.UI.RawUI.WindowTitle = 'Testing: $Target'
            # Output status message to the console
            Write-Host 'Testing Connectivity for: $Target' -ForegroundColor Yellow
            Write-Host '-------------------------------------'
            
            # Perform a standard ICMP ping test
            Write-Host '1. Pinging...' -ForegroundColor Cyan
            Test-Connection -ComputerName '$Target' -Count 4

            # Perform TCP port scans on common service ports
            Write-Host '`n2. Port Scan...' -ForegroundColor Cyan
            `$Ports = 80, 443, 3389
            foreach (`$P in `$Ports) {
                # Test connectivity to specific ports silently
                `$R = Test-NetConnection -ComputerName '$Target' -Port `$P -WarningAction SilentlyContinue
                # Report port status using color-coded results
                if (`$R.TcpTestSucceeded) { Write-Host "    [OPEN] Port `$P" -ForegroundColor Green }
                else { Write-Host "    [CLOSED] Port `$P" -ForegroundColor Red }
            }

            # Pause execution so the user can view the results
            Write-Host '`nDone. Press ENTER to close.' -ForegroundColor White
            Read-Host
"@
        # Save the constructed script to the temporary file location
        $ScriptContent | Out-File $TempScript -Encoding UTF8 -Force

        # 2. Launch the Script in a New Window
        # Start a new PowerShell process to run the temporary diagnostic file
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$TempScript`""
    }
})

# Add the initialized button to the Network Tab's FlowLayout panel
$FlowNet.Controls.Add($Btn_RemoteTest)



# ==============================================================================
#                         BUTTON --> Fix Network Connectivity
# ==============================================================================

# BUTTON: Fix Network Connectivity
# Create a new button object for the network repair utility
$Btn_NetFix = New-Object System.Windows.Forms.Button
# Set the text label displayed on the button
$Btn_NetFix.Text = "Fix Network Connectivity"
# Apply the standard toolkit formatting to the button
Format-Button $Btn_NetFix

# Define the event that occurs when the button is clicked
$Btn_NetFix.Add_Click({
    # Define a warning message explaining the network reset process
    $Msg = "This will reset TCP/IP, Flush DNS, and restart network adapters.`n`nYou will lose internet connection momentarily.`n`nContinue?"
    # Display a confirmation dialog to the user before proceeding
    $Confirm = [System.Windows.Forms.MessageBox]::Show($Msg, "Network Reset", "YesNo", "Warning")

    # Proceed only if the user clicks "Yes"
    if ($Confirm -eq "Yes") {
        # 1. Create the Batch Script
        # Define the content of the repair batch file using a Here-String
        $BatchContent = @"
@echo off
Title Network Connectivity Fixer
color 1f
cls
echo ===========================================
echo     Resetting Network Configurations
echo ===========================================
echo.

echo [1/7] Flushing DNS...
ipconfig /flushdns
echo.

echo [2/7] Releasing & Renewing IP...
ipconfig /release
ipconfig /renew
echo.

echo [3/7] Resetting Winsock Catalog...
netsh winsock reset
echo.

echo [4/7] Resetting TCP/IP Stack...
netsh int ip reset
echo.

echo [5/7] Hard Resetting Network Config...
netcfg -d
echo.

echo [6/7] Resetting Interfaces...
netsh interface ipv4 reset
netsh interface ipv6 reset
netsh interface tcp reset
echo.

echo [7/7] Verifying Connectivity (Ping Google)...
ping 8.8.8.8 -n 4
echo.

echo ===========================================
echo     Process Complete. 
echo     PLEASE RESTART YOUR COMPUTER.
echo ===========================================
pause
"@
        # 2. Save to Temp File
        # Define the temporary path to save the batch script
        $TempPath = "$env:TEMP\Fix-Network.bat"
        # Export the batch content to the file using ASCII encoding
        $BatchContent | Out-File $TempPath -Encoding ASCII -Force

        # 3. Run as Administrator
        try {
            # Execute the batch file with elevated privileges
            Start-Process cmd.exe -ArgumentList "/c `"$TempPath`"" -Verb RunAs
        } catch {
            # Show a message if the elevation process is denied or cancelled
            [System.Windows.Forms.MessageBox]::Show("Process cancelled.", "Info")
        }
    }
})

# Add the initialized repair button to the Network Tab's FlowLayout panel
$FlowNet.Controls.Add($Btn_NetFix)



# ==============================================================================
#                         BUTTON --> Repair Winsock & TCP/IP
# ==============================================================================

# BUTTON: Repair Winsock & TCP/IP
# Initialize a new Button object for the repair utility
$Btn_Winsock = New-Object System.Windows.Forms.Button
# Set the visible label for the button
$Btn_Winsock.Text = "Repair Winsock & TCP/IP"
# Apply standardized toolkit styling to the button
Format-Button $Btn_Winsock

# Define the action to take when the button is clicked
$Btn_Winsock.Add_Click({
    # 1. Warning Confirmation
    # Create the text for the initial warning message
    $Msg = "This will reset the Winsock Catalog and TCP/IP Stack to fix connectivity issues.`n`nA system RESTART will be required immediately.`n`nContinue?"
    # Display a Yes/No warning box to ensure user intent
    $Confirm = [System.Windows.Forms.MessageBox]::Show($Msg, "Confirm Repair", "YesNo", "Warning")

    # Proceed only if the user confirms the action
    if ($Confirm -eq "Yes") {
        # Visual Feedback
        # Change the mouse cursor to a loading/wait icon
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        # Disable the button to prevent multiple simultaneous clicks
        $Btn_Winsock.Enabled = $false
        # Process pending UI events to ensure the cursor change renders immediately
        [System.Windows.Forms.Application]::DoEvents()

        try {
            # 2. Run Commands Hidden
            # Execute the Winsock reset command silently in the background
            Start-Process netsh -ArgumentList "winsock reset" -WindowStyle Hidden -Wait
            # Execute the TCP/IP stack reset command silently in the background
            Start-Process netsh -ArgumentList "int ip reset" -WindowStyle Hidden -Wait

            # 3. Success & Restart Prompt
            # Prepare the success message with a restart instruction
            $RestartMsg = "✅ Repair Complete!`n`nYou must restart the computer for changes to take effect.`n`nRestart Now?"
            # Ask the user if they want to reboot the system immediately
            $Restart = [System.Windows.Forms.MessageBox]::Show($RestartMsg, "Restart Required", "YesNo", "Question")

            # Force an immediate system restart if the user chooses 'Yes'
            if ($Restart -eq "Yes") {
                Restart-Computer -Force
            }
        } catch {
            # Display an error message if any part of the reset process fails
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed")
        } finally {
            # Restore the default mouse cursor after completion or failure
            $Form.Cursor = [System.Windows.Forms.Cursors]::Default
            # Re-enable the button for future use
            $Btn_Winsock.Enabled = $true
        }
    }
})

# Add the initialized repair button to the Network Tab's FlowLayout container
$FlowNet.Controls.Add($Btn_Winsock)



# ==============================================================================
#                         BUTTON --> Get Public IP
# ==============================================================================

# Initialize the button object for Public IP retrieval
$Btn_PubIP = New-Object System.Windows.Forms.Button
# Set the text label displayed on the button
$Btn_PubIP.Text = "Get Public IP"
# Apply the standard toolkit formatting
Format-Button $Btn_PubIP
# Define the action to take when the button is clicked
$Btn_PubIP.Add_Click({
    try {
        # Change the mouse cursor to a 'Wait' icon for visual feedback
        $Form.Cursor = "WaitCursor"
        # Query the ipify API to retrieve the public IP address
        $IP = Invoke-RestMethod -Uri "https://api.ipify.org" -ErrorAction Stop
        # Display the result in a message box
        [System.Windows.Forms.MessageBox]::Show("Your Public IP is: $IP", "Public IP")
    } catch { 
        # Show an error message if the API call fails
        [System.Windows.Forms.MessageBox]::Show("Failed to get IP.", "Error") 
    }
    finally { 
        # Revert the mouse cursor to the default pointer
        $Form.Cursor = "Default" 
    }
})
# Add the button to the network tab layout panel
$FlowNet.Controls.Add($Btn_PubIP)



# ==============================================================================
#                         BUTTON --> Trace Route
# ==============================================================================

# BUTTON: Trace Route
# Initialize a new Button object specifically for the Trace Route utility
$Btn_Trace = New-Object System.Windows.Forms.Button
# Set the display text for the button surface
$Btn_Trace.Text = "Trace Route"
# Apply standardized toolkit formatting using your custom function
Format-Button $Btn_Trace

# Define the action to execute when the button is clicked
$Btn_Trace.Add_Click({
    # Create a secondary popup window to gather user input
    $InputForm = New-Object System.Windows.Forms.Form
    # Set the title of the input window
    $InputForm.Text = "Traceroute Target"
    # Define window dimensions (Width, Height)
    $InputForm.Size = New-Object System.Drawing.Size(300, 150)
    # Ensure the popup appears centered relative to the main toolkit window
    $InputForm.StartPosition = "CenterParent"

    # Create a text label for the input field
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Enter IP or Domain:"
    # Set label coordinates (X, Y)
    $Label.Location = New-Object System.Drawing.Point(10, 10)
    
    # Create an editable text box for user input
    $TextBox = New-Object System.Windows.Forms.TextBox
    # Set the default value to a common domain
    $TextBox.Text = "google.com"
    $TextBox.Location = New-Object System.Drawing.Point(10, 35)
    # Set width to match the form size
    $TextBox.Width = 260

    # Create a confirmation button
    $OKBtn = New-Object System.Windows.Forms.Button
    $OKBtn.Text = "OK"
    $OKBtn.Location = New-Object System.Drawing.Point(100, 70)
    # Assign the standard 'OK' dialog result to this button
    $OKBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Add all created elements to the popup form container
    $InputForm.Controls.AddRange(@($Label, $TextBox, $OKBtn))
    
    # Show the popup and check if the user clicked the 'OK' button
    if ($InputForm.ShowDialog() -eq "OK") {
        # Capture the text entered by the user
        $Target = $TextBox.Text
        # Launch a new PowerShell process to run the diagnostic command in a separate window
        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "Test-NetConnection $Target -TraceRoute"
    }
})

# Add the completed button to the Network Tab's FlowLayout container
$FlowNet.Controls.Add($Btn_Trace)



# ==============================================================================
#                         BUTTON --> DNS Lookup
# ==============================================================================

# BUTTON: DNS Lookup
# Initialize a new Button object for the DNS Lookup utility
$Btn_DNS = New-Object System.Windows.Forms.Button
# Set the visible label text for the button
$Btn_DNS.Text = "DNS Lookup"
# Apply standardized toolkit styling to the button using your custom function
Format-Button $Btn_DNS
# Define the action to execute when the button is clicked
$Btn_DNS.Add_Click({
    # Create a secondary popup window to collect the domain name from the user
    $InputForm = New-Object System.Windows.Forms.Form
    $InputForm.Text = "DNS Lookup"
    # Set the input window dimensions (Width, Height)
    $InputForm.Size = New-Object System.Drawing.Size(300,150)
    # Center the popup relative to the main application window
    $InputForm.StartPosition = "CenterParent"
    # Create a label to instruct the user
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "Enter Domain:"
    $Label.Location = New-Object System.Drawing.Point(10,10)
    # Create an input box with a default value of microsoft.com
    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Text = "microsoft.com"; $TextBox.Location = New-Object System.Drawing.Point(10,35); $TextBox.Width = 260
    # Create an 'OK' button to submit the form
    $OKBtn = New-Object System.Windows.Forms.Button
    $OKBtn.Text = "OK"; $OKBtn.Location = New-Object System.Drawing.Point(100,70); $OKBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    # Add all created UI elements to the popup form
    $InputForm.Controls.AddRange(@($Label, $TextBox, $OKBtn))
    
    # Show the dialog and proceed only if the user clicks 'OK'
    if ($InputForm.ShowDialog() -eq "OK") {
        # Capture the domain name entered by the user
        $Domain = $TextBox.Text
        # Resolve DNS records for the domain and display results in a searchable grid
        Resolve-DnsName $Domain -ErrorAction SilentlyContinue | Out-GridView -Title "DNS Records for $Domain"
    }
})
# Add the finished button to the Network Tab's FlowLayout container
$FlowNet.Controls.Add($Btn_DNS)



# ==============================================================================
#                         BUTTON --> FLUSH DNS
# ==============================================================================

# [BUTTON 12] FLUSH DNS
# Initialize a new Button object for the DNS repair utility
$Btn_FlushDNS = New-Object System.Windows.Forms.Button

# Set the visible label text for the button
$Btn_FlushDNS.Text = "Flush DNS"

# Apply standardized toolkit styling to the button using your custom function
Format-Button $Btn_FlushDNS

# Define the action to execute when the button is clicked
$Btn_FlushDNS.Add_Click({
    # Change the mouse cursor to a loading/wait icon for visual feedback
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    try {
        # Execute the built-in cmdlet to clear the local DNS resolver cache
        Clear-DnsClientCache
        
        # Display a success message to the user upon completion
        [System.Windows.Forms.MessageBox]::Show("✅ DNS Cache flushed !", "Network Success")
    } catch {
        # Provide an error message if the operation fails (e.g., due to permissions)
        [System.Windows.Forms.MessageBox]::Show("❌ Error: $($_.Exception.Message)", "Flush DNS Failed")
    } finally {
        # Restore the default mouse cursor after the process completes
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the initialized flush button to the Network Tab's FlowLayout container
$FlowNet.Controls.Add($Btn_FlushDNS)



# ==============================================================================
#                         BUTTON --> IP RELEASE/RENEW
# ==============================================================================

# BUTTON: IP RELEASE/RENEW
# Initialize a new Button object for the IP reset utility
$Btn_IPReset = New-Object System.Windows.Forms.Button
# Set the visible label for the button
$Btn_IPReset.Text = "IP Release/Renew"
# Apply the standardized toolkit formatting to the button
Format-Button $Btn_IPReset

# Define the action that triggers when the button is clicked
$Btn_IPReset.Add_Click({
    # 1. Define Warning Message
    # Set the text for the user confirmation dialog
    $Msg = "Warning: Releasing and renewing your IP address will temporarily disconnect you from the network.`n`nDo you want to proceed?"
    
    # 2. Show MessageBox with Yes/No options
    # Display the warning box and capture the user's choice (Yes/No)
    $Choice = [System.Windows.Forms.MessageBox]::Show($Msg, "Network Warning", "YesNo", "Warning")
    
    # 3. Run If statement on the result
    # Execute repair logic only if the user explicitly confirms
    if ($Choice -eq "Yes") {
        # Change the mouse cursor to a loading/wait icon
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            # Execute command to drop the current DHCP assigned IP address
            ipconfig /release
            # Pause for 2 seconds to allow network stack to process the release
            Start-Sleep -Seconds 2
            # Request a new IP address from the DHCP server
            ipconfig /renew
            # Show success message once a new IP is assigned
            [System.Windows.Forms.MessageBox]::Show("✅ IP Address successfully renewed!", "Success")
        } catch {
            # Display error details if the renew process fails
            [System.Windows.Forms.MessageBox]::Show("❌ Error: $($_.Exception.Message)", "Reset Failed")
        } finally {
            # Revert the mouse cursor to the default pointer
            $Form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})

# Add the completed button to the Network Tab's FlowLayout container
$FlowNet.Controls.Add($Btn_IPReset)



# ==============================================================================
#                         BUTTON --> IP Config /All
# ==============================================================================

$btnIPConfig = New-Object System.Windows.Forms.Button
$btnIPConfig.Text = "IP Config /All"
Format-Button $btnIPConfig
$btnIPConfig.Add_Click({
    ipconfig /all | Out-GridView -Title "Network Details"
})
$FlowNet.Controls.Add($btnIPConfig)




# BUTTON Set IP/DNS to Auto
$Btn_AutoIP = New-Object System.Windows.Forms.Button
$Btn_AutoIP.Text = "Set IP/DNS to Auto"
Format-Button $Btn_AutoIP

$Btn_AutoIP.Add_Click({
    # Admin Check
    $Principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        [System.Windows.Forms.MessageBox]::Show("Admin privileges required. Please 'Run as Administrator'.", "Access Denied")
        return
    }

    try {
        $Interface = Get-NetIPInterface -AddressFamily IPv4 | Where-Object { $_.ConnectionState -eq "Connected" } | Select-Object -First 1
        Set-NetIPInterface -InterfaceIndex $Interface.InterfaceIndex -DHCP Enabled
        Set-DnsClientServerAddress -InterfaceIndex $Interface.InterfaceIndex -ResetToAutomatic
        [System.Windows.Forms.MessageBox]::Show("Successfully reset $($Interface.InterfaceAlias) to Automatic.", "Success")
    } catch { [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed") }
})
$FlowNet.Controls.Add($Btn_AutoIP)



# ==============================================================================
#                         BUTTON --> Reset ALL Adapters
# ==============================================================================

$Btn_AutoAll = New-Object System.Windows.Forms.Button
$Btn_AutoAll.Text = "Reset ALL Adapters"
Format-Button $Btn_AutoAll

$Btn_AutoAll.Add_Click({
    # Admin Check
    $Principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        [System.Windows.Forms.MessageBox]::Show("Admin privileges required to reset all adapters.", "Access Denied")
        return
    }

    try {
        Get-NetIPInterface -AddressFamily IPv4 | ForEach-Object {
            Set-NetIPInterface -InterfaceIndex $_.InterfaceIndex -DHCP Enabled -ErrorAction SilentlyContinue
            Set-DnsClientServerAddress -InterfaceIndex $_.InterfaceIndex -ResetToAutomatic -ErrorAction SilentlyContinue
        }
        [System.Windows.Forms.MessageBox]::Show("All adapters reset to Automatic.", "Success")
    } catch { [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed") }
})
$FlowNet.Controls.Add($Btn_AutoAll)



# ==============================================================================
#                         BUTTON --> Set IP/DNS to Auto
# ==============================================================================

# [BUTTON: Set IP/DNS to Auto]
# Initialize a new Button object for resetting network settings
$Btn_AutoIP = New-Object System.Windows.Forms.Button
# Set the visible label text for the button
$Btn_AutoIP.Text = "Set IP/DNS to Auto"
# Apply standardized toolkit formatting to the button
Format-Button $Btn_AutoIP

# Define the action to execute when the button is clicked
$Btn_AutoIP.Add_Click({
    # Admin Check: Verify that the toolkit is running with elevated privileges [[2](https://jingyan.baidu.com/article/54b6b9c0e6641d6c593b476c.html)]
    $Principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        # Halt execution and alert the user if they lack permissions
        [System.Windows.Forms.MessageBox]::Show("Admin privileges required. Please 'Run as Administrator'.", "Access Denied")
        return
    }

    try {
        # Identify the first active IPv4 network interface currently connected
        $Interface = Get-NetIPInterface -AddressFamily IPv4 | Where-Object { $_.ConnectionState -eq "Connected" } | Select-Object -First 1
        # Re-enable DHCP for IP address assignment on the found interface
        Set-NetIPInterface -InterfaceIndex $Interface.InterfaceIndex -DHCP Enabled
        # Reset the DNS client to automatically obtain server addresses via DHCP
        Set-DnsClientServerAddress -InterfaceIndex $Interface.InterfaceIndex -ResetToAutomatic
        # Notify the user of the successful reset for the specific adapter
        [System.Windows.Forms.MessageBox]::Show("Successfully reset $($Interface.InterfaceAlias) to Automatic.", "Success")
    } catch { 
        # Display an error message if the interface retrieval or setting change fails
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Failed") 
    }
})

# Add the initialized button to the Network Tab's FlowLayout container
$FlowNet.Controls.Add($Btn_AutoIP)



# ==============================================================================
#                         BUTTON --> Hosts File Manager
# ==============================================================================

# BUTTON: Hosts File Manager
$Btn_HostsManager = New-Object System.Windows.Forms.Button
$Btn_HostsManager.Text = "Hosts File Manager"
Format-Button $Btn_HostsManager

$Btn_HostsManager.Add_Click({
    # 1. Admin Privilege Check
    $Principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        [System.Windows.Forms.MessageBox]::Show("Admin privileges are required to edit the hosts file. Please run as Administrator.", "Access Denied")
        return
    }

    # 2. Open Sub-Window for Options
    $HostsForm = New-Object System.Windows.Forms.Form
    $HostsForm.Text = "Hosts Manager"; $HostsForm.Size = "350,280"; $HostsForm.StartPosition = "CenterParent"
    $Flow = New-Object System.Windows.Forms.FlowLayoutPanel; $Flow.Dock = "Fill"; $HostsForm.Controls.Add($Flow)
    
    # Declare path first
$Path = "$env:windir\System32\drivers\etc\hosts"

$Btn_View = New-Object System.Windows.Forms.Button
$Btn_View.Text = "View Hosts (Read-Only)"
$Btn_View.Width = 320

# Call to function/action inside the script block
$Btn_View.Add_Click({
    $Path = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $Path) {
        # Opens the hosts file directly in Notepad
        Start-Process notepad.exe -ArgumentList $Path
    } else {
        [System.Windows.Forms.MessageBox]::Show("Hosts file not found at $Path")
    }
})

    # Option 2: Add Entry
    $Btn_Add = New-Object System.Windows.Forms.Button; $Btn_Add.Text = "Add Host Entry"; $Btn_Add.Width = 320
    $Btn_Add.Add_Click({
    $IP = [Microsoft.VisualBasic.Interaction]::InputBox("Enter IP:", "Add Entry")
    $URL = [Microsoft.VisualBasic.Interaction]::InputBox("Enter URL:", "Add Entry")
    if ($IP -and $URL) { 
        Add-Content $Path -Value "`n$IP $URL" -Force 
        [System.Windows.Forms.MessageBox]::Show("Host added successfully", "Success") # Confirmation message
    }
})

    # Option 3: Remove Entry
    $Btn_Rem = New-Object System.Windows.Forms.Button; $Btn_Rem.Text = "Remove Host Entry"; $Btn_Rem.Width = 320
    $Btn_Rem.Add_Click({
    $URL = [Microsoft.VisualBasic.Interaction]::InputBox("Enter URL to Remove:", "Remove Entry")
    if ($URL) { 
        $Content = Get-Content $Path
        $NewContent = $Content | Where-Object { $_ -notmatch "\s$URL(\s|$)" }
        $NewContent | Set-Content $Path -Force 
        [System.Windows.Forms.MessageBox]::Show("Host removed successfully", "Success") # Confirmation message
    }
})

    $Flow.Controls.AddRange(@($Btn_View, $Btn_Add, $Btn_Rem))
    $HostsForm.ShowDialog()
})
$FlowNet.Controls.Add($Btn_HostsManager)



# Add Panel to Tab (ONLY ONCE)
$TabNet.Controls.Add($FlowNet)

# Add Tab to Main Control (ONLY ONCE)
$TabControl.TabPages.Add($TabNet)


#*****************************************************************************************************************************************************************************************************

# *****************************************************************************************************************************************************************************************************
# ==============================================================================
#                              LOGS TAB
# This section initializes the Logs Tab, providing a dedicated area for event monitoring and record analysis.
# ==============================================================================

# --- TAB 3: Logs ---
$TabLog = New-Object System.Windows.Forms.TabPage
$TabLog.Text = "Logs"
$TabLog.BackColor = "#2d2d30"

# FlowLayout Setup
$FlowLog = New-Object System.Windows.Forms.FlowLayoutPanel
$FlowLog.Dock = "Fill" 
$FlowLog.AutoScroll = $true

# Double Buffering
Enable-DoubleBuffering $FlowLog

# Background Image
$ImgPath = "$PSScriptRoot\Images\WhiteBG.png"
if (Test-Path $ImgPath) {
    $TabLog.BackgroundImage = [System.Drawing.Image]::FromFile($ImgPath)
    $TabLog.BackgroundImageLayout = "Stretch" # FIXED: Added quotes
}

# Transparency
$FlowLog.BackColor = [System.Drawing.Color]::Transparent



# ==============================================================================
#                         BUTTON --> Get Event Log Errors
# ==============================================================================

# BUTTON: Get Event Log Errors
# Initialize a new Button object for retrieving event logs
$Btn_EvntLog = New-Object System.Windows.Forms.Button
# Set the visible label text for the button
$Btn_EvntLog.Text = "Get Event Log Errors"
# Apply standardized toolkit styling to the button
Format-Button $Btn_EvntLog

# Define the action to execute when the button is clicked
$Btn_EvntLog.Add_Click({
    # Change the mouse cursor to a loading/wait icon for visual feedback
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    # Initialize an empty array to store collected log data
    $LogData = @()
    
    # Check System and Application Logs
    # Define an array of the core log files to search [[5](https://www.tenforums.com/tutorials/25581-open-windows-powershell-windows-10-a.html)]
    $LogNames = @("System", "Application")
    
    # Iterate through each defined log name
    foreach ($Log in $LogNames) {
        # Level 1=Critical, 2=Error
        # Query the top 50 events filtered by severity levels 1 and 2
        $Events = Get-WinEvent -FilterHashtable @{LogName=$Log; Level=1,2} -MaxEvents 50 -ErrorAction SilentlyContinue
        
        # Proceed if events were successfully retrieved
        if ($Events) {
            # Add a SourceLog property to distinguish them in the grid and select relevant fields
            $LogData += $Events | Select-Object @{N='SourceLog';E={$Log}}, TimeCreated, Id, LevelDisplayName, Message
        }
    }

    # Display the compiled log data in an interactive, searchable grid window
    $LogData | Out-GridView -Title "Last 50 Critical/Error Events (System & Application)"
    # Restore the mouse cursor to the default pointer
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# Add the initialized log viewer button to the Logs Tab's FlowLayout container
$FlowLog.Controls.Add($Btn_EvntLog)




# ==============================================================================
#                         BUTTON --> Analyze Startup & Shutdown
# ==============================================================================

# BUTTON: Analyze Startup & Shutdown
# Initialize a new Button object for boot/shutdown diagnostics
$Btn_BootLog = New-Object System.Windows.Forms.Button
# Set the visible label text for the button
$Btn_BootLog.Text = "Analyze Boot/Shutdown"
# Apply standardized toolkit formatting to the button
Format-Button $Btn_BootLog

# Define the action to execute when the button is clicked
$Btn_BootLog.Add_Click({
    # Change the mouse cursor to a loading/wait icon for visual feedback
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Event IDs: 6005 (Boot), 6006 (Clean Shutdown), 6008 (Dirty Shutdown), 1074 (User Initiated)
    # Define an array of specific Event IDs relevant to system power states
    $IDs = @(6005, 6006, 6008, 1074)

    try {
        # Query the 'System' log for the specified IDs, limiting to the 100 most recent events [[5](https://www.tenforums.com/tutorials/25581-open-windows-powershell-windows-10-a.html)]
        $Events = Get-WinEvent -FilterHashtable @{LogName='System'; Id=$IDs} -MaxEvents 100 -ErrorAction Stop
        
        # Transform the raw event data into a more readable format using custom labels
        $Results = $Events | Select-Object TimeCreated, Id, @{N='Type';E={
            if ($_.Id -eq 6005) { "Boot/Startup" }
            elseif ($_.Id -eq 6008) { "Unexpected Shutdown" }
            elseif ($_.Id -eq 1074) { "User Initiated" }
            else { "Normal Shutdown" }
        }}, Message

        # Output the analyzed results to an interactive, searchable grid window
        $Results | Out-GridView -Title "Recent Boot & Shutdown Events"
    } catch {
        # Notify the user if no logs match the criteria or if permissions are missing [[2](https://jingyan.baidu.com/article/54b6b9c0e6641d6c593b476c.html)]
        [System.Windows.Forms.MessageBox]::Show("No recent boot events found or access denied.", "Info")
    } finally {
        # Restore the mouse cursor to the default pointer
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the initialized diagnostic button to the Logs Tab's FlowLayout container
$FlowLog.Controls.Add($Btn_BootLog)



# ==============================================================================
#                         BUTTON --> Get User Login History
# ==============================================================================

# BUTTON: Get User Login History
# Initialize a new Button object for auditing user access
$Btn_LoginHist = New-Object System.Windows.Forms.Button
# Set the visible text displayed on the button
$Btn_LoginHist.Text = "User Login History"
# Apply the standardized toolkit formatting to the button
Format-Button $Btn_LoginHist

# Define the logic to execute when the button is clicked
$Btn_LoginHist.Add_Click({
    # Change the mouse cursor to a loading/wait icon for visual feedback
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # 4624 = Successful Login, 4625 = Failed Login Attempt
    # Define an array containing the specific Security Event IDs to track
    $EventIDs = @(4624, 4625) 

    try {
        # Fetch Events
        # Query the Security log for the defined IDs, limiting results to the most recent 100 entries
        $Events = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=$EventIDs} -MaxEvents 100 -ErrorAction Stop
        
        # Process and Filter Data
        # Iterate through each event to extract and clean human-readable information
        $Results = $Events | ForEach-Object {
            # Extract User Name (Index 5 in the event properties contains the TargetUserName)
            $User = $_.Properties[5].Value
            if (-not $User) { $User = "System/Unknown" }

            # Filter out Computer Accounts (ending in $) and system service accounts for clarity
            if ($User -notlike "*$" -and $User -ne "SYSTEM" -and $User -ne "DWM-1" -and $User -ne "UMFD-0") {
                # Create a custom object with the cleaned data for display
                [PSCustomObject]@{
                    Time   = $_.TimeCreated
                    Status = if ($_.Id -eq 4624) { "✅ Success" } else { "❌ FAILED" }
                    User   = $User
                    ID     = $_.Id
                }
            }
        }

        # Show Results
        # If human logins were found, display them in an interactive grid; otherwise, notify the user
        if ($Results) {
            $Results | Out-GridView -Title "Recent User Login History (Last 100 Attempts)"
        } else {
            [System.Windows.Forms.MessageBox]::Show("No human user logins found in the last 100 events.", "Info")
        }

    } catch {
        # Provide error feedback if the Security log is inaccessible (usually requires Admin rights)
        [System.Windows.Forms.MessageBox]::Show("Error reading Security Log. Run as Administrator.", "Access Denied")
    } finally {
        # Restore the mouse cursor to the default pointer
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the initialized login history button to the Logs Tab's FlowLayout container
$FlowLog.Controls.Add($Btn_LoginHist)



# ==============================================================================
#                         BUTTON --> Search Logs
# ==============================================================================

# BUTTON: Search Logs
# Initialize a new Button object for the log search utility
$Btn_SearchLog = New-Object System.Windows.Forms.Button
# Set the visible label text for the button
$Btn_SearchLog.Text = "Search Logs by Keyword"
# Apply standardized toolkit formatting to the button
Format-Button $Btn_SearchLog

# Define the action to execute when the button is clicked
$Btn_SearchLog.Add_Click({
    # 1. Create native Input Dialog
    # Create a secondary popup form for the search interface
    $InputForm = New-Object System.Windows.Forms.Form
    $InputForm.Text = "Search Logs"
    $InputForm.Size = New-Object System.Drawing.Size(300,150)
    # Center the search box relative to the main toolkit window
    $InputForm.StartPosition = "CenterParent"
    
    # Create a text box for the user to enter their search keyword
    $TextBox = New-Object System.Windows.Forms.TextBox
    # Set default search term and position the box
    $TextBox.Text = "Outlook"; $TextBox.Location = New-Object System.Drawing.Point(10,35); $TextBox.Width = 260
    
    # Create an 'OK' button to confirm and start the search
    $OKBtn = New-Object System.Windows.Forms.Button
    $OKBtn.Text = "Search"; $OKBtn.Location = New-Object System.Drawing.Point(100,70); $OKBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    # Add the input controls to the popup form
    $InputForm.Controls.AddRange(@($TextBox, $OKBtn))

    # 2. Execute if OK is clicked
    # Proceed only if the user confirms the dialog
    if ($InputForm.ShowDialog() -eq "OK") {
        # Store the keyword from the text box
        $Key = $TextBox.Text
        # Change the mouse cursor to a loading/wait icon
        $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        # Identify relevant logs and iterate through them [[5](https://www.tenforums.com/tutorials/25581-open-windows-powershell-windows-10-a.html)]
        Get-WinEvent -ListLog Application, System | ForEach-Object {
            # Query the 2000 most recent events and filter for the user's keyword
            Get-WinEvent -LogName $_.LogName -MaxEvents 2000 -ErrorAction SilentlyContinue | Where-Object { $_.Message -match $Key }
        } | Select-Object TimeCreated, LogName, Id, LevelDisplayName, Message | Out-GridView -Title "Search Results: $Key"
        
        # Restore the mouse cursor to the default pointer after the search finishes
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Add the initialized search button to the Logs Tab's FlowLayout container
$FlowLog.Controls.Add($Btn_SearchLog)



# ==============================================================================
#                         BUTTON --> Export Logs to CSV
# ==============================================================================

# BUTTON: Export Logs to CSV
# Initialize a new Button object for log exportation
$Btn_Export = New-Object System.Windows.Forms.Button
# Set the visible label text for the button
$Btn_Export.Text = "Export Errors to CSV"
# Apply standardized toolkit formatting to the button
Format-Button $Btn_Export

# Define the action to execute when the button is clicked
$Btn_Export.Add_Click({
    # Change the mouse cursor to a loading/wait icon for visual feedback
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # 1. Define Folder and File Path
    # Set the target storage directory for log files
    $Folder = "C:\Errors Logs"
    # Generate a unique filename using the current date and time (YearMonthDay_HourMinute)
    $FileName = "ErrorLog_Export_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    # Combine the folder and filename into a single absolute path
    $Path = Join-Path -Path $Folder -ChildPath $FileName

    # 2. Auto-Create Folder if it doesn't exist
    # Verify if the destination directory exists; if not, create it [[5](https://www.tenforums.com/tutorials/25581-open-windows-powershell-windows-10-a.html)]
    if (-not (Test-Path $Folder)) {
        New-Item -ItemType Directory -Force -Path $Folder | Out-Null
    }

    # 3. Export Data
    try {
        # Query System and Application logs for Critical (1) and Error (2) levels
        Get-WinEvent -FilterHashtable @{LogName='System','Application'; Level=1,2} -MaxEvents 500 -ErrorAction Stop | 
            # Filter the raw event data to include only essential diagnostic columns
            Select-Object TimeCreated, LogName, Id, Message | 
            # Save the filtered results to the CSV file using standard ASCII encoding
            Export-Csv -Path $Path -NoTypeInformation -Encoding ASCII
        
        # Notify the user of the successful export and show them the file location
        [System.Windows.Forms.MessageBox]::Show("✅ Success!`n`nFile Saved at: $Path", "Export Complete")
    } catch {
        # Provide error feedback if the export fails (likely due to missing Admin rights) [[2](https://jingyan.baidu.com/article/54b6b9c0e6641d6c593b476c.html)]
        [System.Windows.Forms.MessageBox]::Show("Failed to export logs.`nMake sure you run as Administrator.", "Error")
    }

    # Restore the mouse cursor to the default pointer after completion
    $Form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# Add the initialized export button to the Logs Tab's FlowLayout container
$FlowLog.Controls.Add($Btn_Export)



# ==============================================================================
#                         BUTTON --> BSOD History
# ==============================================================================

# BUTTON: BSOD History
# Initialize a new Button object for crash diagnostics
$Btn_BSOD = New-Object System.Windows.Forms.Button
# Set the visible text displayed on the button
$Btn_BSOD.Text = "BSOD History"
# Apply the standardized toolkit formatting to the button
Format-Button $Btn_BSOD

# Define the logic to execute when the button is clicked
$Btn_BSOD.Add_Click({
    # Change the mouse cursor to a loading/wait icon for visual feedback
    $Form.Cursor = "WaitCursor"
    
    try {
        # Check System Log for BugCheck (1001)
        # Search the System log for Event ID 1001, which signifies a system crash (BSOD) [[5](https://www.tenforums.com/tutorials/25581-open-windows-powershell-windows-10-a.html)]
        $Events = Get-WinEvent -FilterHashtable @{LogName='System'; Id=1001} -MaxEvents 50 -ErrorAction Stop
        
        # Output found crash events to an interactive, searchable grid window
        $Events | Select-Object TimeCreated, Id, Message | Out-GridView -Title "Blue Screen Events (BugCheck 1001)"
    } catch {
        # This catch block runs if NO events are found or access is denied [[2](https://jingyan.baidu.com/article/54b6b9c0e6641d6c593b476c.html)]
        # Notify the user that the system logs appear healthy with no recorded crashes
        [System.Windows.Forms.MessageBox]::Show("Good news! No Blue Screen (BSOD) events found in the System Log.", "System Healthy")
    } finally {
        # Restore the mouse cursor to the default pointer
        $Form.Cursor = "Default"
    }
})

# Add the initialized diagnostic button to the Logs Tab's FlowLayout container
$FlowLog.Controls.Add($Btn_BSOD)



# ==============================================================================
#                         BUTTON --> Windows Update History
# ==============================================================================

# BUTTON: Windows Update History
# Create a new Button object for the update history tool
$Btn_WinUpd = New-Object System.Windows.Forms.Button
# Set the text displayed on the button
$Btn_WinUpd.Text = "Windows Update History"
# Apply standardized formatting using your custom function
Format-Button $Btn_WinUpd

# Define the action that occurs when the button is clicked
$Btn_WinUpd.Add_Click({
    # Change the mouse cursor to a loading/wait icon for visual feedback
    $Form.Cursor = "WaitCursor"
    
    # Query the 100 most recent events from the Windows Update Client provider
    Get-WinEvent -ProviderName 'Microsoft-Windows-WindowsUpdateClient' -MaxEvents 100 -ErrorAction SilentlyContinue |
        # Filter for specific Event IDs: 19 signifies success, 20 signifies failure
        Where-Object { $_.Id -in 19, 20 } | 
        # Select and format data: map ID 19 to 'Success' and others (20) to 'Failed'
        Select-Object TimeCreated, @{N='Status';E={if($_.Id -eq 19){'Success'}else{'Failed'}}}, Message |
        # Display the formatted results in a searchable GridView window
        Out-GridView -Title "Windows Update History"
    
    # Revert the mouse cursor to the default pointer
    $Form.Cursor = "Default"
})

# Add the initialized button to the Log Tab's layout panel
$FlowLog.Controls.Add($Btn_WinUpd)





# Add Panel to Tab (ONLY ONCE)
$TabLog.Controls.Add($FlowLog)

# Add Tab to Main Control (ONLY ONCE)
$TabControl.Controls.Add($TabLog)


# *****************************************************************************************************************************************************************************************************


#Show Form
$Form.Controls.Add($TabControl)

# Resize Handlers
$Form.Add_ResizeBegin({ $Form.SuspendLayout() })
$Form.Add_ResizeEnd({ $Form.ResumeLayout(); $Form.Refresh() })










# ==============================================================================
#                         MASTER FOOTER (Docking Method)
# ==============================================================================

# Create the main container for the bottom of the application window
$FooterPanel = New-Object System.Windows.Forms.Panel
# Attach the panel firmly to the bottom edge of the form 
$FooterPanel.Dock = "Bottom"
# Set a fixed height of 30 pixels for the footer bar
$FooterPanel.Height = 30
# Allow the main form background to show through the footer
$FooterPanel.BackColor = "Transparent"

# 1. RIGHT SIDE (Links) - Added first to reserve its space on the right edge
$PanelLinks = New-Object System.Windows.Forms.FlowLayoutPanel
# Arrange children from right to left (useful for sticking items to the right)
$PanelLinks.FlowDirection = "RightToLeft"
# Attach the sub-panel to the right side of the footer 
$PanelLinks.Dock = "Right"
# Allow the panel to resize horizontally based on its content
$PanelLinks.AutoSize = $true
# Set the sizing mode to expand or shrink as needed
$PanelLinks.AutoSizeMode = "GrowAndShrink"
# Add 10px top padding to vertically center text and 5px right margin
$PanelLinks.Padding = New-Object System.Windows.Forms.Padding(0, 10, 5, 0)

# Links Components: Support Email Label
$LblSupport = New-Object System.Windows.Forms.Label
$LblSupport.Text = "Support: asasifshaikh668@gmail.com"
$LblSupport.AutoSize = $true
$LblSupport.ForeColor = "DimGray"

# Links Components: Vertical Separator
$Sep1 = New-Object System.Windows.Forms.Label
$Sep1.Text = " | "
$Sep1.AutoSize = $true
$Sep1.ForeColor = "DimGray"

# Add components to the Right Panel (processed in Right-to-Left order)
$PanelLinks.Controls.Add($LblSupport)
$PanelLinks.Controls.Add($Sep1)

# 2. LEFT SIDE (System Info) - Added second to claim the left edge 
$LblUser = New-Object System.Windows.Forms.Label
# Dynamic placeholder for Username and Machine Name info
$LblUser.Text = "$env:USERNAME | $env:COMPUTERNAME"
# Attach the label to the left side of the footer
$LblUser.Dock = "Left"
# Allow the label to grow based on the length of the names
$LblUser.AutoSize = $true
# Align text to the vertical center of the label
$LblUser.TextAlign = "MiddleLeft"
# Use a bold Segoe UI font for visibility
$LblUser.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
# Set text color to match the toolkit's primary blue theme
$LblUser.ForeColor = "#007ACC"
# Align the vertical position with the right-side links
$LblUser.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)

# 3. CENTER (Dev Credits) - Added last to 'Fill' the remaining middle space 
$LblDev = New-Object System.Windows.Forms.Label
$Year = Get-Date -Format yyyy
# Copyright and Developer information string
$LblDev.Text = "Developed By Mohammed Asif Shaikh | v1.0 | © $Year"
# Set to fill all remaining space between Left and Right elements
$LblDev.Dock = "Fill"
# Center the text horizontally within the filled area
$LblDev.TextAlign = "TopCenter"
# Use an italic font for a professional credit style
$LblDev.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$LblDev.ForeColor = "DimGray"
# Align the vertical position with the other footer elements
$LblDev.Padding = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)

# --- ADD TO MAIN PANEL ---
# Order is critical for Docking: Added in sequence to establish stack priority [[2](https://github.com/lazywinadmin/WinFormPS)]
$FooterPanel.Controls.Add($LblDev)   # Added first to become the lowest layer (Fill)
$FooterPanel.Controls.Add($LblUser)  # Claimed Left
$FooterPanel.Controls.Add($PanelLinks) # Claimed Right

# Finalize the footer addition to the main application window
$Form.Controls.Add($FooterPanel)











# [CRITICAL LAYOUT FIX] 
# This specific order ensures the Header and Footer stick to the edges,
# and the Tabs squeeze perfectly into the middle space.
# ==============================================================================

# 1. Send Header and Footer to the "Back" of the stack. 
# This forces Windows to reserve their space (Top 40px and Bottom 30px) first.
$HeaderPanel.SendToBack()
$FooterPanel.SendToBack()


# 2. Bring Tabs to the "Front".
# Since the Top and Bottom are now reserved, "Fill" will only fill the remaining middle.
$TabControl.BringToFront()

# 3. Bring AI Assistant fixed button to the "Front".

$Btn_AiAssistant.BringToFront()



$Form.ShowDialog()