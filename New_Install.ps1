# Auto Installer
cd $PSScriptRoot
@'
#######################################################
#                                                     #
#  PLEASE ONLY USE THIS SCRIPT FOR BRAND NEW MACHINES #
#     WITH NO MONEYWORKS OR AUTOQUOTE DATA ON THEM    #
#                                                     #
#  =================================================  #
#                                                     #
#    1. Script is being run from AQ Drive             #
#    2. Powershell is elevated                        #
#    3. MAKE SURE NO DATA IS ALREADY PRESENT          #
#    4. MAKE SURE NO DATA IS ALREADY PRESENT          #
#    5. MAKE SURE NO DATA IS ALREADY PRESENT          #
#                                                     #
#######################################################

'@

echo 'Ensure you are running this script in Q:\ or the location you need Auto-Quote instlaled'

Start-Sleep -s 5

# Turn off UAC - REQUIRES RESTART
New-ItemProperty -Path HKLM:Software\Microsoft\Windows\CurrentVersion\policies\system -Name EnableLUA -PropertyType DWord -Value 0 -Force

# Extrat AQwin directories to current directory. 
# This will extract Aqwin and New folders into current directory. 
$file = Resolve-Path 'data\Aqwin.zip'
Expand-Archive $file .

# Extract OCX install files into a folder called "OCX Install"
$ocx = Resolve-Path 'data\OCX.zip'
Expand-Archive $ocx .

# Create blank directories
$tempFolders = @('1out', '1po', '1in')

foreach ($p in $tempFolders) {
    if (Test-Path $p) {
        echo 'File Exists'
        } 
    else {
            New-Item -ItemType Directory $p
        }
    }

# Copies ocx files into windows\system
copy "OCX Install\Addocx\*" C:\windows\System

# Registers OCX files in windows\system
cd C:\windows\system
.\regocx.bat
cd $PSScriptRoot

# Runs ocx install exe
cd ".\OCX Install\AQWinOcx\"
.\setup.exe
cd $PSScriptRoot

# Setting Shortdate format
Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sShortDate -Value "dd/MMM/yyyy"

# Loading AQwin pathing hive
.\data\S_Hive.reg

# Install ORM OCX files
.\data\Ins_Modules\ORM\SETUP.EXE

# Installing Pnet msi
& '.\data\Ins_Modules\AQaami Installer.msi'

# Register the VFupdate
cd .\Aqwin\vfupdater
.\reg.bat
cd $PSScriptRoot

# Create shortcut on desktop
$Shell = New-Object -ComObject ("WScript.Shell")
$Shortcut = $Shell.CreateShortcut($env:PUBLIC + "\Desktop\Auto-Quote Windows.lnk")
$Shortcut.TargetPath = "Q:\aqwin\aqwin.exe"
$Shortcut.WorkingDirectory = "Q:\aqwin\"
$Shortcut.Save()

# Moneyworks Install
function gold{
    $version = Read-Host -Prompt 'Version 5, 6 or 7? _> '
    switch -Regex ($version)
        {
        '5|6|7'{ cd .\data\Moneyworks; $x = ".\setupV$($version).exe"; Start-Process $x } 
        'b'{Menu}
        }
}

function dc{
    $version = Read-Host 'Version 5, 6 or 7? _> '
    switch -Regex ($version)
        {
        '5|6|7'{ cd .\data\Moneyworks; $x = ".\setupV$($version).exe"; Start-Process $x }
        'b'{Menu}
        }

}

function mwopt{
    New-ItemProperty -Path "HKCU:\Software\VB and VBA Program Settings\AQW\Set Up" -Name DontUseMW -Value "Y"
}

function Menu {
    echo $m_menu
    $input = Read-Host -Prompt "_> "

    switch($input)
    {
        '1'{gold}
        '2'{dc}
        'n'{echo 'No Install'; mwopt}
    }
}

$m_menu = @'

    1. Gold
    2. DC
    
    n. No instal

'@

Menu

# Adding Custom Reports
cd $PSScriptRoot
copy .\data\Moneyworks\Custom_Reports "$($env:APPDATA)\Cognito\Moneyworks Gold"


echo @'
    Installation of Auto-Quote Windows is complete
        1. Auto-Quote Opens into OVD.
        2. Moneyworks opens AND connects.
        3. Moneyworks Import maps are correctly copied.

'@
pause
echo 'Goodbye Asshat'

Set-ExecutionPolicy Restricted

#TODO :: Actually add custom import maps.
