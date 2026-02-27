"""
Build interactive installer
Use NSIS or Inno Setup to create Windows installer
"""
import os
import subprocess
import sys


def create_installer_config():
    """创建 pynsist 配置文件"""
    config_content = """[Application]
name=Windows Port Manager
version=1.0.0
publisher=Port Manager Team
entry_point=main:main
icon=icon.ico
license_file=LICENSE

[Python]
version=3.11.0
include_msvcrt=true

[Include]
packages=tkinter

files=config.json
    README.md

[Command portmanager]
entry_point=main:main

[Build]
installer_name=PortManager_Setup.exe
nsi_template=installer.nsi
"""
    with open("installer.cfg", "w", encoding="utf-8") as f:
        f.write(config_content)


def create_nsis_template():
    """Create NSIS template with admin privileges"""
    nsi_content = r'''!include "MUI2.nsh"
!include "FileFunc.nsh"

; Basic info
Name "Windows Port Manager"
OutFile "..\installer_output\PortManager_Setup_NSIS.exe"
InstallDir "$PROGRAMFILES\PortManager"
RequestExecutionLevel admin

; UI settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Language
!insertmacro MUI_LANGUAGE "English"

; Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

Section "Install"
    SetOutPath "$INSTDIR"

    ; Copy files
    File /r "dist\*.*"
    File "config.json"
    File "README.md"

    ; Create uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"

    ; Create Start Menu shortcuts
    CreateDirectory "$SMPROGRAMS\Port Manager"
    CreateShortcut "$SMPROGRAMS\Port Manager\Port Manager.lnk" "$INSTDIR\PortManager.exe" "" "$INSTDIR\PortManager.exe" 0
    CreateShortcut "$SMPROGRAMS\Port Manager\Uninstall.lnk" "$INSTDIR\Uninstall.exe"

    ; Create Desktop shortcut
    CreateShortcut "$DESKTOP\Port Manager.lnk" "$INSTDIR\PortManager.exe"

    ; Write registry (Programs and Features)
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PortManager" "DisplayName" "Windows Port Manager"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PortManager" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PortManager" "DisplayIcon" "$INSTDIR\PortManager.exe"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PortManager" "Publisher" "Port Manager Team"
    WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PortManager" "DisplayVersion" "1.0.0"

    ; Calculate install size
    ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
    IntFmt $0 "0x%08X" $0
    WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PortManager" "EstimatedSize" "$0"
SectionEnd

Section "Uninstall"
    ; Delete files
    RMDir /r "$INSTDIR"

    ; Delete Start Menu
    RMDir /r "$SMPROGRAMS\Port Manager"

    ; Delete Desktop shortcut
    Delete "$DESKTOP\Port Manager.lnk"

    ; Delete registry
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\PortManager"
SectionEnd
'''
    with open("installer.nsi", "w", encoding="utf-8") as f:
        f.write(nsi_content)


def create_simple_installer():
    """Create installer script using Inno Setup format"""
    iss_content = r'''[Setup]
AppName=Windows Port Manager
AppVersion=1.0.0
AppPublisher=Port Manager Team
AppPublisherURL=https://github.com/yourusername/windows-port-manager
AppSupportURL=https://github.com/yourusername/windows-port-manager/issues
AppUpdatesURL=https://github.com/yourusername/windows-port-manager/releases
DefaultDirName={autopf}\PortManager
DefaultGroupName=Port Manager
OutputDir=installer_output
OutputBaseFilename=PortManager_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\PortManager.exe
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional options:"

[Files]
Source: "dist\PortManager.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\Port Manager"; Filename: "{app}\PortManager.exe"
Name: "{group}\Uninstall Port Manager"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Port Manager"; Filename: "{app}\PortManager.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\PortManager.exe"; Description: "Launch Port Manager"; Flags: nowait postinstall skipifsilent runascurrentuser
'''
    os.makedirs("installer_output", exist_ok=True)
    with open("installer.iss", "w", encoding="utf-8") as f:
        f.write(iss_content)


def main():
    print("Creating installer configuration...")

    # Create output directory
    os.makedirs("installer_output", exist_ok=True)

    # Create Inno Setup script (more common)
    create_simple_installer()

    # Create NSIS script (alternative)
    create_nsis_template()

    # Check if Inno Setup is installed
    inno_paths = [
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
    ]

    inno_exe = None
    for path in inno_paths:
        if os.path.exists(path):
            inno_exe = path
            break

    if inno_exe:
        print(f"Found Inno Setup: {inno_exe}")
        print("Building installer...")
        result = subprocess.run([inno_exe, "installer.iss"], capture_output=True, text=True)
        if result.returncode == 0:
            print("Installer built successfully!")
            print("Output: installer_output/PortManager_Setup.exe")
        else:
            print(f"Build failed: {result.stderr}")
    else:
        print("Inno Setup not found, installer.iss config file generated")
        print("Install Inno Setup 6 and run: ISCC.exe installer.iss")
        print("Download: https://jrsoftware.org/isdl.php")

        # Try using NSIS
        nsis_paths = [
            r"C:\Program Files (x86)\NSIS\makensis.exe",
            r"C:\Program Files\NSIS\makensis.exe",
        ]

        for path in nsis_paths:
            if os.path.exists(path):
                print(f"\nFound NSIS: {path}")
                print("Building with NSIS...")
                result = subprocess.run([path, "installer.nsi"], capture_output=True, text=True)
                if result.returncode == 0:
                    print("NSIS installer built successfully!")
                break


if __name__ == "__main__":
    main()
