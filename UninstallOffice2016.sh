#!/bin/bash

# source https://support.microsoft.com/en-us/office/uninstall-office-for-mac-eefa1199-5b58-43af-8a3d-b73dc1a8cae3

USERNAME=$(ls -l /dev/console | awk '{print $3}')

if [[ $EUID -ne 0 ]]; then
    echo "You must have ROOT priviledges to run this script."
    exit 1
else

    echo "Removing Office 2016 apps..."
    rm -rf "/Applications/Microsoft Excel.app"
    rm -rf "/Applications/Microsoft OneNote.app"
    rm -rf "/Applications/Microsoft Outlook.app"
    rm -rf "/Applications/Microsoft PowerPoint.app"
    rm -rf "/Applications/Microsoft Word.app"
    rm -rf "/Applications/OneDrive.app"
    # rm -rf "/Applications/Microsoft Teams.app"

    echo "Cleaning ~/$USERNAME/Library/Containers directory..."
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.errorreporting
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.Excel
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.netlib.shipassertprocess
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.Office365ServiceV2
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.Outlook
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.Powerpoint
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.RMS-XPCService
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.Word
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.onenote*
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.OneDrive*
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.SkyDriveLauncher
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.openxml.excel.app
    rm -rf /Users/$USERNAME/Library/Containers/com.microsoft.outlook.profilemanager

    rm -rf /Users/$USERNAME/Library/Application\ Support/com.microsoft.OneDrive
    rm -rf /Users/$USERNAME/Library/Application\ Support/com.microsoft.OneDriveStandaloneUpdater
    rm -rf /Users/$USERNAME/Library/Application\ Support/com.microsoft.OneDriveUpdater
    # rm -rf /Users/$USERNAME/Library/Application\ Support/com.microsoft.teams

    echo "Cleaning ALL Outlook data..."
    rm -rf "/Users/$USERNAME/Library/Group Containers/UBF8T346G9.ms"
    rm -rf "/Users/$USERNAME/Library/Group Containers/UBF8T346G9.Office"
    rm -rf "/Users/$USERNAME/Library/Group Containers/UBF8T346G9.OfficeOsfWebHost"
    rm -rf "/Users/$USERNAME/Library/Group Containers/UBF8T346G9.OfficeOneDriveSyncIntegration"
    
    echo "Cleaning /Library directories..."
    rm -rf /Library/Application Support/Microsoft/MAU2.0
    rm -rf /Library/Fonts/Microsoft
    rm -rf /Library/LaunchDaemons/com.microsoft.office.licensing.helper.plist
    rm -rf /Library/LaunchDaemons/com.microsoft.office.licensingV2.helper.plist
    rm -rf /Library/LaunchDaemons/com.microsoft.OneDriveStandaloneUpdaterDaemon.plist
    rm -rf /Library/LaunchDaemons/com.microsoft.OneDriveUpdaterDaemon.plist
    rm -rf /Library/LaunchDaemons/com.microsoft.autoupdate.helper.plist
    rm -rf /Library/Preferences/com.microsoft.Excel.plist
    rm -rf /Library/Preferences/com.microsoft.office.plist
    rm -rf /Library/Preferences/com.microsoft.office.setupassistant.plist
    rm -rf /Library/Preferences/com.microsoft.outlook.databasedaemon.plist
    rm -rf /Library/Preferences/com.microsoft.outlook.office_reminders.plist
    rm -rf /Library/Preferences/com.microsoft.Outlook.plist
    rm -rf /Library/Preferences/com.microsoft.PowerPoint.plist
    rm -rf /Library/Preferences/com.microsoft.Word.plist
    rm -rf /Library/Preferences/com.microsoft.office.licensingV2.plist
    rm -rf /Library/Preferences/com.microsoft.autoupdate2.plist
    rm -rf /Library/Preferences/com.microsoft.OneDriveUpdater
    rm -rf /Library/Preferences/com.microsoft.OneDriveStandaloneUpdater
    rm -rf /Library/Preferences/com.microsoft.OneDrive.plist
    # rm -rf /Library/Preferences/com.microsoft.teams
    # rm -rf /Library/LaunchDaemons/com.microsoft.teams.TeamsUpdaterDaemon.plist
    rm -rf /Library/Preferences/ByHost/com.microsoft
    rm -rf /Library/Receipts/Office2016_*
    rm -rf /Library/PrivilegedHelperTools/com.microsoft.office.licensing.helper
    rm -rf /Library/PrivilegedHelperTools/com.microsoft.office.licensingV2.helper

    echo "Forgetting Office 2016 packages..."
    pkgutil --forget com.microsoft.package.Fonts
    pkgutil --forget com.microsoft.package.Microsoft_AutoUpdate.app
    pkgutil --forget com.microsoft.package.DFonts
    pkgutil --forget com.microsoft.package.Microsoft_Excel.app
    pkgutil --forget com.microsoft.package.Microsoft_OneNote.app
    pkgutil --forget com.microsoft.onenote.mac
    pkgutil --forget com.microsoft.package.Microsoft_Outlook.app
    pkgutil --forget com.microsoft.package.Microsoft_PowerPoint.app
    pkgutil --forget com.microsoft.package.Microsoft_Word.app
    pkgutil --forget com.microsoft.OneDrive
    pkgutil --forget com.microsoft.package.Frameworks
    # pkgutil --forget com.microsoft.teams
    pkgutil --forget com.microsoft.package.Proofing_Tools
    pkgutil --forget com.microsoft.pkg.licensing

    echo "Done!"
fi
