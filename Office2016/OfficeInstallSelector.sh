#!/bin/sh
###################################################
### Created by: Greg Fisher						###
### Created on: 5-30-2017						###
### Modified on: 5-30-2017                      ###
### This script creates a popup dialog box		###
### with a drop down menu of all the			###
### Individual Office 2016 Product Installers. 	###
### This script relies on JAMF Pro management	###
### system to run and properly execute.			###
###################################################

## Make list of Office Products and URLs
writelist(){
/bin/cat <<EOF > /private/tmp/Installers.csv
Office 2016 Suite,https://go.microsoft.com/fwlink/?linkid=525133
Word 2016,https://go.microsoft.com/fwlink/?linkid=525134
Excel 2016,https://go.microsoft.com/fwlink/?linkid=525135
PowerPoint 2016,https://go.microsoft.com/fwlink/?linkid=525136
Outlook 2016,https://go.microsoft.com/fwlink/?linkid=525137
OneNote 2016,https://go.microsoft.com/fwlink/?linkid=820886
Skype for Business,https://go.microsoft.com/fwlink/?linkid=832978
EOF
}

## Make CofC Icon .b64 file
picture(){
/bin/cat <<EOF >/private/tmp/microsofticon.b64
iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAAAXNSR0IArs4c6QAAAAlwSFlzAAALEwAACxMBAJqcGAAAAsBJREFUeAHtmDFBg1EYxPgxgAiGGqgRWPCCoy6w1QAqGBCBAYqDdL/krfeWy2X6jt+359uD6D1dfg
6qe76+q3g8EoyyfQIJsL8xNkwAxLMfJsD+xtgwARDPfpgA+xtjwwRAPPthAuxvjA0TAPHshwmwvzE2TADEsx8mwP7G2DABEM9+mAD7G2PDBEA8+2EC7G+MDRMA8eyHCbC/MTZMAMSzHybA/sbYMAEQz36YAPsbY8MEQDz7YQLs
b4wNEwDx7IcJsL8xNkwAxLMfJsD+xtgwARDPfpgA+xtjwwRAPPthAuxvjA0TAPHshwmwvzE2TADEsx8mwP7G2DABEM9+mAD7G2PDBEA8+2EC7G+MDRMA8eyHCbC/MTZMAMSzHybA/sbYMAEQz36YAPsbY8MEQDz7YQLsb4wNEw
Dx7IcJsL8xNkwAxLMfJsD+xtgwARDPfpgA+xtjwwRAPPthAuxvjA0TAPHshwmwvzE2TADEUxiBCEQgAhGIQAQiEIEIRCACEYhABHYIHC8f37edOvebfL6eDvr193VS8egSSDYIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUT
gOgIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjE
wVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjEwVE4DoCLIEEIxMFROA6AiyBBCMTBUTgOgIsgQQjEwVE4DoCLIE
EIxMFROA6Aiyfw9sDjk0OpLcAAAAAElFTkSuQmCC
EOF
}
## Calling the function which writes out file
picture

## Decode .b64
/usr/bin/base64 -D -i /private/tmp/microsofticon.b64 -o /private/tmp/microsofticon.png

## Convert .png to .icns
/usr/bin/sips -s format icns /private/tmp/microsofticon.png --out /private/tmp/microsofticon.icns

#############################################################
########### USER DEFINED VARIABLES BELOW ####################
#############################################################
### Path to cocoaDialog (custom location for CofC)
cdPath="/Library/Application Support/JAMF/bin/cocoaDialog.app/Contents/MacOS/cocoaDialog"

### Path for the icon used in the cocoaDialog box (College of Charleston Logo)
iconPath="/private/tmp/microsofticon.png"

### Path for the JAMF binary
jamfbin="/usr/local/bin/jamf"

### Custom JAMF Trigger name to install cocoaDialog if not installed
custTrigger="InstallcocoaDialog"
############################################################
############################################################
writelist
## Set Internal Field Separator to end of line so deptName is displayed correctly
IFS=$'\n'

## Determine OS version
osvers=$(sw_vers -productVersion | awk -F. '{print $2}')

## Set Full Department Name to display to user in Dialog dropdown
instName=($(/usr/bin/awk -F , '{print $1}' /private/tmp/Installers.csv))

## Check for cocoaDialog and install if it has not been installed 
if [ ! -f "$cdPath" ]; then
	echo "Installing CocoaDialog..."
	$jamfbin policy -event $custTrigger
fi

## cocoadialog dropdown menu promt for user to select department
dialogAnswr=$( "$cdPath" dropdown --float --icon-file "$iconPath" --title "Microsoft Office 2016 Product Installer" --text "Select the Office 2016 Product to Install:" \
--items "${instName[@]}" --button1 "Install" --button2 "Cancel" --string-output)

## Separate out the button pressed name from the value selected from the dropdown menu.
buttonClicked=$(echo "$dialogAnswr" | awk 'NR==1{print}')
instSelected=$(echo "$dialogAnswr" | awk 'NR>1{print}')

## Pull department code from selection user made from cocoadialog
installerURL=($(/usr/bin/awk -F , "/$instSelected/{print \$2}" /private/tmp/Installers.csv))
## Set Internal Field Separator back to space
IFS=$' \t\n'

## Evaluate Button click and name machine
if [ "$buttonClicked" == "Cancel" ]; then
    echo "User pressed Cancel."
    exit
else
	echo "User selected an installer"
	## Install Office Product.
	## Specify name of downloaded disk image

	installer_pkg="/tmp/officeinstaller.pkg"

	if [[ ${osvers} -lt 10 ]]; then
	  echo "Office 2016 is not compatible on Mac OS X 10.9.5 or below."
	fi

	if [[ ${osvers} -ge 10 ]]; then
 
 	   # Download the latest Office Installer software package
  	   # Silently download, retry 5 times if error, redirect to location of installer as url
 	   # from Microsoft redirects to the actual download, output the package as named
  	   # from the URL specified above
 	   /usr/bin/curl --silent --retry 5 --location --output "$installer_pkg" "$installerURL"
    
 	   # Before installation on Mac OS X 10.10.x and later, the installer's
 	   # developer certificate is checked to see if it has been signed by
 	   # Microsofts developer certificate. Once the certificate check has been
	    # passed, the package is then installed.

	    if [[ ${installer_pkg} != "" ]]; then
	       if [[ ${osvers} -ge 10 ]]; then
	         signature_check=`/usr/sbin/pkgutil --check-signature "$installer_pkg" | awk /'Developer ID Installer/{ print $5 }'`
 	         if [[ ${signature_check} = "Microsoft" ]]; then
	           # Install Office 2016 from the installer package stored inside the disk image
	           /usr/sbin/installer -dumplog -verbose -pkg "${installer_pkg}" -target "/"
	         fi
	       fi
	    fi

	    # Remove the downloaded disk image
	    /bin/rm -rf "$installer_pkg"
	fi
fi
## Remove tmp files
rm /private/tmp/microsofticon.icns
rm /private/tmp/microsofticon.b64
rm /private/tmp/microsofticon.png
rm /private/tmp/Installers.csv
exit 0