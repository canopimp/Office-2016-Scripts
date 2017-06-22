#!/bin/bash

# This script downloads and installs the latest Office OneNote 2016 for compatible Macs
# Script was pieced together by Greg Fisher with inspiration from Rich Trouton's 
# Adobe Flash installer script. Also help from Paul Bowden from Microsoft.

# Determine OS version
osvers=$(sw_vers -productVersion | awk -F. '{print $2}')

# Specify the complete address of the SKUless Office 2016 OneNote Installer
# installer package

installerURL="https://go.microsoft.com/fwlink/?linkid=820886"

# Specify name of downloaded disk image

installer_pkg="/tmp/onenoteinstaller.pkg"

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

exit 0