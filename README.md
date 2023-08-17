# AutoMID

A tool for checking the insurance status of multiple vehicles via [the UK's Motor Insurance Database website](https://ownvehicle.askMID.com).  This is intended to be an easy way for individuals and small business to check the statuses of their own vehicles.

# Prerequisites

This script uses AutoHotKey version 2.  Please download and install AutoHotKey version 2 from the [AutoHotKey website](https://www.autohotkey.com/).  Microsoft Excel must also be installed.

# Usage

With AutoHotKey installed, double-click on the AutoMID.ahk file to lauch the script.  The user will be prompted to select an Excel spreadsheet containing the vehicle data.  The script will parse the spreadsheet, launch the askMID website, and attempt to enter the vehicle data into the format.

Once the script completes, it will launch a report of what it was able to find.  This report is saved on the user's desktop.

## Be Patient

Web pages take time to load.  Sometimes, the ponder for a bit over the question of whether or not you are a robot.  In order to allow for uncertainties like these, the script spends a great deal of time waiting to ensure that the web page is ready to proceed.  For the most part, just wait and the script will go about its merry way. Things work best if you refrain from using the mouse or keyboard while the script is running.

## Captchas

When challenged, the script cannot prove that it's not a robot (because it **is** a robot).  If a Captcha prompt appears, the script will (hopefully) pause and wait for it to be completed.  Complete the Captcha so that the "I'm not a robot" box is checked and then wait.

## The "Check this vehicle" button

Unless the script asks you to do so, do not click the "Check this vehicle" button.  This may cause the script to be out of sync with the active element on the webpage.

# Troubleshooting

The declarations of which columns contain each vehicle's registration number, make, and model are set as variables at the beginning of the AutoMID.ahk file.  If the script does not properly parse your vehicle spreadsheet, open the AutoMID.ahk file using a plain text editor (such as Notepad.exe) and adjust the column variables as needed to fit your data.

Sometimes, the script will accidentally "click" a link on the page.  If this happens, close the page that is launched.  Do not close the askMID website.

# Disclaimer

No warranties or guarantees are included with this script.  Use entirely at your own risk.  While care was taken in its writing, it is a hobby project.  Also note that this script depends heavily on the design of the askMID website at the time that it was written.  Changes to the website's content or formatting may cause the script to fail.

# License

This script is released under the Creative Commons "CC0 1.0 Universal" license.
https://creativecommons.org/publicdomain/zero/1.0
