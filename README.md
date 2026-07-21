# Key Club Hour Automator Purpose
This is a tool meant to assist (Key Club) secretaries automate the manual process of inputing data into a large membership database table. <br> In greater depth, the Key Club Hour Automator (which will be refered to as KcHA from this point forward) takes several lists divided by member position (officer/member) then grade level provided in an officer restricted Google Sheet, then searches several sheets organized by member position (officer/member) then grade level in a seperate public Google Sheet for tracking membership hours, finally a new column is made on each sheet inside the public Google Sheet and for each name found on both Google Sheets is given hours based on the private Google Sheet.

## History
The issue KcHA wanted to solve was several complaints from my Key Club Secretaries about the time consuming process of searching & inputting volenteer attendance data into our membership database, occasionally there were some events that allowed non-members therefore it would take more time searching for a person who doesn't exist. Then began KcHA started as my hobby side project originating back in late 2024, attempting to solve the searching, formating, and overall reducing the total time it took to input these weekly event/attendance data. KcHA was then deployed in early 2025 in which the project has been, and for the forseeable future, maintained/updated by the Key Club Secretaries. A majority of the original code is found inside the "PreviousDesign" folder; however, it has been updated, refactored, and simplified into the scripts outside of "PreviousDesign" folder. 

### Two main scripts
Check our [wiki](https://github.com/todybo/Key-Club-Hour-Implementation/wiki) for a more indepth explaination of the functions of each script.
- HourMain.js <br>
This is the where the main functionality of the project is ran from, reads a list of names and searches for those names on another sheet then gives the person a set number.
	- GiveHourList.js <br>
	This script is filled with helper functions for HourMain.js
- GeneralMeetings.js <br>
This is functionally similar to HourMain.js except that it instead fills in true or false based on names found on the officer restricted sheet.

### Replicating/Contributing
KcHA was developed directly on Google App Script however ported to GitHub through the use of [clasp](https://github.com/google/clasp) (using clasp allows for local development/testing using your own IDE). This project uses two seperate Google Sheets, an officer restricted and a public member hour tracker. To create a usable copy of this project, download each .js file, changing the variable sheets ID's and prefered settings (located at the top of each script), then import each .js file into app script by pasting or by using clasp. Feel free to create branches to fix bugs or to improve the effectiveness/efficiency of this project. Check the [wiki](https://github.com/todybo/Key-Club-Hour-Implementation/wiki) for more assistance 
