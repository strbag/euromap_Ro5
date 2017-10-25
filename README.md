# euromap_Ro5

--- euromap_v3.0 ---
created by the Swiss OFSP, section for Environmental Radiation


How to use:

1. Copy your Ro5 Excel sheets into the Excel Folder of this application.

2. Launch euromap_v3.0.exe and follow the instructions.

3. The KML Files can be found in the Kml Folder of this application.


What it does:

The application generates a Kml file for each Excel file you placed in the Excel folder.
The Placemarks will be colored according to their Activity concentration: 
from green (0) to red (RedLine) in 512 steps leading through yellow (0.5*RedLine).
The RedLine value can be set by the user.
Any Placemark with an Activity concentration above the RedLine value will be colored red.


The Kml Files can then be imported either into My Maps by Google https://www.google.com/maps/about/mymaps/ (Google account needed) or in Google Earth

Because the Excel files can contain more than one measurement per location,
the kml files will be generated so that the highest Concentration Values will be displayed per location.


Known Issues:

- The coordinates must be entered in Excel as decimal values (i.e. 46.712, NOT in 46,712°N)
- The implemented unit conversion for µBq, mBq and Bq does not work reliably as of now. Please enter the values in mBq in Excel!  (Otherwise the Color-code will be off)


Examples:

![alt text](https://github.com/strbag/euromap_Ro5/edit/master/Example_Map.png)
