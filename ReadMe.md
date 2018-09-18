# Geology Profile Tool

ESRI ArcGIS Pro 2.x Add-In which produces various data outputs for the generation of geologic profiles which can be used to create digital cross sections.
***
### Directions for Use: 
**Installation**
1.	Download the .esriAddinX file (this is found in the "bin" --> "Release" folder above) and save to a known location. 
2.	Open ArcGIS Pro and either start a new project or open an existing one. 
2.	Click the blue ‘Project’ button. 
3.	Click ‘Add-In Manager’ in the menu.
4.	Click ‘Options’ tab and ‘Add Folder’ button. 
5.	Navigate to the folder location of the .esriAddinX file.
6.	Go back to the main window displaying the map, etc. and confirm that a new tab called “KGS Tools" has been added to the display.

**Running the Tool**
1.	To use the Add-In, click the “Build Profile” button. 
2.	Once the sketch tool (cursor will be crosshairs) is fully enabled, click any two points on the shapefile or feature class being used for geologic map unit information and approve the sketch once satisfied with the line. 
3.	Fill in the form(s) options and click “Create Profile” button on the form to run the tool. Detailed progress bar information should appear while the tool runs. If errors occur which impede the completion of the tool, message boxes detailing the error will open. If the tool successfully runs to completion, a message box will open informing the user of the location of the files created. 
4.	To disable the Sketch tool upon completion of the tool, click the “Explore” button under the ‘Map’ tab, shown below.

For further questions about running the tool and the products produced, consult "Profile Tool v1.2.1.pdf." 

**Credits**

* Tool Originally created 24Jul07 by: James L. Poelstra, Email: james.poelstra@umontana.edu
* This code is directly based on work by: Michael Moex Maxelon, Geologisches Institut, ETH Zentrum,8092 Zürich, Switzerland
Dec 14 2004 http://e-collection.ethbib.ethz.ch/show?type=bericht&nr=377
* With additional code contributions by: Chris Rae's VBA Code 19/5/99 Archive - http://chrisrae.com/vba 

* Originally written in VBA
* Converted to VB.net for use in ArcMap 10.x (version 2.0), April 2015 by: 
Kristen Jordan- Koenig, Kansas Data Access and Support Center, Kansas Geological Survey	Email: kristen@kgs.ku.edu
* Update for ArcGIS Pro (version 3.0) performed by Emily Bunse, GRA Cartographic Services, Kansas Geological Survey (Aug. 2017-May 2018), Email: egbunse@gmail.com


