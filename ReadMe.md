# Geology Profile Tool

ESRI ArcGIS Pro 2.x Add-In which produces various data outputs for the generation of geologic profiles which can be used to create digital cross sections.
***
### Directions for Use: 
**Installation**
1.	Download the .esriAddinX file listed above. 
2.	Open ArcGIS Pro and either start a new project or open an existing one. 
2.	Click the blue ‘Project’ button. 
3.	Click ‘Add-In Manager’ in the menu.
4.	Click ‘Options’ tab and ‘Add Folder’ button. 
5.	Navigate to the folder location of the .esriAddinX file.
![](images/Add-In_Manager.PNG)

**Running the Tool**
6.	Go back to the main window displaying the map, etc. where there should be a tab called “KGS Tools” as shown below.
![](images/Toolbar.PNG)

7.	To use the Add-In, click the “Build Profile” button. 
8.	Once the sketch tool (cursor will be crosshairs and the tool menu shown in Fig. 3 should appear) is fully enabled, click any two points on the shapefile or feature class being used for geologic map unit information and approve the sketch once satisfied with the line. 
![](images/sketch_toolbar.PNG)

9.	Fill in the form(s) options and click “Create Profile” button on the form to run the tool. Detailed progress bar information should appear while the tool runs. If errors occur which impede the completion of the tool, message boxes detailing the error will open. If the tool successfully runs to completion, a message box will open informing the user of the location of the files created. 
![](images/Pg1.PNG)
![](images/Pg2.PNG)
![](images/Pg3.PNG)
![](images/Pg4.PNG)

10.	To disable the Sketch tool upon completion of the tool, click the “Explore” button under the ‘Map’ tab, shown below.
![](images/Explore.PNG)

**Credits**

Tool Originally created 24Jul07 by: James L. Poelstra
Email: james.poelstra@umontana.edu
This code is directly based on work by:
Michael Moex Maxelon, Geologisches Institut, ETH Zentrum,8092 Zürich, Switzerland
Dec 14 2004 http://e-collection.ethbib.ethz.ch/show?type=bericht&nr=377
With additional code contributions by: Chris Rae's VBA Code 19/5/99 Archive - http://chrisrae.com/vba 

Originally written in VBA
Converted to VB.net for use in ArcMap 10.x (version 2.0), April 2015 by: Kristen Jordan- Koenig, Kansas Data Access and Support Center	Email: kristen@kgs.ku.edu
Update for ArcGIS Pro (version 3.0) performed by Emily Bunse, GRA Cartographic Services (Aug. 2017-May 2018), Email: egbunse@gmail.com