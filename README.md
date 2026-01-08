# BCFtoACADstudy (command: BCFpanel)
Use at own risk: Exporting BCF from ACC and making use of the data in AutoCAD.
This is a study on how to use BCF exported from ACC in AutoCAD.
With this you can load a BCF file and display the issues information, also the thumbnail image.
Zooming into the right place only works if you are in the file that contains the selected object.
Acutally the zooming part of the code does not really work, but a selection is made using the "selected component" (see later), so with manually executing _zoom command (and choosing "object" option) you are getting closer to it.
The code is mostly ai generated and there is some deactivated part that was trying to create the view based on the camera information from the BCF.
This did not work and seems to be quite complicated, that's why the active code only selects the "selected component" to signal visually where the issue is located.

How to improve the code?
Camera information in the BCF only contains camera position, view direction, camera up direction and FOV. Target is missing, but target information can be read via entity handle (IfcGuid) supplied in the same file in the BCF (zip):  
<Component IfcGuid="1952" ...
With above approach the view saved in BCF file could potentially be reproduced.
The improvement would be that you could zoom to the issue location, even you are not in the file that contains the "selected component", for example if you have a central DWG that references many DWGs. Of course there are many more improvements needed to run this as a tool, maybe most important one to supply and calculate with correction values if the issues where set in offset positions. Note that the file information is only automatically added to the issue in ACC when it is generated from a clash, that's a big shortcoming..

Finally working with ACC issues side by side with AutoCAD should work anyway nicely, the BCF reader code would be only needed if the AutoCAD user has no access to ACC.


Here is the original prompt that I used in github copilot to generate the code, as described above there were a lot modifications done on the zoom/select part:

Hi, Please write a small BCF client for AutoCAD in C#. Based on a command a resizeable panel opens with button in it for file selection (*.bcf).
Once the BCF file (it is a common zip format) is selected, the information from the file is read. The file contains folders, one per issue.
the folder contains a bcfv file, a bcf file and a png file.
After reading all this information into memory, the panel should now show the contents of the bcf files in a row, but the tags are stripped.
It is showing all issues (folders) bcf text in a row with some seperation element. At the end of the text per issue there is a button that sets the AutoCAD camera
just as it is specified in the bcfv file and at the same time it shows the png image in a seperate resizeable window (also image shall resize with the window)


