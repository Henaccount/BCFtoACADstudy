# BCFtoACADstudy (command: BCFclient)
Use at own risk: Exporting BCF from ACC and making use of the data in AutoCAD.
This is a study on how to use BCF exported from ACC in AutoCAD.
With this you can load a BCF file and display the issues information, also the thumbnail image.
Zooming into the right place only works if you are in the file that contains the selected object.
Acutally the zooming part of the code does not really work, but a selection is made using the "selected component" (see later), so with manually executing _zoom command (and choosing "object" option) you are getting closer to it.
The code is mostly ai generated and there is some deactivated part that was trying to create the view based on the camera information from the BCF.
This did not work and seems to be quite complicated, that's why the active code only selects the "selected component" to signal visually where the issue is located.

How to improve the code?
Camera information in the BCF only contains camera position, view direction, camera up direction and FOV. Target is missing, but target information can be read via entity handle (IfcGuid) supplied in the same file in the BCF (zip):  
<Selection><Component IfcGuid="1952" ...
With above approach the view saved in BCF file could potentially be reproduced.
The improvement would be that you could zoom to the issue location, even you are not in the file that contains the "selected component", for example if you have a central DWG that references many DWGs. Of course there are many more improvements needed to run this as a tool, maybe most important one to supply and calculate with correction values if the issues where set in offset positions.

Finally working with ACC issues side by side with AutoCAD should work anyway nicely, the BCF reader code would be only needed if the AutoCAD user has no access to ACC.


