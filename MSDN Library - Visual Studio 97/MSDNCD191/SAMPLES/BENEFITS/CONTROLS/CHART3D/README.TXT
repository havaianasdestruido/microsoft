README.TXT for Corporate Benefits Sample [ActiveX 3D Chart Control]
Visual Studio '97 - 20 December 1996

Use It on the Web
---------------------------------
It's all set up and compiled into a cabinet (.cab) file for you.
Just take the test HTML page and the .cab file and copy them to your Web server.

This control uses the MSChart control, which requires a license (.lpk) file. The .cab file installs the .lpk file in the \windows\system directory along with the MSChart.ocx. You will need to add the License Manager Object to your page like this:

<OBJECT 
	CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
	<PARAM NAME="LPKPath" VALUE="BenChart3d.lpk">
</OBJECT>

The Value for the "LPKPath" is relative to the URL of the page. So, if the page is in c:\temp (for instance), the relative path for the LPKPath would be "..\windows\system\BenChart3d.lpk". The control will not work on machines that are not licensed to use the control unless you provide the .lpk file.

Change the codebase in the .alx file to point to the proper cab URL, like this:

    <OBJECT ID="chartMain"
     CLASSID="CLSID:E8D74292-3386-11D0-BADA-000000000000" STYLE="TOP:25pt;LEFT:25pt;WIDTH:350pt;HEIGHT:200pt;ZINDEX:0;"
     CODEBASE="--INSERT URL PATH HERE--benchart.cab#version=1,0,0,0">


Alternatively, you can use the ActiveX&#153; Control Pad to script this component for you.
Because it is an ActiveX control, you can program it from any control-compatible container.
