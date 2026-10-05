README.TXT for Corporate Benefits Sample [Visual C++ Servers Component]
Visual Studio '97

What Are the Different Files for?
----------------------------------
bencpp.dll  -	In-process server
bencpp.tlb  - 	Type library for in-process server
bencppr.exe  -	Out-of-process server
bencppr.tlb  - 	Type library for out-of-process server

How Do I Install the Servers?
------------------------------
Run setup.exe from the benefits\clients\vbclient\install directory.


How Do I Configure the Servers for DCOM Operation?
---------------------------------------------------
On the Client machine:
From RACManager, select Remote Registration and check the DCOM check box for all objects from the BenefitsCPPRemote server. Enter the appropriate machine names in the network address and select the network protocol.

On the Remote machine:
If the server is an intermediate link with a connection to another remote server, select Remote Registration for all objects from BenefitsCPPRemote. If the server is the terminating remote server, select Local Registration for all objects from BenefitsCPPRemote. 


How Do I Configure the Servers for Remote Automation Operation?
----------------------------------------------------------------
Configuration for Remote Automation operation is similar to DCOM configuration except for the choice of Remote Automation as the transport option. Remote Automation also requires the Automation manager to be running on all machines acting as remote servers.

How Do I Build the Servers?
----------------------------
The workspace files for the in- and out-of-process servers are bencpp.dsw and bencppr.dsw respectively.  The "r" post-fix notation is used throughout the project to designate "remote" or out-of-process versions of the COM APIs.
