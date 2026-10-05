README.TXT for Corporate Benefits Sample [Visual Basic Servers Component]
Visual Studio '97


What Are The Different Files For?
---------------------------------
benvb.dll   - In-process version of VB server
benvbr.exe  - Out-of-process version of the VB server
benvbr.tlb  - Type library for out-of-process version of the VB server


How Do I Install The Servers?
-------------------------------
Run setup.exe from the benefits\servers directory.


How Do I Configure The Servers For DCOM Operation?
------------------------------------------------
On the Client machine:
From RACManager, select Remote Registration and check the DCOM check box for all objects from the BenefitsVBRemote server. Enter the appropriate machine names in the network address and select the network protocol.

On the Remote machine:
If the server is an intermediate link with a connection to another remote server, select Remote Registration for all objects from BenefitsVBRemote. If the server is the terminating remote server, select Local Registration for all objects from BenefitsVBRemote. 


How Do I Configure The Servers For Remote Automation Operation?
---------------------------------------------------------------
Configuration for Remote Automation operation is similar to DCOM configuration except for the choice of Remote Automation as the transport option. Remote Automation also requires the Automation manager to be running on all machines acting as remote servers.
