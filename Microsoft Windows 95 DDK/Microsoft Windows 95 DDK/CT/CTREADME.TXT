CTREADME.TXT - WHQL Compatibility Tests README

WHQL distributes Windows 95 compatibility tests (CTs) and instructions for using them. Third-party engineers use the CTs to test hardware devices and drivers at their site before submitting them to WHQL for compatibility testing. All the compatibility tests are used by WHQL during the testing process at Microsoft Corporation.

The CTs are not included on this release of the Windows 95 DDK, and you must *not* use any of the CTs distributed on either of the two previous releases of the Windows 95 DDK.

Where To Get the Latest CTs 
===========================
The current CTs and documentation kits are now distributed on the Internet, at two sites:

-- An FTP site (ftp://ftp.microsoft.com/services/whql).
-- A Web site (http://www.microsoft.com/windows/thirdparty/whql.htm).

The files available for downloading are always the same on both sites. The most recent version of the compatibility tests for different device classes, instructions for using the tests, device class-specific Setup programs, and any special test tools you need to carry out the tests are in compressed files on both sites. 

Just to give you an idea of the CTs available on the Internet sites, when this version of the Windows 95 DDK was released, compatibility tests were on the sites for the following device classes:

-- Audio adapters.
-- Display adapters and monitors.
-- Gaming devices.
-- Input devices.
-- ISDN and WAN.
-- Keyboards.
-- Modems.
-- Network LAN.
-- Printers.
-- Storage devices and adapters.
-- System (complete platforms).
-- Video adapters.

Other Sources of Information
============================  
You can obtain step-by-step instructions for how to proceed toward WHQL certification of your device without accessing the Internet by calling the U.S. number 206-635-2222. The instructions on how to proceed will be automatically faxed to you. 

If you have any questions about WHQL's program or the compatibility tests after looking at the FTP sites and/or reading the faxed-back instructions, e-mail whqlinfo@microsoft.com or fax your questions to 206-703-3872.

Contact Microsoft Developer Support Services to answer technical questions about device driver development, not WHQL. For example, sites that have a support contract with Microsoft, such as a Premier, Priority, or MSDN Level II support contract, should call their regular support number and ask for DDK support for OEMs and IHVs. After you have submitted your device and device driver for WHQL testing, WHQL test engineers will take the initiative to contact you about specific problems your device is having with the WHQL tests.

If your site does not have a support contract with Microsoft, or for answers to general questions about the Microsoft support options, telephone 1-800-936-3500.

DDK Documentation Errors About the CTs
======================================
If the documentation for this release of the Windows 95 DDK states that the WHQL tests and tools are included in the DDK, that is not true. Download the compatibility tests, test documentation, and test tools from one of the Internet sites. That way, you know you are always getting the most recent versions. 

If the DDK documentation states that you must run a Setup program called Ctsetup.exe from the \Ct directory of the DDK to install the compatibility tests and tools, that is not true. Ctsetup.exe is not included in this version of the DDK. When you download the compatibility tests for a device class from one of the Internet sites, a device class-specific Setup program is downloaded along with the tests, test instructions, and tools. Run that Setup program to install the tests and tools for your device class. 

If the DDK documentation states that Ct setup instructions are on the DDK, this is not true.

