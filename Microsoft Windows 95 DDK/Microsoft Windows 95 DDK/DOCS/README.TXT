README for April MSDN Level 2 Release of the Windows 95 DDK
===========================================================

Before beginning, please read Getstart.doc or the Getting Started topic
in the DDK help file (Ddpr.hlp). Getstart.doc and the Getting Started
topic use Ddk.hlp as the filename of the DDK help file; the correct
filename is Ddpr.hlp.

Driver Compatibility Tests
--------------------------
Unlike previous releases of the Windows 95 DDK, the \Ct subdirectory 
on the DDK distribution CD-ROM does not contain the Windows 
Hardware Quality Lab (WHQL) compatibility tests. The tests are 
available on the Internet. See \Ct\Ctreadme.txt for the latest
information about the WHQL compatibility tests, including the 
URLs of the FTP and Web sites WHQL uses to distribute the
compatibility tests.

DDK Support
-----------
Contact Microsoft Developer Support Services to answer technical questions 
about device driver development. Sites that have a support contract with 
Microsoft, such as a Premier, Priority, or MSDN Level II support contract, 
should call their regular support number and ask for DDK support for OEMs 
and IHVs.

If your site does not have a support contract with Microsoft, or for 
answers to general questions about the Microsoft support options, 
telephone 1-800-936-3500.

Driver Build Environment
------------------------
LINK.EXE and RC.EXE are tools that have both a 32-bit and a 16-bit
version. The LINK.EXE that is provided with MSVC20 (or later versions)
is a 32-bit tool. When building VxDs (32-bit) care must be
taken to not have an executable named LINK.EXE on the path ahead of
the MSVC20 LINK.EXE. Similarly, when adding version resources to a
VxD one must ensure that the 16-bit RC.EXE (from BINW16) is directly
referenced, as VxDs are built in a 32-bit environment (which typically 
contains a 32-bit RC.EXE).

Include and lib paths should be checked to ensure that they do not
contain invalid duplicate copies of include files and libraries that 
the driver or test application may be including.


