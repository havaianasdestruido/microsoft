README for MSDN Level 2 Release of the Windows 95 DDK
=====================================================

New Global Makefile Sets Up the Driver Build Environment
--------------------------------------------------------
This latest release of the Windows DDK contains a Makefile that 
sets up the build environment for all the included samples. This 
file is called MASTER.MK and is installed in the root location 
of the DDK. Once you have edited MASTER.MK to indicate the paths 
to the build tools on your machine, and you have set two 
environment variables, you can build any sample driver in the DDK
by simply entering the command 'nmake'. To build a sample driver, 
run nmake.exe in the appropriate directory or sub-directory (by 
entering the command 'nmake'). 

The two MS-DOS commands that set the required environment 
variables are:

   set MASTER_MAKE=1
   set DDKROOT=c:\win95ddk, where c:\win95ddk is an example of 
                            a path to the directory where the DDK 
                            is installed; substitute the path 
                            where you installed the DDK.

Note that the procedures documented in GETSTART.DOC for setting 
up the sample driver build environment are still valid, but 
using the alternative procedure documented above is easier. 
                              
Other Notes on the Driver Build Environment
-------------------------------------------
LINK.EXE and RC.EXE are tools that have both a 32-bit and a 16-bit
version. The LINK.EXE that is provided with MSVC20 (or later versions)
is a 32-bit tool. When building VxDs (32-bit) care must be
taken to not have an executable named LINK.EXE on the path ahead of
the MSVC20 LINK.EXE. When adding version resources to a VxD one must 
ensure that the 16-bit RC.EXE (from BINW16) is directly referenced, 
as VxDs are built in a 32-bit environment (which typically contains 
a 32-bit RC.EXE).

Include and lib paths should be checked to ensure that they do not
contain invalid duplicate copies of include files and libraries that 
the driver or test application may be including.

Before you install the DDK, you should have installed MASM, 
version 6.11. You must upgrade your MASM installation by copying all
the files from the \Masm611c\Bin directory on the DDK CD-ROM to 
your MASM_ROOT directory. These files are necessary for building
VxDs for Windows 95.

Installing the DDK Debug Components
-----------------------------------
The debug components supplied by this DDK are from the August 1995
release of Windows 95. If your machine has any later drivers
installed (either through updates acquired by OEMs or through 
new retail releases), overlaying these drivers could cause
unpredictable results and/or prevent your system from booting.
Therefore, it is STRONGLY recommended that you back up your
system before installing these components.

Documentation Errors
--------------------
The DDK reference documentation incorrectly documents the 
PwdChangePassword function. The correct function definition is
in the Pwdspi.h file distributed with this release of the DDK.

In GETSTART.DOC, it is stated that the Windows Hardware Quality
Labs (WHQL) compatibility tests are distributed on the DDK
CD-ROM and can be installed by running Ctsetup.exe from the 
\Ct directory on the CD-ROM. This is not true. For the correct
information, see the topic "Driver Compatibility Tests Are
Distributed on the Internet" in this Readme.txt file.

Driver Compatibility Tests Are Distributed on the Internet
----------------------------------------------------------
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