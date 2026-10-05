# Microsoft Developer Studio Project File - Name="BenefitsCPPr" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 5.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=BenefitsCPPr - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "BenefitsCPPr.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "BenefitsCPPr.mak" CFG="BenefitsCPPr - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "BenefitsCPPr - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "BenefitsCPPr - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "BenefitsCPPr - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir ".."
# PROP Intermediate_Dir ""
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /YX /FD /c
# ADD CPP /nologo /W3 /Gi /GX /Od /Gy /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "UNICODE" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /incremental:yes /machine:I386 /out:"bencppr.exe"
# Begin Special Build Tool
SOURCE=$(InputPath)
PostBuild_Cmds=copy bencppr.tlb ..\bencppr.tlb
# End Special Build Tool

!ELSEIF  "$(CFG)" == "BenefitsCPPr - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Benefits"
# PROP BASE Intermediate_Dir "Benefits"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ".."
# PROP Intermediate_Dir ""
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /YX /FD /c
# ADD CPP /nologo /W3 /Gm /Gi /GX /Zi /Od /Gy /D "_DEBUG" /D "WIN32" /D "_WINDOWS" /D "UNICODE" /YX /FD /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo /o"BenefitsCPPr.bsc"
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /pdb:"BenCPPr.pdb" /debug /machine:I386 /out:"..\bencppr.exe" /pdbtype:sept
# SUBTRACT LINK32 /pdb:none /incremental:no
# Begin Special Build Tool
SOURCE=$(InputPath)
PostBuild_Cmds=copy bencppr.tlb ..\bencppr.tlb
# End Special Build Tool

!ENDIF 

# Begin Target

# Name "BenefitsCPPr - Win32 Release"
# Name "BenefitsCPPr - Win32 Debug"
# Begin Source File

SOURCE=.\benefitsr.idl

!IF  "$(CFG)" == "BenefitsCPPr - Win32 Release"

# Begin Custom Build
InputPath=.\benefitsr.idl

BuildCmds= \
	midl benefitsr.idl /header benefitsr.h /tlb bencppr.tlb

"bencppr.tlb" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"benefitsr.h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"benefitsr_i.c" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "BenefitsCPPr - Win32 Debug"

# PROP Output_Dir "c:\work\projects\benefits\servers\cpp"
# Begin Custom Build
InputPath=.\benefitsr.idl

BuildCmds= \
	midl benefitsr.idl /header benefitsr.h /tlb bencppr.tlb

"bencppr.tlb" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"benefitsr.h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"benefitsr_i.c" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\benefitsr.rc

!IF  "$(CFG)" == "BenefitsCPPr - Win32 Release"

!ELSEIF  "$(CFG)" == "BenefitsCPPr - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\benefitsr_i.c
# End Source File
# Begin Source File

SOURCE=.\benexe.cpp
# End Source File
# Begin Source File

SOURCE=.\cadmin.cpp
# End Source File
# Begin Source File

SOURCE=.\cbank.cpp
# End Source File
# Begin Source File

SOURCE=.\cbenefit.cpp
# End Source File
# Begin Source File

SOURCE=.\cbenefitlink.cpp
# End Source File
# Begin Source File

SOURCE=.\cdependent.cpp
# End Source File
# Begin Source File

SOURCE=.\cemployee.cpp
# End Source File
# Begin Source File

SOURCE=.\cienum.cpp
# End Source File
# Begin Source File

SOURCE=.\ckeyll.cpp
# End Source File
# Begin Source File

SOURCE=.\clist.cpp
# End Source File
# Begin Source File

SOURCE=.\clistbank.cpp
# End Source File
# Begin Source File

SOURCE=.\clistbenefit.cpp
# End Source File
# Begin Source File

SOURCE=.\clistbenefitlink.cpp
# End Source File
# Begin Source File

SOURCE=.\clistdependant.cpp
# End Source File
# Begin Source File

SOURCE=.\clistemployee.cpp
# End Source File
# Begin Source File

SOURCE=.\csaferay.cpp
# End Source File
# Begin Source File

SOURCE=.\cstrfunc.cpp
# End Source File
# Begin Source File

SOURCE="c:\Program Files\DevStudio\Vc\lib\Ole32.lib"
# End Source File
# Begin Source File

SOURCE=.\support.cpp
# End Source File
# End Target
# End Project
