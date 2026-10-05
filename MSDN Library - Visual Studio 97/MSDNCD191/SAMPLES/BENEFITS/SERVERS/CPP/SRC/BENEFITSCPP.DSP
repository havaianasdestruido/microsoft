# Microsoft Developer Studio Project File - Name="BenefitsCPP" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 5.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=BenefitsCPP - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "benefitscpp.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "benefitscpp.mak" CFG="BenefitsCPP - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "BenefitsCPP - Win32 Release" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE "BenefitsCPP - Win32 Debug" (based on\
 "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "BenefitsCPP - Win32 Release"

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
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /YX /FD /c
# ADD CPP /nologo /MT /W3 /Gi /GX /Od /D "NDEBUG" /D "_DLL" /D "WIN32" /D "_WINDOWS" /D "UNICODE" /D "_BENDLL" /YX /FD /c
# SUBTRACT CPP /Fr
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /fo"./benefits.res" /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /incremental:yes /machine:I386 /out:"..\bencpp.dll"
# Begin Special Build Tool
SOURCE=$(InputPath)
PostBuild_Desc=Copy Type Library
PostBuild_Cmds=copy bencpp.tlb ..\bencpp.tlb
# End Special Build Tool

!ELSEIF  "$(CFG)" == "BenefitsCPP - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir ".."
# PROP Intermediate_Dir ""
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /YX /FD /c
# ADD CPP /nologo /MTd /W3 /Gm /Gi /GX /Zi /Od /Gy /D "_DEBUG" /D "WIN32" /D "_WINDOWS" /D "UNICODE" /D "_BENDLL" /YX /FD /c
# SUBTRACT CPP /Fr
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:windows /dll /debug /machine:I386 /out:"..\bencpp.dll" /pdbtype:sept
# SUBTRACT LINK32 /incremental:no
# Begin Special Build Tool
SOURCE=$(InputPath)
PostBuild_Desc=Copy Type Library
PostBuild_Cmds=copy bencpp.tlb ..\bencpp.tlb
# End Special Build Tool

!ENDIF 

# Begin Target

# Name "BenefitsCPP - Win32 Release"
# Name "BenefitsCPP - Win32 Debug"
# Begin Source File

SOURCE=.\benefits.def
# End Source File
# Begin Source File

SOURCE=.\benefits.idl

!IF  "$(CFG)" == "BenefitsCPP - Win32 Release"

# Begin Custom Build - Performing MIDL Compile...
InputPath=.\benefits.idl

BuildCmds= \
	midl benefits.idl /header benefits.h /tlb bencpp.tlb

"benefits.h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"bencpp.tlb" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"benefits_i.c" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ELSEIF  "$(CFG)" == "BenefitsCPP - Win32 Debug"

# Begin Custom Build - Performing MIDL Compile...
InputPath=.\benefits.idl

BuildCmds= \
	midl benefits.idl /header benefits.h /tlb bencpp.tlb

"benefits.h" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"bencpp.tlb" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)

"benefits_i.c" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
   $(BuildCmds)
# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\benefits.rc

!IF  "$(CFG)" == "BenefitsCPP - Win32 Release"

!ELSEIF  "$(CFG)" == "BenefitsCPP - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\benefits_i.c
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

SOURCE=.\dllfuncs.cpp
# End Source File
# Begin Source File

SOURCE=.\support.cpp
# End Source File
# End Target
# End Project
