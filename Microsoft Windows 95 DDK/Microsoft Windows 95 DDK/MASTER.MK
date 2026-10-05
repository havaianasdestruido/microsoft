#******************************************************************************
#
# This is the environment file for building all the DDK samples.  If the
# following environment is setup on a machine you can build any sample in
# the DDK.	The documentation for each item is provided below.
#
# Created
#
#	 06/01/96
#
#******************************************************************************

#
# You may or may not want to use the following environment variables.  If you
# do simply set the variables as instructed below.	If not, you'll need to
# modify the path, include and lib marco's below.
#
#MASM_ROOT=
#C16_ROOT=
#C32_ROOT=
#SDKROOT=
#DDKROOT=

!ifndef MASM_ROOT
!MESSAGE >
!MESSAGE > Must have MASM_ROOT set to the base location where MASM is installed.
!MESSAGE > (ie. SET MASM_ROOT=c:\masm611) This variable is case sensitive.
!MESSAGE >
!ERROR > MASM_ROOT
!endif
!ifndef C16_ROOT
!MESSAGE >
!MESSAGE > Must have C16_ROOT set to the base location where the C16 compiler can
!MESSAGE > be found.  The expected directory structure expects to find a C16\BIN,
!MESSAGE > C16\INCLUDE and C16\LIB. (ie. c:\msvc) This variable is case sensitive.
!MESSAGE >
!ERROR > C16_ROOT
!endif
!ifndef C32_ROOT
!MESSAGE >
!MESSAGE > Must have C32_ROOT set to the base location where the C32 compiler can
!MESSAGE > be found.  The expected directory structure expects to find a C32\BIN,
!MESSAGE > C32\INCLUDE and C32\LIB. (ie. c:\msdev) This variable is case sensitive.
!MESSAGE >
!ERROR > C32_ROOT
!endif
!ifndef SDKROOT
!MESSAGE >
!MESSAGE > Must have SDKROOT set to the base location where the Windows SDK is
!MESSAGE > installed. (ie c:\mstools) This variable is case sensitive.
!ERROR > SDKROOT
!endif
!ifndef DDKROOT
!MESSAGE >
!MESSAGE > Must have DDKROOT set to the base location where the Windwos DDK is installed.
!MESSAGE > (ie c:\ddk) This variable is case sensitive.
!MESSAGE >
!ERROR > DDKROOT
!endif

!ifndef TMP
!ERROR > The TMP environment variable is not set.  It must be.
!endif
!ifndef TEMP
!ERROR > The TEMP environment variable is not set.	It must be.
!endif

!ifdef DIRS
#
# This target is used for a recursive build.
#
all: $(DIRS)

clean: all

$(DIRS)::
	  @cd $@
	  @$(CMDLINE)
	  @cd ..

!else

!ifndef BUILD_TYPE
#
# Each class of samples might need specific header files, thus we special
# case this.
#
!ERROR > Must set BUILD_TYPE in local makefile.
!endif

!if [set include=;]
!endif
!if [set path=;]
!endif
!if [set lib=;]
!endif

# setup masm environment for both 16 and 32 bit
PATH		= $(MASM_ROOT)\BIN;\
			  $(MASM_ROOT)\BINR;\
			  $(PATH)

LIB 		= $(MASM_ROOT)\LIB;\
			  $(LIB)

INCLUDE 	= $(MASM_ROOT)\include;\
			  $(INCLUDE)

!if "$(BUILD_BITS)" == "16"
#
# For the 16 and 32 bit build environments, we must be able to search the tool
# directories in a specific order.	This order is listed here and in the
# BUILD_BITS section below.
#
# Note: the default PATH is assumed to be something like:
#
#	 PATH=c:windows;c:\windows\command;
#
# If it contains some of these directories you may experiance building
# problems.
#
# Also, if you do not wish to use the Global environment variables - MASM_ROOT,
# C16_ROOT, C32_ROOT, SDKROOT and DDKROOT you must explicitly define the
# appripriate directories here.  (Keep in mind that some samples require
# DDKROOT to be set).
#

PATH		= $(DDKROOT)\$(BUILD_TYPE)\bin;\
			  $(DDKROOT)\bin;\
			  $(SDKROOT)\binw16;\
			  $(SDKROOT)\bin;\
			  $(C16_ROOT)\bin;\
			  $(PATH);

LIB 		= $(DDKROOT)\lib;\
			  $(SDKROOT)\lib16;\
			  $(C16_ROOT)\LIB;\
			  $(C16_ROOT)\MFC\LIB;\
			  $(LIB)

INCLUDE 	= $(DDKROOT)\$(BUILD_TYPE)\inc;\
			  $(DDKROOT)\inc16;\
			  $(SDKROOT)\inc16;\
			  $(C16_ROOT)\include;\
			  $(C16_ROOT)\MFC\include;\
			  $(INCLUDE);\
			  $(DDKROOT)\inc32;\
			  $(SDKROOT)\include;

!elseif "$(BUILD_BITS)" == "32"
#
# This is the order we need to get for 32 bit apps
#
PATH		= $(DDKROOT)\$(BUILD_TYPE)\bin;\
			  $(DDKROOT)\bin;\
			  $(SDKROOT)\bin;\
			  $(C32_ROOT)\BIN;\
			  $(C32_ROOT)\BIN\WIN95;\
			  $(PATH);

INCLUDE 	= $(DDKROOT)\$(BUILD_TYPE)\inc;\
			  $(DDKROOT)\inc32;\
			  $(SDKROOT)\include;\
			  $(C32_ROOT)\include;\
			  $(C32_ROOT)\MFC\include;\
			  $(INCLUDE);\
			  $(DDKROOT)\inc16;\
			  $(SDKROOT)\inc16;

LIB 		= $(DDKROOT)\lib;\
			  $(SDKROOT)\lib;\
			  $(C32_ROOT)\LIB;\
			  $(C32_ROOT)\MFC\LIB;\
			  $(LIB)

!endif
#
# removeextra characters
#
PATH		= $(PATH:;;=;)
PATH		= $(PATH: =)

LIB 		= $(LIB:;;=;)
LIB 		= $(LIB: =)

INCLUDE 	= $(INCLUDE:;;=;)
INCLUDE 	= $(INCLUDE: =)

!ifndef BUILD_BITS
SET_BUILD_BITS_32:
	-@SET PATH=$(DDKROOT)\$(BUILD_TYPE)\bin;$(DDKROOT)\bin;$(SDKROOT)\bin;$(C32_ROOT)\BIN;$(C32_ROOT)\BIN\WIN95;$(PATH);
	-@SET INCLUDE=$(DDKROOT)\$(BUILD_TYPE)\inc;$(DDKROOT)\inc32;$(SDKROOT)\include;$(C32_ROOT)\include;$(C32_ROOT)\MFC\include;$(INCLUDE);$(DDKROOT)\inc16;$(SDKROOT)\inc16;
	-@SET LIB=$(DDKROOT)\lib;$(SDKROOT)\lib;$(C32_ROOT)\LIB;$(C32_ROOT)\MFC\LIB;$(LIB)
	-@echo 32 bit env set

SET_BUILD_BITS_16:
	-@SET PATH=$(DDKROOT)\$(BUILD_TYPE)\bin;$(DDKROOT)\bin;$(SDKROOT)\binw16;$(SDKROOT)\bin;$(C16_ROOT)\bin;$(PATH);
	-@SET LIB=$(DDKROOT)\lib;$(SDKROOT)\lib16;$(C16_ROOT)\LIB;$(C16_ROOT)\MFC\LIB;$(LIB)
	-@SET INCLUDE=$(DDKROOT)\$(BUILD_TYPE)\inc;$(DDKROOT)\inc16;$(SDKROOT)\inc16;$(C16_ROOT)\include;$(C16_ROOT)\MFC\include;$(INCLUDE);$(DDKROOT)\inc32;$(SDKROOT)\include;
	-@echo 16 bit env set
!endif

!ifdef SHOWALL
#
# If you experiance build problems within a particular sample try setting
# SHOWALL=1 and verify that the environment is set up correctly.
#
!MESSAGE > include=$(INCLUDE)
!MESSAGE > lib=$(LIB)
!MESSAGE > path=$(PATH)
!endif

!endif #ends the DIRS conditional
