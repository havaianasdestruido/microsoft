To build serialui.dll properly you must pay close attention to your tools and 
your build environment.  

Tools:
1. msvc 1.5x or compatible compiler
2. masm6.11 or compatible assembler
3. Win95 DDK

Remember that if you use a compatible compiler or assembler you will probably 
write your own makefile as well.

Environment:
1. run msvcvars.bat or make sure you have the PATH, INCLUDE and LIB environment
   variables set correctly
2. make sure that your assembler is included in the PATH
3. run the following batch file from the ddk with the following command line
   options:

   ddkenv 16 comm

This will make sure your environment variables and tools will be properly configured
to build serialui.dll.

This sample is provided in the ddk as an example of a property page provider, adding
proprietary property pages, and adding context sensitive help to a property page.  This
is *NOT* the source to the property page provider supplied in Windows 95 or any other 
Microsoft product.  