REM compile and create the .net resource file

@echo on

REM SETLOCAL

REM if exist X:\Bldtools\version.txt (
  REM --- Get build version number ---
REM  FOR /F %%i in (X:\Bldtools\version.txt) do set BldVersion=%%i
REM )
REM if defined BldVersion (
REM  set BldVersionDelimited=%BldVersion:~0,2%.%BldVersion:~2,2%.%BldVersion:~4,2%.%BldVersion:~6,4%
REM ) else (
REM  set BldVersionDelimited=00.00.00.0001
REM )
REM Echo BldVersionDelimited=%BldVersionDelimited%

Rem *** Required for ResGen.exe and AL.exe ***
REM Echo "Using Visual Studio 2010 variables from: %VS100COMNTOOLS%vsvars32.bat"
REM call "%VS100COMNTOOLS%vsvars32.bat"

Rem *** Set paths for ResGen.exe and AL.exe ***
REM set resgpath=%WindowsSdkDir%bin\NETFX 4.0 Tools\
REM set alpath=%WindowsSdkDir%bin\NETFX 4.0 Tools\

REM if not exist X:\Container\Bin\Assemblies\Debug\en-US mkdir X:\Container\Bin\Assemblies\Debug\en-US
REM if not exist X:\Container\Bin\Assemblies\Release\en-US mkdir X:\Container\Bin\Assemblies\Release\en-US

REM "%resgpath%resgen.exe" M:\SharedContent\Src\RefData\Rules\SectionLibraryCalculator\en-US\SectionLibraryCalculator.resx X:\Container\Bin\Assemblies\Debug\en-US\SectionLibraryCalculator.resources 
REM "%alpath%al.exe" /embed:X:\Container\Bin\Assemblies\Debug\en-US\SectionLibraryCalculator.resources,SectionLibraryCalculator.resources /out:X:\Container\Bin\Assemblies\Debug\en-US\SectionLibraryCalculator.resources.dll /c:en-US /fileversion:%BldVersionDelimited%

REM copy X:\Container\Bin\Assemblies\Debug\en-US\SectionLibraryCalculator.resources X:\Container\Bin\Assemblies\Release\en-US\SectionLibraryCalculator.resources
REM copy X:\Container\Bin\Assemblies\Debug\en-US\SectionLibraryCalculator.resources.dll X:\Container\Bin\Assemblies\Release\en-US\SectionLibraryCalculator.resources.dll

REM ENDLOCAL
