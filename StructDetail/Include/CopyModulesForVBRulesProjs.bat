@echo off
:-------------------------------------------------------------------------------------------
: File:
:     CopyModulesforVBRulesProjs.bat
:
: Description:
:     This batch script is used during builds to copy standard modules found
:     outside an app's folder tree into a location within the app's folder tree.  This
:     is done for VB rules projects which must be able to be opened and compiled
:     on both development and end-user machines.
:
:     If an included module is on a different drive than the VB rule project
:     (typical on a development machine), then the absolute path to the module
:     is coded in .vbp file.  On an end-user machine, the module won't be found
:     and an error will occur when opening the VB rule project.  If the included
:     module is within the app's folder tree, then the relative path is coded in the
:     .vbp file and the module will be found when opening the rule project on
:     either a development or end-user machine.
:
:     Rather than maintaining a permanent copy of common modules (e.g.,
:     X:\Container\Include\CoreTraderKeys.bas) within several apps, a rule
:     project .vbp file maintains a permanent reference to a "local" copy, but the
:     local copy is created at the time of the app's build.
:-------------------------------------------------------------------------------------------

: StructDetail
set DestDir=S:\StructDetail\Data\Include\


set SourceDir=X:\Container\Include\

set ModuleFile=CoreTraderKeys.bas
copy %SourceDir%%ModuleFile% %DestDir%%ModuleFile% /y 
set ModuleFile=winerror.bas
copy %SourceDir%%ModuleFile% %DestDir%%ModuleFile% /y 


set SourceDir=X:\Shared\Include\

set ModuleFile=MCoreRegistry.bas
copy %SourceDir%%ModuleFile% %DestDir%%ModuleFile% /y 
set ModuleFile=MVbContants.bas
copy %SourceDir%%ModuleFile% %DestDir%%ModuleFile% /y 


set SourceDir=M:\CommonShip\Client\Include\

set ModuleFile=ClientFunctions.bas
copy %SourceDir%%ModuleFile% %DestDir%%ModuleFile% /y
set ModuleFile=PropertyPageHelper2.bas
copy %SourceDir%%ModuleFile% %DestDir%%ModuleFile% /y
