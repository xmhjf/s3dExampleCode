Attribute VB_Name = "HgrSupSymbolModule"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   HgrSupSymbolModule.bas
'   ProgID:         
'   Author:         Sundar
'   Creation Date:  14.Jun.2002
'   Description:
'    
'
'   Change History:
'	14.Jun.2002     Sundar            Creation Date
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

'   -------------------------------------------
'   Support Symbols related module information.
'   -------------------------------------------


' This enum is created to support Symbols wizard. Based on the representation type
' ports will be added to symbol.
Public Enum RepresentationType
    DetailedRepresentation
    SymbolicRepresentation
    MaintenanceRepresentation
End Enum



