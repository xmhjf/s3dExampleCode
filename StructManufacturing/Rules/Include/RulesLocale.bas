Attribute VB_Name = "Locale"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'  Project:
'  Abstract:    Locale.bas
'  History:
'       Bhaskar Sanga     Created    Oct 27, 2006
'******************************************************************
Option Explicit

Public Function InitializeLocalizer(strPath As String, strFileName As String) As IJLocalizer
    Const METHOD = "InitializeLocalizer"
    On Error GoTo ErrorHandler
    
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = CreateObject("IMSLocalizer.Localizer")
    
    'Dim oContext As IJContext
    'Set oContext = GetJContext()
    
    'Const CONTEXTSTRING = "OLE_SERVER"
    'Const STRMFGRULESLOCATION = "\bin\StructManufacturing\Resource\"
    Dim strDirectory As String
    strDirectory = strPath + strFileName
    'strDirectory = oContext.GetVariable(CONTEXTSTRING) & STRMFGRULESLOCATION & strFileName

    oLocalizer.Initialize strDirectory
    
    Set InitializeLocalizer = oLocalizer
    Exit Function
ErrorHandler:
   Err.Clear
End Function

Public Function IsInDebugMode() As Boolean
    On Error GoTo ErrorHandler
'     If the program is compiled, the following
'     Debug statement has been removed so it will
'     not generate an error.
    Debug.Print 1 / 0
    IsInDebugMode = False
    Exit Function
ErrorHandler:
'   We got an error so the Debug statement must
'   be working.
    IsInDebugMode = True
   Err.Clear
End Function

