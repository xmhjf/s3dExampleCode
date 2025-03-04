VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImportAssembly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation. All rights reserved.
'
'  Project     : Planning\Data\Rules\XLSImport\
'  File        : ImportAssembly.cls
'
'  Description :
'
'  History     :
'   6th Feb 2012      Siva     Initial creation
'**************************************************************************************

Option Explicit
Private Const MODULE = "ImportAssembly"

Implements IJPlnXLSImport

Private Sub IJPlnXLSImport_Execute(ByVal eAction As GSCADPlnIntHelper.PlanningXLSActionType, ByVal bstrWorkBook As String, ByVal bstrWorkSheet As String, _
                                   ByVal bstrLogFile As String, Optional ByVal bstrAttribute1 As String = "", Optional ByVal bstrAttribute2 As String = "", _
                                   Optional ByVal bstrAttribute3 As String = "", Optional ByVal pRootAssembly As Object)
    Const METHOD = "IJPlnXLSImport_Execute"
    On Error GoTo ErrorHandler
    
    ' Check the inputs
    If Not eAction = IMPORT_ASSEMBLY Or bstrWorkBook = "" Or bstrWorkSheet = "" Then
        Err.Description = "Invaild inputs to the rule"
        Exit Sub
    End If
    
    ' Call the method to read the work sheet and create the assebly hierarchy
    ImportAssemblyHierarchy bstrWorkBook, bstrWorkSheet, bstrLogFile, bstrAttribute1, pRootAssembly
        
cleanup:
    
Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
