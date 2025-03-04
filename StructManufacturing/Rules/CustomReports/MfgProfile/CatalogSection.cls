VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CatalogSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        : CatalogSection.cls
''
''  Description :
''
''
''  Author      : Intergraph
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************

Implements IJDCustomReport

Private Const Module = "MfgProfileCustomReports.CatalogSection"

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
       
    Dim oProfileSectionService As IJStrMfgProfileSectionService
    Dim oXMLDom As DOMDocument
    Dim loadSuccessful As Boolean
    
    Set oProfileSectionService = New StrMfgProfileSectionService
    
    Dim strXML As String
    strXML = oProfileSectionService.GetProfileSectionData("", "", "ShipShapes", DATABASE_UNITS)
    
    Set oXMLDom = New DOMDocument
    loadSuccessful = oXMLDom.loadXML(strXML)
        
    If loadSuccessful = True Then
            oXMLDom.save strFileName
            eCustomReportStatus = StrMfgProcessFinished
    Else
            eCustomReportStatus = StrMfgErrorUnknown
    End If
    
Wrapup:
    Set oProfileSectionService = Nothing
    Set oXMLDom = Nothing
    Exit Sub
    
ErrorHandler:
    Set oProfileSectionService = Nothing
    Set oXMLDom = Nothing
    eCustomReportStatus = StrMfgErrorUnknown
    Err.Raise Err.Number, , Err.Description
End Sub

