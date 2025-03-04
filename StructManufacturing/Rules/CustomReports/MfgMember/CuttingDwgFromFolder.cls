VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuttingDwg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgMemberCustomReports
''  File        : CuttingDwgFromFile.cls
''
''  Description :
''
''
''  Author      : Intergraph
''
''  History     :
''               Initial Creation   -
''**************************************************************************************

Implements IJDCustomReport

Private Const MODULE = "MfgMemberCustomReports.CuttingDwgFromFile"
Dim m_oFileSystmeObject As FileSystemObject
Dim m_strLogFile As String

Private Sub Class_Initialize()
    Set m_oFileSystmeObject = New FileSystemObject
    Dim sTemp As String
    sTemp = Environ("TEMP")

    m_strLogFile = sTemp & "\GenerateMemberCuttingDwgFromFolder.log"
    
    Open m_strLogFile For Append As #1
    
    Print #1, "Start - " & CStr(Now)
    Print #1, ""
End Sub

Private Sub Class_Terminate()
    Set m_oFileSystmeObject = Nothing
    Print #1, ""
    Print #1, "End - " & CStr(Now)
    Print #1, ""
    Close #1
End Sub

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
Const METHOD = "IJDCustomReport_Generate"
On Error GoTo ErrorHandler

'    Dim lastPos As Long
'    Dim strFolderName As String
'    If m_oFileSystmeObject.FolderExists(strFileName) = True Then
'        strFolderName = strFileName
'    Else
'
'        lastPos = InStrRev(strFileName, "\")
'
'        strFolderName = Left$(strFileName, lastPos - 1)
'
'        If m_oFileSystmeObject.FolderExists(strFolderName) = False Then
'            eCustomReportStatus = StrMfgNoFileTypeMatch
'            Print #1, strFolderName + " doesn't exist"
'            Exit Sub
'        End If
'    End If
'
'    Dim strMemberCuttingDwgProgID As String
'    strMemberCuttingDwgProgID = "MfgMemberCuttingDrawings.DXF"
'
'    Dim oMemberCuttingDwg As IGenerateDocument
'    Set oMemberCuttingDwg = CreateObject(strMemberCuttingDwgProgID)
'
'    If Not oMemberCuttingDwg Is Nothing Then
'        Dim strSharedContentPath As String
'        strSharedContentPath = GetSharedContentPath
'
'        Dim strControlInputXML As String
'        strControlInputXML = strSharedContentPath + "\" + "StructManufacturing\SMS_SCHEMA\ControlInput.xml"
'
'        If m_oFileSystmeObject.FileExists(strControlInputXML) = False Then
'            Print #1, strControlInputXML + " doesn't exist"
'            eCustomReportStatus = StrMfgProcessFinished
'            Exit Sub
'        Else
'
'            'Validate if it has CONTROLINPUTS node
'            Dim oContronInputXMLDocument As IXMLDOMDocument
'            Set oContronInputXMLDocument = New DOMDocument
'
'            Dim bSucess As Boolean
'            bSucess = oContronInputXMLDocument.Load(strControlInputXML)
'            If bSucess = False Then
'                Print #1, strControlInputXML + " is not valid XML"
'                eCustomReportStatus = StrMfgProcessFinished
'                Exit Sub
'            End If
'
'            Dim oCONTROLINPUTSNode As IXMLDOMNode
'            Set oCONTROLINPUTSNode = oContronInputXMLDocument.selectSingleNode("//CONTROLINPUTS")
'
'            If oCONTROLINPUTSNode Is Nothing Then
'                Print #1, strControlInputXML + "," + " CONTROLINPUTS node doesn't exist in the xml"
'                eCustomReportStatus = StrMfgNoFileTypeMatch
'                Exit Sub
'            Else
'                Dim bFileExist As Boolean
'                bFileExist = False
'                'Validate lua script location
'                Dim strLuaScriptLocation As String
'                strLuaScriptLocation = oContronInputXMLDocument.selectSingleNode("//MACROSCRIPTS/@PATH").Text
'
'                If m_oFileSystmeObject.FolderExists(strLuaScriptLocation) = True Then
'                   bFileExist = True
'                ElseIf m_oFileSystmeObject.FolderExists(strSharedContentPath + "\" + strLuaScriptLocation) = True Then
'                    bFileExist = True
'                End If
'
'                If bFileExist = False Then
'                    Print #1, strLuaScriptLocation + ", " + " lua script location doesn't exist"
'                    eCustomReportStatus = StrMfgNoFileTypeMatch
'                    Exit Sub
'                End If
'
'                'Validate detail scallop dxf location
'                bFileExist = False
'                Dim strDetailScallopLocation As String
'                strDetailScallopLocation = oContronInputXMLDocument.selectSingleNode("//DETAILVIEWTEMPLATE/@PATH").Text
'
'                If m_oFileSystmeObject.FileExists(strDetailScallopLocation) = True Then
'                   bFileExist = True
'                ElseIf m_oFileSystmeObject.FileExists(strSharedContentPath + "\" + strDetailScallopLocation) = True Then
'                    bFileExist = True
'                End If
'
'                If bFileExist = False Then
'                    Print #1, strDetailScallopLocation + ", " + " Detail scallop dxf template file doesn't exist"
'                    eCustomReportStatus = StrMfgNoFileTypeMatch
'                    Exit Sub
'                End If
'            End If
'            Set oContronInputXMLDocument = Nothing
'        End If
'
'        oMemberCuttingDwg.Initialize strControlInputXML, strSharedContentPath
'
'        Dim oFolder As Folder
'        Set oFolder = m_oFileSystmeObject.GetFolder(strFolderName)
'
'        Dim oFile As File
'        Dim partPath() As String
'        Dim nCount As Integer
'        nCount = 0
'
'        For Each oFile In oFolder.Files
'
'            nCount = nCount + 1
'            Dim strXMLFileName As String
'            strXMLFileName = oFile.Name
'
'            If InStr(strXMLFileName, ".xml") > 0 Then
'
'                ReDim partPath(1)
'                partPath(0) = strFolderName + "\" + strXMLFileName
'
'                Dim oXMLDocument  As IXMLDOMDocument
'                Set oXMLDocument = New DOMDocument
'
'                bSucess = oXMLDocument.Load(partPath(0))
'                If bSucess = False Then
'                   Print #1, partPath(0) + " is not valid XML"
'                   GoTo Continue
'                End If
'
'                Dim oSMS_MACROSNode As IXMLDOMNode
'                Set oSMS_MACROSNode = oXMLDocument.selectSingleNode("//SMS_MACROS")
'
'                If oSMS_MACROSNode Is Nothing Then
'                    Print #1, partPath(0) + "," + " SMS_MACROS node doesn't exist in the xml"
'                   GoTo Continue
'                End If
'
'                Dim strOutputFileName As String
'                lastPos = InStrRev(strXMLFileName, ".xml")
'                strOutputFileName = Left$(strXMLFileName, lastPos - 1)
'
'                Dim strDrawingTemplate As String
'                strDrawingTemplate = strSharedContentPath + "\" + "StructManufacturing\DrawingTemplate\drawing_template_mcp_single_member_beam_CommonQty.dxf"
'                If m_oFileSystmeObject.FileExists(strDrawingTemplate) = False Then
'                    Print #1, strDrawingTemplate + " Drawing Template doesn't exist"
'                    GoTo Continue
'                End If
'
'                Dim strReportName As String
'                strReportName = oMemberCuttingDwg.Generate(partPath, strDrawingTemplate, True, strOutputFileName)
'
'                 'Delete log files
'                If m_oFileSystmeObject.FileExists(strReportName) = True Then
'                    m_oFileSystmeObject.DeleteFile strReportName
'                End If
'
'                Print #1, CStr(nCount) + ". " + partPath(0) + "," & strOutputFileName & ".dxf File Is generated"
'                Set oXMLDocument = Nothing
'            Else
'                Print #1, CStr(nCount) + ". " + strXMLFileName + " is not XML"
'            End If
'Continue:
'            Set oFile = Nothing
'
'        Next
'        Set oFolder = Nothing
'    Else
'        eCustomReportStatus = StrMfgNoObjectFound
'        Print #1, "MfgMemberCuttingDrawings.DXF is not registered or can't be created"
'    End If
'
'    Set oMemberCuttingDwg = Nothing
    eCustomReportStatus = StrMfgProcessFinished
    
Wrapup:
    Set m_oFileSystmeObject = Nothing
    Exit Sub
    
ErrorHandler:
    eCustomReportStatus = StrMfgErrorUnknown
    Set m_oFileSystmeObject = Nothing
    Close #1
    Err.Raise Err.Number, , Err.Description
End Sub
Private Function GetSharedContentPath() As String
Const METHOD = "GetSharedContentPath"
On Error GoTo ErrorHandler

    Dim oContext As IJContext
    Set oContext = GetJContext()
    
    Dim strContextString As String
    strContextString = "OLE_SERVER"
    
    Dim strSymbolShare As String
    
    If Not oContext Is Nothing Then
        strSymbolShare = oContext.GetVariable(strContextString)
    End If
    
    GetSharedContentPath = strSymbolShare
    
    Set oContext = Nothing
    
Cleanup:
    Exit Function
ErrorHandler:
    GoTo Cleanup
End Function
