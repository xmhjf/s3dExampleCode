VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberCuttingDwg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJMfgOutputDocument
Private m_oControlInputDoc As IXMLDOMDocument
'Refer to CSCAM library
'Private m_oCSCAMMemberCuttingDwg As IGenerateDocument
Private m_strSharedContentPath As String
Private m_bstrOutputOption As String

Private Sub Class_Initialize()
    Set m_oControlInputDoc = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_oControlInputDoc = Nothing
    'Set m_oCSCAMMemberCuttingDwg = Nothing
End Sub

Private Function IJMfgOutputDocument_Generate(ByVal pObject As Object, ByVal pOutputXMLObject As Object, ByVal bstrBaseOutputName As String) As Object
Const METHOD = "IJMfgOutputDocument_Generate"
On Error GoTo ErrorHandler


'    Dim strOutputDocumentName As String
'    Dim oXMLDoc As DOMDocument
'    Set oXMLDoc = pOutputXMLObject
'
'    Dim strPartGuid As String
'    Dim strModelPartGuid As String

'
'    Dim oFileSystemObj As FileSystemObject
'    Set oFileSystemObj = New FileSystemObject
'
'    If Not oXMLDoc Is Nothing Then
'        Dim oNodeList As IXMLDOMNodeList
'        Dim oNode As IXMLDOMNode
'        Set oNodeList = oXMLDoc.selectNodes("//SMS_PLATE | //SMS_PROFILE")
'
'        For Each oNode In oNodeList
'            strModelPartGuid = oNode.selectSingleNode(".//SMS_PART_INFO/@MODEL_PART_GUID").nodeValue
'            strPartGuid = oNode.selectSingleNode(".//SMS_PART_INFO/@PART_GUID").nodeValue
'        Next oNode
'
'        Dim strExtension As String
'        strExtension = GetRuleValue(m_bstrOutputOption, "DocumentType")
'
'        strOutputDocumentName = bstrBaseOutputName + "_" + "{" + strModelPartGuid + "}" + strExtension
'
'        Dim strDrawingTemplate As String
'        strDrawingTemplate = GetRuleValue(m_bstrOutputOption, "String7")
'        strDrawingTemplate = m_strSharedContentPath + "\" + strDrawingTemplate
'
'        'For the demonstration
'        Dim partPath() As String
'        If Not m_oCSCAMMemberCuttingDwg Is Nothing Then
'           ReDim partPath(1)
'           partPath(0) = "D:\03_ENEST\Samples\Others\Sample1XML_AfterModification-WithSectionInfo.xml"
'           Dim outputReport As String
'           outputReport = m_oCSCAMMemberCuttingDwg.Generate(partPath, strDrawingTemplate, True, strOutputDocumentName)
'           If oFileSystemObj.FileExists(outputReport) = True Then
'                oFileSystemObj.DeleteFile outputReport
'            End If
'            Set oFileSystemObj = Nothing
'        End If
'
'       'Generate Output ReportDocument
'        Dim oOutputReportDoc As New DOMDocument
'        Dim oOutputReportNode As IXMLDOMNode
'        Set oOutputReportNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputReport", "")
'        oOutputReportDoc.appendChild oOutputReportNode
'
'        Dim oOutputDocumentNode As IXMLDOMNode
'        Set oOutputDocumentNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputDocument", "")
'
'        Dim oOutputDocumentElem As IXMLDOMElement
'        Set oOutputDocumentElem = oOutputDocumentNode
'
'        oOutputDocumentElem.setAttribute "ID", "{" + strPartGuid + "}"
'
'        Dim lastPos As Long
'        lastPos = InStrRev(strOutputDocumentName, "\")
'        Dim strFileName As String
'
'        strFileName = Trim(Right$(strOutputDocumentName, Len(strOutputDocumentName) - lastPos))
'        oOutputDocumentElem.setAttribute "NAME", strFileName
'
'        oOutputReportNode.appendChild oOutputDocumentNode
'        Set IJMfgOutputDocument_Generate = oOutputReportDoc
'    End If
'
'
CleanUp:
'    Set m_oCSCAMMemberCuttingDwg = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function


Private Sub IJMfgOutputDocument_Initialize(ByVal pControlXMLObject As Object, ByVal bstrOutputOption As String, ByVal bstrSharedContentsPath As String)
Const METHOD = "IJMfgOutputDocument_Initialize"
On Error GoTo ErrorHandler

'    'Get Information from
'    m_bstrOutputOption = bstrOutputOption
'    m_strSharedContentPath = bstrSharedContentsPath
'
'    Dim strString6 As Variant
'    Dim strControlInputXMLPath As String
'    Dim iLoc As Integer
'    strString6 = GetRuleValue(m_bstrOutputOption, "String6")
'
'    iLoc = InStr(strString6, "|")
'    If iLoc > 1 Then
'        strControlInputXMLPath = Mid(strString6, iLoc + 1)
'    End If
'
'    strControlInputXMLPath = m_strSharedContentPath + "\" + strControlInputXMLPath
'
'    Set m_oCSCAMMemberCuttingDwg = CreateObject("MfgMemberCuttingDrawings.DXF")
'
'    If Not m_oCSCAMMemberCuttingDwg Is Nothing Then
'        m_oCSCAMMemberCuttingDwg.Initialize strControlInputXMLPath, m_strSharedContentPath
'    End If


CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Sub
