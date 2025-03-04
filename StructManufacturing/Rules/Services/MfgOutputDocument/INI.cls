VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJMfgOutputDocument

Private m_oControlInputDoc As IXMLDOMDocument
Private m_bstrOutputOption As String
Private m_strSharedContentPath As String

Private Sub Class_Initialize()
    Set m_oControlInputDoc = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_oControlInputDoc = Nothing
End Sub

Private Function IJMfgOutputDocument_Generate(ByVal pObject As Object, ByVal pOutputXMLObject As Object, ByVal bstrBaseOutputName As String) As Object
Const METHOD = "IJMfgOutputDocument_Generate"
On Error GoTo ErrorHandler

    Dim strModelPartGuid As String
    Dim strOutputDocumentName As String
    Dim strID As String
    Dim strExtension As String
    Dim strPartName As String
    Dim strPartGuid As String
    strExtension = GetRuleValue(m_bstrOutputOption, "DocumentType")
   
    If TypeOf pObject Is IJMfgPlatePart Then
        Dim oXMLDOC As DOMDocument
        Set oXMLDOC = pOutputXMLObject
   
        If Not oXMLDOC Is Nothing Then
             Dim oNodeList As IXMLDOMNodeList
             Dim oNode As IXMLDOMNode
             Set oNodeList = oXMLDOC.selectNodes("//SMS_PLATE | //SMS_PROFILE")
    
             For Each oNode In oNodeList
                 strPartName = oNode.selectSingleNode(".//SMS_PART_INFO/@PART_NAME").nodeValue
                 strModelPartGuid = oNode.selectSingleNode(".//SMS_PART_INFO/@MODEL_PART_GUID").nodeValue
                 strPartGuid = oNode.selectSingleNode(".//SMS_PART_INFO/@PART_GUID").nodeValue
            Next oNode
    
            strID = "{" + strPartGuid + "}"
            strOutputDocumentName = bstrBaseOutputName + "_" + "{" + strModelPartGuid + "}" + strExtension
            Open strOutputDocumentName For Output As #1
            Print #1, "Part NAME : " & strPartName
            Print #1, "Part GUID : " & strPartGuid
            Close #1
    
            Set IJMfgOutputDocument_Generate = GenerateOutputDocument(strID, strOutputDocumentName)
        End If
    ElseIf TypeOf pObject Is IJAssemblyChild Then
    
        Dim oAssembly As IJAssemblyBase
        Dim oAssemblyChild As IJAssemblyChild
        
        Set oAssemblyChild = pObject
        Set oAssembly = oAssemblyChild.Parent
        
        strModelPartGuid = GetOID(pObject)
        strPartGuid = CleanGUID(strModelPartGuid)
        strOutputDocumentName = bstrBaseOutputName + "_" + strModelPartGuid + strExtension
        
        Dim oNamedItem As IJNamedItem
        Set oNamedItem = pObject
        strPartName = oNamedItem.Name
        Set oNamedItem = Nothing
        
        If Not oAssembly Is Nothing Then
        
            If oAssembly.Type = 7 Then  'PreAssembly
                Exit Function
            Else
                Open strOutputDocumentName For Output As #1
                Print #1, "Part NAME : " & strPartName
                Print #1, "Part GUID : " & strModelPartGuid
                Close #1
                
                strID = "{" + strPartGuid + "}"
                Set IJMfgOutputDocument_Generate = GenerateOutputDocument(strID, strOutputDocumentName)
            End If
        End If
    End If

CleanUp:
    Set m_oControlInputDoc = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Sub IJMfgOutputDocument_Initialize(ByVal pControlXMLObject As Object, ByVal bstrOutputOption As String, ByVal bstrSharedContentsPath As String)
    m_bstrOutputOption = bstrOutputOption
    m_strSharedContentPath = bstrSharedContentsPath
End Sub

Private Function GenerateOutputDocument(ByVal strID As String, ByVal strOutputDocumentName As String) As DOMDocument
Const METHOD = "GenerateOutputDocument"
On Error GoTo ErrorHandler

    'Generate Output ReportDocument
    Dim oOutputReportDoc As New DOMDocument
    Dim oOutputReportNode As IXMLDOMNode
    Set oOutputReportNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputReport", "")
    oOutputReportDoc.appendChild oOutputReportNode

    Dim oOutputDocumentNode As IXMLDOMNode
    Set oOutputDocumentNode = oOutputReportDoc.createNode(NODE_ELEMENT, "S3DOutputDocument", "")

    Dim oOutputDocumentElem As IXMLDOMElement
    Set oOutputDocumentElem = oOutputDocumentNode

    oOutputDocumentElem.setAttribute "ID", strID

    Dim lastPos As Long
    lastPos = InStrRev(strOutputDocumentName, "\")
    Dim strFileName As String
    strFileName = Trim(Right$(strOutputDocumentName, Len(strOutputDocumentName) - lastPos))
    oOutputDocumentElem.setAttribute "NAME", strFileName

    oOutputReportNode.appendChild oOutputDocumentNode
    
    Set GenerateOutputDocument = oOutputReportDoc
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function


