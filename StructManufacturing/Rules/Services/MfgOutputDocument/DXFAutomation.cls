VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DXFAutomation"
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
   
    Dim oDXFService As IJStrMfgDXFService
    Set oDXFService = New StrMfgDXFServiceLib.CStrMfgDXFService

    'Initialize the StyleService
    Dim strStyleXMLPath As String
    strStyleXMLPath = m_strSharedContentPath + "\" + GetRuleValue(m_bstrOutputOption, "String3")


    Dim oOutputReportDoc As IXMLDOMDocument
    Set oOutputReportDoc = oDXFService.Generate(pObject, pOutputXMLObject, m_bstrOutputOption, bstrBaseOutputName, strStyleXMLPath)
                                                                                   
    Set IJMfgOutputDocument_Generate = oOutputReportDoc

CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp

End Function
Private Sub IJMfgOutputDocument_Initialize(ByVal pControlXMLObject As Object, ByVal bstrOutputOption As String, ByVal bstrSharedContentsPath As String)
Const METHOD = "IJMfgOutputDocument_Initialize"
On Error GoTo ErrorHandler

    m_bstrOutputOption = bstrOutputOption
    m_strSharedContentPath = bstrSharedContentsPath
    If Not pControlXMLObject Is Nothing Then
        If TypeOf pControlXMLObject Is IXMLDOMDocument Then
            Set m_oControlInputDoc = pControlXMLObject
        End If
    End If

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Sub



