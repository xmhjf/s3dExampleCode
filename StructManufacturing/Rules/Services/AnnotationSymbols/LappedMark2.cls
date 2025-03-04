VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LappedMark2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\NewLapAnnotation.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\AnnotationSymbols.vbp
'
'
' Abstract:
'   Create custom profile location flange symbols
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom flange direction symbol on profile location marks
'
' History:
' 02/16/2011    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.LappedMark2"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
'Creating Constructor Class
'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_sControlPoint = "ll"
    
End Sub

Private Sub IJDMfgOutputAnnotation_SetArguments(ByVal sSettingsXML As String)
    Const METHOD = "IJDMfgOutputAnnotation_SetArguments"
    On Error GoTo ErrorHandler

    Dim oXMLDomDoc As New DOMDocument
    Dim oAttributeNodeList As IXMLDOMNodeList
    Dim oXMLElement As IXMLDOMElement
    Dim sAttrName As String
    Dim sAttrValue As String
    Dim vTemp As Variant

    If Not oXMLDomDoc.loadXML(sSettingsXML) Then GoTo CleanUp
    
    Set oAttributeNodeList = oXMLDomDoc.getElementsByTagName("SMS_GEOM_ARG")
    
    If oAttributeNodeList Is Nothing Then GoTo CleanUp
    
    For Each oXMLElement In oAttributeNodeList
        sAttrName = ""
        sAttrValue = ""
        vTemp = Trim(oXMLElement.getAttribute("NAME"))
        sAttrName = IIf(VarType(vTemp) = vbString, vTemp, "")
        If Not sAttrName = "" Then
            Select Case sAttrName
                Case "ControlPoint"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sControlPoint = sAttrValue
                Case "TextSize"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextSize = Val(sAttrValue)
                    End If
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
            End Select
        End If
    Next oXMLElement
    
CleanUp:
    Set oXMLDomDoc = Nothing
    Set oAttributeNodeList = Nothing
    Set oXMLElement = Nothing
    

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Sub

Private Function IJDMfgOutputAnnotation_Evaluate(ByVal pStartPoint As IJDPosition, ByVal pOrientation As IJDVector, ByVal sAttributeXML As String) As String
    Const METHOD = "IJDMfgOutputAnnotation_Evaluate"
    On Error GoTo ErrorHandler
    
    '               #2|\
    '                 | \#4
    '                 |
    '                 |
    '         ________|#1_______Marking Line
    '                 |
    '                 |
    '                 | /#5
    '               #3|/
    '
    '
    'Distance between #2 & #3 = 4 * Textsize
    'Vertical Distance between #2 & #4 = Textsize
    'Horizontal Distance between #2 & #4 = 0.3 * Textsize
    'Vertical Distance between #3 & #5 = Textsize
    'Horizontal Distance between #3 & #5 = 0.3 * Textsize

    Dim oStartPoint                 As IJDPosition          '#1
    Dim oVertLineStartPoint         As IJDPosition          '#2
    Dim oVertLineEndPoint           As IJDPosition          '#3
    Dim oSlantLine1EndPoint         As IJDPosition          '#4
    Dim oSlantLine2EndPoint         As IJDPosition          '#5
   
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPoint = New DPosition                 '#1
    Set oVertLineStartPoint = New DPosition         '#2
    Set oVertLineEndPoint = New DPosition           '#3
    Set oSlantLine1EndPoint = New DPosition         '#4
    Set oSlantLine2EndPoint = New DPosition         '#5
    
    '*** Set All Points ***'
    oStartPoint.Set 0, 0, 0                                         '#1
    oVertLineStartPoint.Set 0, 2 * m_dTextSize, 0                   '#2
    oVertLineEndPoint.Set 0, -2 * m_dTextSize, 0                    '#3
    oSlantLine1EndPoint.Set 0.3 * m_dTextSize, m_dTextSize, 0       '#4
    oSlantLine2EndPoint.Set 0.3 * m_dTextSize, -m_dTextSize, 0      '#5
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oVertLineStartPoint, pOrientation, pStartPoint
    TranslatePoint oVertLineEndPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine1EndPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine2EndPoint, pOrientation, pStartPoint
    
    '*** Create the lines ***'
    'Vertical Line having Endpoints as #2 and #3
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVertLineStartPoint, oVertLineEndPoint

    'Slant Line 1 having Endpoints as #2 and #4
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVertLineStartPoint, oSlantLine1EndPoint
    
    'Slant Line 2 having Endpoints as #3 and #5
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVertLineEndPoint, oSlantLine2EndPoint
    '************************'
    
    Set oTempEdgeElement = Nothing
    
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:
    Set oStartPoint = Nothing                       '#1
    Set oVertLineStartPoint = Nothing               '#2
    Set oVertLineEndPoint = Nothing                 '#3
    Set oSlantLine1EndPoint = Nothing               '#4
    Set oSlantLine2EndPoint = Nothing               '#5
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function
