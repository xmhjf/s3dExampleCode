VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrowMark1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\ArrowMark1.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
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
' 11/03/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.ArrowMark1"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dMinorAngle                   As Double
Private m_dArrowLen                     As Double
Private m_dMajorAngle                   As Double
Private m_dVerticalLen                  As Double
Private m_dOffset                       As Double
Private m_dTextRightOffset              As Double
Private m_dTextBottomOffset             As Double
Private m_sPrefix                       As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dVerticalLen = 1.575 * m_dTextSize
    m_dMinorAngle = 0.262
    m_dMajorAngle = 1.912
    m_dArrowLen = 0.375 * m_dTextSize
    m_dOffset = 0.25 * m_dTextSize
    m_dTextRightOffset = 0.1 * m_dTextSize
    m_dTextBottomOffset = 0.333 * m_dTextSize
    m_sControlPoint = "ll"
    m_sPrefix = "2A"
    
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
        'sAttrName = trim(oXMLElement.getAttribute("NAME"))
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
                Case "MinorAngle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMinorAngle = Val(sAttrValue)
                    End If
                Case "MajorAngle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMajorAngle = Val(sAttrValue)
                    End If
                Case "VerticalLen"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalLen = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ArrowLen"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowLen = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextRightOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextRightOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextBottomOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextBottomOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "Offset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dOffset = Val(sAttrValue) * m_dTextSize
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
 
    Dim oStartPoint            As IJDPosition          '#0
    Dim oArrowStart             As IJDPosition          '#1
    Dim oArrowEnd               As IJDPosition          '#2
    Dim oHorizLineStart         As IJDPosition          '#3
    Dim oHorizLineEnd           As IJDPosition          '#4
    Dim oTextPoint              As IJDPosition          '#5
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    Dim oPlateMarkElem              As IXMLDOMAttribute
    
    
    
    Dim strText                     As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
    
    strText = GetAttributeValueFromXML(sAttributeXML, "DESCRIPTION")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")
    
    'normalize the vector
    pOrientation.length = 1

    Set oStartPoint = New DPosition
    Set oArrowStart = New DPosition                    '#1
    Set oArrowEnd = New DPosition                  '#2
    Set oHorizLineStart = New DPosition               '#3
    Set oHorizLineEnd = New DPosition               '#4
    Set oTextPoint = New DPosition                  '#5
    
    '*** Set All Points ***'
    
    oStartPoint.Set 0, 0, 0
    oArrowStart.Set 0, m_dOffset, 0                                                                 '#1
    oArrowEnd.Set (m_dArrowLen * Cos(3.142 - m_dMajorAngle - m_dMinorAngle)), (m_dOffset + (m_dArrowLen * Sin(3.142 - m_dMajorAngle - m_dMinorAngle))), 0                             '#2
    oHorizLineStart.Set ((m_dVerticalLen - m_dOffset) / Tan(3.142 - m_dMajorAngle)), m_dVerticalLen, 0                            '#3
    oHorizLineEnd.Set (((m_dVerticalLen - m_dOffset) / Tan(3.142 - m_dMajorAngle)) + (Len(strText) * m_dTextSize) + m_dTextRightOffset), m_dVerticalLen, 0         '#4
    oTextPoint.Set ((m_dVerticalLen - m_dOffset) / Tan(3.142 - m_dMajorAngle)), m_dVerticalLen + m_dTextBottomOffset, 0 '#5
    
    
    '**********************'
    'If pOrientation.x < 0 Then
        'pOrientation.x = -1 * pOrientation.x
    'End If
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint              '#1
    TranslatePoint oArrowStart, pOrientation, pStartPoint              '#1
    TranslatePoint oArrowEnd, pOrientation, pStartPoint          '#2
    TranslatePoint oHorizLineStart, pOrientation, pStartPoint           '#3
    TranslatePoint oHorizLineEnd, pOrientation, pStartPoint          '#4
    TranslatePoint oTextPoint, pOrientation, pStartPoint         '#5

    
    
    
    
    '*** Create the lines ***'
    'From 1 to 2 -- Line 1
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStart, oArrowEnd

    'From 1 to 3 -- Line 2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStart, oHorizLineStart

    'From 3 to 4 -- Line 3
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oHorizLineStart, oHorizLineEnd
    
    '************************'
  
    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing
    
    '*** Create the XML Text ***'

    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint, strText, _
                                              m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                               m_dTextSize, 0, "DESCRIPTION"
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:
    Set oStartPoint = Nothing
    Set oArrowStart = Nothing                    '#1
    Set oArrowEnd = Nothing                  '#2
    Set oHorizLineStart = Nothing               '#3
    Set oHorizLineEnd = Nothing                  '#4
    Set oTextPoint = Nothing               '#5
  
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function

'Below code is for adding a line to the output for testing purposes:
'Dim oTestPos As IJDPosition
'Set oTestPos = New DPosition
'Dim oTestPos2 As IJDPosition
'Set oTestPos2 = New DPosition
'oTestPos.Set 0, 150, 0
'TranslatePoint oTestPos, pOrientation, pStartPoint
''MsgBox "startx: " & pStartPoint.x & ", starty: " & pStartPoint.y & vbNewLine _
''    & ", endx: " & oTestPos.x & ", endy: " & oTestPos.y
'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, pStartPoint, oTestPos, "plate_thickness_annotation")
'If Not oTempCurveElement Is Nothing Then oTempEdgeElement.appendChild oTempCurveElement
'
'    oOutputElem.appendChild oTempEdgeElement






