VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndFittingMark2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\EndFittingMark2.cls
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
' 08/26/2010    Ninad           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.EndFittingMark2"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dArrowHorizDist             As Double
Private m_dArrowLength                  As Double
Private m_dArrowVerticalDist                  As Double
Private m_sPrefix                     As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dArrowLength = 2 * m_dTextSize
    m_dArrowHorizDist = 0.3 * m_dTextSize
    m_dArrowVerticalDist = 0.5 * m_dTextSize
    m_sControlPoint = "ur"
    m_sPrefix = "$"
    
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
                Case "ArrowLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ArrowHorizDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowHorizDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ArrowVerticalDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowVerticalDist = Val(sAttrValue) * m_dTextSize
                    End If
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
    
    
'            #2
'            | #5
'            | $
'         #3 |   #4
'            |   /
'            |  /
'            | /
'    ________|/_________
'            #1
    

    
    Dim oVerticalLineStartPoint     As IJDPosition          '#1
    Dim oVerticalLineEndPoint       As IJDPosition          '#2
    Dim oVerticalLineMidPoint       As IJDPosition          '#3
    Dim oArrowStartPoint            As IJDPosition          '#1
    Dim oArrowEndPoint              As IJDPosition          '#4
    Dim oTextPoint                  As IJDPosition          '#5
        
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim strPART_SIDE                 As String
    Dim strGrinder                As String
    Dim strUpside                   As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_SIDE will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
   
    strPART_SIDE = GetAttributeValueFromXML(sAttributeXML, "PART_SIDE")
    strGrinder = GetAttributeValueFromXML(sAttributeXML, "PREFIX")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oVerticalLineStartPoint = New DPosition             '#1
    Set oVerticalLineEndPoint = New DPosition               '#2
    Set oVerticalLineMidPoint = New DPosition               '#3
    Set oArrowStartPoint = New DPosition                    '#1
    Set oArrowEndPoint = New DPosition                      '#4
    Set oTextPoint = New DPosition
    
    '*** Set All Points ***'
    oVerticalLineStartPoint.Set 0, 0, 0                     '#1
    oVerticalLineEndPoint.Set 0, m_dArrowLength, 0         '#2
    oVerticalLineMidPoint.Set 0, m_dArrowLength / 2, 0     '#3
    oArrowStartPoint.Set 0, 0, 0           '#3
    oArrowEndPoint.Set m_dArrowHorizDist, m_dArrowVerticalDist, 0          '#4
    oTextPoint.Set 0, m_dArrowLength, 0                     '#5
    '**********************'
    
    Select Case UCase(strPART_SIDE)
        Case "R"
            oVerticalLineEndPoint.Set 0, -m_dArrowLength, 0         '#2
            oVerticalLineMidPoint.Set 0, -(m_dArrowLength / 2), 0     '#3
            oArrowEndPoint.Set m_dArrowHorizDist, -m_dArrowVerticalDist, 0          '#4
            oTextPoint.Set 0, -m_dArrowLength / 2, 0           '#5
    End Select
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oVerticalLineStartPoint, pOrientation, pStartPoint
    TranslatePoint oVerticalLineEndPoint, pOrientation, pStartPoint
    TranslatePoint oVerticalLineMidPoint, pOrientation, pStartPoint
        
    TranslatePoint oArrowStartPoint, pOrientation, pStartPoint
    TranslatePoint oArrowEndPoint, pOrientation, pStartPoint
    TranslatePoint oTextPoint, pOrientation, pStartPoint
    
    

    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVerticalLineStartPoint, oVerticalLineEndPoint

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStartPoint, oArrowEndPoint

    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing
    
    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint, strGrinder, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 1.5707, "PREFIX"
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oVerticalLineStartPoint = Nothing
    Set oVerticalLineEndPoint = Nothing
    Set oVerticalLineMidPoint = Nothing
    Set oArrowStartPoint = Nothing
    Set oArrowEndPoint = Nothing
  
    
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
'oTestPos.Set 0, 180, 0
'TranslatePoint oTestPos, pOrientation, pStartPoint
''MsgBox "startx: " & pStartPoint.x & ", starty: " & pStartPoint.y & vbNewLine _
''    & ", endx: " & oTestPos.x & ", endy: " & oTestPos.y
'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, pStartPoint, oTestPos, "plate_thickness_annotation")
'If Not oTempCurveElement Is Nothing Then oTempEdgeElement.appendChild oTempCurveElement
'
'    oOutputElem.appendChild oTempEdgeElement








