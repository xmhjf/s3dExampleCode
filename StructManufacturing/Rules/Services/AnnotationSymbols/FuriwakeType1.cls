VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuriwakeType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\FuriwakeType1.cls
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
' 02/8/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.FuriwakeType1"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dDistToCenterBottom               As Double
Private m_dThinPlateThickness               As Double
Private m_dThinPlateWidth                  As Double
Private m_dThickPlateThickness                  As Double
Private m_dThickPlateWidth                   As Double
Private m_dTextLeftOffset                   As Double
Private m_dBottomDifference               As Double
Private m_dTextBottomOffset          As Double


Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dDistToCenterBottom = 1 * m_dTextSize
    m_dThinPlateThickness = 0.4 * m_dTextSize
    m_dThinPlateWidth = 3 * m_dTextSize
    m_dThickPlateThickness = 1 * m_dTextSize
    m_dThickPlateWidth = 2 * m_dTextSize
    m_dBottomDifference = 0.3 * m_dTextSize
    m_dTextLeftOffset = 0.5 * m_dTextSize
    m_dTextBottomOffset = 0.1 * m_dTextSize
    m_sControlPoint = "ul"
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
                Case "DistToCenterBottom"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dDistToCenterBottom = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ThinPlateThickness"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dThinPlateThickness = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ThinPlateWidth"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dThinPlateWidth = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ThickPlateThickness"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dThickPlateThickness = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ThickPlateWidth"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dThickPlateWidth = Val(sAttrValue) * m_dTextSize
                    End If
                Case "BottomDifference"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dBottomDifference = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextLeftOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextLeftOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextBottomOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextBottomOffset = Val(sAttrValue) * m_dTextSize
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
    
    
    '                    M
    '                 #10
    '    1.5    #6___________#7
    '  #9        |
    '____________|
    '#5        #4|
    '            |
    '____________|
    '#3        #2|
    '            |___________#8
    '            #1
    '
    
    

    
    Dim oThinPlateBottomStartPoint  As IJDPosition          '#2
    Dim oThinPlateBottomEndPoint    As IJDPosition          '#3
    Dim oThinPlateTopStartPoint     As IJDPosition      '#4
    Dim oThinPlateTopEndPoint       As IJDPosition        '#5
    Dim oThickPlateBottomStartPoint As IJDPosition         '#1
    Dim oThickPlateBottomEndPoint   As IJDPosition         '#8
    Dim oThickPlateTopStartPoint    As IJDPosition     '#6
    Dim oThickPlateTopEndPoint      As IJDPosition       '#7
    Dim oThinPlateTextOffsetPoint   As IJDPosition     '#9
    Dim oThickPlateTextOffsetPoint  As IJDPosition     '#10
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    Dim oPlateMarkDom               As New DOMDocument
    
'    Dim strPART_DIR                 As String
    Dim sPartDir                    As String
    Dim strThickness                As String
    Dim strUpside                   As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
    
    
'    strPART_DIR = GetAttributeValueFromXML(sAttributeXML, "PART_DIR")
    strThickness = GetAttributeValueFromXML(sAttributeXML, "THICKNESS")
    strUpside = GetAttributeValueFromXML(sAttributeXML, "UPSIDE")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1

    Set oThinPlateBottomStartPoint = New DPosition
    Set oThinPlateBottomEndPoint = New DPosition
    Set oThinPlateTopStartPoint = New DPosition
    Set oThinPlateTopEndPoint = New DPosition
    
    Set oThickPlateBottomStartPoint = New DPosition
    Set oThickPlateBottomEndPoint = New DPosition
    Set oThickPlateTopStartPoint = New DPosition
    Set oThickPlateTopEndPoint = New DPosition
    
    Set oThinPlateTextOffsetPoint = New DPosition
    Set oThickPlateTextOffsetPoint = New DPosition
    
    
'    Select Case sPartDir
'        Case "R", "r"
'            oStartPos.Set m_dRightStartX * m_dTextSize, -m_dRightStartY * m_dTextSize, 0
'        Case "L", "l"
'            oStartPos.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
'        Case Else
'            oStartPos.Set -m_dCenteredStartX * m_dTextSize, -m_dCenteredStartY * m_dTextSize, 0
'    End Select
'
    
    '*** Set All Points ***'
    oThickPlateBottomStartPoint.Set 0, m_dDistToCenterBottom, 0   '#1
    oThickPlateBottomEndPoint.Set m_dThickPlateWidth, oThickPlateBottomStartPoint.Y, 0  '#8
    
    oThinPlateBottomStartPoint.Set 0, m_dDistToCenterBottom + m_dBottomDifference, 0    '#2
    oThinPlateBottomEndPoint.Set -m_dThinPlateWidth, oThinPlateBottomStartPoint.Y, 0    '#3
    
    oThinPlateTopStartPoint.Set 0, m_dDistToCenterBottom + m_dBottomDifference + m_dThinPlateThickness, 0    '#4
    oThinPlateTopEndPoint.Set -m_dThinPlateWidth, oThinPlateTopStartPoint.Y, 0    '#5
    
    oThickPlateTopStartPoint.Set 0, m_dDistToCenterBottom + m_dThickPlateThickness, 0   '#6
    oThickPlateTopEndPoint.Set m_dThickPlateWidth, oThickPlateTopStartPoint.Y, 0  '#7
    
    Dim sTextElem()             As String
    sTextElem = VBA.Split(strThickness, ".")
    
    If UBound(sTextElem) > LBound(sTextElem) Then
        If sTextElem(UBound(sTextElem)) = "0" Then
            oThinPlateTextOffsetPoint.Set oThinPlateTopEndPoint.X, oThinPlateTopEndPoint.Y + m_dTextBottomOffset, 0  '#9
        Else
            oThinPlateTextOffsetPoint.Set oThinPlateTopEndPoint.X - m_dTextSize, oThinPlateTopEndPoint.Y + m_dTextBottomOffset, 0 '#9
        End If
    Else
        oThinPlateTextOffsetPoint.Set oThinPlateTopEndPoint.X, oThinPlateTopEndPoint.Y + m_dTextBottomOffset, 0  '#9
    End If
    
    oThickPlateTextOffsetPoint.Set oThickPlateTopStartPoint.X, oThickPlateTopStartPoint.Y + m_dTextBottomOffset, 0  '#10
    '**********************'
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oThickPlateBottomStartPoint, pOrientation, pStartPoint
    TranslatePoint oThickPlateBottomEndPoint, pOrientation, pStartPoint
    TranslatePoint oThickPlateTopStartPoint, pOrientation, pStartPoint
    TranslatePoint oThickPlateTopEndPoint, pOrientation, pStartPoint
    
    TranslatePoint oThinPlateBottomStartPoint, pOrientation, pStartPoint
    TranslatePoint oThinPlateBottomEndPoint, pOrientation, pStartPoint
    TranslatePoint oThinPlateTopStartPoint, pOrientation, pStartPoint
    TranslatePoint oThinPlateTopEndPoint, pOrientation, pStartPoint
    
    TranslatePoint oThinPlateTextOffsetPoint, pOrientation, pStartPoint
    TranslatePoint oThickPlateTextOffsetPoint, pOrientation, pStartPoint
    
    
    
    
    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oThickPlateBottomStartPoint, oThickPlateBottomEndPoint
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oThickPlateTopStartPoint, oThickPlateTopEndPoint
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oThinPlateBottomStartPoint, oThinPlateBottomEndPoint

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oThinPlateTopStartPoint, oThinPlateTopEndPoint

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oThickPlateBottomStartPoint, oThickPlateTopStartPoint

    

    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing
    
    '*** Create the XML Text ***'
     SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oThinPlateTextOffsetPoint, strThickness, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 0, "THICKNESS"
                                               
     SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oThickPlateTextOffsetPoint, strUpside, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 0, "UPSIDE"
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oThinPlateBottomStartPoint = Nothing           '#2
    Set oThinPlateBottomEndPoint = Nothing           '#3
    Set oThinPlateTopStartPoint = Nothing        '#4
    Set oThinPlateTopEndPoint = Nothing             '#5

    
    Set oThickPlateBottomStartPoint = Nothing         '#1
    Set oThickPlateBottomEndPoint = Nothing          '#8
    Set oThickPlateTopStartPoint = Nothing         '#6
    Set oThickPlateTopEndPoint = Nothing            '#7
    
    Set oThinPlateTextOffsetPoint = Nothing          '#9
    Set oThickPlateTextOffsetPoint = Nothing           '#10
  
    
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


