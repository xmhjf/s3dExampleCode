VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GrindType3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\GrindType3.cls
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
Private Const MODULE = "MfgCustomAnnotation.GrindType3"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dArrowLength               As Double
Private m_dArrowHeadDimension              As Double
Private m_dVerticalDist                  As Double
Private m_dTextVerticalDist                  As Double
Private m_dTextHorizDist                  As Double
Private m_sGrinder                  As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dArrowLength = 4 * m_dTextSize
    m_dArrowHeadDimension = 1 * m_dTextSize
    m_dVerticalDist = 4 * m_dTextSize
    m_dTextVerticalDist = 1 * m_dTextSize
    m_dTextHorizDist = 1 * m_dTextSize
    m_sControlPoint = "ll"
    m_sGrinder = "G"
    
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
                Case "VerticalDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ArrowLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ArrowHeadDimension"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowHeadDimension = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextVerticalDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextVerticalDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextHorizDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextHorizDist = Val(sAttrValue) * m_dTextSize
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
    
    
    
'        #2
'        |  #5
'        | /
'      #3|/___________#6  G
'        |\               #7
'        | \#4
'        |
'        #1
    

    
    Dim oVerticalLineStartPoint     As IJDPosition          '#1
    Dim oVerticalLineEndPoint       As IJDPosition          '#2
    Dim oVerticalLineMidPoint       As IJDPosition          '#3
    Dim oArrowStartPoint            As IJDPosition          '#3
    Dim oArrowEndPoint              As IJDPosition          '#6
    Dim oArrowHeadPoint1            As IJDPosition          '#4
    Dim oArrowHeadPoint2            As IJDPosition          '#5
    Dim oTextPoint                  As IJDPosition          '#7
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim strPART_DIR                 As String
    Dim sPartDir                    As String
    Dim strGrinder                  As String
    Dim strUpside                   As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp

    'strPART_DIR = GetAttributeValueFromXML(sAttributeXML, "PART_DIR")
    strGrinder = GetAttributeValueFromXML(sAttributeXML, "GRINDER")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oVerticalLineStartPoint = New DPosition
    Set oVerticalLineEndPoint = New DPosition
    Set oVerticalLineMidPoint = New DPosition
    Set oArrowStartPoint = New DPosition
    Set oArrowEndPoint = New DPosition
    Set oArrowHeadPoint1 = New DPosition
    Set oArrowHeadPoint2 = New DPosition
    Set oTextPoint = New DPosition
    
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
    oVerticalLineStartPoint.Set 0, 0, 0                     '#1
    oVerticalLineEndPoint.Set 0, m_dVerticalDist, 0         '#2
    oVerticalLineMidPoint.Set 0, m_dVerticalDist / 2, 0     '#3
    oArrowStartPoint.Set 0, m_dVerticalDist / 2, 0           '#3
    oArrowHeadPoint1.Set 0 + m_dArrowHeadDimension, m_dArrowHeadDimension, 0        '#4
    oArrowHeadPoint2.Set 0 + m_dArrowHeadDimension, (m_dVerticalDist / 2) + m_dArrowHeadDimension, 0      '#5
    oArrowEndPoint.Set 0 + m_dArrowLength, m_dVerticalDist / 2, 0         '#6
    oTextPoint.Set m_dArrowLength + m_dTextHorizDist, m_dTextVerticalDist, 0
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oVerticalLineStartPoint, pOrientation, pStartPoint
    TranslatePoint oVerticalLineEndPoint, pOrientation, pStartPoint
    TranslatePoint oVerticalLineMidPoint, pOrientation, pStartPoint
        
    TranslatePoint oArrowStartPoint, pOrientation, pStartPoint
    TranslatePoint oArrowHeadPoint1, pOrientation, pStartPoint
    TranslatePoint oArrowHeadPoint2, pOrientation, pStartPoint
    TranslatePoint oArrowEndPoint, pOrientation, pStartPoint
    
    TranslatePoint oTextPoint, pOrientation, pStartPoint
        
    
    
    
    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVerticalLineStartPoint, oVerticalLineEndPoint

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStartPoint, oArrowEndPoint

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStartPoint, oArrowHeadPoint1

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStartPoint, oArrowHeadPoint2

    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing
    
    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint, strGrinder, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 0, "GRINDER"
                                        
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
    Set oArrowHeadPoint1 = Nothing
    Set oArrowHeadPoint2 = Nothing
    Set oTextPoint = Nothing
  
    
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




