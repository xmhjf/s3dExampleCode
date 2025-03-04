VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\SeamControl3.cls
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
' 09/10/2010    Ninad           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.SeamControl3"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dVerticalDist                  As Double
Private m_dHorizDist                  As Double
Private m_dArrowLength               As Double
Private m_dArrowHeadDimension              As Double
Private m_SeamText                  As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dVerticalDist = 4 * m_dTextSize
    m_dHorizDist = 0.6 * m_dTextSize
    m_sControlPoint = "ll"
    m_dArrowLength = 2 * m_dTextSize
    m_dArrowHeadDimension = 0.5 * m_dTextSize
    m_SeamText = "100"
    
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
                Case "HorizDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizDist = Val(sAttrValue) * m_dTextSize
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
                    
'                #7
'                .
'            #10/|\#11  #13
'                |     .100
'             #8 | #9
'               \|/
'            #2  . #4
'             \ #6 /
'              \  /
'        _______\/_______
'             #1/\
'              /  \
'             /    \
'            #5    #3
             
    Dim oCenterPoint                As IJDPosition          '#1
    Dim oSlantLine1TopPoint         As IJDPosition          '#2
    Dim oSlantLine1BottomPoint      As IJDPosition          '#3
    Dim oSlantLine2TopPoint         As IJDPosition          '#4
    Dim oSlantLine2BottomPoint      As IJDPosition          '#5

    '*** ARROW ***'
    Dim oArrowStartPoint            As IJDPosition          '#6
    Dim oArrowEndPoint              As IJDPosition          '#7
    Dim oArrowHeadPoint1            As IJDPosition          '#8
    Dim oArrowHeadPoint2            As IJDPosition          '#9
    Dim oArrowHeadPoint3            As IJDPosition          '#10
    Dim oArrowHeadPoint4            As IJDPosition          '#11
    Dim oTextPoint                  As IJDPosition          '#12


    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim vTemp                       As Variant
    Dim strPART_DIR                 As String
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
  
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oCenterPoint = New DPosition                        '#1
    Set oSlantLine1TopPoint = New DPosition                 '#2
    Set oSlantLine1BottomPoint = New DPosition              '#3
    Set oSlantLine2TopPoint = New DPosition                 '#4
    Set oSlantLine2BottomPoint = New DPosition              '#5
    
    Set oArrowStartPoint = New DPosition                    '#6
    Set oArrowEndPoint = New DPosition                      '#7
    Set oArrowHeadPoint1 = New DPosition                    '#8
    Set oArrowHeadPoint2 = New DPosition                    '#9
    Set oArrowHeadPoint3 = New DPosition                    '#10
    Set oArrowHeadPoint4 = New DPosition                    '#11
    Set oTextPoint = New DPosition
    
    '*** Set All Points ***'
    oCenterPoint.Set 0, 0, 0                                                '#1
    oSlantLine1TopPoint.Set -m_dHorizDist / 2, m_dVerticalDist / 2, 0       '#2
    oSlantLine1BottomPoint.Set m_dHorizDist / 2, -m_dVerticalDist / 2, 0    '#3
    oSlantLine2TopPoint.Set m_dHorizDist / 2, m_dVerticalDist / 2, 0        '#4
    oSlantLine2BottomPoint.Set -m_dHorizDist / 2, -m_dVerticalDist / 2, 0   '#5
    
    oArrowStartPoint.Set 0, m_dVerticalDist / 2, 0                          '#6
    oArrowEndPoint.Set 0, (m_dVerticalDist / 2) + m_dArrowLength, 0                               '#7
    oArrowHeadPoint1.Set 0 - m_dArrowHeadDimension, (m_dVerticalDist / 2) + m_dArrowHeadDimension, 0      '#8
    oArrowHeadPoint2.Set 0 + m_dArrowHeadDimension, (m_dVerticalDist / 2) + m_dArrowHeadDimension, 0      '#9
    oArrowHeadPoint3.Set 0 - m_dArrowHeadDimension, (m_dVerticalDist / 2) + m_dArrowLength - m_dArrowHeadDimension, 0           '#8
    oArrowHeadPoint4.Set 0 + m_dArrowHeadDimension, (m_dVerticalDist / 2) + m_dArrowLength - m_dArrowHeadDimension, 0       '#9
    oTextPoint.Set m_dHorizDist / 2, (m_dVerticalDist / 2) + (m_dArrowLength / 2), 0
    '**********************'
    
'    Select Case UCase(strPART_DIR)
'        Case "R"
                oArrowStartPoint.Set 0, -(m_dVerticalDist / 2), 0                          '#6
                oArrowEndPoint.Set 0, -((m_dVerticalDist / 2) + m_dArrowLength), 0                                '#7
                oArrowHeadPoint1.Set 0 - m_dArrowHeadDimension, -((m_dVerticalDist / 2) + m_dArrowHeadDimension), 0     '#8
                oArrowHeadPoint2.Set 0 + m_dArrowHeadDimension, -((m_dVerticalDist / 2) + m_dArrowHeadDimension), 0      '#9
                oArrowHeadPoint3.Set 0 - m_dArrowHeadDimension, -((m_dVerticalDist / 2) + m_dArrowLength - m_dArrowHeadDimension), 0           '#8
                oArrowHeadPoint4.Set 0 + m_dArrowHeadDimension, -((m_dVerticalDist / 2) + m_dArrowLength - m_dArrowHeadDimension), 0       '#9
                oTextPoint.Set m_dHorizDist / 2, -((m_dVerticalDist / 2) + (m_dArrowLength / 2)), 0
'    End Select
    
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oCenterPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine1TopPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine1BottomPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine2TopPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine2BottomPoint, pOrientation, pStartPoint
    
    TranslatePoint oArrowStartPoint, pOrientation, pStartPoint
    TranslatePoint oArrowEndPoint, pOrientation, pStartPoint
    TranslatePoint oArrowHeadPoint1, pOrientation, pStartPoint
    TranslatePoint oArrowHeadPoint2, pOrientation, pStartPoint
    TranslatePoint oArrowHeadPoint3, pOrientation, pStartPoint
    TranslatePoint oArrowHeadPoint4, pOrientation, pStartPoint
    
    TranslatePoint oTextPoint, pOrientation, pStartPoint
    
    '*** Create the lines ***'
    
    'Slant Line 1
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSlantLine1TopPoint, oSlantLine1BottomPoint

    'Slant Line 2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSlantLine2TopPoint, oSlantLine2BottomPoint
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStartPoint, oArrowEndPoint

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStartPoint, oArrowHeadPoint1

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowStartPoint, oArrowHeadPoint2

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowEndPoint, oArrowHeadPoint3

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrowEndPoint, oArrowHeadPoint4

    '************************'

    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint, m_SeamText, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, -1.5707, "DISPLACEMENT"

    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oCenterPoint = Nothing                         '#1
    Set oSlantLine1TopPoint = Nothing                  '#2
    Set oSlantLine1BottomPoint = Nothing               '#3
    Set oSlantLine2TopPoint = Nothing                  '#4
    Set oSlantLine2BottomPoint = Nothing               '#5
  
    
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








