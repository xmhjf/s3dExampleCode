VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CarDeckMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\CarDeckMark.cls
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
' 11/04/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.CarDeckMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dPlateHeight                  As Double
Private m_dWebLength                  As Double
Private m_dCornerRadius                  As Double
Private m_dFilletRadius                  As Double
Private m_dFlangeLength                  As Double
Private m_dWebThickness                  As Double
Private m_dFlangeThickness                  As Double
Private m_dPlateThickness                  As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dPlateHeight = 3 * m_dTextSize
    m_dWebLength = 2.5 * m_dTextSize
    m_dCornerRadius = 0.125 * m_dTextSize
    m_dFilletRadius = 0.25 * m_dTextSize
    m_dFlangeLength = 1.625 * m_dTextSize
    m_dWebThickness = 0.25 * m_dTextSize
    m_dFlangeThickness = 0.25 * m_dTextSize
    m_dPlateThickness = 0.333 * m_dTextSize
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
                Case "PlateHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dPlateHeight = Val(sAttrValue) * m_dTextSize
                    End If
                Case "WebLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dWebLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "CornerRadius"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dCornerRadius = Val(sAttrValue) * m_dTextSize
                    End If
                Case "FilletRadius"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFilletRadius = Val(sAttrValue) * m_dTextSize
                    End If
                Case "FlangeLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlangeLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "FlangeThickness"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlangeThickness = Val(sAttrValue) * m_dTextSize
                    End If
                Case "WebThickness"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dWebThickness = Val(sAttrValue) * m_dTextSize
                    End If
                Case "PlateThickness"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dPlateThickness = Val(sAttrValue) * m_dTextSize
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
      
    Dim oStartPoint           As IJDPosition          '#1
    Dim oLine1RightPoint         As IJDPosition          '#2
    Dim oLine2TopPoint      As IJDPosition          '#3
    Dim oLine3TopPoint      As IJDPosition          '#4
    Dim oArc1StartPoint         As IJDPosition          '#5
    Dim oCenterPoint1         As IJDPosition          '#6
    Dim oCenterPoint2         As IJDPosition          '#7
    Dim oArc2StartPoint         As IJDPosition          '#8
    Dim oLine5EndPoint         As IJDPosition          '#9
    Dim oLine6BottomPoint         As IJDPosition          '#10
    Dim oLine8BottomPoint         As IJDPosition          '#11
    Dim oLine8TopPoint         As IJDPosition          '#12
    Dim oLine6TopPoint         As IJDPosition          '#13
    Dim oArc2EndPoint         As IJDPosition          '#14
    

    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim vTemp                       As Variant
    
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
    
    Set oStartPoint = New DPosition                    '#1
    Set oLine1RightPoint = New DPosition                  '#2
    Set oLine2TopPoint = New DPosition               '#3
    Set oLine3TopPoint = New DPosition               '#4
    Set oArc1StartPoint = New DPosition                  '#5
    Set oCenterPoint1 = New DPosition                  '#6
    Set oCenterPoint2 = New DPosition                  '#7
    Set oArc2StartPoint = New DPosition                  '#8
    Set oLine5EndPoint = New DPosition                  '#9
    Set oLine6BottomPoint = New DPosition                  '#10
    Set oLine8BottomPoint = New DPosition                  '#11
    Set oLine8TopPoint = New DPosition                  '#12
    Set oLine6TopPoint = New DPosition                  '#13
    Set oArc2EndPoint = New DPosition                  '#14
    

    
    '*** Set All Points ***'
    
    oStartPoint.Set 0, 0, 0                                    '#1
    oLine1RightPoint.Set m_dWebThickness, 0, 0                    '#2
    oLine2TopPoint.Set 0, m_dWebLength, 0                '#3
    oLine3TopPoint.Set m_dWebThickness, (m_dWebLength - m_dCornerRadius - m_dFlangeThickness), 0               '#4
    oArc1StartPoint.Set (m_dWebThickness + m_dCornerRadius), (m_dWebLength - m_dFlangeThickness), 0    '#5
    oCenterPoint1.Set (m_dWebThickness + m_dCornerRadius), (m_dWebLength - m_dCornerRadius - m_dFlangeThickness), 0        '#6
    oCenterPoint2.Set (m_dFlangeLength - m_dFilletRadius), (m_dWebLength + m_dFilletRadius - m_dFlangeThickness), 0       '#7
    oArc2StartPoint.Set (m_dFlangeLength - m_dFilletRadius), (m_dWebLength - m_dFlangeThickness), 0    '#8
    oLine5EndPoint.Set m_dFlangeLength, m_dWebLength, 0        '#9
    oLine6BottomPoint.Set m_dFlangeLength, 0, 0        '#10
    oLine8BottomPoint.Set (m_dFlangeLength + m_dPlateThickness), 0, 0      '#11
    oLine8TopPoint.Set (m_dFlangeLength + m_dPlateThickness), m_dPlateHeight, 0      '#12
    oLine6TopPoint.Set m_dFlangeLength, m_dPlateHeight, 0        '#13
    oArc2EndPoint.Set m_dFlangeLength, (m_dWebLength + m_dFilletRadius - m_dFlangeThickness), 0        '#14
    
    
    '**********************'
'    If pOrientation.x < 0 Then
'        pOrientation.x = -1 * pOrientation.x
'    End If
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oLine1RightPoint, pOrientation, pStartPoint
    TranslatePoint oLine2TopPoint, pOrientation, pStartPoint
    TranslatePoint oLine3TopPoint, pOrientation, pStartPoint
    TranslatePoint oArc1StartPoint, pOrientation, pStartPoint
    TranslatePoint oCenterPoint1, pOrientation, pStartPoint
    TranslatePoint oCenterPoint2, pOrientation, pStartPoint
    TranslatePoint oArc2StartPoint, pOrientation, pStartPoint
    TranslatePoint oLine5EndPoint, pOrientation, pStartPoint
    TranslatePoint oLine6BottomPoint, pOrientation, pStartPoint
    TranslatePoint oLine8BottomPoint, pOrientation, pStartPoint
    TranslatePoint oLine8TopPoint, pOrientation, pStartPoint
    TranslatePoint oLine6TopPoint, pOrientation, pStartPoint
    TranslatePoint oArc2EndPoint, pOrientation, pStartPoint
   
    'Creating the Output
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    '*** Create the lines ***'
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    'From 1 to 2 -- Line 1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine1RightPoint

    'From 1 to 3 -- Line 1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine2TopPoint

    'From 2 to 4 of Line 3
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1RightPoint, oLine3TopPoint
    
    'From 3 to 9 of Line 4
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2TopPoint, oLine5EndPoint
    
    'From 5 to 8 of Line 5
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArc1StartPoint, oArc2StartPoint
    
    'From 10 to 13 of Line 6
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine6BottomPoint, oLine6TopPoint
    
    'From 13 to 12 of Line 7
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine6TopPoint, oLine8TopPoint
    
    'From 11 to 12 of Line 8
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine8TopPoint, oLine8BottomPoint
    
    '*** Create the Arcs ***'
    
    'From 5 to 4 -- Arc1
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1StartPoint, oLine3TopPoint, oCenterPoint1
    
    'From 8 to 14 -- Arc2
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc2StartPoint, oArc2EndPoint, oCenterPoint2
'
    '************************'
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

    
CleanUp:
    Set oStartPoint = Nothing                    '#1
    Set oLine1RightPoint = Nothing                  '#2
    Set oLine2TopPoint = Nothing               '#3
    Set oLine3TopPoint = Nothing                  '#4
    Set oArc1StartPoint = Nothing               '#5
    Set oCenterPoint1 = Nothing               '#6
    Set oCenterPoint2 = Nothing               '#7
    Set oArc2StartPoint = Nothing               '#8
    Set oLine5EndPoint = Nothing               '#9
    Set oLine6BottomPoint = Nothing               '#10
    Set oLine8BottomPoint = Nothing               '#11
    Set oLine8TopPoint = Nothing               '#12
    Set oLine6TopPoint = Nothing               '#13
    Set oArc2EndPoint = Nothing               '#14
  
    
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




