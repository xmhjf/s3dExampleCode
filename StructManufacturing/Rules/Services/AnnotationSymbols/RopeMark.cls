VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RopeMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\RopeMark.cls
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
Private Const MODULE = "MfgCustomAnnotation.RopeMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dVerticalDist                  As Double
Private m_dThickness                  As Double
Private m_dTopOffset                  As Double
Private m_dDiameter                  As Double


Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dVerticalDist = 2 * m_dTextSize
    m_dThickness = 0.8 * m_dTextSize
    m_dDiameter = m_dTextSize
    m_dTopOffset = 0.125 * m_dTextSize
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
                Case "VerticalDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "Diameter"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dDiameter = Val(sAttrValue) * m_dTextSize
                    End If
                Case "Thickness"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dThickness = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TopOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTopOffset = Val(sAttrValue) * m_dTextSize
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
 


'
'          #5      #2
'          |     |
'          |     |
'      ____|_____|#1_____
'          #4    |
'                |
'                |
'                #3
'
             
    Dim oStartPoint           As IJDPosition          '#1
    Dim oLine1RightPoint         As IJDPosition          '#2
    Dim oLine2TopPoint      As IJDPosition          '#3
    Dim oLine3LeftPoint      As IJDPosition          '#4
    Dim oLine5TopPoint         As IJDPosition          '#5
    Dim oCenterPoint         As IJDPosition          '#6
    Dim oArc1EndPoint         As IJDPosition          '#7
    Dim oArc2EndPoint         As IJDPosition          '#8

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
    Set oLine3LeftPoint = New DPosition               '#4
    Set oLine5TopPoint = New DPosition                  '#5
    Set oCenterPoint = New DPosition                  '#6
    Set oArc1EndPoint = New DPosition                  '#7
    Set oArc2EndPoint = New DPosition                  '#8

    '*** Set All Points ***'
    
    oStartPoint.Set 0, 0, 0                                    '#1
    oLine1RightPoint.Set m_dThickness, 0, 0                    '#2
    oLine2TopPoint.Set m_dThickness, m_dVerticalDist, 0                '#3
    oLine3LeftPoint.Set 0, m_dVerticalDist, 0                   '#4
    oLine5TopPoint.Set 0, (m_dVerticalDist + m_dTopOffset), 0      '#5
    oCenterPoint.Set 0, (m_dVerticalDist + m_dTopOffset - (m_dDiameter / 2)), 0    '#6
    oArc1EndPoint.Set -(m_dDiameter / 2), (m_dVerticalDist + m_dTopOffset - (m_dDiameter / 2)), 0        '#7
    oArc2EndPoint.Set 0, (m_dVerticalDist + m_dTopOffset - m_dDiameter), 0        '#8
    
    
    '**********************'
'    If pOrientation.x < 0 Then
'        pOrientation.x = -1 * pOrientation.x
'    End If
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oLine1RightPoint, pOrientation, pStartPoint
    TranslatePoint oLine2TopPoint, pOrientation, pStartPoint
    TranslatePoint oLine3LeftPoint, pOrientation, pStartPoint
    TranslatePoint oLine5TopPoint, pOrientation, pStartPoint
    TranslatePoint oCenterPoint, pOrientation, pStartPoint
    TranslatePoint oArc1EndPoint, pOrientation, pStartPoint
    TranslatePoint oArc2EndPoint, pOrientation, pStartPoint

    '*** Create the lines ***'
    'From 1 to 2 -- Line 1
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine1RightPoint

    'From 2 to 3 -- Line 2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1RightPoint, oLine2TopPoint

    'From 3 to 4 -- Line 3
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2TopPoint, oLine3LeftPoint
    
    'From 4 to 1 -- Line 4
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine3LeftPoint, oStartPoint
    
    'From 4 to 5 -- Line 5
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine3LeftPoint, oLine5TopPoint
    
    '*** Create the Arcs ***'
    'From 5 to 7 -- Arc 1
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oLine5TopPoint, oArc1EndPoint, oCenterPoint
    
    'From 7 to 8 -- Arc 2
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1EndPoint, oArc2EndPoint, oCenterPoint

    '************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPoint = Nothing                    '#1
    Set oLine1RightPoint = Nothing                  '#2
    Set oLine2TopPoint = Nothing               '#3
    Set oLine3LeftPoint = Nothing                  '#4
    Set oLine5TopPoint = Nothing               '#5
    Set oCenterPoint = Nothing               '#6
    Set oArc1EndPoint = Nothing               '#7
    Set oArc2EndPoint = Nothing               '#8

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




