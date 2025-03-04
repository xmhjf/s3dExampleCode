VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CorrugateMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\CorrugateMark.cls
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
Private Const MODULE = "MfgCustomAnnotation.CorrugateMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dVerticalDist                 As Double
Private m_dMinorHorizDist               As Double
Private m_dMajorHorizDist               As Double
Private m_dSideHorizDist                As Double
Private m_dOffset                       As Double
Private m_sPrefix                       As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dVerticalDist = 1.875 * m_dTextSize
    m_dMinorHorizDist = 1.5 * m_dTextSize
    m_dMajorHorizDist = 3 * m_dTextSize
    m_dSideHorizDist = 1.5 * m_dTextSize
    m_dOffset = 0.333 * m_dTextSize
    m_sControlPoint = "lr"
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
                Case "MinorHorizDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMinorHorizDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "MajorHorizDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMajorHorizDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "SideHorizDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSideHorizDist = Val(sAttrValue) * m_dTextSize
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
 

'        #8$      '
'  #7_______ #6   '   #5 _______ #4
'           \     '     /
'            \    '    /
'             \   '   /
'            #3\__'__/#2
'                 '#1
'                 '
'                 '
'
             
    Dim oCenterPoint           As IJDPosition          '#1
    Dim oLine1RightPoint         As IJDPosition          '#2
    Dim oLine1LeftPoint      As IJDPosition          '#3
    Dim oLine3RightPoint      As IJDPosition          '#4
    Dim oLine3LeftPoint         As IJDPosition          '#5
    Dim oLine4RightPoint         As IJDPosition          '#6
    Dim oLine4LeftPoint         As IJDPosition          '#7
    Dim oTextPoint                  As IJDPosition          '#8
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim strGrinder                As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
    
    
    strGrinder = GetAttributeValueFromXML(sAttributeXML, "PREFIX")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")
    
    'normalize the vector
    pOrientation.length = 1
    
    Set oCenterPoint = New DPosition                    '#1
    Set oLine1RightPoint = New DPosition                  '#2
    Set oLine1LeftPoint = New DPosition               '#3
    Set oLine3RightPoint = New DPosition               '#4
    Set oLine3LeftPoint = New DPosition                  '#5
    Set oLine4RightPoint = New DPosition                  '#6
    Set oLine4LeftPoint = New DPosition                  '#7
    Set oTextPoint = New DPosition                  '#8
    
    '*** Set All Points ***'
    
    oCenterPoint.Set 0, 0, 0                                                                 '#1
    oLine1RightPoint.Set 0, m_dMinorHorizDist / 2, 0                                     '#2
    oLine1LeftPoint.Set 0, -m_dMinorHorizDist / 2, 0                                        '#3
    oLine3RightPoint.Set -m_dVerticalDist, m_dSideHorizDist + (m_dMajorHorizDist / 2), 0               '#4
    oLine3LeftPoint.Set -m_dVerticalDist, m_dMajorHorizDist / 2, 0                                '#5
    oLine4RightPoint.Set -m_dVerticalDist, -m_dMajorHorizDist / 2, 0                                  '#6
    oLine4LeftPoint.Set -m_dVerticalDist, -(m_dSideHorizDist + (m_dMajorHorizDist / 2)), 0        '#7
    oTextPoint.Set -(m_dVerticalDist + m_dOffset), -(m_dMajorHorizDist / 2), 0   '#8
    
    
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oCenterPoint, pOrientation, pStartPoint              '#1
    TranslatePoint oLine1RightPoint, pOrientation, pStartPoint          '#2
    TranslatePoint oLine1LeftPoint, pOrientation, pStartPoint           '#3
    TranslatePoint oLine3RightPoint, pOrientation, pStartPoint          '#4
    TranslatePoint oLine3LeftPoint, pOrientation, pStartPoint            '#5
    TranslatePoint oLine4RightPoint, pOrientation, pStartPoint         '#6
    TranslatePoint oLine4LeftPoint, pOrientation, pStartPoint         '#7
    TranslatePoint oTextPoint, pOrientation, pStartPoint         '#8

    'Create the Output
    
    
    '*** Create the lines ***'
    'From 3 to 2 -- Line 1
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1LeftPoint, oLine1RightPoint

    'From 2 to 5 -- Line 2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1RightPoint, oLine3LeftPoint

    'From 5 to 4 -- Line 3
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine3LeftPoint, oLine3RightPoint
    
    'From 6 to 7 -- Line 4
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine4RightPoint, oLine4LeftPoint
    
    'From 6 to 3 -- Line 5
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine4RightPoint, oLine1LeftPoint
    '************************'

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
    Set oCenterPoint = Nothing                    '#1
    Set oLine1RightPoint = Nothing                  '#2
    Set oLine1LeftPoint = Nothing               '#3
    Set oLine3RightPoint = Nothing                  '#4
    Set oLine3LeftPoint = Nothing               '#5
    Set oLine4RightPoint = Nothing               '#6
    Set oLine4LeftPoint = Nothing               '#7
    Set oTextPoint = Nothing               '#8
  
    
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




