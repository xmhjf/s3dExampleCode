VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Centerlinemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\Centerlinemark.cls
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
Private Const MODULE = "MfgCustomAnnotation.CenterLineMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dDiameter                     As Double
Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 30
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dDiameter = m_dTextSize
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
                Case "Diameter"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dDiameter = Val(sAttrValue) * m_dTextSize
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
             
    Dim oCenterPoint           As IJDPosition          '#1
    Dim oCircle1Point1         As IJDPosition          '#2
    Dim oCircle1Point2         As IJDPosition          '#3
    Dim oCircle1Point3         As IJDPosition          '#4
    Dim oCircle1Point4         As IJDPosition          '#5
    Dim oArc1Point1            As IJDPosition          '#6
    Dim oArc1Point2            As IJDPosition          '#7
    Dim oArc2Point1            As IJDPosition          '#8
    Dim oArc2Point2            As IJDPosition          '#9
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    
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
    
    Set oCenterPoint = New DPosition                    '#1
    Set oCircle1Point1 = New DPosition                  '#2
    Set oCircle1Point2 = New DPosition                  '#3
    Set oCircle1Point3 = New DPosition                  '#4
    Set oCircle1Point4 = New DPosition                  '#5
    Set oArc1Point1 = New DPosition                     '#6
    Set oArc1Point2 = New DPosition                     '#7
    Set oArc2Point1 = New DPosition                     '#8
    Set oArc2Point2 = New DPosition                     '#9
    
    '*** Set All Points ***'
    
    oCenterPoint.Set 0, 0, 0                                      '#1
    oCircle1Point1.Set -m_dDiameter / 2, 0, 0                     '#2
    oCircle1Point2.Set 0, m_dDiameter / 2, 0                      '#3
    oCircle1Point3.Set m_dDiameter / 2, 0, 0                      '#4
    oCircle1Point4.Set 0, -m_dDiameter / 2, 0                     '#5
    oArc1Point1.Set m_dDiameter / 2, m_dDiameter / 2, 0           '#6
    oArc1Point2.Set -m_dDiameter / 2, m_dDiameter / 2, 0          '#7
    oArc2Point1.Set m_dDiameter / 2, -m_dDiameter / 2, 0          '#8
    oArc2Point2.Set -m_dDiameter / 2, -m_dDiameter / 2, 0         '#9
    
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oCenterPoint, pOrientation, pStartPoint
    
    TranslatePoint oCircle1Point1, pOrientation, pStartPoint
    TranslatePoint oCircle1Point2, pOrientation, pStartPoint
    TranslatePoint oCircle1Point3, pOrientation, pStartPoint
    TranslatePoint oCircle1Point4, pOrientation, pStartPoint
    
    TranslatePoint oArc1Point1, pOrientation, pStartPoint
    TranslatePoint oArc1Point2, pOrientation, pStartPoint

    TranslatePoint oArc2Point1, pOrientation, pStartPoint
    TranslatePoint oArc2Point2, pOrientation, pStartPoint

    'Creating the Output
    
  
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    '*** Create the Curves ***'
    'From 3 to 2 of circle
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1Point2, oCircle1Point1, oCenterPoint

    'From 4 to 3 of circle
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1Point3, oCircle1Point2, oCenterPoint
    
    'From 5 to 4 of circle
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1Point4, oCircle1Point3, oCenterPoint

    'From 2 to 5 of circle
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1Point1, oCircle1Point4, oCenterPoint


    '*** Create the Arcs ***'
    'From 1 to 6 of arc1
     SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCenterPoint, oArc1Point1, oCircle1Point2
     
     'From 1 to 7 of arc1
     SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1Point2, oCenterPoint, oCircle1Point2

     'From 1 to 8 of arc2
     SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc2Point1, oCenterPoint, oCircle1Point4

     'From 1 to 9 of arc2
     SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCenterPoint, oArc2Point2, oCircle1Point4
     
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oCenterPoint = Nothing                 '#1
    Set oCircle1Point1 = Nothing               '#2
    Set oCircle1Point2 = Nothing               '#3
    Set oCircle1Point3 = Nothing               '#4
    Set oCircle1Point4 = Nothing               '#5
    Set oArc1Point1 = Nothing                  '#6
    Set oArc1Point2 = Nothing                  '#7
    Set oArc2Point1 = Nothing                  '#8
    Set oArc2Point2 = Nothing                  '#9
    
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    
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


















