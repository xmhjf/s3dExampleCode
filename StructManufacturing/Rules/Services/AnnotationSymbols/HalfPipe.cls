VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HalfPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\HalfPipe.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\AnnotationSymbols.vbp
'
'
' Abstract:
'   Create custom Pipe Symbol for Profile Landing Curve Mark in Template Drawings
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a Half Pipe Symbol on Profile Landing Curve marks in Template Drawings
'
' History:
' 09/15/2010    Santosh Ram           Created
'***************************************************************************


Option Explicit
Private Const MODULE = "AnnotationSymbols.HalfPipe"


'General Properties
Private m_sControlPoint                 As String
Private m_dOffset                       As Double
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dHorizLineLen                 As Double
Private m_dDiameter                     As Double


Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 30
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dDiameter = m_dTextSize
    m_dHorizLineLen = m_dTextSize * 0.4
    m_dOffset = m_dTextSize * 3
    m_sControlPoint = "ll"
    
End Sub

Private Sub IJDMfgOutputAnnotation_SetArguments(ByVal sSettingsXML As String)
    Const METHOD = "IJDMfgOutputAnnotation_SetArguments"
    On Error GoTo ErrorHandler

    'All the Geom Args of the Annotation are set in this SubRoutine
    
    Dim oXMLDomDoc                  As New DOMDocument
    Dim oAttributeNodeList          As IXMLDOMNodeList
    Dim oXMLElement                 As IXMLDOMElement
    Dim sAttrName                   As String
    Dim sAttrValue                  As String
    Dim vTemp                       As Variant

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
                Case "HorizLineLen"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizLineLen = Val(sAttrValue) * m_dTextSize
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
                Case "Diameter"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_dDiameter = Val(sAttrValue) * m_dTextSize
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
    
    'All the Text Args of the Annotation are set in this Function
             
    Dim oStartPoint                 As IJDPosition          '#1
    Dim oOffsetPoint                As IJDPosition          '#2
    Dim oArc1Point1                 As IJDPosition          '#3
    Dim oArc1Point2                 As IJDPosition          '#4
    Dim oArc2Point2                 As IJDPosition          '#5
    Dim oLine1Point1                As IJDPosition          '#6
    Dim oLine1Point2                As IJDPosition          '#7
    Dim oLine2Point1                As IJDPosition          '#8
    Dim oLine2Point2                As IJDPosition          '#9
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPoint = New DPosition                     '#1
    Set oOffsetPoint = New DPosition                    '#2
    Set oArc1Point1 = New DPosition                     '#3
    Set oArc1Point2 = New DPosition                     '#4
    Set oArc2Point2 = New DPosition                     '#5
    Set oLine1Point1 = New DPosition                    '#6
    Set oLine1Point2 = New DPosition                    '#7
    Set oLine2Point1 = New DPosition                    '#8
    Set oLine2Point2 = New DPosition                    '#9
    
    '*** Set All Points ***'
    m_dOffset = Round(m_dOffset, 1)
    m_dHorizLineLen = Round(m_dHorizLineLen, 1)
    m_dDiameter = Round(m_dDiameter, 1)

    oStartPoint.Set 0, 0, 0                                                 '#1
    oOffsetPoint.Set -1 * m_dOffset, 0, 0                                   '#2
    oArc1Point1.Set -m_dOffset, -m_dDiameter / 2, 0                         '#3
    oArc1Point2.Set (m_dDiameter / 2) - m_dOffset, 0, 0                     '#4
    oArc2Point2.Set -m_dOffset, m_dDiameter / 2, 0                          '#5
    oLine1Point1.Set -m_dOffset, -(m_dDiameter - m_dHorizLineLen) / 2, 0    '#6
    oLine1Point2.Set -m_dOffset, -(m_dDiameter + m_dHorizLineLen) / 2, 0    '#7
    oLine2Point1.Set -m_dOffset, (m_dDiameter + m_dHorizLineLen) / 2, 0     '#8
    oLine2Point2.Set -m_dOffset, (m_dDiameter - m_dHorizLineLen) / 2, 0     '#9
    
     '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oOffsetPoint, pOrientation, pStartPoint
    TranslatePoint oArc1Point1, pOrientation, pStartPoint
    TranslatePoint oArc1Point2, pOrientation, pStartPoint
    TranslatePoint oArc2Point2, pOrientation, pStartPoint
    TranslatePoint oLine1Point1, pOrientation, pStartPoint
    TranslatePoint oLine1Point2, pOrientation, pStartPoint
    TranslatePoint oLine2Point1, pOrientation, pStartPoint
    TranslatePoint oLine2Point2, pOrientation, pStartPoint

    
    
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    '*** Create the Arcs ***'
    'From 3 to 4  Center 1 of arc1
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1Point1, oArc1Point2, oOffsetPoint
  
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1Point2, oArc2Point2, oOffsetPoint
    
    '*** Create the lines ***'
    'From 2 to 3 -- Line 1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1Point2, oLine1Point1

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2Point1, oLine2Point2

    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPoint = Nothing               '#1
    Set oOffsetPoint = Nothing              '#2
    Set oArc1Point1 = Nothing               '#3
    Set oArc1Point2 = Nothing               '#4
    Set oArc2Point2 = Nothing               '#5
    Set oLine1Point1 = Nothing              '#6
    Set oLine1Point2 = Nothing              '#7
    Set oLine2Point1 = Nothing              '#8
    Set oLine2Point2 = Nothing              '#9
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function





