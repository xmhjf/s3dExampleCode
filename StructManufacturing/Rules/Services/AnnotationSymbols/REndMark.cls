VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REndMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\REndMark.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\AnnotationSymbols.vbp
'
'
' Abstract:
'   Create custom Seam Symbols for Rend Mark in Template Drawings
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a REnd Symbol on all Rend marks in Template Drawings
'
' History:
' 09/13/2010    Santosh Ram           Created
'***************************************************************************


Option Explicit
Private Const MODULE = "AnnotationSymbols.REndMark"


'General Properties
Private m_sControlPoint                 As String
Private m_dMarkLength                   As Double
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dRadius                       As Double

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dMarkLength = 30
    m_dMarkLength = 30
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dRadius = m_dMarkLength
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
                Case "MarkLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMarkLength = Val(sAttrValue)
                    End If
                Case "TextSize"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextSize = Val(sAttrValue)
                    End If
                Case "Radius"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dRadius = Val(sAttrValue) * m_dMarkLength
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
    
    'All the Text Args of the Annotation are set in this SubRoutine
             
    Dim oStartPoint                 As IJDPosition          '#1
    Dim oArc1Point1                 As IJDPosition          '#6
    Dim oArc1Point2                 As IJDPosition          '#7
    Dim oArc2Point2                 As IJDPosition          '#8
    Dim oArc3Point2                 As IJDPosition          '#9
    Dim oArc4Point2                 As IJDPosition          '#4
    Dim oCircle2Center              As IJDPosition          '#5
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If

    m_dMarkLength = GetAttributeValueFromXML(sAttributeXML, "TextSize")
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPoint = New DPosition                     '#1
    Set oArc1Point1 = New DPosition                     '#4
    Set oArc1Point2 = New DPosition                     '#5
    Set oArc2Point2 = New DPosition                     '#6
    Set oArc3Point2 = New DPosition                     '#7
    Set oArc4Point2 = New DPosition                     '#8
    Set oCircle2Center = New DPosition                  '#9
    
    '*** Set All Points ***'
    
    m_dRadius = Round(m_dRadius, 1)
    
    oStartPoint.Set -m_dRadius, 0, 0                                     '#1
    oArc1Point1.Set -2 * m_dRadius, 0, 0                          '#2
    oArc1Point2.Set -m_dRadius, -m_dRadius, 0                            '#3
    oArc2Point2.Set 0, 0, 0                             '#4
    oArc3Point2.Set m_dRadius, -m_dRadius, 0                '#5
    oArc4Point2.Set 2 * m_dRadius, 0, 0                         '#6
    oCircle2Center.Set m_dRadius, 0, 0                      '#7
    
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oArc1Point1, pOrientation, pStartPoint
    TranslatePoint oArc1Point2, pOrientation, pStartPoint
    TranslatePoint oArc2Point2, pOrientation, pStartPoint
    TranslatePoint oArc3Point2, pOrientation, pStartPoint
    TranslatePoint oArc4Point2, pOrientation, pStartPoint
    TranslatePoint oCircle2Center, pOrientation, pStartPoint
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    '*** Create the Arcs ***'
    'From 3 to 4  Center 1 of arc1
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1Point1, oArc1Point2, oStartPoint

    'From 4 to 5 Center 1 of arc2
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1Point2, oArc2Point2, oStartPoint

    'From 1 to 8 of arc2
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc2Point2, oArc3Point2, oCircle2Center

    'From 1 to 9 of arc2
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc3Point2, oArc4Point2, oCircle2Center

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPoint = Nothing               '#1
    Set oArc4Point2 = Nothing               '#4
    Set oCircle2Center = Nothing               '#5
    Set oArc1Point1 = Nothing                  '#6
    Set oArc1Point2 = Nothing                  '#7
    Set oArc2Point2 = Nothing                  '#8
    Set oArc3Point2 = Nothing                  '#9
    
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function




