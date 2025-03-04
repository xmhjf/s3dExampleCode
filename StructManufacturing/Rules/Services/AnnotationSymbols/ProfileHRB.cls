VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileHRB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\ProfileHRB.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\AnnotationSymbols.vbp
'
'
' Abstract:
'   Create custom Profile Symbols for Profile Mark in Template Drawings
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a Seam Symbol on all seam marks in Template Drawings
'
' History:
' 09/13/2010    Santosh Ram           Created
'***************************************************************************


Option Explicit
Private Const MODULE = "MfgCustomAnnotation.ProfileHRB"


'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dVericalDist                  As Double
Private m_dRadius                       As Double

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 30
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dRadius = m_dTextSize
    m_dVericalDist = m_dTextSize
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
                Case "Radius"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dRadius = Val(sAttrValue) * m_dTextSize
                    End If
                Case "VerticalDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVericalDist = Val(sAttrValue) * m_dTextSize
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
    Dim oVertLineEndPt              As IJDPosition          '#2
    Dim oCenterPoint                As IJDPosition          '#3
    Dim oArc1Point2                 As IJDPosition          '#4
    Dim oArc3Point2                 As IJDPosition

    
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
    Set oVertLineEndPt = New DPosition                   '#2
    Set oCenterPoint = New DPosition                   '#3
    Set oArc1Point2 = New DPosition                     '#4
    Set oArc3Point2 = New DPosition                     '#4

    
    '*** Set All Points ***'
    
    oStartPoint.Set 0, 0, 0                                '#1
'    oVertLineEndPt.Set 0, m_dVericalDist, 0                '#2
    oVertLineEndPt.Set m_dVericalDist, 0, 0                '#2
    oCenterPoint.Set m_dVericalDist / 2, 0, 0              '#3
    oArc1Point2.Set m_dVericalDist / 2, m_dRadius, 0   '#4
    oArc3Point2.Set m_dVericalDist / 2, -m_dRadius, 0   '#4
    
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oVertLineEndPt, pOrientation, pStartPoint
    TranslatePoint oCenterPoint, pOrientation, pStartPoint
    TranslatePoint oArc1Point2, pOrientation, pStartPoint
    TranslatePoint oArc3Point2, pOrientation, pStartPoint

    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    '*** Create the Arcs ***'
    'From 3 to 4  Center 1 of arc1
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oVertLineEndPt, oArc1Point2, oCenterPoint

    'From 4 to 5 Center 1 of arc2
    SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oArc1Point2, oStartPoint, oCenterPoint
    
    '*** Create the line ***'
    'From 2 to 3 -- Line 1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVertLineEndPt, oStartPoint
    
    
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPoint = Nothing               '#1
    Set oVertLineEndPt = Nothing            '#2
    Set oCenterPoint = Nothing              '#3
    Set oArc1Point2 = Nothing               '#4
 
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function





