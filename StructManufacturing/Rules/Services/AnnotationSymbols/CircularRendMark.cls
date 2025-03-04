VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CircularRendMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\CircularRendMark.cls
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
' 11/12/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.CircularRendMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dDiameter                      As Double

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dDiameter = m_dTextSize
    m_sControlPoint = "ur"
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
                Case "Diameter"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dDiameter = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextSize"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextSize = Val(sAttrValue)
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
    
    Dim oStartPoint               As IJDPosition          '#1
    Dim oCircle1BottomPoint       As IJDPosition          '#2
    Dim oCircle1RightPoint       As IJDPosition          '#3
    Dim oCircle1TopPoint            As IJDPosition          '#4
    Dim oCircle1CenterPoint              As IJDPosition          '#5
    Dim oCircle2TopPoint            As IJDPosition          '#6
    Dim oCircle2LeftPoint       As IJDPosition          '#7
    Dim oCircle2BottomPoint       As IJDPosition          '#8
    Dim oCircle2CenterPoint              As IJDPosition          '#9
        
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim strTYPE                 As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
    
    strTYPE = GetAttributeValueFromXML(sAttributeXML, "REND_TYPE")
    
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPoint = New DPosition             '#1
    Set oCircle1BottomPoint = New DPosition               '#2
    Set oCircle1RightPoint = New DPosition               '#3
    Set oCircle1TopPoint = New DPosition                    '#4
    Set oCircle1CenterPoint = New DPosition                      '#5
    Set oCircle2TopPoint = New DPosition                      '#6
    Set oCircle2LeftPoint = New DPosition                      '#7
    Set oCircle2BottomPoint = New DPosition                      '#8
    Set oCircle2CenterPoint = New DPosition                      '#9
    
    '*** Set All Points ***'
    oStartPoint.Set 0, 0, 0                     '#1
    oCircle1BottomPoint.Set m_dDiameter / 2, -m_dDiameter / 2, 0     '#2
    oCircle1RightPoint.Set m_dDiameter, 0, 0     '#3
    oCircle1TopPoint.Set m_dDiameter / 2, m_dDiameter / 2, 0           '#4
    oCircle1CenterPoint.Set m_dDiameter / 2, 0, 0          '#5
    oCircle2TopPoint.Set -m_dDiameter / 2, m_dDiameter / 2, 0         '#6
    oCircle2LeftPoint.Set -m_dDiameter, 0, 0        '#7
    oCircle2BottomPoint.Set -m_dDiameter / 2, -m_dDiameter / 2, 0      '#8
    oCircle2CenterPoint.Set -m_dDiameter / 2, 0, 0      '#9
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    'Still need to add the rend type as an attribute on the SMS_ANNOTATION XML node!
    If Not oOutputElem Is Nothing Then
        oOutputElem.setAttribute "REND_TYPE", strTYPE
    End If
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint       '1
    TranslatePoint oCircle1BottomPoint, pOrientation, pStartPoint         '2
    TranslatePoint oCircle1RightPoint, pOrientation, pStartPoint         '3
    TranslatePoint oCircle1TopPoint, pOrientation, pStartPoint          '4
    TranslatePoint oCircle1CenterPoint, pOrientation, pStartPoint       '5
    TranslatePoint oCircle2TopPoint, pOrientation, pStartPoint      '6
    TranslatePoint oCircle2LeftPoint, pOrientation, pStartPoint         '7
    TranslatePoint oCircle2BottomPoint, pOrientation, pStartPoint       '8
    TranslatePoint oCircle2CenterPoint, pOrientation, pStartPoint       '9
    
    'Creating the Output
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    Select Case strTYPE
        Case "Type1"
        'Create the Arcs
            'Arc1 -- 1 to 2 Center 5
                       
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oStartPoint, oCircle1BottomPoint, oCircle1CenterPoint
            
            'Arc2 -- 2 to 3 Center 5
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1BottomPoint, oCircle1RightPoint, oCircle1CenterPoint
            
            'Arc7 -- 7 to 8 Center 9
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle2LeftPoint, oCircle2BottomPoint, oCircle2CenterPoint
            
            'Arc8 -- 8 to 1 Center 9
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle2BottomPoint, oStartPoint, oCircle2CenterPoint
            
            'Arc3 -- 3 to 4 Center 5
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1RightPoint, oCircle1TopPoint, oCircle1CenterPoint
            
            'Arc4 -- 4 to 1 Center 5
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1TopPoint, oStartPoint, oCircle1CenterPoint
            
            'Arc5 -- 1 to 6 Center 9
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oStartPoint, oCircle2TopPoint, oCircle2CenterPoint
            
            'Arc6 -- 6 to 7 Center 9
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle2TopPoint, oCircle2LeftPoint, oCircle2CenterPoint
            
        Case "Type2"
            'Create the Arcs
            'Arc1 -- 3 to 4 Center 5
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1RightPoint, oCircle1TopPoint, oCircle1CenterPoint
            
            'Arc2 -- 4 to 1 Center 5
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle1TopPoint, oStartPoint, oCircle1CenterPoint
            
            'Arc5 -- 1 to 6 Center 9
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oStartPoint, oCircle2TopPoint, oCircle2CenterPoint
            
            'Arc6 -- 6 to 7 Center 9
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oCircle2TopPoint, oCircle2LeftPoint, oCircle2CenterPoint
        
    End Select
    
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    

CleanUp:
    Set oStartPoint = Nothing
    Set oCircle1RightPoint = Nothing
    Set oCircle1BottomPoint = Nothing
    Set oCircle1TopPoint = Nothing
    Set oCircle1CenterPoint = Nothing
    Set oCircle2BottomPoint = Nothing
    Set oCircle2CenterPoint = Nothing
    Set oCircle2LeftPoint = Nothing
    Set oCircle2TopPoint = Nothing
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function
    
    
    
