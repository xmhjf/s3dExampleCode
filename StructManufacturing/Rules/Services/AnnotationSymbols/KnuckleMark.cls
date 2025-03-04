VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckleMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\KnuckleMark.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\AnnotationSymbols.vbp
'
'
' Abstract:
'   Create custom Knuckle Symbols for Seam Mark in Template Drawings
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a Seam Symbol on all Knuckle marks in Template Drawings
'
' History:
' 09/13/2010    Santosh Ram           Created
'***************************************************************************


Option Explicit
Private Const MODULE = "MfgCustomAnnotation.KnuckleMark"


'General Properties
Private m_sControlPoint                 As String
Private m_dMarkLength                   As Double
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dHorizLength                  As Double
Private m_dAngle                        As Double
Private m_dSlantLength1                 As Double
Private m_dSlantLength2                 As Double


Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dMarkLength = 30
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dHorizLength = m_dMarkLength / 2
    m_dSlantLength1 = m_dMarkLength
    m_dSlantLength2 = m_dMarkLength
    m_dTextSize = 30
    m_dAngle = 45
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
                Case "HorizHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizLength = Val(sAttrValue) * m_dMarkLength
                    End If
                Case "SlantLength1"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlantLength1 = Val(sAttrValue) * m_dMarkLength
                    End If
                Case "SlantLength2"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlantLength2 = Val(sAttrValue) * m_dMarkLength
                    End If
                Case "Angle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dAngle = Val(sAttrValue)
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
    
    '        #2     #5    #6
    '         \    /\    /
    '          \  /  \  /
    '  _______#1\/____\/#8______ Mark
    '           /\    /\
    '          /  \  /  \
    '         /    \/    \
    '        #4     #3    #7
    '
             
    Dim oStartPoint                 As IJDPosition          '#1
    Dim oLine1Point1                As IJDPosition          '#2
    Dim oLine1Point2                As IJDPosition          '#3
    Dim oLine2Point1                As IJDPosition          '#4
    Dim oLine2Point2                As IJDPosition          '#5
    Dim oLine3Point2                As IJDPosition          '#6
    Dim oLine4Point2                As IJDPosition          '#7
    Dim oRefPoint                   As IJDPosition          '#8
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    m_dMarkLength = GetAttributeValueFromXML(sAttributeXML, "MarkLength")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPoint = New DPosition                     '#1
    Set oLine1Point1 = New DPosition                    '#4
    Set oLine1Point2 = New DPosition                    '#5
    Set oLine2Point1 = New DPosition                    '#6
    Set oLine2Point2 = New DPosition                    '#7
    Set oLine3Point2 = New DPosition                    '#8
    Set oLine4Point2 = New DPosition                    '#9
    Set oRefPoint = New DPosition                       '#10
    
    '*** Set All Points ***'
    
    m_dHorizLength = Round(m_dHorizLength, 1)
    m_dSlantLength2 = Round(m_dSlantLength2, 1)
    m_dSlantLength1 = Round(m_dSlantLength1, 1)
    
    oStartPoint.Set -m_dHorizLength / 2, 0, 0                                   '#1
    oLine1Point1.Set oStartPoint.X + Round(-(m_dSlantLength1 / 2) * Cos(m_dAngle), 1), oStartPoint.Y + Round((m_dSlantLength1 / 2) * Sin(m_dAngle), 1), 0 '#4
    oLine1Point2.Set oStartPoint.X + Round((m_dSlantLength1 / 2) * Cos(m_dAngle), 1), oStartPoint.Y + Round(-(m_dSlantLength1 / 2) * Sin(m_dAngle), 1), 0     '#5
    oLine2Point1.Set oStartPoint.X + Round(-(m_dSlantLength1 / 2) * Sin(m_dAngle), 1), oStartPoint.Y + Round(-(m_dSlantLength1 / 2) * Cos(m_dAngle), 1), 0                           '#6
    oLine2Point2.Set oStartPoint.X + Round((m_dSlantLength1 / 2) * Sin(m_dAngle), 1), oStartPoint.Y + Round((m_dSlantLength1 / 2) * Cos(m_dAngle), 1), 0                '#7
    oRefPoint.Set m_dHorizLength / 2, 0, 0
    oLine3Point2.Set (oRefPoint.X + Round((m_dSlantLength2 / 2) * Sin(m_dAngle), 1)), oRefPoint.Y + (Round((m_dSlantLength2 / 2) * Cos(m_dAngle), 1)), 0 '#8
    oLine4Point2.Set (oRefPoint.X + Round((m_dSlantLength2 / 2) * Sin(m_dAngle), 1)), oRefPoint.Y - (Round((m_dSlantLength2 / 2) * Cos(m_dAngle), 1)), 0                     '#9
    
    '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oLine1Point1, pOrientation, pStartPoint
    TranslatePoint oLine1Point2, pOrientation, pStartPoint
    TranslatePoint oLine2Point1, pOrientation, pStartPoint
    TranslatePoint oLine2Point2, pOrientation, pStartPoint
    TranslatePoint oLine3Point2, pOrientation, pStartPoint
    TranslatePoint oLine4Point2, pOrientation, pStartPoint
    TranslatePoint oRefPoint, pOrientation, pStartPoint

    
    
    '*** Create the lines ***'
       
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1Point1, oLine1Point2
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2Point1, oLine2Point2
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1Point2, oLine3Point2
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2Point2, oLine4Point2

    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPoint = Nothing       '#1
    Set oLine3Point2 = Nothing      '#4
    Set oLine4Point2 = Nothing      '#5
    Set oLine1Point1 = Nothing      '#6
    Set oLine1Point2 = Nothing      '#7
    Set oLine2Point1 = Nothing      '#8
    Set oLine2Point2 = Nothing      '#9
    Set oRefPoint = Nothing
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function




