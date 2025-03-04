VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TomariType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\TomariType1.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MFGCustomAnnotation.vbp
'
'
' Abstract:
'   Create custom profile location Tomari symbols
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom mark symbol
'
' History:
' 02/8/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MFGCustomAnnotation.TomariType1"

'General Properties
Private m_dTextSize As Double
Private m_dFlatLineLength As Double
Private m_dVerticalLineLength As Double
Private m_dSlantLineHLength As Double
Private m_dSlantLineVLength As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    m_dTextSize = 40
    m_dFlatLineLength = 4
    m_dVerticalLineLength = 2
    m_dSlantLineHLength = 2
    m_dSlantLineVLength = 0.6
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
        vTemp = Trim(oXMLElement.getAttribute("NAME"))
        sAttrName = IIf(VarType(vTemp) = vbString, vTemp, "")
        If Not sAttrName = "" Then
            Select Case sAttrName
                Case "TextSize"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextSize = Val(sAttrValue)
                    End If
                Case "FlatLineLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlatLineLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "VerticalLineLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalLineLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "SlantLineHLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlantLineHLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "SlantLineVLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlantLineVLength = Val(sAttrValue) * m_dTextSize
                    End If
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

    Dim oStartPos               As IJDPosition
    Dim oFlatStart              As IJDPosition
    Dim oFlatEnd                As IJDPosition
    Dim oSlantStart             As IJDPosition
    Dim oSlantEnd               As IJDPosition
    Dim oVertStart              As IJDPosition
    Dim oVertEnd                As IJDPosition
    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim oTempCurveElement       As IXMLDOMElement
        
    If pStartPoint Is Nothing Or pOrientation Is Nothing Then
        GoTo CleanUp
    End If
    
    
''                |
''                |
''                |
''             \  |
''              \ |
''   Mark  ______\|______
''                |\
''                | \
''                |  \
''                |
''                |
              
    
    'normalize the vector
    pOrientation.length = 1

    Set oStartPos = New DPosition
    Set oFlatStart = New DPosition
    Set oFlatEnd = New DPosition
    Set oVertStart = New DPosition
    Set oVertEnd = New DPosition
    Set oSlantStart = New DPosition
    Set oSlantEnd = New DPosition

    oVertStart.Set 0, 0, 0
    oVertEnd.Set (m_dVerticalLineLength), 0, 0
    oFlatStart.Set -m_dFlatLineLength / 2, 0, 0
    oFlatEnd.Set m_dFlatLineLength / 2, 0, 0
    oSlantStart.Set -m_dSlantLineHLength / 2, -m_dSlantLineVLength / 2, 0
    oSlantEnd.Set m_dSlantLineHLength / 2, m_dSlantLineVLength / 2, 0
    
'    oSlantStart.Set ((m_dVerticalLineLength + m_dSlantLineVLength) / 2#), _
'                    (m_dSlantLineHLength / 2#), 0
'    oSlantEnd.Set ((m_dVerticalLineLength - m_dSlantLineVLength) / 2#), _
'                  -(m_dSlantLineHLength / 2#), 0

    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPos, pOrientation, pStartPoint
    TranslatePoint oFlatStart, pOrientation, pStartPoint
    TranslatePoint oVertStart, pOrientation, pStartPoint
    TranslatePoint oVertEnd, pOrientation, pStartPoint
    TranslatePoint oFlatEnd, pOrientation, pStartPoint
    TranslatePoint oSlantStart, pOrientation, pStartPoint
    TranslatePoint oSlantEnd, pOrientation, pStartPoint

    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlatStart, oFlatEnd
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSlantStart, oSlantEnd

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPos = Nothing
    Set oFlatStart = Nothing
    Set oFlatEnd = Nothing
    Set oSlantStart = Nothing
    Set oSlantEnd = Nothing
    Set oVertStart = Nothing
    Set oVertEnd = Nothing

    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function


