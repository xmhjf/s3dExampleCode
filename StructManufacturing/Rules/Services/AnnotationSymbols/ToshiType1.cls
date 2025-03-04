VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToshiType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\ToshiType1.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MFGCustomAnnotation.vbp
'
'
' Abstract:
'   Create custom profile location Toshi symbols
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom mark symbol
'
' History:
' 02/8/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MFGCustomAnnotation.ToshiType1"

'General Properties
Private m_dTextSize                 As Double
Private m_dFlatLineLength           As Double
Private m_dSlantLineHLength         As Double
Private m_dSlantLineVLength         As Double

Private m_sText                     As String
Private m_sControlPoint             As String
Private m_sTextFont                 As String
Private m_sTextStyle                As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    m_dTextSize = 40
    m_dFlatLineLength = 4
    m_dSlantLineVLength = 1
    m_dSlantLineHLength = 0.6
    m_sText = "Default"
    m_sControlPoint = "ll"
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
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
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_sTextFont = sAttrValue
                    End If
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
                Case "FlatLineLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlatLineLength = Val(sAttrValue)
                    End If
                Case "SlantLineVLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlantLineVLength = Val(sAttrValue)
                    End If
                Case "SlantLineHLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlantLineHLength = Val(sAttrValue)
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
    '               '#2   /#4
    '               '   /
    '               '  /
    '               ' /
    '_______________'/_________________________
    '               /#1
    '              /'
    '             / '
    '          #5/  '#3
    Dim oStartPos               As IJDPosition             '#1
    Dim oFlatStart              As IJDPosition             '#2
    Dim oFlatEnd                As IJDPosition             '#3
    Dim oSlantStart             As IJDPosition             '#4
    Dim oSlantEnd               As IJDPosition             '#5
    
    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim oTempCurveElement       As IXMLDOMElement
    
    Dim strText1                As String
    Dim oCVGTextElem1           As IXMLDOMElement
    
'    Dim strTYPE                 As String
    
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Then
        GoTo CleanUp
    End If
    
    strText1 = GetAttributeValueFromXML(sAttributeXML, "MARKING_NAME")
    '    m_sControlPoint = GetAttributeValueFromXML(sAttributeXML, "ControlPoint")


    'Create the main

    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""

    Set oStartPos = New DPosition
    Set oFlatStart = New DPosition
    Set oFlatEnd = New DPosition
    Set oSlantStart = New DPosition
    Set oSlantEnd = New DPosition
    

    oStartPos.Set 0, 0, 0
    oFlatStart.Set 0, m_dFlatLineLength * m_dTextSize / 2#, 0
    oFlatEnd.Set 0, -(m_dFlatLineLength * m_dTextSize / 2#), 0
    oSlantStart.Set (m_dSlantLineHLength * m_dTextSize / 2#), _
                    (m_dSlantLineVLength * m_dTextSize / 2#), 0
    oSlantEnd.Set -(m_dSlantLineHLength * m_dTextSize / 2#), _
                  -(m_dSlantLineVLength * m_dTextSize / 2#), 0
    

    TranslatePoint oStartPos, pOrientation, pStartPoint
    TranslatePoint oFlatStart, pOrientation, pStartPoint
    TranslatePoint oFlatEnd, pOrientation, pStartPoint
    TranslatePoint oSlantStart, pOrientation, pStartPoint
    TranslatePoint oSlantEnd, pOrientation, pStartPoint
    

    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlatStart, oFlatEnd

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSlantStart, oSlantEnd

    
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oStartPos, strText1, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 0, "MARKING_NAME"

    
    
    Dim sOutputXML As String
    sOutputXML = Replace(oOutputElem.xml, "><", ">" & vbNewLine & "<")
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPos = Nothing
    Set oFlatStart = Nothing
    Set oFlatEnd = Nothing
    Set oSlantStart = Nothing
    Set oSlantEnd = Nothing
    

    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function
