VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MarginType1.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
'
' Abstract:
'   Create custom Seam Control Mark
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom Margin mark symbol.
'
' History:
' 06/02/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.SeamControl1"

'General Properties
Private m_dTextSize                     As Double
Private m_dDiagonalLength                   As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_dDiagonalLength = 1.5
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
                Case "TextSize"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextSize = Val(sAttrValue)
                    End If
                Case "DiagonalLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dDiagonalLength = m_dTextSize * Val(sAttrValue)
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
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Sub

Private Function IJDMfgOutputAnnotation_Evaluate(ByVal pStartPoint As IJDPosition, ByVal pOrientation As IJDVector, ByVal sAttributeXML As String) As String
    Const METHOD = "IJDMfgOutputAnnotation_Evaluate"
    On Error GoTo ErrorHandler

    Dim oStartPos               As IJDPosition
    Dim oSeg1Start              As IJDPosition
    Dim oSeg2Start              As IJDPosition
    Dim oSeg3Start              As IJDPosition
    Dim oSeg4start              As IJDPosition

    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim oCVGTextElem            As IXMLDOMElement
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then GoTo CleanUp

    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPos = New DPosition
    Set oSeg1Start = New DPosition
    Set oSeg2Start = New DPosition
    Set oSeg3Start = New DPosition
    Set oSeg4start = New DPosition
    
    oStartPos.Set 0, 0, 0
    oSeg1Start.Set m_dDiagonalLength, 0, 0
    oSeg2Start.Set 0, -m_dDiagonalLength, 0
    oSeg3Start.Set -m_dDiagonalLength, 0, 0
    oSeg4start.Set 0, m_dDiagonalLength, 0
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""

    TranslatePoint oStartPos, pOrientation, pStartPoint
    TranslatePoint oSeg1Start, pOrientation, pStartPoint
    TranslatePoint oSeg2Start, pOrientation, pStartPoint
    TranslatePoint oSeg3Start, pOrientation, pStartPoint
    TranslatePoint oSeg4start, pOrientation, pStartPoint

    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSeg1Start, oSeg2Start

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSeg2Start, oSeg3Start

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSeg3Start, oSeg4start

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSeg4start, oSeg1Start

''For testing purposes, should comment or delete for actual implementation
'Dim oSampleMarkLeft As IJDPosition
'Dim oSampleMarkRight As IJDPosition
'Set oSampleMarkLeft = New DPosition
'Set oSampleMarkRight = New DPosition
'oSampleMarkLeft.Set -dLength / 2, 0, 0
'oSampleMarkRight.Set dLength / 2, 0, 0
'TranslatePoint oSampleMarkLeft, pOrientation, pStartPoint
'TranslatePoint oSampleMarkRight, pOrientation, pStartPoint
'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, oSampleMarkLeft, oSampleMarkRight, "margin_annotation")
'If Not oTempCurveElement Is Nothing Then oOutputElem.appendChild oTempCurveElement
'Set oTempCurveElement = Nothing
''End for testing purposes

    Set oTempEdgeElement = Nothing
    
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPos = Nothing
    Set oSeg1Start = Nothing
    Set oSeg2Start = Nothing
    Set oSeg3Start = Nothing
    Set oSeg4start = Nothing
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oCVGTextElem = Nothing


    Exit Function
ErrorHandler:
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function


