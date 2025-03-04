VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarginType2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MarginType2.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
'
' Abstract:
'   Create custom Margin Marks
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom Margin mark symbol.
'
' History:
' 05/24/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.MarginType2"

'General Properties
Private m_dTextSize                     As Double
Private m_dLargeTotLength               As Double
Private m_dLargeHorizSegLength          As Double
Private m_dTotalVertLength              As Double
Private m_dSmallTotalLength             As Double
Private m_dSmallHorizSegLength          As Double
Private m_dTextDistFromSymbol           As Double
Private m_dLength                       As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_dLargeTotLength = 200
    m_dLargeHorizSegLength = 50
    m_dTotalVertLength = 14
    m_dSmallTotalLength = 0.5
    m_dSmallHorizSegLength = 0.25
    m_dTextDistFromSymbol = 0.6667
    m_dLength = 200
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
                Case "LargeTotalLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dLargeTotLength = Val(sAttrValue)
                    End If
                Case "LargeHorizSegLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dLargeHorizSegLength = Val(sAttrValue)
                    End If
                Case "TotalVertLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTotalVertLength = Val(sAttrValue)
                    End If
                Case "SmallTotalLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSmallTotalLength = Val(sAttrValue)
                    End If
                Case "SmallHorizSegLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSmallHorizSegLength = Val(sAttrValue)
                    End If
                Case "TextDistFromSymbol"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextDistFromSymbol = Val(sAttrValue)
                    End If
                Case "Length"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dLength = Val(sAttrValue)
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
    Dim oSeg3End                As IJDPosition
    Dim oTextPos                As IJDPosition
   
    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oTempEdgeElement       As IXMLDOMElement
'    Dim oCVGTextElem            As IXMLDOMElement
    
    Dim sStage                  As String
    'Dim dLength                 As Double
    
    
    
'    #2                     #4
'    \              #1     /\
' __ ___\___________.__/_______\_________
'          \       /              \
'             \/                     \
'            #3                      #5
    
    
    
    
    
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then GoTo CleanUp
    
    '*** Fill Bevel Values From XML ***'
    
    sStage = GetAttributeValueFromXML(sAttributeXML, "STAGE")
    'dLength = GetAttributeValueFromXML(sAttributeXML, "LENGTH")
    '**********************************'
    
    If m_dLength <= 0 Then GoTo CleanUp

    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    Set oStartPos = New DPosition
    Set oSeg1Start = New DPosition
    Set oSeg2Start = New DPosition
    Set oSeg3Start = New DPosition
    Set oSeg3End = New DPosition
    Set oTextPos = New DPosition
    
    If m_dLength >= 400 Then
        oStartPos.Set 0, 0, 0
        oSeg1Start.Set -m_dLargeHorizSegLength * 2, m_dTotalVertLength / 2, 0
        oSeg2Start.Set -m_dLargeHorizSegLength, -m_dTotalVertLength / 2, 0
        oSeg3Start.Set m_dLargeHorizSegLength, m_dTotalVertLength / 2, 0
        oSeg3End.Set m_dLargeHorizSegLength * 2, -m_dTotalVertLength / 2, 0
    Else
        oStartPos.Set 0, 0, 0
        oSeg1Start.Set -m_dLength / 2 * m_dSmallHorizSegLength * 2, m_dTotalVertLength / 2, 0
        oSeg2Start.Set -m_dLength / 2 * m_dSmallHorizSegLength, -m_dTotalVertLength / 2, 0
        oSeg3Start.Set m_dLength / 2 * m_dSmallHorizSegLength, m_dTotalVertLength / 2, 0
        oSeg3End.Set m_dLength / 2 * m_dSmallHorizSegLength * 2, -m_dTotalVertLength / 2, 0
    End If
    
    oTextPos.Set 0, -m_dTextSize * (1 + m_dTextDistFromSymbol), 0

    TranslatePoint oStartPos, pOrientation, pStartPoint
    TranslatePoint oSeg1Start, pOrientation, pStartPoint
    TranslatePoint oSeg2Start, pOrientation, pStartPoint
    TranslatePoint oSeg3Start, pOrientation, pStartPoint
    TranslatePoint oSeg3End, pOrientation, pStartPoint
    TranslatePoint oTextPos, pOrientation, pStartPoint

    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSeg1Start, oSeg2Start

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSeg2Start, oSeg3Start

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSeg3Start, oSeg3End

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
    
    'Create the text
    'Create the XML Text
    SMS_NodeText oOutputDom, oOutputElem, "ll", oTextPos, sStage, _
                                      "Arial", pOrientation, "regular", "partmon", _
                                      m_dTextSize, 0, "STAGE"
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPos = Nothing
    Set oSeg1Start = Nothing
    Set oSeg2Start = Nothing
    Set oSeg3Start = Nothing
    Set oSeg3End = Nothing
    Set oTextPos = Nothing

    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
        


    Exit Function
ErrorHandler:
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function
