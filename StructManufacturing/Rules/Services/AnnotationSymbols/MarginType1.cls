VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarginType1"
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
'   Create Default Margin Marks
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   the default Margin mark symbol.
'
' History:
' 06/14/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.MarginType1"

'General Properties
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_dSymbolHeight                 As Double
Private m_dSymbolWidth                  As Double
Private m_dTextGap                      As Double
Private m_sControlPoint                 As String
Private m_sTextStyle                    As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_dSymbolHeight = 1
    m_dSymbolWidth = 1
    m_dTextGap = 0.2
    m_sTextFont = "Arial"
    m_sControlPoint = "bl"
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
                    
                Case "ControlPoint"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sControlPoint = sAttrValue
                    
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                    
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
                    
                Case "SymbolHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSymbolHeight = Val(sAttrValue)
                    End If
                    
                Case "SymbolWidth"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSymbolWidth = Val(sAttrValue)
                    End If
                    
                Case "TextGap"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextGap = Val(sAttrValue)
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

    Dim oSymb1Start             As IJDPosition
    Dim oSymb2Start             As IJDPosition
    Dim oSymb1TopVertex         As IJDPosition
    Dim oSymb2TopVertex         As IJDPosition
    Dim oSymb1BottomVertex      As IJDPosition
    Dim oSymb2BottomVertex      As IJDPosition
    Dim oSymb1TextPos           As IJDPosition
    Dim oSymb2TextPos           As IJDPosition
    Dim oVector1                As IJDVector

    Dim oOutputDom              As New DOMDocument
    
    Dim oOutputElem             As IXMLDOMElement
    Dim oOutputElem2            As IXMLDOMElement
    Dim oTempEdgeElement       As IXMLDOMElement
    Dim oCVGTextElem            As IXMLDOMElement
    
    Dim sOutputXML              As String
    
    Dim dStartVal               As Double
    Dim dExtraRot               As Double
    Dim dEndVal                 As Double
    Dim dX1                     As Double
    Dim dY1                     As Double
    Dim dX2                     As Double
    Dim dY2                     As Double
    
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then GoTo CleanUp
    
    'FillDefaultMarginValuesFromXML sAttributeXML, strGUID, dStartVal, dEndVal, dX1, dY1, dX2, dY2
    
    '*** Fill Bevel Values From XML ***'
    dStartVal = GetAttributeValueFromXML(sAttributeXML, "START_VALUE")
    dEndVal = GetAttributeValueFromXML(sAttributeXML, "END_VALUE")
    dX1 = GetAttributeValueFromXML(sAttributeXML, "SX")
    dY1 = GetAttributeValueFromXML(sAttributeXML, "SY")
    dX2 = GetAttributeValueFromXML(sAttributeXML, "EX")
    dY2 = GetAttributeValueFromXML(sAttributeXML, "EY")
    '**********************************'
    
    'normalize the vector
    pOrientation.length = 1
    
    dExtraRot = -PI / 2
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    If Abs(dEndVal - dStartVal) < 0.00001 Or dEndVal < 0.000001 Then
        'Single Margin Symbol

        Set oSymb1Start = New DPosition
        Set oSymb1TopVertex = New DPosition
        Set oSymb1BottomVertex = New DPosition
        Set oSymb1TextPos = New DPosition
        
        oSymb1Start.Set 0, 0, 0
        oSymb1TopVertex.Set m_dSymbolHeight * m_dTextSize, m_dSymbolWidth * m_dTextSize / 2, 0
        oSymb1BottomVertex.Set m_dSymbolHeight * m_dTextSize, -m_dSymbolWidth * m_dTextSize / 2, 0
        oSymb1TextPos.Set (m_dSymbolHeight + m_dTextGap) * m_dTextSize, 0, 0
        
        TranslatePoint oSymb1Start, pOrientation, pStartPoint
        TranslatePoint oSymb1TopVertex, pOrientation, pStartPoint
        TranslatePoint oSymb1BottomVertex, pOrientation, pStartPoint
        TranslatePoint oSymb1TextPos, pOrientation, pStartPoint
        
        'Create the XML Lines
        'Create the lines
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb1Start, oSymb1TopVertex
    
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb1TopVertex, oSymb1BottomVertex

        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb1BottomVertex, oSymb1Start
        
        'Create the XML Text
        SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oSymb1TextPos, Round(dStartVal), _
                                           m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                           m_dTextSize, dExtraRot, "START_VALUE"
        
        If oOutputElem.childNodes.length > 0 Then
            sOutputXML = Replace(oOutputElem.xml, "><", ">" & vbNewLine & "<")
        Else
            sOutputXML = ""
        End If
        IJDMfgOutputAnnotation_Evaluate = sOutputXML
        GoTo CleanUp
    Else
        'Two Symbols

        Set oOutputElem2 = oOutputElem.cloneNode(True)
        oOutputElem2.setAttribute "SX", Round(dX2, 5)
        oOutputElem2.setAttribute "SY", Round(dY2, 5)

        Set oVector1 = New DVector
        Set oSymb1Start = New DPosition
        Set oSymb1TopVertex = New DPosition
        Set oSymb1BottomVertex = New DPosition
        Set oSymb1TextPos = New DPosition
        Set oSymb2Start = New DPosition
        
        'When 2 symbols are needed we use the input point as a the first orientation vector
        '   the input vector as the second orientation vector and the CVG_POINT values
        '   from the SMS_PART_MARGIN_INFO node as the start points for the 2 symbols
        oVector1.Set pStartPoint.X, pStartPoint.Y, 0
        oVector1.length = 1 'Normalize
        
        'Don't Translate these, use them for translation:
        oSymb1Start.Set dX1, dY1, 0
        oSymb2Start.Set dX2, dY2, 0
        
        'Translate these:
        oSymb1TopVertex.Set m_dSymbolHeight * m_dTextSize, m_dSymbolWidth * m_dTextSize / 2, 0
        oSymb1BottomVertex.Set m_dSymbolHeight * m_dTextSize, -m_dSymbolWidth * m_dTextSize / 2, 0
        oSymb1TextPos.Set (m_dSymbolHeight + m_dTextGap) * m_dTextSize, 0, 0
        Set oSymb2TopVertex = oSymb1TopVertex.Clone
        Set oSymb2BottomVertex = oSymb1BottomVertex.Clone
        Set oSymb2TextPos = oSymb1TextPos.Clone
        
        'First symbol
        TranslatePoint oSymb1TopVertex, oVector1, oSymb1Start
        TranslatePoint oSymb1BottomVertex, oVector1, oSymb1Start
        TranslatePoint oSymb1TextPos, oVector1, oSymb1Start
        
        'Second symbol
        TranslatePoint oSymb2TopVertex, pOrientation, oSymb2Start
        TranslatePoint oSymb2BottomVertex, pOrientation, oSymb2Start
        TranslatePoint oSymb2TextPos, pOrientation, oSymb2Start
        
        'Symbol 1
        'Create the XML Lines
        'Create the lines
       
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb1Start, oSymb1TopVertex

        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb1TopVertex, oSymb1BottomVertex
        
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb1BottomVertex, oSymb1Start
        
        'Create the XML Text
        SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oSymb1TextPos, Round(dStartVal), _
                                           m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                           m_dTextSize, dExtraRot, "START_VALUE"
        
        'Symbol 2
        'Create the XML Lines
        'Create the lines

        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb2Start, oSymb2TopVertex

        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb2TopVertex, oSymb2BottomVertex

        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSymb2BottomVertex, oSymb2Start
        
        'Create the XML Text
        SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oSymb2TextPos, Round(dEndVal), _
                                           m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                           m_dTextSize, dExtraRot, "END_VALUE"

        If oOutputElem.childNodes.length > 0 And oOutputElem2.childNodes.length > 0 Then
            sOutputXML = Replace(oOutputElem.xml, "><", ">" & vbNewLine & "<") & vbNewLine & _
                         Replace(oOutputElem2.xml, "><", ">" & vbNewLine & "<") & vbNewLine
        Else
            sOutputXML = ""
        End If
        IJDMfgOutputAnnotation_Evaluate = sOutputXML
    End If
    
    

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

CleanUp:
    Set oSymb1Start = Nothing
    Set oSymb2Start = Nothing
    Set oSymb1TopVertex = Nothing
    Set oSymb2TopVertex = Nothing
    Set oSymb1BottomVertex = Nothing
    Set oSymb2BottomVertex = Nothing
    Set oSymb1TextPos = Nothing
    Set oSymb2TextPos = Nothing
    Set oVector1 = Nothing

    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oOutputElem2 = Nothing
    Set oTempEdgeElement = Nothing
    Set oCVGTextElem = Nothing


    Exit Function
ErrorHandler:
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function


