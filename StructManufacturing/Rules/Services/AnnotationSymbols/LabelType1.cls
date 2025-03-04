VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabelType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\LabelType1.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
'
' Abstract:
'   Create custom Label Marks
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom Label mark symbol.
'
' History:
' 05/24/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.LabelType1"

'General Properties
Private m_dTextSize                     As Double
Private m_dNewLineGap                   As Double
Private m_dHorizontalGap                As Double
Private m_sTextFont                     As String


Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_dNewLineGap = 0.4
    m_dHorizontalGap = 1
    m_sTextFont = "Arial"
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
    
    Set oAttributeNodeList = oXMLDomDoc.getElementsByTagName("SMS_OUTPUT_ARG")
    
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
            Case "NewLineGap"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dNewLineGap = Val(sAttrValue)
                End If
            Case "HorizontalGap"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dHorizontalGap = Val(sAttrValue)
                End If
            Case "TextFont"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_sTextFont = sAttrValue
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

    Dim oCurrPos                As IJDPosition
    Dim oTextPositions()        As IJDPosition
    Dim sTxtValues()            As String
    Dim iTextNodes              As Integer
    Dim iCount                  As Integer
    Dim vTemp                   As Variant
    Dim dLineWidth              As Double
    Dim dMaxWidth               As Double
    Dim iGaplessLines           As Integer
    Dim iGapLines               As Integer
    Dim dLabelHeight            As Double
    Dim sControlPoint           As String
    Dim dConversionFactor       As Double
    Dim oTempCurveElement       As IXMLDOMElement

    Dim oOutputDom              As New DOMDocument
    Dim oTextDom                As New DOMDocument
    Dim oTextNode               As IXMLDOMElement
    Dim oOutputElem             As IXMLDOMElement
    Dim oCVGTextElem            As IXMLDOMElement
    Dim oNodeList               As IXMLDOMNodeList
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then GoTo CleanUp
    
    If Not oTextDom.loadXML(sAttributeXML) Then GoTo CleanUp
    
    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    dummyForm.FontName = m_sTextFont
    dummyForm.FontSize = m_dTextSize

    dummyForm.ScaleMode = vbMillimeters
    dConversionFactor = 3.92042723072738

       
    If oTextDom.hasChildNodes Then
        Set oTextNode = oTextDom.firstChild
        'Determine the justification of the symbol
        vTemp = oTextNode.getAttribute("CONTROL_POINT")
        sControlPoint = IIf(VarType(vTemp) = vbString, vTemp, "ul")
        Set oTextNode = Nothing
        If oTextDom.firstChild.hasChildNodes Then
            iGaplessLines = 1
            iGapLines = 0
            'Determine the number of textnodes to be created and initialize
            '   the dynamic arrays
            iTextNodes = oTextDom.firstChild.childNodes.length
            ReDim oTextPositions(1 To iTextNodes)
            ReDim sTxtValues(1 To iTextNodes)
            'Determine the number of lines with no leading gap
            Set oNodeList = oTextDom.selectNodes("//GAPLESS_NEWLINE")
            If Not oNodeList Is Nothing Then
                iGaplessLines = iGaplessLines + oNodeList.length
            End If
            Set oNodeList = Nothing
            'Determine the number of lines with a leading gap
            Set oNodeList = oTextDom.selectNodes("//NEWLINE")
            If Not oNodeList Is Nothing Then
                iGapLines = iGapLines + oNodeList.length
            End If
            Set oNodeList = Nothing
            'Calculate the total height of the label
            dLabelHeight = m_dTextSize * (iGaplessLines + iGapLines * (1 + m_dNewLineGap))
            'Alternate form of above equation:
            'dLabelHeight = (iGaplessLines * m_dTextSize) + (iGapLines * (m_dTextSize + m_dNewLineGap * m_dTextSize))
        Else: GoTo CleanUp
        End If
    Else: GoTo CleanUp
    End If
    
    dMaxWidth = 0
    dLineWidth = 0
    Set oCurrPos = New DPosition
    Set oTextNode = oTextDom.firstChild.firstChild
    Select Case oTextNode.nodeName
        Case "NEWLINE", "GAPLESS_NEWLINE"
            oCurrPos.Set 0, 0, 0
        Case Else
            oCurrPos.Set 0, -(1 + TEXT_ADJUST_BOTTOM) * m_dTextSize, 0
    End Select
    
    'Create all the points for the text nodes in a "ul" justification
    For iCount = 1 To iTextNodes
        Set oTextNode = oTextDom.firstChild.childNodes(iCount - 1)
        Set oTextPositions(iCount) = New DPosition
        vTemp = oTextNode.getAttribute("TEXT")
        sTxtValues(iCount) = IIf(VarType(vTemp) = vbString, vTemp, "")
        Select Case oTextNode.nodeName
        Case "NEWLINE"
            oCurrPos.Y = oCurrPos.Y - (1 + m_dNewLineGap) * m_dTextSize
            oTextPositions(iCount).Set 0, oCurrPos.Y, 0
            If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
            dLineWidth = dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
            oCurrPos.X = dLineWidth
        Case "GAPLESS_NEWLINE"
            oCurrPos.Y = oCurrPos.Y - m_dTextSize
            oTextPositions(iCount).Set 0, oCurrPos.Y, 0
            If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
            dLineWidth = dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
            oCurrPos.X = dLineWidth
        Case "SAMELINE"
            oCurrPos.X = oCurrPos.X + m_dHorizontalGap * m_dTextSize
            oTextPositions(iCount).Set oCurrPos.X, oCurrPos.Y, oCurrPos.Z
            dLineWidth = dLineWidth + dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor _
                         + m_dHorizontalGap * m_dTextSize
            oCurrPos.X = dLineWidth
        Case "GAPLESS_SAMELINE"
            oTextPositions(iCount).Set oCurrPos.X, oCurrPos.Y, oCurrPos.Z
            dLineWidth = dLineWidth + dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
            oCurrPos.X = dLineWidth
        End Select
        
    Next iCount
    
    If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
    
    For iCount = 1 To iTextNodes
        Select Case sControlPoint 'adjust for justification
        Case "ul" 'in "ul" case do nothing
        Case "ll"
            oTextPositions(iCount).Y = oTextPositions(iCount).Y + dLabelHeight
        Case "lm"
            oTextPositions(iCount).Y = oTextPositions(iCount).Y + dLabelHeight
            oTextPositions(iCount).X = oTextPositions(iCount).X - dMaxWidth / 2
        Case "lr"
            oTextPositions(iCount).Y = oTextPositions(iCount).Y + dLabelHeight
            oTextPositions(iCount).X = oTextPositions(iCount).X - dMaxWidth
        Case "ml"
            oTextPositions(iCount).Y = oTextPositions(iCount).Y + dLabelHeight / 2
        Case "mm"
            oTextPositions(iCount).Y = oTextPositions(iCount).Y + dLabelHeight / 2
            oTextPositions(iCount).X = oTextPositions(iCount).X - dMaxWidth / 2
        Case "mr"
            oTextPositions(iCount).Y = oTextPositions(iCount).Y + dLabelHeight / 2
            oTextPositions(iCount).X = oTextPositions(iCount).X - dMaxWidth
        Case "um"
            oTextPositions(iCount).X = oTextPositions(iCount).X - dMaxWidth / 2
        Case "ur"
            oTextPositions(iCount).X = oTextPositions(iCount).X - dMaxWidth
        End Select
        TranslatePoint oTextPositions(iCount), pOrientation, pStartPoint
        'Create the XML Text
        SMS_NodeText oOutputDom, oOutputElem, "ll", oTextPositions(iCount), sTxtValues(iCount), _
                                            m_sTextFont, pOrientation, "regular", "partmon", _
                                            m_dTextSize, 0, "label_annotation"
    Next iCount
    
''For testing purposes, should delete for actual implementation
'Dim oTest1 As IJDPosition
'Dim oTest2 As IJDPosition
'Set oTest1 = New DPosition
'Set oTest2 = New DPosition
'oTest1.Set 0, 0, 0
'oTest2.Set 0, -m_dTextSize, 0
'TranslatePoint oTest1, pOrientation, pStartPoint
'TranslatePoint oTest2, pOrientation, pStartPoint
'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, oTest1, oTest2, "label_annotation")
'If Not oTempCurveElement Is Nothing Then oOutputElem.appendChild oTempCurveElement
'Set oTempCurveElement = Nothing
''End for testing purposes

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oCurrPos = Nothing
    Erase oTextPositions
    Erase sTxtValues
    Set oOutputDom = Nothing
    Set oTextDom = Nothing
    Set oTextNode = Nothing
    Set oOutputElem = Nothing
    Set oCVGTextElem = Nothing
    Set oNodeList = Nothing
'

    Exit Function
ErrorHandler:
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    GoTo CleanUp
End Function



