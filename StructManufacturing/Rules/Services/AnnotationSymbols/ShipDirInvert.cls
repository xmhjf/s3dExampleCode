VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShipDirInvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\ShipDirectionInvert.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
' Abstract:
'   Create custom Label Marks
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom Label mark symbol.
'
' History:
' 11/05/2010    Ninad           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.ShipDirectionInvert"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_dNewLineGap                   As Double
Private m_dHorizontalGap                As Double
Private m_sFontName                     As String
Private m_sFontStyle                     As String
Private m_dShipDirectionVOffset                   As Double
Private m_dShipDirectionOffset                   As Double
Private m_dShipDirectionUOffset                   As Double
Private m_dShipNumberOffset                   As Double
Private m_dPartNameOffset                   As Double
Private m_dMarkingSideOffset                   As Double
Private m_dNoOfPartsOffset                   As Double
Private m_dErectionStageOffset                   As Double
Private m_dPlateThicknessOffset                   As Double
Private m_dGradeOffset                   As Double
Private m_dPrimerOffset                   As Double
Private m_dGrinderOffset                   As Double

'Get the NewLine/SameLine Values
Private m_sShipDirectionV                   As String
Private m_sShipDirection                    As String
Private m_sShipDirectionU                    As String
Private m_sShipNumber                    As String
Private m_sPartName                    As String
Private m_sMarkingSide                    As String
Private m_sNoOfParts                    As String
Private m_sErectionStage                    As String
Private m_sPlateThickness                    As String
Private m_sGrade                    As String
Private m_sPrimer                    As String
Private m_sGrinder                    As String

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_dNewLineGap = 0.4
    m_dHorizontalGap = 1
    m_sFontName = "Arial"
    m_sControlPoint = "ul"
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
        'sAttrName = oXMLElement.getAttribute("NAME")
        vTemp = oXMLElement.getAttribute("NAME")
        sAttrName = IIf(VarType(vTemp) = vbString, vTemp, "")
        If Not sAttrName = "" Then
            Select Case sAttrName
            Case "ControlPoint"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                m_sControlPoint = sAttrValue
            Case "TextSize"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dTextSize = CDbl(sAttrValue)
                End If
            Case "NewLineGap"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dNewLineGap = CDbl(sAttrValue)
                End If
            Case "HorizontalGap"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dHorizontalGap = CDbl(sAttrValue)
                End If
            Case "FontName"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_sFontName = sAttrValue
                End If
            Case "FontStyle"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_sFontStyle = CDbl(sAttrValue)
                End If
                
            Case "ShipDirectionV"
                m_sShipDirectionV = oXMLElement.getAttribute("VALUE")
            
            Case "ShipDirectionVOffset"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dShipDirectionVOffset = CDbl(sAttrValue)
                End If
                
            Case "ShipDirection"
                m_sShipDirection = oXMLElement.getAttribute("VALUE")

            Case "ShipDirectionOffset"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dShipDirectionOffset = CDbl(sAttrValue)
                End If
                
            Case "ShipDirectionU"
                 m_sShipDirectionU = oXMLElement.getAttribute("VALUE")

            Case "ShipDirectionUOffset"
                vTemp = oXMLElement.getAttribute("VALUE")
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_dShipDirectionUOffset = sAttrValue
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
    Dim sAttribValues()         As String
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
    Dim oInputDom                As New DOMDocument
    Dim oTextNode               As IXMLDOMElement
    Dim oOutputElem             As IXMLDOMElement
    Dim oCVGTextElem            As IXMLDOMElement
    Dim oSettingElement         As IXMLDOMElement
    Dim oPlateIdElem            As IXMLDOMElement
    Dim oNodeList               As IXMLDOMNodeList
    Dim oAttributeNodeList      As IXMLDOMNodeList
    
    Dim sShipNumber As String
    Dim sPartName As String
    Dim sShipDirectionU As String
    Dim sShipDirectionV As String
    Dim sShipDirU As String
    Dim sShipDirV As String
    Dim sMarkingSide As String
    Dim sPlateThickness As String
    Dim sGrade As String
    Dim sNoOfParts As String
    Dim sGrinder As String
    Dim sPrimer As String
    Dim sErectionStage As String
  
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then GoTo CleanUp
    
    If Not oInputDom.loadXML(sAttributeXML) Then GoTo CleanUp
            
    sShipDirectionU = GetAttributeValueFromXML(sAttributeXML, "USIDE")
    sShipDirU = GetDirectionValue(sShipDirectionU)
    
    sShipDirectionV = GetAttributeValueFromXML(sAttributeXML, "VSIDE")
    sShipDirV = GetDirectionValue(sShipDirectionV)
    
    '*** Create oTextDom for processing ***'
    Dim sTextString As Variant
    
    sTextString = "<SMS_ANNOTATION CONTROL_POINT=""" & m_sControlPoint & """ >"
    sTextString = sTextString & vbNewLine & "<" & m_sShipDirectionU & " ATTRIBUTE=" & """USIDE""" & " TEXT=""" & sShipDirU & """ />"
    sTextString = sTextString & vbNewLine & "<" & m_sShipDirectionV & " ATTRIBUTE=" & """VSIDE""" & " TEXT=""" & sShipDirV & """ />"
    If sShipDirV <> "" And sShipDirU <> "" Then
        sTextString = sTextString & vbNewLine & "<" & m_sShipDirection & " ATTRIBUTE=" & """DEVELOPMENT_DIR""" & " TEXT=""+"" />"
    End If
    sTextString = sTextString & vbNewLine & "</SMS_ANNOTATION>"

    oTextDom.loadXML (sTextString)    
    '**************************************'
    Set oPlateIdElem = oInputDom.selectSingleNode("SMS_OUTPUT_ANNOTATION")

    Set oOutputElem = oOutputDom.createElement("SMS_ANNOTATION")
    oOutputElem.setAttribute "TYPE", oPlateIdElem.getAttribute("TYPE")
    oOutputElem.setAttribute "MARKED_SIDE", "Marking"
    oOutputElem.setAttribute "TEXT_SIZE", m_dTextSize
    oOutputElem.setAttribute "SX", Round(pStartPoint.x, 5)
    oOutputElem.setAttribute "SY", Round(pStartPoint.y, 5)
    
    dummyForm.FontName = m_sFontName
    dummyForm.FontSize = m_dTextSize

    dummyForm.ScaleMode = vbMillimeters
    dConversionFactor = 3.92042723072738

    'normalize the vector
    pOrientation.length = 1
    
    'oTextDom is each node in input (part) XML i.e. SMS_PROD_INFO
    If oTextDom.hasChildNodes Then
        If oTextDom.firstChild.hasChildNodes Then
            iGaplessLines = 1
            iGapLines = 0
            'Determine the number of textnodes to be created and initialize
            '   the dynamic arrays
            iTextNodes = oTextDom.firstChild.childNodes.length
            ReDim oTextPositions(1 To iTextNodes)
            ReDim sTxtValues(1 To iTextNodes)
            ReDim sAttribValues(1 To iTextNodes)
            
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
        
        'Get Attributes i.e. "VSIDE", "USIDE", "MARKED_SIDE" etc
        vTemp = oTextNode.getAttribute("ATTRIBUTE")
        sAttribValues(iCount) = IIf(VarType(vTemp) = vbString, vTemp, "")
        
        'Get the attribute values i.e. U, V, <NoBlk>1.1 etc
        vTemp = oTextNode.getAttribute("TEXT")
        sTxtValues(iCount) = IIf(VarType(vTemp) = vbString, vTemp, "")
                
        If sTxtValues(iCount) <> "" Then
            Select Case oTextNode.nodeName
            Case "NEWLINE"
                oCurrPos.y = oCurrPos.y - (1 + m_dNewLineGap) * m_dTextSize
                oTextPositions(iCount).Set 0, oCurrPos.y, 0
                If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
                dLineWidth = dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
                oCurrPos.x = dLineWidth
            Case "GAPLESS_NEWLINE"
                oCurrPos.y = oCurrPos.y - m_dTextSize
                oTextPositions(iCount).Set 0, oCurrPos.y, 0
                If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
                dLineWidth = dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
                oCurrPos.x = dLineWidth
            Case "SAMELINE"
                oCurrPos.x = oCurrPos.x + m_dHorizontalGap * m_dTextSize
                oTextPositions(iCount).Set oCurrPos.x, oCurrPos.y, oCurrPos.z
                dLineWidth = dLineWidth + dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor _
                             + m_dHorizontalGap * m_dTextSize
                oCurrPos.x = dLineWidth
            Case "GAPLESS_SAMELINE"
                oTextPositions(iCount).Set oCurrPos.x, oCurrPos.y, oCurrPos.z
                dLineWidth = dLineWidth + dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
                oCurrPos.x = dLineWidth
            End Select
        End If
    Next iCount
    
    If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
    
    
    For iCount = 1 To iTextNodes
        Select Case m_sControlPoint 'adjust for justification
        Case "ul" 'in "ul" case do nothing
            oTextPositions(iCount).y = oTextPositions(iCount).y - m_dTextSize
        Case "ll"
            oTextPositions(iCount).y = oTextPositions(iCount).y + dLabelHeight
        Case "lm"
            oTextPositions(iCount).y = oTextPositions(iCount).y + dLabelHeight
            oTextPositions(iCount).x = oTextPositions(iCount).x - dMaxWidth / 2
        Case "lr"
            oTextPositions(iCount).y = oTextPositions(iCount).y + dLabelHeight
            oTextPositions(iCount).x = oTextPositions(iCount).x - dMaxWidth
        Case "ml"
            oTextPositions(iCount).y = oTextPositions(iCount).y + dLabelHeight / 2
        Case "mm"
            oTextPositions(iCount).y = oTextPositions(iCount).y + dLabelHeight / 2
            oTextPositions(iCount).x = oTextPositions(iCount).x - dMaxWidth / 2
        Case "mr"
            oTextPositions(iCount).y = oTextPositions(iCount).y + dLabelHeight / 2
            oTextPositions(iCount).x = oTextPositions(iCount).x - dMaxWidth
        Case "um"
            oTextPositions(iCount).x = oTextPositions(iCount).x - dMaxWidth / 2
            oTextPositions(iCount).y = oTextPositions(iCount).y - m_dTextSize
        Case "ur"
            oTextPositions(iCount).x = oTextPositions(iCount).x - dMaxWidth
            oTextPositions(iCount).y = oTextPositions(iCount).y - m_dTextSize
        End Select
        TranslatePoint oTextPositions(iCount), pOrientation, pStartPoint
        
        'Create the XML Text
        If sTxtValues(iCount) <> "" Then
            Set oCVGTextElem = CreateTextNode(oOutputDom, "ll", oTextPositions(iCount), sTxtValues(iCount), _
                                          m_sFontName, pOrientation, "regular", "identification", _
                                          m_dTextSize, 0, sAttribValues(iCount))
            oOutputElem.appendChild oCVGTextElem
        End If
        
        Set oCVGTextElem = Nothing
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
    If oOutputElem.childNodes.length > 0 Then
        sOutputXML = Replace(oOutputElem.xml, "><", ">" & vbNewLine & "<")
    Else
        sOutputXML = ""
    End If
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

    Exit Function
ErrorHandler:
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    GoTo CleanUp
End Function







