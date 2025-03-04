VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileId1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\ProfileId1.cls
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
Private Const MODULE = "MfgCustomAnnotation.ProfileId1"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_dNewLineGap                   As Double
Private m_dHorizontalGap                As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dShipDirectionVOffset                   As Double
Private m_dShipDirectionOffset                   As Double
Private m_dShipDirectionUOffset                   As Double
Private m_dShipNumberOffset                   As Double
Private m_dPartNameOffset                   As Double
Private m_dMarkingSideOffset                   As Double
Private m_dNoOfPartsOffset                   As Double
Private m_dErectionStageOffset                   As Double
Private m_dProfileThicknessOffset                   As Double
Private m_dGradeOffset                   As Double
Private m_dPrimerOffset                   As Double
Private m_dGrinderOffset                   As Double
Private m_dCurvatureOffset                   As Double

'Get the NewLine/SameLine Values
Private m_sShipDirectionV                   As String
Private m_sShipDirection                    As String
Private m_sShipDirectionU                    As String
Private m_sShipNumber                    As String
Private m_sPartName                    As String
Private m_sMarkingSide                    As String
Private m_sNoOfParts                    As String
Private m_sErectionStage                    As String
Private m_sProfileThickness                    As String
Private m_sGrade                    As String
Private m_sCurvature                    As String
Private m_sPrimer                    As String
Private m_sGrinder                    As String


Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_dNewLineGap = 0.4
    m_dHorizontalGap = 1
    m_sTextFont = "Arial"
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
        'sAttrName = trim(oXMLElement.getAttribute("NAME"))
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
            Case "TextStyle"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_sTextStyle = Val(sAttrValue)
                End If
                
            Case "ShipDirectionV"
                m_sShipDirectionV = Trim(oXMLElement.getAttribute("VALUE"))
            
            Case "ShipDirectionVOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dShipDirectionVOffset = Val(sAttrValue)
                End If
                
            Case "ShipDirection"
                m_sShipDirection = Trim(oXMLElement.getAttribute("VALUE"))

            Case "ShipDirectionOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dShipDirectionOffset = Val(sAttrValue)
                End If
                
            Case "ShipDirectionU"
                 m_sShipDirectionU = Trim(oXMLElement.getAttribute("VALUE"))

            Case "ShipDirectionUOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_dShipDirectionUOffset = sAttrValue
                End If
                
            Case "ShipNumber"
                m_sShipNumber = Trim(oXMLElement.getAttribute("VALUE"))

            Case "ShipNumberOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dShipNumberOffset = Val(sAttrValue)
                End If
                
            Case "PartName"
                m_sPartName = Trim(oXMLElement.getAttribute("VALUE"))

            Case "PartNameOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dPartNameOffset = Val(sAttrValue)
                End If
                
            Case "MarkingSide"
                m_sMarkingSide = Trim(oXMLElement.getAttribute("VALUE"))

            Case "MarkingSideOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_dMarkingSideOffset = sAttrValue
                End If
                
            Case "NoOfParts"
                m_sNoOfParts = Trim(oXMLElement.getAttribute("VALUE"))

            Case "NoOfPartsOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dNoOfPartsOffset = Val(sAttrValue)
                End If
            
            Case "ErectionStage"
                m_sErectionStage = Trim(oXMLElement.getAttribute("VALUE"))
            
            Case "ErectionStageOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dErectionStageOffset = Val(sAttrValue)
                End If
                
            Case "ProfileThickness"
                m_sProfileThickness = Trim(oXMLElement.getAttribute("VALUE"))
                
            Case "ProfileThicknessOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_dProfileThicknessOffset = sAttrValue
                End If
                
            Case "Grade"
                m_sGrade = Trim(oXMLElement.getAttribute("VALUE"))
                
            Case "GradeOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dGradeOffset = Val(sAttrValue)
                End If
            
            Case "Primer"
                m_sPrimer = Trim(oXMLElement.getAttribute("VALUE"))

            Case "PrimerOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If IsNumeric(sAttrValue) Then
                    m_dPrimerOffset = Val(sAttrValue)
                End If
            
            Case "Grinder"
                m_sGrinder = Trim(oXMLElement.getAttribute("VALUE"))
            
            Case "GrinderOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_dGrinderOffset = sAttrValue
                End If
                
            Case "Curvature"
                m_sCurvature = Trim(oXMLElement.getAttribute("VALUE"))
            
            Case "CurvatureOffset"
                vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                If sAttrValue <> "" Then
                    m_dCurvatureOffset = sAttrValue
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
    
    Dim oTextNode               As IXMLDOMElement
    Dim oOutputElem             As IXMLDOMElement
    Dim oCVGTextElem            As IXMLDOMElement
    Dim oSettingElement         As IXMLDOMElement
    
    Dim oNodeList               As IXMLDOMNodeList
    Dim oAttributeNodeList      As IXMLDOMNodeList
    
    
    Dim sShipNumber As String
    Dim sPartName As String
    Dim sShipDirectionU As String
    Dim sShipDirectionV As String
    Dim sShipDirU As String
    Dim sShipDirV As String
    Dim sMarkingSide As String
    Dim sProfileThickness As String
    Dim sGrade As String
    Dim sNoOfParts As String
    Dim sGrinder As String
    Dim sCurvature As String
    Dim sPrimer As String
    Dim sErectionStage As String
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then GoTo CleanUp
    
    'oTextDom is each node in input (part) XML i.e. SMS_PROD_INFO
    
    sShipNumber = GetAttributeValueFromXML(sAttributeXML, "SHIP_NUMBER")
'    sProfileThickness = GetAttributeValueFromXML(sAttributeXML, "THICKNESS")
'    sGrade = GetAttributeValueFromXML(sAttributeXML, "MATERIAL_GRADE")
    
    vTemp = GetAttributeValueFromXML(sAttributeXML, "PART_NAME")
        sPartName = IIf(VarType(vTemp) = vbString, vTemp, "")
        sPartName = Replace(sPartName, "<", "&lt;")
        sPartName = Replace(sPartName, ">", "&gt;")
    
'    sMarkingSide = GetAttributeValueFromXML(sAttributeXML, "MARKED_SIDE")
        
    sShipDirectionU = GetAttributeValueFromXML(sAttributeXML, "DEVELOPMENT_DIR_U")
    sShipDirU = GetDirectionValue(sShipDirectionU)
    
'    sShipDirectionV = GetAttributeValueFromXML(sAttributeXML, "DEVELOPMENT_DIR_V")
'    sShipDirV = GetDirectionValue(sShipDirectionV)
    
'    sNoOfParts = GetAttributeValueFromXML(sAttributeXML, "NO_OF_PARTS")
'    sGrinder = GetAttributeValueFromXML(sAttributeXML, "GRINDER")
'
'    vTemp = GetAttributeValueFromXML(sAttributeXML, "PRIMER")
'        sPrimer = IIf(VarType(vTemp) = vbString, vTemp, "")
'        sPrimer = Replace(sPrimer, "<", "&lt;")
'        sPrimer = Replace(sPrimer, ">", "&gt;")
    
    vTemp = GetAttributeValueFromXML(sAttributeXML, "ERECTION_STAGE")
        sErectionStage = IIf(VarType(vTemp) = vbString, vTemp, "")
        sErectionStage = Replace(sErectionStage, "<", "&lt;")
        sErectionStage = Replace(sErectionStage, ">", "&gt;")
        
    sCurvature = GetAttributeValueFromXML(sAttributeXML, "CURVATURE")
            
'    If oInputDom.hasChildNodes Then
'
'        On Error Resume Next
'
'        Dim oInnerInputElement               As IXMLDOMElement
'
'        'if oTextNode.hasChildNodes then
'        Set oInnerInputElement = oInputDom.firstChild
'
'        Set oInnerInputElement = oInnerInputElement.selectSingleNode("//SMS_PART_INFO")
'
'        sShipNumber = oInnerInputElement.getAttribute("SHIP_NUMBER")
'        sProfileThickness = oInnerInputElement.getAttribute("THICKNESS")
'        sGrade = oInnerInputElement.getAttribute("MATERIAL_GRADE")
'
'        vTemp = oInnerInputElement.getAttribute("PART_NAME")
'        sPartName = IIf(VarType(vTemp) = vbString, vTemp, "")
'        sPartName = Replace(sPartName, "<", "&lt;")
'        sPartName = Replace(sPartName, ">", "&gt;")
'
'        sMarkingSide = oInnerInputElement.getAttribute("MARKED_SIDE")
'
'        sShipDirectionU = oInnerInputElement.getAttribute("DEVELOPMENT_DIR_U")
'        sShipDirU = GetDirectionValue(sShipDirectionU)
'
'        sShipDirectionV = oInnerInputElement.getAttribute("DEVELOPMENT_DIR_V")
'        sShipDirV = GetDirectionValue(sShipDirectionV)
'
'        sNoOfParts = oInnerInputElement.getAttribute("NO_OF_PARTS")
'        sGrinder = oInnerInputElement.getAttribute("GRINDER")
'
'        vTemp = oInnerInputElement.getAttribute("PRIMER")
'        sPrimer = IIf(VarType(vTemp) = vbString, vTemp, "")
'        sPrimer = Replace(sPrimer, "<", "&lt;")
'        sPrimer = Replace(sPrimer, ">", "&gt;")
'
'        vTemp = oInnerInputElement.getAttribute("ERECTION_STAGE")
'        sErectionStage = IIf(VarType(vTemp) = vbString, vTemp, "")
'        sErectionStage = Replace(sErectionStage, "<", "&lt;")
'        sErectionStage = Replace(sErectionStage, ">", "&gt;")
'
'
'    End If
    
    '*** Create oTextDom for processing ***'
    Dim sTextString As Variant
    
    sTextString = "<SMS_ANNOTATION CONTROL_POINT=""ur"" >"
    sTextString = sTextString & vbNewLine & "<" & m_sShipDirectionU & " ATTRIBUTE=" & """DEVELOPMENT_DIR_U""" & " TEXT=""" & sShipDirU & """ />"
    sTextString = sTextString & vbNewLine & "<" & m_sShipNumber & " ATTRIBUTE=" & """SHIP_NUMBER""" & " TEXT=""" & sShipNumber & """ />"
    sTextString = sTextString & vbNewLine & "<" & m_sPartName & " ATTRIBUTE=" & """PART_NAME""" & " TEXT=""" & sPartName & """ />"
    sTextString = sTextString & vbNewLine & "<" & m_sErectionStage & " ATTRIBUTE=" & """ERECTION_STAGE""" & " TEXT=""" & sErectionStage & """ />"
    If sCurvature <> "" Then
        sTextString = sTextString & vbNewLine & "<" & m_sCurvature & " ATTRIBUTE=" & """CURVATURE""" & " TEXT=""" & sCurvature & """ />"
    End If
    sTextString = sTextString & vbNewLine & "</SMS_ANNOTATION>"

    oTextDom.loadXML (sTextString)

    '**************************************'
    
    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""

    dummyForm.FontName = m_sTextFont
    dummyForm.FontSize = m_dTextSize

    dummyForm.ScaleMode = vbMillimeters
    dConversionFactor = 3.92042723072738

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
        
        'Get Attributes i.e. "DEVELOPMENT_DIR_V", "DEVELOPMENT_DIR_U", "MARKED_SIDE" etc
        vTemp = oTextNode.getAttribute("ATTRIBUTE")
        sAttribValues(iCount) = IIf(VarType(vTemp) = vbString, vTemp, "")
        
        vTemp = oTextNode.getAttribute("TEXT")
        sTxtValues(iCount) = IIf(VarType(vTemp) = vbString, vTemp, "")
        
        If sTxtValues(iCount) <> "" Then
            Select Case oTextNode.nodeName
            Case "NEWLINE"
                oCurrPos.Y = oCurrPos.Y - (1 + m_dNewLineGap) * m_dTextSize
                oTextPositions(iCount).Set 0, oCurrPos.Y, 0
                If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
                If (dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor) > (Len(sTxtValues(iCount)) * m_dTextSize) Then
                    dLineWidth = dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
                Else
                    dLineWidth = (Len(sTxtValues(iCount)) * m_dTextSize)
                End If
                oCurrPos.X = dLineWidth
            Case "GAPLESS_NEWLINE"
                oCurrPos.Y = oCurrPos.Y - m_dTextSize
                oTextPositions(iCount).Set 0, oCurrPos.Y, 0
                If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
                dLineWidth = dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
                oCurrPos.X = dLineWidth
            Case "SAMELINE"
                If sTxtValues(iCount - 1) = sShipDirU And iCount = 2 Then
                    oCurrPos.X = oCurrPos.X + (m_dShipNumberOffset * m_dTextSize)
                Else
                    oCurrPos.X = oCurrPos.X + (m_dHorizontalGap * m_dTextSize)
                End If
                oTextPositions(iCount).Set oCurrPos.X, oCurrPos.Y, oCurrPos.Z
                If (dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor) > (Len(sTxtValues(iCount)) * m_dTextSize) Then
                    dLineWidth = dLineWidth + (dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor) + (m_dHorizontalGap * m_dTextSize)
                Else
                    dLineWidth = oCurrPos.X + (Len(sTxtValues(iCount)) * m_dTextSize * PART_TEXT_ADJUST)
                End If
                oCurrPos.X = dLineWidth
            Case "GAPLESS_SAMELINE"
                oTextPositions(iCount).Set oCurrPos.X, oCurrPos.Y, oCurrPos.Z
                dLineWidth = dLineWidth + dummyForm.TextWidth(sTxtValues(iCount)) * dConversionFactor
                oCurrPos.X = dLineWidth
            End Select
        End If
    Next iCount
    
    If dLineWidth > dMaxWidth Then dMaxWidth = dLineWidth
    
    For iCount = 1 To iTextNodes
        Select Case m_sControlPoint 'adjust for justification
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
        If sTxtValues(iCount) <> "" Then
            SMS_NodeText oOutputDom, oOutputElem, "ll", oTextPositions(iCount), sTxtValues(iCount), _
                                                m_sTextFont, pOrientation, "regular", "identification", _
                                                m_dTextSize, 0, sAttribValues(iCount)
        End If
        
'        Set oCVGTextElem = Nothing
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



