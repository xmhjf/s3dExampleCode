VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GrindType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\GrindType1.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
'
' Abstract:
'   Create custom Grind symbols
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom Grind symbol on contours
'
' History:
' 02/8/2010    Nathan Bruner           Created
'***************************************************************************
Option Explicit

Private Const MODULE = "MfgCustomAnnotation.GrindType1"

'General Properties
Private m_lTextSize                     As Double
Private m_lOrientation                  As Double
Private m_dDistanceFromPoint            As Double
Private m_sMarkingChar                  As String
Private m_sAntiMarkChar                 As String
Private m_sChamferAngleChar             As String
Private m_bGenICutSymbols               As Boolean
Private m_dExtraRotation                As Double
Private m_sTextFont                     As String

'grind_radius properties
Private m_dRadiusMHeight                As Double
Private m_dRadiusUMHeight               As Double

'VCut Properties
Private m_dVCutTxtOffFromLines          As Double
Private m_bVCutInvert                   As Boolean
Private m_dVCutDistFromPoint            As Double
Private m_dVCutEndOffset                As Double
Private m_dVCutNoseEndOffset            As Double
Private m_dVCutNoseHeight               As Double
      
'General Grind Symbol Properties
Private m_dThicknessTextOffsetFromBase  As Double
Private m_dAngleDistanceFromCL          As Double
Private m_dAngleDistanceFromLines       As Double
Private m_dThicknessTextStartDist       As Double
Private m_dShortLineLength              As Double
Private m_dAngledNoseEndOffset          As Double
Private m_dNoseHeight                   As Double
Private m_dBAndDEndOffset               As Double
Private m_dBAndDHeight                  As Double
Private m_dAAndEEndOffset               As Double
Private m_dAAndEHeight                  As Double
Private m_dChamferEndOffset             As Double
Private m_dChamferHeight                As Double
Private m_dMarkingCharLowerDist         As Double
Private m_dMarkingCharStartDist         As Double


Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values

    'General Properties
    m_lTextSize = 40
    m_sTextFont = "Arial"
    m_lOrientation = 1
    m_dDistanceFromPoint = 0
    m_sMarkingChar = "x"
    m_sAntiMarkChar = "z"
    m_sChamferAngleChar = "C"
    m_bGenICutSymbols = True
    m_dExtraRotation = 0
    
    'grind_radius properties
    m_dRadiusMHeight = 1.5
    m_dRadiusUMHeight = 1.25
    
    'VCut Properties
    m_dVCutTxtOffFromLines = 0.1
    m_bVCutInvert = True
    m_dVCutDistFromPoint = 0.33
    m_dVCutEndOffset = 3.75
    m_dVCutNoseEndOffset = 0.83
    m_dVCutNoseHeight = 0.6
          
    'General Grind Symbol Properties
    m_dThicknessTextOffsetFromBase = 0.16
    m_dAngleDistanceFromLines = 0.16
    m_dAngleDistanceFromCL = 0.5
    m_dThicknessTextStartDist = 2
    m_dShortLineLength = 2
    m_dAngledNoseEndOffset = 0.4
    m_dNoseHeight = 1.32
    m_dBAndDEndOffset = 0.9
    m_dBAndDHeight = 1.32
    m_dAAndEEndOffset = 0
    m_dAAndEHeight = 0
    m_dChamferEndOffset = 2
    m_dChamferHeight = 1.32
    m_dMarkingCharLowerDist = 0.16
    m_dMarkingCharStartDist = 1.5
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
                        m_lTextSize = Val(sAttrValue)
                    End If
                    
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                    
                Case "Orientation"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If "2" = sAttrValue Then
                        m_lOrientation = 2
                        m_dExtraRotation = -PI / 2
                    Else
                        m_lOrientation = 1
                    End If
                Case "DistFromPoint"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dDistanceFromPoint = Val(sAttrValue)
                    End If
                Case "MarkingCharacter"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If Len(sAttrValue) = 1 Then
                        m_sMarkingChar = sAttrValue
                    End If
                Case "AntiMarkingCharacter"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If Len(sAttrValue) = 1 Then
                        m_sAntiMarkChar = sAttrValue
                    End If
                Case "ChamferAngleCharacter"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If Len(sAttrValue) = 1 Then
                        m_sChamferAngleChar = sAttrValue
                    End If
                Case "GenerateICutSymbol"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_bGenICutSymbols = ("1" = sAttrValue)
                    
                Case "RadiusMHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dRadiusMHeight = Val(sAttrValue)
                    End If
                    
                Case "RadiusUMHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dRadiusUMHeight = Val(sAttrValue)
                    End If
                    
                Case "VCutTextOffsetFromLines"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVCutTxtOffFromLines = Val(sAttrValue)
                    End If
                Case "VCutInvert"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_bVCutInvert = ("1" = sAttrValue)
                Case "VCutDistFromPoint"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVCutDistFromPoint = Val(sAttrValue)
                    End If
                Case "VCutEndOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVCutEndOffset = Val(sAttrValue)
                    End If
                Case "VCutNoseEndOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVCutNoseEndOffset = Val(sAttrValue)
                    End If
                Case "VCutNoseHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVCutNoseHeight = Val(sAttrValue)
                    End If
                Case "ThicknessTextOffsetFromBase"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dThicknessTextOffsetFromBase = Val(sAttrValue)
                    End If
                Case "AngleDistanceFromLine"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dAngleDistanceFromLines = Val(sAttrValue)
                    End If
                Case "AngleDistanceFromCL"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dAngleDistanceFromCL = Val(sAttrValue)
                    End If
                Case "ThicknessTextStartDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dThicknessTextStartDist = Val(sAttrValue)
                    End If
                Case "ShortLineLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dShortLineLength = Val(sAttrValue)
                    End If
                Case "AngledNoseEndOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dAngledNoseEndOffset = Val(sAttrValue)
                    End If
                Case "NoseHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dNoseHeight = Val(sAttrValue)
                    End If
                Case "BAndDEndOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dBAndDEndOffset = Val(sAttrValue)
                    End If
                Case "BAndDHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dBAndDHeight = Val(sAttrValue)
                    End If
                Case "AAndEEndOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dAAndEEndOffset = Val(sAttrValue)
                    End If
                Case "AAndEHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dAAndEHeight = Val(sAttrValue)
                    End If
                Case "ChamferEndOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dChamferEndOffset = Val(sAttrValue)
                    End If
                Case "ChamferHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dChamferHeight = Val(sAttrValue)
                    End If
                Case "MarkingCharLowerDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMarkingCharLowerDist = Val(sAttrValue)
                    End If
                Case "MarkingCharStartDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMarkingCharStartDist = Val(sAttrValue)
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
    'Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", , , "RULES")
    Err.Raise Err.Number
    Resume Next
End Sub

Private Function IJDMfgOutputAnnotation_Evaluate(ByVal pStartPoint As IJDPosition, ByVal pOrientation As IJDVector, ByVal sAttributeXML As String) As String
    Const METHOD = "IJDMfgOutputAnnotation_Evaluate"
    On Error GoTo ErrorHandler
    Dim strGUID                 As String
    Dim strTYPE                 As String
    Dim dRadiusM                As Double
    Dim dRadiusUM               As Double
    Dim dChamferAngleM          As Double
    Dim dChamferDepthM          As Double
    Dim dChamferAngleUM         As Double
    Dim dChamferDepthUM         As Double
    Dim dAngleA                 As Double
    Dim dAngleB                 As Double
    Dim dAngleN                 As Double
    Dim dAngleD                 As Double
    Dim dAngleE                 As Double
    Dim dDepthA                 As Double
    Dim dDepthB                 As Double
    Dim dDepthN                 As Double
    Dim dDepthD                 As Double
    Dim dDepthE                 As Double
    
    Dim dTotalDepth             As Double
    Dim oTopRightPos            As IJDPosition
    Dim oBottomRightPos         As IJDPosition
    Dim oSymbolMidpoint         As IJDPosition
    Dim oTempPos                As IJDPosition
    Dim oTempRearPos            As IJDPosition
    Dim dEndXCoord              As Double
    Dim dTemp                   As Double
    Dim dSymbolHeight           As Double

    Dim oNoseAngTxtPos          As IJDPosition
    Dim oNoseThkTxtPos          As IJDPosition
    Dim oMChamTopS              As IJDPosition
    Dim oMChamTopE              As IJDPosition
    Dim oMChamAngTxtPos         As IJDPosition
    Dim oMChamThkTxtPos         As IJDPosition
    Dim oATopS                  As IJDPosition
    Dim oATopE                  As IJDPosition
    Dim oAAngTxtPos             As IJDPosition
    Dim oAThkTxtPos             As IJDPosition
    Dim oBTopS                  As IJDPosition
    Dim oBTopE                  As IJDPosition
    Dim oBAngTxtPos             As IJDPosition
    Dim oBThkTxtPos             As IJDPosition
    Dim oNoseBottomS            As IJDPosition
    Dim oNoseBottomE            As IJDPosition
    Dim oNoseTopS               As IJDPosition
    Dim oNoseTopE               As IJDPosition
    Dim oDBottomS               As IJDPosition
    Dim oDBottomE               As IJDPosition
    Dim oDThkTxtPos             As IJDPosition
    Dim oDAngTxtPos             As IJDPosition
    Dim oEBottomS               As IJDPosition
    Dim oEBottomE               As IJDPosition
    Dim oEAngTxtPos             As IJDPosition
    Dim oEThkTxtPos             As IJDPosition
    Dim oUMChamBottomS          As IJDPosition
    Dim oUMChamBottomE          As IJDPosition
    Dim oUMChamAngTxtPos        As IJDPosition
    Dim oUMChamThkTxtPos        As IJDPosition
    Dim oMarkingTextPos         As IJDPosition
    Dim oCenterPoint            As IJDPosition
    Dim oCenterPoint2           As IJDPosition

    Dim sVCutText               As String
    Dim sDepthChamfM            As String
    Dim sAngleChamfM            As String
    Dim sDepthA                 As String
    Dim sAngleA                 As String
    Dim sDepthB                 As String
    Dim sAngleB                 As String
    Dim sDepthNose              As String
    Dim sAngleNose              As String
    Dim sDepthD                 As String
    Dim sAngleD                 As String
    Dim sDepthE                 As String
    Dim sAngleE                 As String
    Dim sDepthChamfUM           As String
    Dim sAngleChamfUM           As String
    Dim bRoundUp                As Boolean
    Dim bSingleThickness        As Boolean
    Dim dVCutAngle              As Double
    Dim smarkedSide             As String
    Dim sMarkingSide            As String
    
    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oCVGTextElem            As IXMLDOMElement
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim oTempCurveElement       As IXMLDOMElement
    Dim oTempVertexElement      As IXMLDOMElement
    Dim dAVal                   As Double
    Dim dBVal                   As Double
    Dim dCVal                   As Double
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'FillGrindValuesFromXML sAttributeXML, strTYPE, strGUID, dRadiusM, dRadiusUM, _
                      dChamferAngleM, dChamferDepthM, _
                      dChamferAngleUM, dChamferDepthUM, dAngleA, _
                      dAngleB, dAngleN, dAngleD, dAngleE, dDepthA, _
                      dDepthB, dDepthN, dDepthD, dDepthE
                      
    
    '*** Fill Bevel Values From XML ***'
    strTYPE = GetAttributeValueFromXML(sAttributeXML, "GRIND_TYPE")
    dRadiusM = GetAttributeValueFromXML(sAttributeXML, "RADIUS_M")
    dRadiusUM = GetAttributeValueFromXML(sAttributeXML, "RADIUS_UM")
    dChamferAngleM = GetAttributeValueFromXML(sAttributeXML, "CHAMFER_ANGLE_M")
    dChamferDepthM = GetAttributeValueFromXML(sAttributeXML, "CHAMFER_DEPTH_M")
    dChamferAngleUM = GetAttributeValueFromXML(sAttributeXML, "CHAMFER_ANGLE_UM")
    dChamferDepthUM = GetAttributeValueFromXML(sAttributeXML, "CHAMFER_DEPTH_UM")
    dAngleA = GetAttributeValueFromXML(sAttributeXML, "ANGLE_A")
    dDepthA = GetAttributeValueFromXML(sAttributeXML, "DEPTH_A")
    dAngleB = GetAttributeValueFromXML(sAttributeXML, "ANGLE_B")
    dDepthB = GetAttributeValueFromXML(sAttributeXML, "DEPTH_B")
    dAngleN = GetAttributeValueFromXML(sAttributeXML, "ANGLE_N")
    dDepthN = GetAttributeValueFromXML(sAttributeXML, "DEPTH_N")
    dAngleD = GetAttributeValueFromXML(sAttributeXML, "ANGLE_D")
    dDepthD = GetAttributeValueFromXML(sAttributeXML, "DEPTH_D")
    dAngleE = GetAttributeValueFromXML(sAttributeXML, "ANGLE_E")
    dDepthE = GetAttributeValueFromXML(sAttributeXML, "DEPTH_E")
    sMarkingSide = GetAttributeValueFromXML(sAttributeXML, "MARKED_SIDE")
    '**********************************'

    oOutputDom.preserveWhiteSpace = True

    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    'Still need to add the grind type as an attribute on the SMS_ANNOTATION XML node!
    If Not oOutputElem Is Nothing Then
        oOutputElem.setAttribute "GRIND_TYPE", strTYPE
    End If
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    Select Case strTYPE
    Case "grind_radius"
        If dRadiusM < 0.000001 And dRadiusUM < 0.000001 Then
            'MsgBox "Invalid radius values. Cannot create annotation. Please provide a valid radius value."
            GoTo CleanUp
        End If
        
        'Create the text
        sDepthChamfM = ""
        sDepthChamfUM = ""
        
        Set oNoseTopS = New DPosition
        Set oNoseTopE = New DPosition
        Set oNoseBottomS = New DPosition
        Set oNoseBottomE = New DPosition
        Set oTopRightPos = New DPosition
        Set oNoseThkTxtPos = New DPosition
        Set oMarkingTextPos = New DPosition
        
        Set oSymbolMidpoint = New DPosition
        
        oTopRightPos.Set (m_dShortLineLength + m_dRadiusMHeight + m_dVCutDistFromPoint) * m_lTextSize, _
                         m_dRadiusMHeight * m_lTextSize, 0
        oMarkingTextPos.X = oTopRightPos.X - ((m_dMarkingCharStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
        oMarkingTextPos.Y = oTopRightPos.Y + ((m_dMarkingCharLowerDist - TEXT_ADJUST_BOTTOM) * m_lTextSize)
        oMarkingTextPos.Z = 0
        
        If dRadiusUM < 0.000001 Then
            sDepthChamfM = Trim(Str(CInt(Abs(dRadiusM) + 0.499999)))
            smarkedSide = m_sMarkingChar
        ElseIf dRadiusM < 0.000001 Then
            sDepthChamfUM = Trim(Str(CInt(Abs(dRadiusUM) + 0.499999)))
            smarkedSide = m_sAntiMarkChar
        Else
            smarkedSide = m_sMarkingChar
            sDepthChamfM = Trim(Str(CInt(Abs(dRadiusM) + 0.499999)))
            sDepthChamfUM = Trim(Str(Int(Abs(dRadiusUM))))
        End If
        
        oNoseTopS.Set (m_dRadiusMHeight + m_dVCutDistFromPoint) * m_lTextSize, oTopRightPos.Y, 0
        oNoseTopE.Set oTopRightPos.X, oTopRightPos.Y, 0
        oNoseBottomS.Set m_dVCutDistFromPoint * m_lTextSize, 0, 0
        oNoseBottomE.Set oTopRightPos.X, 0, 0
        oNoseThkTxtPos.Set oNoseTopS.X + ((m_dVCutTxtOffFromLines - TEXT_ADJUST_LEFT) * m_lTextSize), (m_dVCutTxtOffFromLines - TEXT_ADJUST_BOTTOM) * m_lTextSize, 0
        
        Set oCenterPoint = New DPosition
        oCenterPoint.Set oNoseTopS.X, oNoseBottomS.Y, 0
            
        If sDepthChamfM <> "" And sDepthChamfUM <> "" Then
            Set oDBottomS = New DPosition
            Set oDBottomE = New DPosition
            Set oDThkTxtPos = New DPosition
            oDBottomS.Set (m_dVCutDistFromPoint + m_dRadiusUMHeight) * m_lTextSize, -m_dRadiusUMHeight * m_lTextSize, 0
            oDBottomE.Set oTopRightPos.X, oDBottomS.Y, 0
            oDThkTxtPos.Set oNoseThkTxtPos.X - TEXT_ADJUST_LEFT * m_lTextSize, oDBottomS.Y + ((m_dVCutTxtOffFromLines - TEXT_ADJUST_BOTTOM) * m_lTextSize), 0
            oSymbolMidpoint.Set (m_dVCutDistFromPoint * m_lTextSize + oTopRightPos.X) / 2, (oTopRightPos.Y + oDBottomS.Y) / 2, 0
            dSymbolHeight = oTopRightPos.Y - oDBottomS.Y
            Set oCenterPoint2 = New DPosition
            oCenterPoint2.Set oDBottomS.X, oNoseBottomS.Y, 0
        Else
            oSymbolMidpoint.Set (m_dVCutDistFromPoint * m_lTextSize + oTopRightPos.X) / 2, (oTopRightPos.Y + oNoseBottomS.Y) / 2, 0
            dSymbolHeight = oTopRightPos.Y - oNoseBottomS.Y
            
        End If
               
        If m_lOrientation = 2 Then
            MakePerpindicular oNoseTopS, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oNoseTopE, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oNoseBottomS, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oNoseBottomE, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oTopRightPos, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oDBottomS, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oDBottomE, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oNoseThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oDThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oCenterPoint, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oCenterPoint2, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
            MakePerpindicular oMarkingTextPos, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
        End If
        
        'Translate the symbol points based on the starting point and
        '   orientation vector
        TranslatePoint oNoseTopS, pOrientation, pStartPoint
        TranslatePoint oNoseTopE, pOrientation, pStartPoint
        TranslatePoint oNoseBottomS, pOrientation, pStartPoint
        TranslatePoint oNoseBottomE, pOrientation, pStartPoint
        TranslatePoint oTopRightPos, pOrientation, pStartPoint
        TranslatePoint oDBottomS, pOrientation, pStartPoint
        TranslatePoint oDBottomE, pOrientation, pStartPoint
        TranslatePoint oNoseThkTxtPos, pOrientation, pStartPoint
        TranslatePoint oDThkTxtPos, pOrientation, pStartPoint
        TranslatePoint oCenterPoint, pOrientation, pStartPoint
        TranslatePoint oCenterPoint2, pOrientation, pStartPoint
        TranslatePoint oMarkingTextPos, pOrientation, pStartPoint
        
        'Create the lines
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseTopS, oNoseTopE

        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseBottomS, oNoseBottomE
        
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseBottomS, oNoseBottomE
        
        'Create the arcs
        SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oNoseTopS, oNoseBottomS, oCenterPoint
        If Not oCenterPoint2 Is Nothing Then
            SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oNoseBottomS, oDBottomS, oCenterPoint2
        End If

        Set oTempCurveElement = Nothing
        Set oTempEdgeElement = Nothing
        
        'Create the XML Text
        If sDepthChamfM <> "" Then
            SMS_NodeText oOutputDom, oOutputElem, "ll", oNoseThkTxtPos, sDepthChamfM, _
                               m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "CHAMFER_DEPTH_M"

            SMS_NodeText oOutputDom, oOutputElem, "ll", oDThkTxtPos, sDepthChamfUM, _
                               m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "CHAMFER_DEPTH_UM"
        Else
            SMS_NodeText oOutputDom, oOutputElem, "ll", oNoseThkTxtPos, sDepthChamfUM, _
                               m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "CHAMFER_DEPTH_UM"
        End If
        
        SMS_NodeText oOutputDom, oOutputElem, "ll", oMarkingTextPos, smarkedSide, _
                               m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "MARKED_SIDE"
        
        
        'Finished creating the symbol
        'IJDMfgOutputAnnotation_Evaluate = oOutputElem.xml
        'Set oOutputElem = Nothing
        GoTo CleanUp
        
    Case "grind_flat"
    
        dTotalDepth = dDepthA + dDepthB + dDepthN + dDepthD + dDepthE + dChamferDepthM + dChamferAngleUM
        
        bSingleThickness = False
        dVCutAngle = 0
        
        'determine whether this is a multi-section Grind or single-section Grind
        Select Case dTotalDepth
        Case dDepthA
            dVCutAngle = dAngleA
            bSingleThickness = True
        Case dDepthB
            dVCutAngle = dAngleB
            bSingleThickness = True
        Case dDepthN
            dVCutAngle = dAngleN
            bSingleThickness = True
        Case dDepthD
            dVCutAngle = dAngleD
            bSingleThickness = True
        Case dDepthE
            dVCutAngle = dAngleE
            bSingleThickness = True
        Case Else
            bSingleThickness = False
        End Select
        
        
        If bSingleThickness Then
        
            If dVCutAngle = 0 Then
    '*******************************************************************************************
    '*****************************I-Cut, No Chamfer*********************************************
    '*******************************************************************************************
                'I-Cut, no chamfer
                'IJDMfgOutputAnnotation_Evaluate = oOutputElem.xml
                GoTo CleanUp
            Else
            
    '*******************************************************************************************
    '*****************************V-Cut, No Chamfer*********************************************
    '*******************************************************************************************
    
                Set oNoseBottomS = New DPosition
                Set oNoseBottomE = New DPosition
                Set oNoseTopS = New DPosition
                Set oNoseThkTxtPos = New DPosition
                Set oNoseAngTxtPos = New DPosition
                Set oSymbolMidpoint = New DPosition
                
                If dVCutAngle > 0.000001 Then
                    sMarkingSide = m_sMarkingChar
                Else
                    sMarkingSide = m_sAntiMarkChar
                End If
                
                'Calculate the symbol points
                oNoseBottomS.X = m_dVCutDistFromPoint * m_lTextSize
                oNoseBottomS.Y = 0
                oNoseBottomS.Z = 0
                
                oNoseBottomE.X = m_dVCutEndOffset * m_lTextSize
                oNoseBottomE.Y = 0
                oNoseBottomE.Z = 0
    
                oNoseTopS.X = m_dVCutNoseEndOffset * m_lTextSize
                oNoseTopS.Y = m_dVCutNoseHeight * m_lTextSize
                oNoseTopS.Z = 0
                
                oNoseThkTxtPos.X = oNoseTopS.X + ((m_dVCutTxtOffFromLines - TEXT_ADJUST_LEFT) * m_lTextSize)
                oNoseThkTxtPos.Y = (m_dVCutTxtOffFromLines - TEXT_ADJUST_BOTTOM) * m_lTextSize
                oNoseThkTxtPos.Z = 0
                
                oNoseAngTxtPos.X = oNoseThkTxtPos.X + ((Len(sMarkingSide) + 0.3) * m_lTextSize)
                oNoseAngTxtPos.Y = oNoseThkTxtPos.Y
                oNoseAngTxtPos.Z = 0
                
                oSymbolMidpoint.X = (oNoseBottomS.X + oNoseBottomE.X) / 2
                oSymbolMidpoint.Y = oNoseTopS.Y / 2
                
                dSymbolHeight = oNoseTopS.Y
                
                If m_lOrientation = 2 Then
                    MakePerpindicular oNoseBottomS, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
                    MakePerpindicular oNoseBottomE, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
                    MakePerpindicular oNoseTopS, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
                    MakePerpindicular oNoseThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
                    MakePerpindicular oNoseAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dVCutDistFromPoint * m_lTextSize
                End If
                
                'Translate the symbol points based on the starting point and
                '   orientation vector
                TranslatePoint oNoseBottomS, pOrientation, pStartPoint
                TranslatePoint oNoseBottomE, pOrientation, pStartPoint
                TranslatePoint oNoseTopS, pOrientation, pStartPoint
                TranslatePoint oNoseThkTxtPos, pOrientation, pStartPoint
                TranslatePoint oNoseAngTxtPos, pOrientation, pStartPoint
                
                'Determine the text
                sVCutText = ""
                sVCutText = CInt(dTotalDepth)

                'Creat the XML
                
                'Create the lines
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseTopS, oNoseBottomS

                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseBottomS, oNoseBottomE

                
                'Create the XML Text
                SMS_NodeText oOutputDom, oOutputElem, "ll", oNoseThkTxtPos, sMarkingSide, _
                               m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "MARKED_SIDE"
                               
            SMS_NodeText oOutputDom, oOutputElem, "ll", oNoseAngTxtPos, sVCutText, _
                               m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "ANGLE_N"
                
                'Finished creating the symbol
                'IJDMfgOutputAnnotation_Evaluate = oOutputElem.xml
                'Set oOutputElem = Nothing
                GoTo CleanUp
            End If
        Else
        
    '*******************************************************************************************
    '*****************************General (Multi-section) Grind Symbol**************************
    '*******************************************************************************************
    
            'Multi-section Grind symbol
            'Calculate starting points for horizontal lines
            Set oTopRightPos = New DPosition
            oTopRightPos.X = 0
            oTopRightPos.Y = 0
            oTopRightPos.Z = 0
            
            If dDepthN > 0.000001 Then
                Set oNoseBottomS = New DPosition
                Set oNoseTopS = New DPosition
                Set oNoseAngTxtPos = New DPosition
                
                If dAngleN = 0 Then
                    oNoseBottomS.X = m_dDistanceFromPoint * m_lTextSize
                    oNoseBottomS.Y = 0
                    oNoseTopS.X = m_dDistanceFromPoint * m_lTextSize
                    oNoseTopS.Y = m_dNoseHeight * m_lTextSize
                    oNoseAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                    oNoseAngTxtPos.Y = (m_dNoseHeight - m_dAngleDistanceFromLines - TEXT_ADJUST_BOTTOM) * m_lTextSize
                ElseIf dAngleN > 0.000001 Then
                    oNoseBottomS.X = m_dDistanceFromPoint * m_lTextSize
                    oNoseBottomS.Y = 0
                    oNoseTopS.X = (m_dDistanceFromPoint + m_dAngledNoseEndOffset) * m_lTextSize
                    oNoseTopS.Y = m_dNoseHeight * m_lTextSize
                    oNoseAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                    oNoseAngTxtPos.Y = (m_dNoseHeight - m_dAngleDistanceFromLines - TEXT_ADJUST_BOTTOM) * m_lTextSize
                Else
                    oNoseBottomS.X = (m_dDistanceFromPoint + m_dAngledNoseEndOffset) * m_lTextSize
                    oNoseBottomS.Y = 0
                    oNoseTopS.X = m_dDistanceFromPoint * m_lTextSize
                    oNoseTopS.Y = m_dNoseHeight * m_lTextSize
                    oNoseAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                    oNoseAngTxtPos.Y = (m_dAngleDistanceFromLines + TEXT_ADJUST_TOP) * m_lTextSize
                End If
                oTopRightPos.X = oNoseTopS.X
                oTopRightPos.Y = oNoseTopS.Y
                oTopRightPos.Z = 0
            Else 'If nose does not exist, still need bottom nose line
                Set oNoseBottomS = New DPosition
                oNoseBottomS.X = m_dDistanceFromPoint * m_lTextSize
                oNoseBottomS.Y = 0
                oTopRightPos.X = oNoseBottomS.X
                oTopRightPos.Y = 0
            End If
            
            If dDepthB > 0.000001 Then
                Set oBTopS = New DPosition
                Set oBAngTxtPos = New DPosition
                oBTopS.X = (m_dDistanceFromPoint + m_dBAndDEndOffset) * m_lTextSize
                oBTopS.Y = oTopRightPos.Y + (m_dBAndDHeight * m_lTextSize)
                oBTopS.Z = 0
                oBAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                oBAngTxtPos.Y = oBTopS.Y - ((m_dAngleDistanceFromLines + TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oBAngTxtPos.Z = 0
                oTopRightPos.X = oBTopS.X
                oTopRightPos.Y = oBTopS.Y
            End If
            
            If dDepthA > 0.000001 Then
                Set oATopS = New DPosition
                Set oAAngTxtPos = New DPosition
                oATopS.X = (m_dAAndEEndOffset + m_dDistanceFromPoint) * m_lTextSize
                oATopS.Y = oTopRightPos.Y + (m_dAAndEHeight * m_lTextSize)
                oATopS.Z = 0
                oAAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                oAAngTxtPos.Y = oATopS.Y - ((m_dAngleDistanceFromLines + TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oAAngTxtPos.Z = 0
                oTopRightPos.X = oATopS.X
                oTopRightPos.Y = oATopS.Y
            End If
            
            If dChamferDepthM > 0.000001 Then
                Set oMChamTopS = New DPosition
                Set oMChamAngTxtPos = New DPosition
                oMChamTopS.X = (m_dChamferEndOffset + m_dDistanceFromPoint) * m_lTextSize
                oMChamTopS.Y = oTopRightPos.Y + (m_dChamferHeight * m_lTextSize)
                oMChamTopS.Z = 0
                oMChamAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                oMChamAngTxtPos.Y = oMChamTopS.Y - ((m_dAngleDistanceFromLines + TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oMChamAngTxtPos.Z = 0
                oTopRightPos.X = oMChamTopS.X
                oTopRightPos.Y = oMChamTopS.Y
            End If
            
            'We have found the top level by now, but we also need to find the bottom level
            '   in order to find the end of the Grind symbol.
            
            Set oBottomRightPos = New DPosition
            oBottomRightPos.X = oNoseBottomS.X
            oBottomRightPos.Y = oNoseBottomS.Y
            oBottomRightPos.Z = 0
            
            
            If dDepthD > 0.000001 Then
                Set oDBottomS = New DPosition
                Set oDAngTxtPos = New DPosition
                oDBottomS.X = (m_dDistanceFromPoint + m_dBAndDEndOffset) * m_lTextSize
                oDBottomS.Y = oBottomRightPos.Y - (m_dBAndDHeight * m_lTextSize)
                oDBottomS.Z = 0
                oDAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                oDAngTxtPos.Y = oDBottomS.Y + ((m_dAngleDistanceFromLines + TEXT_ADJUST_TOP) * m_lTextSize)
                oDAngTxtPos.Z = 0
                oBottomRightPos.X = oDBottomS.X
                oBottomRightPos.Y = oDBottomS.Y
            End If
            
            If dDepthE > 0.000001 Then
                Set oEBottomS = New DPosition
                Set oEAngTxtPos = New DPosition
                oEBottomS.X = (m_dDistanceFromPoint + m_dAAndEEndOffset) * m_lTextSize
                oEBottomS.Y = oBottomRightPos.Y - (m_dAAndEHeight * m_lTextSize)
                oEBottomS.Z = 0
                oEAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                oEAngTxtPos.Y = oEBottomS.Y + ((m_dAngleDistanceFromLines + TEXT_ADJUST_TOP) * m_lTextSize)
                oEAngTxtPos.Z = 0
                oBottomRightPos.X = oEBottomS.X
                oBottomRightPos.Y = oEBottomS.Y
            End If
            
            If dChamferDepthUM > 0.000001 Then
                Set oUMChamBottomS = New DPosition
                Set oUMChamAngTxtPos = New DPosition
                oUMChamBottomS.X = (m_dDistanceFromPoint + m_dChamferEndOffset) * m_lTextSize
                oUMChamBottomS.Y = oBottomRightPos.Y - (m_dChamferHeight * m_lTextSize)
                oUMChamBottomS.Z = 0
                oUMChamAngTxtPos.X = (m_dAngleDistanceFromCL - TEXT_ADJUST_LEFT) * m_lTextSize
                oUMChamAngTxtPos.Y = oUMChamBottomS.Y + ((m_dAngleDistanceFromLines + TEXT_ADJUST_TOP) * m_lTextSize)
                oUMChamAngTxtPos.Z = 0
                oBottomRightPos.X = oUMChamBottomS.X
                oBottomRightPos.Y = oUMChamBottomS.Y
            End If
            
            'By now, we know what the rightmost point is on the front of the Grind symbol,
            '   so we can calculate the top-right position of the Grind symbol and set
            '   the end points for the horizontal lines and the points for the thickness
            '   text and the marking symbol text
            If oTopRightPos.X >= oBottomRightPos.X Then
                oTopRightPos.X = oTopRightPos.X + (m_dShortLineLength * m_lTextSize)
                oBottomRightPos.X = oTopRightPos.X
            Else
                oBottomRightPos.X = oBottomRightPos.X + (m_dShortLineLength * m_lTextSize)
                oTopRightPos.X = oBottomRightPos.X
            End If
            
            Set oMarkingTextPos = New DPosition
            oMarkingTextPos.X = oTopRightPos.X - ((m_dMarkingCharStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
            oMarkingTextPos.Y = oTopRightPos.Y + ((m_dMarkingCharLowerDist - TEXT_ADJUST_BOTTOM) * m_lTextSize)
            oMarkingTextPos.Z = 0
            
            
            'Find all the endpoints
            Set oNoseBottomE = New DPosition
            oNoseBottomE.X = oTopRightPos.X
            oNoseBottomE.Y = oNoseBottomS.Y
            oNoseBottomE.Z = 0
            
            If Not oNoseTopS Is Nothing Then
                Set oNoseTopE = New DPosition
                Set oNoseThkTxtPos = New DPosition
                oNoseTopE.X = oTopRightPos.X
                oNoseTopE.Y = oNoseTopS.Y
                oNoseTopE.Z = 0
                oNoseThkTxtPos.X = oNoseTopE.X - ((m_dThicknessTextStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
                oNoseThkTxtPos.Y = oNoseTopE.Y - ((m_dThicknessTextOffsetFromBase - TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oNoseThkTxtPos.Z = 0
            End If
            
            If Not oMChamTopS Is Nothing Then
                Set oMChamTopE = New DPosition
                Set oMChamThkTxtPos = New DPosition
                oMChamTopE.X = oTopRightPos.X
                oMChamTopE.Y = oMChamTopS.Y
                oMChamTopE.Z = 0
                oMChamThkTxtPos.X = oTopRightPos.X - ((m_dThicknessTextStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
                oMChamThkTxtPos.Y = oMChamTopS.Y - ((m_dThicknessTextOffsetFromBase - TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oMChamThkTxtPos.Z = 0
            End If
            
            If Not oATopS Is Nothing Then
                Set oATopE = New DPosition
                Set oAThkTxtPos = New DPosition
                oATopE.X = oTopRightPos.X
                oATopE.Y = oATopS.Y
                oATopE.Z = 0
                oAThkTxtPos.X = oTopRightPos.X - ((m_dThicknessTextStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
                oAThkTxtPos.Y = oATopS.Y - ((m_dThicknessTextOffsetFromBase - TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oAThkTxtPos.Z = 0
            End If
            
            If Not oBTopS Is Nothing Then
                Set oBTopE = New DPosition
                Set oBThkTxtPos = New DPosition
                oBTopE.X = oTopRightPos.X
                oBTopE.Y = oBTopS.Y
                oBTopE.Z = 0
                oBThkTxtPos.X = oTopRightPos.X - ((m_dThicknessTextStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
                oBThkTxtPos.Y = oBTopS.Y - ((m_dThicknessTextOffsetFromBase - TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oBThkTxtPos.Z = 0
            End If
            
            If Not oDBottomS Is Nothing Then
                Set oDBottomE = New DPosition
                Set oDThkTxtPos = New DPosition
                oDBottomE.X = oTopRightPos.X
                oDBottomE.Y = oDBottomS.Y
                oDBottomE.Z = 0
                oDThkTxtPos.X = oTopRightPos.X - ((m_dThicknessTextStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
                oDThkTxtPos.Y = oDBottomS.Y + ((m_dThicknessTextOffsetFromBase - TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oDThkTxtPos.Z = 0
            End If
            
            If Not oEBottomS Is Nothing Then
                Set oEBottomE = New DPosition
                Set oEThkTxtPos = New DPosition
                oEBottomE.X = oTopRightPos.X
                oEBottomE.Y = oEBottomS.Y
                oEBottomE.Z = 0
                oEThkTxtPos.X = oTopRightPos.X - ((m_dThicknessTextStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
                oEThkTxtPos.Y = oEBottomS.Y + ((m_dThicknessTextOffsetFromBase - TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oDThkTxtPos.Z = 0
            End If
            
            If Not oUMChamBottomS Is Nothing Then
                Set oUMChamBottomE = New DPosition
                Set oUMChamThkTxtPos = New DPosition
                oUMChamBottomE.X = oTopRightPos.X
                oUMChamBottomE.Y = oUMChamBottomS.Y
                oUMChamBottomE.Z = 0
                oUMChamThkTxtPos.X = oTopRightPos.X - ((m_dThicknessTextStartDist + TEXT_ADJUST_LEFT) * m_lTextSize)
                oUMChamThkTxtPos.Y = oUMChamBottomS.Y + ((m_dThicknessTextOffsetFromBase - TEXT_ADJUST_BOTTOM) * m_lTextSize)
                oUMChamThkTxtPos.Z = 0
            End If
            
    
            
            'Create all the text
            bRoundUp = True
            If Not oMChamTopS Is Nothing Then
                sAngleChamfM = CInt(Abs(dChamferAngleM) + 0.499999)
                If bRoundUp Then
                    sDepthChamfM = CInt(Abs(dChamferDepthM) + 0.499999)
                Else
                    sDepthChamfM = Int(Abs(dChamferDepthM))
                End If
                bRoundUp = Not bRoundUp
            End If
            
            If Not oATopS Is Nothing Then
                sAngleA = CInt(Abs(dAngleA) + 0.499999)
                If bRoundUp Then
                    sDepthA = CInt(Abs(dDepthA) + 0.499999)
                Else
                    sDepthA = Int(Abs(dDepthA))
                End If
                bRoundUp = Not bRoundUp
            End If
            
            If Not oBTopS Is Nothing Then
                sAngleB = CInt(Abs(dAngleB) + 0.499999)
                If bRoundUp Then
                    sDepthB = CInt(Abs(dDepthB) + 0.499999)
                Else
                    sDepthB = Int(Abs(dDepthB))
                End If
                bRoundUp = Not bRoundUp
            End If
            
            If Not oNoseTopS Is Nothing Then
                sAngleNose = CInt(Abs(dAngleN) + 0.499999)
                If bRoundUp Then
                    sDepthNose = CInt(Abs(dDepthN) + 0.499999)
                Else
                    sDepthNose = Int(Abs(dDepthN))
                End If
                bRoundUp = Not bRoundUp
            End If
            
            If Not oDBottomS Is Nothing Then
                sAngleD = CInt(Abs(dAngleD) + 0.499999)
                If bRoundUp Then
                    sDepthD = CInt(Abs(dDepthD) + 0.499999)
                Else
                    sDepthD = Int(Abs(dDepthD))
                End If
                bRoundUp = Not bRoundUp
            End If
            
            If Not oEBottomS Is Nothing Then
                sAngleE = CInt(Abs(dAngleE) + 0.499999)
                If bRoundUp Then
                    sDepthE = CInt(Abs(dDepthE) + 0.499999)
                Else
                    sDepthE = Int(Abs(dDepthE))
                End If
                bRoundUp = Not bRoundUp
            End If
            
            If Not oUMChamBottomS Is Nothing Then
                sAngleChamfUM = CInt(Abs(dChamferAngleUM) + 0.499999)
                If bRoundUp Then
                    sDepthChamfUM = CInt(Abs(dChamferDepthUM) + 0.499999)
                Else
                    sDepthChamfUM = Int(Abs(dChamferDepthUM))
                End If
            End If
            
            Set oSymbolMidpoint = New DPosition
            oSymbolMidpoint.X = m_dDistanceFromPoint * m_lTextSize
            oSymbolMidpoint.Y = (oTopRightPos.Y + oBottomRightPos.Y) / 2
                
            dSymbolHeight = oTopRightPos.Y - oBottomRightPos.Y
            
            If m_lOrientation = 2 Then
                'Rotate the symbol 90 degrees
                MakePerpindicular oNoseAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oNoseThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oMChamTopS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oMChamTopE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oMChamAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oMChamThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oATopS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oATopE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oAAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oAThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oBTopS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oBTopE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oBAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oBThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oNoseBottomS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oNoseBottomE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oNoseTopS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oNoseTopE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oDBottomS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oDBottomE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oDThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oDAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oEBottomS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oEBottomE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oEAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oEThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oUMChamBottomS, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oUMChamBottomE, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oUMChamAngTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oUMChamThkTxtPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
                MakePerpindicular oMarkingTextPos, oSymbolMidpoint, dSymbolHeight, m_dDistanceFromPoint * m_lTextSize
            End If
            
            
            'Translate all the points
            TranslatePoint oNoseAngTxtPos, pOrientation, pStartPoint
            TranslatePoint oNoseThkTxtPos, pOrientation, pStartPoint
            TranslatePoint oMChamTopS, pOrientation, pStartPoint
            TranslatePoint oMChamTopE, pOrientation, pStartPoint
            TranslatePoint oMChamAngTxtPos, pOrientation, pStartPoint
            TranslatePoint oMChamThkTxtPos, pOrientation, pStartPoint
            TranslatePoint oATopS, pOrientation, pStartPoint
            TranslatePoint oATopE, pOrientation, pStartPoint
            TranslatePoint oAAngTxtPos, pOrientation, pStartPoint
            TranslatePoint oAThkTxtPos, pOrientation, pStartPoint
            TranslatePoint oBTopS, pOrientation, pStartPoint
            TranslatePoint oBTopE, pOrientation, pStartPoint
            TranslatePoint oBAngTxtPos, pOrientation, pStartPoint
            TranslatePoint oBThkTxtPos, pOrientation, pStartPoint
            TranslatePoint oNoseBottomS, pOrientation, pStartPoint
            TranslatePoint oNoseBottomE, pOrientation, pStartPoint
            TranslatePoint oNoseTopS, pOrientation, pStartPoint
            TranslatePoint oNoseTopE, pOrientation, pStartPoint
            TranslatePoint oDBottomS, pOrientation, pStartPoint
            TranslatePoint oDBottomE, pOrientation, pStartPoint
            TranslatePoint oDThkTxtPos, pOrientation, pStartPoint
            TranslatePoint oDAngTxtPos, pOrientation, pStartPoint
            TranslatePoint oEBottomS, pOrientation, pStartPoint
            TranslatePoint oEBottomE, pOrientation, pStartPoint
            TranslatePoint oEAngTxtPos, pOrientation, pStartPoint
            TranslatePoint oEThkTxtPos, pOrientation, pStartPoint
            TranslatePoint oUMChamBottomS, pOrientation, pStartPoint
            TranslatePoint oUMChamBottomE, pOrientation, pStartPoint
            TranslatePoint oUMChamAngTxtPos, pOrientation, pStartPoint
            TranslatePoint oUMChamThkTxtPos, pOrientation, pStartPoint
            TranslatePoint oMarkingTextPos, pOrientation, pStartPoint
            
            'Create the XML
    '        Set oTempEdgeElement = oOutputDom.createElement("SMS_EDGE")
    '        oTempEdgeElement.setAttribute "TYPE", "grind_annotation"
            
            Set oTempPos = oNoseBottomS
            Set oTempRearPos = oNoseBottomE
            
            'Nose Lines
            'Bottom Nose Line
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseBottomS, oNoseBottomE
            
            If Not oNoseTopS Is Nothing Then
                'Top Nose Line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseTopS, oNoseTopE
                
                'Front nose line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oNoseTopS, oNoseBottomS
                
                Set oTempPos = oNoseTopS
                Set oTempRearPos = oNoseTopE
            End If
            
            If Not oBTopS Is Nothing Then
                'Top B line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oBTopS, oBTopE
                
                'Front B line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oBTopS, oTempPos
                
                Set oTempPos = oBTopS
                Set oTempRearPos = oBTopE
            End If
            
            If Not oATopS Is Nothing Then
                'Top A line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oATopS, oATopE
                
                'Front A line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oATopS, oTempPos
                
                Set oTempPos = oATopS
                Set oTempRearPos = oATopE
            End If
            
            If Not oMChamTopS Is Nothing Then
                'Top MCham line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oMChamTopS, oMChamTopE
                
                'Front MCham line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oMChamTopS, oTempPos
            End If
            
            Set oTempPos = oNoseBottomS
            Set oTempRearPos = oNoseBottomE
            
            If Not oDBottomS Is Nothing Then
                'Bottom D line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oDBottomS, oDBottomE
                
                'Front D line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oDBottomS, oTempPos
                
                Set oTempPos = oDBottomS
                Set oTempRearPos = oDBottomE
            End If
            
            If Not oEBottomS Is Nothing Then
                'Bottom E line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oEBottomS, oEBottomE
                
                'Front E line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oEBottomS, oTempPos
                
                Set oTempPos = oEBottomS
                Set oTempRearPos = oEBottomE
            End If
            
            If Not oUMChamBottomS Is Nothing Then
                'Bottom UM Chamfer line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oUMChamBottomS, oUMChamBottomE
                
                'Front UM Chamfer line
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oUMChamBottomS, oTempPos
            End If
            
    ''For testing purposes, should delete for actual implementation
    'oTopRightPos.x = 0
    'oTopRightPos.y = oTopRightPos.y + (1.5 * m_lTextSize)
    'oBottomRightPos.x = 0
    'oBottomRightPos.y = oBottomRightPos.y - (1.5 * m_lTextSize)
    'TranslatePoint oTopRightPos, pOrientation, pStartPoint
    'TranslatePoint oBottomRightPos, pOrientation, pStartPoint
    'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, oTopRightPos, oBottomRightPos, "grind_annotation")
    'If Not oTempCurveElement Is Nothing Then oOutputElem.appendChild oTempCurveElement
    'Set oTempCurveElement = Nothing
    ''End for testing purposes
            
            
            'Create Text XML
            'Marking character
            SMS_NodeText oOutputDom, oOutputElem, "ll", oMarkingTextPos, m_sMarkingChar, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "MARKED_SIDE"
            
            'M Chamfer text
            If Not oMChamTopS Is Nothing Then
                SMS_NodeText oOutputDom, oOutputElem, "ll", oMChamThkTxtPos, sDepthChamfM, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "CHAMFER_DEPTH_M"
                SMS_NodeText oOutputDom, oOutputElem, "ll", oMChamAngTxtPos, sAngleChamfM, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "CHAMFER_ANGLE_M"
            End If
            
            'A text
            If Not oATopS Is Nothing Then
                SMS_NodeText oOutputDom, oOutputElem, "ll", oAThkTxtPos, sDepthA, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "DEPTH_A"
                SMS_NodeText oOutputDom, oOutputElem, "ll", oAAngTxtPos, sAngleA, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "ANGLE_A"
            End If
            
            'B text
            If Not oBTopS Is Nothing Then
                SMS_NodeText oOutputDom, oOutputElem, "ll", oBThkTxtPos, sDepthB, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "DEPTH_B"
                SMS_NodeText oOutputDom, oOutputElem, "ll", oBAngTxtPos, sAngleB, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "ANGLE_B"
            End If
            
            'Nose text
            If Not oNoseTopS Is Nothing Then
                SMS_NodeText oOutputDom, oOutputElem, "ll", oNoseThkTxtPos, sDepthNose, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "DEPTH_N"
                If dAngleN > 0.000001 Then
                    SMS_NodeText oOutputDom, oOutputElem, "ll", oNoseAngTxtPos, sAngleNose, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "ANGLE_N"
                ElseIf dAngleN < 0 Then
                    SMS_NodeText oOutputDom, oOutputElem, "ll", oNoseAngTxtPos, sAngleNose, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "ANGLE_N"
                End If
            End If
            
            'D text
            If Not oDBottomS Is Nothing Then
                SMS_NodeText oOutputDom, oOutputElem, "ll", oDThkTxtPos, sDepthD, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "DEPTH_D"
                SMS_NodeText oOutputDom, oOutputElem, "ll", oDAngTxtPos, sAngleD, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "ANGLE_D"
            End If
            
            'E text
            If Not oEBottomS Is Nothing Then
                SMS_NodeText oOutputDom, oOutputElem, "ll", oEThkTxtPos, sDepthE, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "DEPTH_E"
                SMS_NodeText oOutputDom, oOutputElem, "ll", oEAngTxtPos, sAngleE, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "ANGLE_E"
            End If
            
            'UM Chamfer text
            If Not oUMChamBottomS Is Nothing Then
                SMS_NodeText oOutputDom, oOutputElem, "ll", oUMChamThkTxtPos, sDepthChamfUM, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "CHAMFER_DEPTH_UM"
                SMS_NodeText oOutputDom, oOutputElem, "ll", oUMChamAngTxtPos, sAngleChamfUM, _
                                   m_sTextFont, pOrientation, "regular", "partmon", m_lTextSize, m_dExtraRotation, "CHAMFER_ANGLE_UM"
            End If
            
            'Done Creating Symbol
            'IJDMfgOutputAnnotation_Evaluate = oOutputElem.xml
            GoTo CleanUp
        End If
    Case Else
        GoTo CleanUp
    End Select
CleanUp:
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

    Set oAAngTxtPos = Nothing
    Set oAThkTxtPos = Nothing
    Set oATopE = Nothing
    Set oATopS = Nothing
    Set oBAngTxtPos = Nothing
    Set oBottomRightPos = Nothing
    Set oBThkTxtPos = Nothing
    Set oBTopE = Nothing
    Set oBTopS = Nothing
    Set oCVGTextElem = Nothing
    Set oDAngTxtPos = Nothing
    Set oDBottomE = Nothing
    Set oDBottomS = Nothing
    Set oDThkTxtPos = Nothing
    Set oEAngTxtPos = Nothing
    Set oEBottomE = Nothing
    Set oEBottomS = Nothing
    Set oEThkTxtPos = Nothing
    Set oMarkingTextPos = Nothing
    Set oMChamAngTxtPos = Nothing
    Set oMChamThkTxtPos = Nothing
    Set oMChamTopE = Nothing
    Set oMChamTopS = Nothing
    Set oNoseAngTxtPos = Nothing
    Set oNoseBottomE = Nothing
    Set oNoseBottomS = Nothing
    Set oNoseThkTxtPos = Nothing
    Set oNoseTopE = Nothing
    Set oNoseTopS = Nothing
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oSymbolMidpoint = Nothing
    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempPos = Nothing
    Set oTempRearPos = Nothing
    Set oTempVertexElement = Nothing
    Set oTopRightPos = Nothing
    Set oUMChamAngTxtPos = Nothing
    Set oUMChamBottomE = Nothing
    Set oUMChamBottomS = Nothing
    Set oUMChamThkTxtPos = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function
