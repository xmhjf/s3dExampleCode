VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlangeType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\FlangeType1.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
'
' Abstract:
'   Create custom profile location flange symbols
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom flange direction symbol on profile location marks
'
' History:
' 02/8/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.FlangeType1"

'General Properties
Private m_lTextSize                     As Double
Private m_lDistFromBubble               As Double
Private m_lDistFromLocMark              As Double
Private m_lSymbolLength                 As Double
Private m_lSymbolHeight                 As Double
Private m_lBulbCloseDist                As Double
Private m_lTick1Start                   As Double
Private m_lTick1End                     As Double
Private m_lTick2Start                   As Double
Private m_lTick2End                     As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_lTextSize = 40
    m_lDistFromBubble = 70
    m_lDistFromLocMark = 15
    m_lSymbolLength = 100
    m_lSymbolHeight = 30
    m_lBulbCloseDist = 25
    m_lTick1Start = 50
    m_lTick1End = 55
    m_lTick2Start = 57
    m_lTick2End = 62
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
                        m_lTextSize = Val(sAttrValue)
                    End If
                Case "DistFromBubble"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lDistFromBubble = CInt(sAttrValue)
                    End If
                Case "DistFromLocationMark"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lDistFromLocMark = CInt(sAttrValue)
                    End If
                Case "SymbolLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lSymbolLength = CInt(sAttrValue)
                    End If
                Case "SymbolHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lSymbolHeight = CInt(sAttrValue)
                    End If
                Case "BulbCloseDistance"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lBulbCloseDist = CInt(sAttrValue)
                    End If
                Case "Tick1Start"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lTick1Start = CInt(sAttrValue)
                    End If
                Case "Tick1End"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lTick1End = CInt(sAttrValue)
                    End If
                Case "Tick2Start"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lTick2Start = CInt(sAttrValue)
                    End If
                Case "Tick2End"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_lTick2End = CInt(sAttrValue)
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
    Dim oEndPos                 As IJDPosition
    Dim oArmEnd                 As IJDPosition
    Dim oBulbEnd                As IJDPosition
    Dim oTick1Start             As IJDPosition
    Dim oTick2Start             As IJDPosition
    Dim oTick1End               As IJDPosition
    Dim oTick2End               As IJDPosition

    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim oTempCurveElement       As IXMLDOMElement

    Dim dAVal                   As Double
    Dim dBVal                   As Double
    Dim dCVal                   As Double

    
    Dim sXSection               As String
    Dim sThicknessDir           As String
    Dim sFlangeDir              As String

    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    
    sXSection = GetAttributeValueFromXML(sAttributeXML, "XSECTION_NAME")
    sThicknessDir = GetAttributeValueFromXML(sAttributeXML, "THICKNESS_DIR")
    sFlangeDir = GetAttributeValueFromXML(sAttributeXML, "FLANGE_DIR")
        
    If sXSection = "" Or sThicknessDir = "" Or sFlangeDir = "" Then GoTo CleanUp
    
    'normalize the vector
    pOrientation.length = 1

    sXSection = Left$(sXSection, 2)
    
    'Determine whether the cross section type is one that should be drawn
'    Select Case sXSection
'        Case "UA", "EA", "BF", "B ", "BP"
'            'continue
'        Case Else
'            GoTo CleanUp
'    End Select

    Set oStartPos = New DPosition
    Set oEndPos = New DPosition
    Set oArmEnd = New DPosition

    'Create the endpoints for the long horizontal line
    oStartPos.Set m_lDistFromLocMark, m_lDistFromBubble, 0
    oEndPos.Set m_lDistFromLocMark, m_lDistFromBubble + m_lSymbolLength, 0

'    If sFlangeDir = sThicknessDir Then  'Draw the flange dir in the same direction
'                                        '   as the thickness direction
'        oArmEnd.Set m_lDistFromLocMark + m_lSymbolHeight, m_lDistFromBubble, 0
    If sFlangeDir <> sThicknessDir Then 'Draw the flange dir in the opposite direction from the thickness direction
        oArmEnd.Set m_lDistFromLocMark - m_lSymbolHeight, m_lDistFromBubble, 0
        Set oTick1Start = New DPosition
        Set oTick2Start = New DPosition
        Set oTick1End = New DPosition
        Set oTick2End = New DPosition
        oTick1Start.Set m_lDistFromLocMark, m_lDistFromBubble + m_lTick1Start, 0
        oTick2Start.Set m_lDistFromLocMark, m_lDistFromBubble + m_lTick2Start, 0
        oTick1End.Set 0, m_lDistFromBubble + m_lTick1End, 0
        oTick2End.Set 0, m_lDistFromBubble + m_lTick2End, 0
    Else
        GoTo CleanUp
    End If

    If sXSection = "BF" Or sXSection = "B " Or sXSection = "BP" Then
        'Draw bulb line
        Set oBulbEnd = New DPosition
        oBulbEnd.Set m_lDistFromLocMark, m_lDistFromBubble + m_lBulbCloseDist, 0
    End If

    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPos, pOrientation, pStartPoint
    TranslatePoint oEndPos, pOrientation, pStartPoint
    TranslatePoint oArmEnd, pOrientation, pStartPoint
    TranslatePoint oBulbEnd, pOrientation, pStartPoint
    TranslatePoint oTick1Start, pOrientation, pStartPoint
    TranslatePoint oTick2Start, pOrientation, pStartPoint
    TranslatePoint oTick1End, pOrientation, pStartPoint
    TranslatePoint oTick2End, pOrientation, pStartPoint

    'Create the XML
    
    
    
    'Create the lines
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPos, oEndPos

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPos, oArmEnd

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArmEnd, oBulbEnd

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oTick1Start, oTick1End

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oTick2Start, oTick2End

'Dim oTestPos As IJDPosition
'Set oTestPos = New DPosition
'Dim oTestPos2 As IJDPosition
'Set oTestPos2 = New DPosition
'oTestPos.Set 0, 180, 0
'TranslatePoint oTestPos, pOrientation, pStartPoint
''MsgBox "startx: " & pStartPoint.x & ", starty: " & pStartPoint.y & vbNewLine _
''    & ", endx: " & oTestPos.x & ", endy: " & oTestPos.y
'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, pStartPoint, oTestPos, "flange_annotation")
'If Not oTempCurveElement Is Nothing Then oTempEdgeElement.appendChild oTempCurveElement
'
'    oOutputElem.appendChild oTempEdgeElement

    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing

    'Attempt to format the nodes so that each node is on it's own line
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oArmEnd = Nothing
    Set oBulbEnd = Nothing
    Set oTick1Start = Nothing
    Set oTick2Start = Nothing
    Set oTick1End = Nothing
    Set oTick2End = Nothing

    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function
