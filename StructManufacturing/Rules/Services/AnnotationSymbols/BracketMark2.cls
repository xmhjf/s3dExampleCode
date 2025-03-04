VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\BracketMark2.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
' Abstract:
'   Create custom profile location flange symbols
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a custom flange direction symbol on profile location marks
'
' History:
' 09/08/2010    Ninad           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.BracketMark2"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dVerticalDist                  As Double
Private m_dHorizDist                  As Double

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dVerticalDist = 4 * m_dTextSize
    m_dHorizDist = 0.6 * m_dTextSize
    m_sControlPoint = "ll"
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
                Case "VerticalDist"
                    vTemp = oXMLElement.getAttribute("VALUE")
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalDist = CDbl(sAttrValue) * m_dTextSize
                    End If
                Case "HorizDist"
                    vTemp = oXMLElement.getAttribute("VALUE")
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizDist = CDbl(sAttrValue) * m_dTextSize
                    End If
                Case "TextFont"
                    vTemp = oXMLElement.getAttribute("VALUE")
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                Case "TextStyle"
                    vTemp = oXMLElement.getAttribute("VALUE")
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
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
    
    
    
'            #2    #4
'             \    /
'              \  /
'        _______\/_______
'             #1/\
'              /  \
'             /    \
'            #5    #3
             
    Dim oCenterPoint                As IJDPosition          '#1
    Dim oSlantLine1TopPoint         As IJDPosition          '#2
    Dim oSlantLine1BottomPoint      As IJDPosition          '#3
    Dim oSlantLine2TopPoint         As IJDPosition          '#4
    Dim oSlantLine2BottomPoint      As IJDPosition          '#5

    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    Dim oPlateMarkDom               As New DOMDocument
    Dim oPlateMarkElem              As IXMLDOMAttribute
    Dim oSeamElement                As IXMLDOMElement
    
    Dim strGUID                     As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
    
    oPlateMarkDom.loadXML sAttributeXML
    Set oSeamElement = oPlateMarkDom.selectSingleNode("SMS_OUTPUT_ANNOTATION")
    
    
    strGUID = GetAttributeValueFromXML(sAttributeXML, "GUID")
    m_dTextSize = GetAttributeValueFromXML(sAttributeXML, "TextSize")
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    Set oOutputElem = oOutputDom.createElement("SMS_ANNOTATION")
    oOutputElem.setAttribute "TYPE", oSeamElement.getAttribute("TYPE")
    oOutputElem.setAttribute "MARKED_SIDE", "Marking"
    oOutputElem.setAttribute "GUID", strGUID
    oOutputElem.setAttribute "TEXT_SIZE", m_dTextSize
    oOutputElem.setAttribute "SX", Round(pStartPoint.x, 5)
    oOutputElem.setAttribute "SY", Round(pStartPoint.y, 5)
    
    
    'MsgBox "Round x:" & Round(pStartPoint.x, 5) & "  And y:" & Round(pStartPoint.y, 5)

    'normalize the vector
    pOrientation.length = 1
    
    Set oCenterPoint = New DPosition                         '#1
    Set oSlantLine1TopPoint = New DPosition                  '#2
    Set oSlantLine1BottomPoint = New DPosition               '#3
    Set oSlantLine2TopPoint = New DPosition                  '#4
    Set oSlantLine2BottomPoint = New DPosition               '#5
    
'    Select Case sPartDir
'        Case "R", "r"
'            oStartPos.Set m_dRightStartX * m_dTextSize, -m_dRightStartY * m_dTextSize, 0
'        Case "L", "l"
'            oStartPos.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
'        Case Else
'            oStartPos.Set -m_dCenteredStartX * m_dTextSize, -m_dCenteredStartY * m_dTextSize, 0
'    End Select
'
    
    '*** Set All Points ***'
    oCenterPoint.Set 0, 0, 0                                                '#1
    oSlantLine1TopPoint.Set -m_dHorizDist / 2, m_dVerticalDist / 2, 0       '#2
    oSlantLine1BottomPoint.Set m_dHorizDist / 2, -m_dVerticalDist / 2, 0    '#3
    oSlantLine2TopPoint.Set m_dHorizDist / 2, m_dVerticalDist / 2, 0        '#4
    oSlantLine2BottomPoint.Set -m_dHorizDist / 2, -m_dVerticalDist / 2, 0    '#5
    '**********************'
    
    TranslatePoint oCenterPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine1TopPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine1BottomPoint, pOrientation, pStartPoint
        
    TranslatePoint oSlantLine2TopPoint, pOrientation, pStartPoint
    TranslatePoint oSlantLine2BottomPoint, pOrientation, pStartPoint
   
    
        
    
    '*** Create the lines ***'
    'Slant Line 1
    Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, oSlantLine1TopPoint, oSlantLine1BottomPoint, "BracketStiffener")
    If Not oTempCurveElement Is Nothing Then oOutputElem.appendChild oTempCurveElement

    'Slant Line 2
    Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, oSlantLine2TopPoint, oSlantLine2BottomPoint, "BracketStiffener")
    If Not oTempCurveElement Is Nothing Then oOutputElem.appendChild oTempCurveElement
    '************************'

    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing
    
    '*** Create the XML Text ***'
    'Dim oCVGTextElem            As IXMLDOMElement
    'Set oCVGTextElem = CreateTextNode(oOutputDom, m_sControlPoint, oTextPoint, strGrinder, _
                                      m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                      m_dTextSize, 0, m_sGrinder)
                                      
    'oOutputElem.appendChild oCVGTextElem
    
    'Set oCVGTextElem = Nothing
    '****************************'

    Dim sOutputXML As String
    If oOutputElem.childNodes.length > 0 Then
        sOutputXML = Replace(oOutputElem.xml, "><", ">" & vbNewLine & "<")
    Else
        sOutputXML = ""
    End If
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:
    Set oCenterPoint = Nothing                         '#1
    Set oSlantLine1TopPoint = Nothing                  '#2
    Set oSlantLine1BottomPoint = Nothing               '#3
    Set oSlantLine2TopPoint = Nothing                  '#4
    Set oSlantLine2BottomPoint = Nothing               '#5
  
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function

'Below code is for adding a line to the output for testing purposes:
'Dim oTestPos As IJDPosition
'Set oTestPos = New DPosition
'Dim oTestPos2 As IJDPosition
'Set oTestPos2 = New DPosition
'oTestPos.Set 0, 180, 0
'TranslatePoint oTestPos, pOrientation, pStartPoint
''MsgBox "startx: " & pStartPoint.x & ", starty: " & pStartPoint.y & vbNewLine _
''    & ", endx: " & oTestPos.x & ", endy: " & oTestPos.y
'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, pStartPoint, oTestPos, "plate_thickness_annotation")
'If Not oTempCurveElement Is Nothing Then oTempEdgeElement.appendChild oTempCurveElement
'
'    oOutputElem.appendChild oTempEdgeElement












