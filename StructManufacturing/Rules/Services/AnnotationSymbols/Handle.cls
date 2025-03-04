VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Handle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\LappedMark.cls
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
' 11/02/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.LappedMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dVerticalOffset                  As Double
Private m_dHorizOffset                    As Double
Dim m_dTextOffset                       As Double
Dim m_dHorizLength                      As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 21
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dVerticalOffset = 25
    m_dHorizOffset = 25
    m_sControlPoint = "lm"
    m_dTextOffset = 0.6667 * m_dTextSize
    m_dHorizLength = 250
    
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
                Case "VerticalOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalOffset = Val(sAttrValue)
                    End If
                Case "HorizontalOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizOffset = Val(sAttrValue)
                    End If
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
                Case "TextOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "HoriziontalLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizLength = Val(sAttrValue)
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

'
'          #5            #2
'          |            |
'          |            |
'      ____|____________|_____#1
'      #4  |            |
'          |            |
'          |            |
'           #6           #3
'
             
    Dim oHorizLineStartPos     As IJDPosition          '#1
    Dim oHorizLineEndPos     As IJDPosition          '#4
    Dim oLine1TopPoint         As IJDPosition          '#2
    Dim oLine1BottomPoint      As IJDPosition          '#3
    Dim oLine2BottomPoint      As IJDPosition          '#6
    Dim oLine2TopPoint         As IJDPosition          '#5
    Dim oStartPoint            As IJDPosition
    Dim oTextPoint            As IJDPosition
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    Dim strGrade                    As String
    
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
   
    
    strGrade = GetAttributeValueFromXML(sAttributeXML, "HANDLE_TYPE")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")
    
    m_dTextOffset = Round(m_dTextOffset, 1)

    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPoint = New DPosition                  '#5
    Set oTextPoint = New DPosition                  '#5
    Set oHorizLineStartPos = New DPosition                    '#1
    Set oHorizLineEndPos = New DPosition                    '#4
    Set oLine1TopPoint = New DPosition                  '#2
    Set oLine1BottomPoint = New DPosition               '#3
    Set oLine2BottomPoint = New DPosition               '#4
    Set oLine2TopPoint = New DPosition                  '#5
    

    '*** Set All Points ***'
    
    oStartPoint.Set 0, 0, 0
    oTextPoint.Set (-Len(strGrade) * m_dTextSize) / 2, m_dTextOffset + m_dVerticalOffset, 0
    oHorizLineStartPos.Set m_dHorizLength / 2, m_dVerticalOffset, 0                                  '#1
    oHorizLineEndPos.Set -m_dHorizLength / 2, m_dVerticalOffset, 0
    oLine1TopPoint.Set (m_dHorizLength / 2) - m_dHorizOffset, 2 * m_dVerticalOffset, 0                '#2
    oLine1BottomPoint.Set (m_dHorizLength / 2) - m_dHorizOffset, 0, 0                '#3
    oLine2BottomPoint.Set -((m_dHorizLength / 2) - m_dHorizOffset), 0, 0                   '#4
    oLine2TopPoint.Set -((m_dHorizLength / 2) - m_dHorizOffset), 2 * m_dVerticalOffset, 0      '#5
    
    
    '**********************'
'    If pOrientation.x < 0 Then
'        pOrientation.x = -1 * pOrientation.x
'    End If
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oTextPoint, pOrientation, pStartPoint
    TranslatePoint oHorizLineStartPos, pOrientation, pStartPoint
    TranslatePoint oHorizLineEndPos, pOrientation, pStartPoint
    TranslatePoint oLine1TopPoint, pOrientation, pStartPoint
    TranslatePoint oLine1BottomPoint, pOrientation, pStartPoint
        
    TranslatePoint oLine2TopPoint, pOrientation, pStartPoint
    TranslatePoint oLine2BottomPoint, pOrientation, pStartPoint
   
    
    
    
    '*** Create the lines ***'
    'From 1 to 2 of Line 1
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oHorizLineEndPos, oHorizLineStartPos

    'From 1 to 3 of Line 1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1TopPoint, oLine1BottomPoint

    'From 4 to 5 of Line 2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2TopPoint, oLine2BottomPoint

    '************************'
    
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint, strGrade, _
                                              m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                               m_dTextSize, 0, "HANDLE_TYPE"
    
    
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oHorizLineStartPos = Nothing                    '#1
    Set oLine1TopPoint = Nothing                  '#2
    Set oLine1BottomPoint = Nothing               '#3
    Set oLine2TopPoint = Nothing                  '#4
    Set oLine2BottomPoint = Nothing               '#5
  
    
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




