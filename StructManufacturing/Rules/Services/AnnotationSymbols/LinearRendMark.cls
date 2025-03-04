VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinearRendMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\EndFittingMark2.cls
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
' 08/26/2010    Ninad           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.LinearRendMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dVerticalDist                     As Double
Private m_dHorizDist                     As Double

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dVerticalDist = 0.5 * m_dTextSize
    m_dHorizDist = 0.6 * m_dTextSize
    m_sControlPoint = "ur"
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
                Case "VerticalDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "HorizDist"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizDist = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextSize"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextSize = Val(sAttrValue)
                    End If
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
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
    
    Dim oStartPoint               As IJDPosition          '#1
    Dim oLine1BottomPoint       As IJDPosition          '#2
    Dim oLine1RightPoint       As IJDPosition          '#3
    Dim oLine1TopPoint            As IJDPosition          '#4
    Dim oLine2TopPoint            As IJDPosition          '#5
    Dim oLine2LeftPoint       As IJDPosition          '#6
    Dim oLine2BottomPoint       As IJDPosition          '#7
        
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim strTYPE                 As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
    
    strTYPE = GetAttributeValueFromXML(sAttributeXML, "REND_TYPE")
        
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")
    
    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    'Still need to add the rend type as an attribute on the SMS_ANNOTATION XML node!
    If Not oOutputElem Is Nothing Then
        oOutputElem.setAttribute "REND_TYPE", strTYPE
    End If
    
    Set oStartPoint = New DPosition             '#1
    Set oLine1BottomPoint = New DPosition               '#2
    Set oLine1RightPoint = New DPosition               '#3
    Set oLine1TopPoint = New DPosition                    '#4
    
    Set oLine2TopPoint = New DPosition                      '#5
    Set oLine2LeftPoint = New DPosition                      '#6
    Set oLine2BottomPoint = New DPosition                      '#7
    
    
    '*** Set All Points ***'
    oStartPoint.Set 0, 0, 0                     '#1
    oLine1BottomPoint.Set m_dHorizDist / 2, -m_dVerticalDist, 0     '#2
    oLine1RightPoint.Set m_dHorizDist, 0, 0     '#3
    oLine1TopPoint.Set m_dHorizDist / 2, m_dVerticalDist, 0           '#4
    oLine2TopPoint.Set -m_dHorizDist / 2, m_dVerticalDist, 0         '#5
    oLine2LeftPoint.Set -m_dHorizDist, 0, 0        '#6
    oLine2BottomPoint.Set -m_dHorizDist / 2, -m_dVerticalDist, 0      '#7
    
    '**********************'
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint       '1
    TranslatePoint oLine1BottomPoint, pOrientation, pStartPoint         '2
    TranslatePoint oLine1RightPoint, pOrientation, pStartPoint         '3
    TranslatePoint oLine1TopPoint, pOrientation, pStartPoint          '4
    
    TranslatePoint oLine2TopPoint, pOrientation, pStartPoint      '5
    TranslatePoint oLine2LeftPoint, pOrientation, pStartPoint         '6
    TranslatePoint oLine2BottomPoint, pOrientation, pStartPoint       '7
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    Select Case strTYPE
        Case "Type1"
        'Create the Lines
            'Line1 -- 1 to 2
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine1BottomPoint
            
            'Line2 -- 2 to 3
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1BottomPoint, oLine1RightPoint
            
            'Line7 -- 6 to 7
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2LeftPoint, oLine2BottomPoint
            
            'Line8 -- 7 to 1
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2BottomPoint, oStartPoint
            
            'Line3 -- 3 to 4
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1RightPoint, oLine1TopPoint
            
            'Line4 -- 4 to 1
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1TopPoint, oStartPoint
            
            'Line5 -- 1 to 5
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine2TopPoint
            
            'Line6 -- 5 to 6
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2TopPoint, oLine2LeftPoint
            
        Case "Type2"
        'Create the Lines
             'Line1 -- 3 to 4
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1RightPoint, oLine1TopPoint
            
            'Line2 -- 4 to 1
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1TopPoint, oStartPoint
            
            'Line3 -- 1 to 5
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine2TopPoint
            
            'Line4 -- 5 to 6
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2TopPoint, oLine2LeftPoint
            
    End Select
    
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPoint = Nothing
    Set oLine1RightPoint = Nothing
    Set oLine1BottomPoint = Nothing
    Set oLine1TopPoint = Nothing
    
    Set oLine2BottomPoint = Nothing
    
    Set oLine2LeftPoint = Nothing
    Set oLine2TopPoint = Nothing
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function
    
    
    

