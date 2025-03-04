VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DashigiriType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\DashigiriType.cls
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
' 11/11/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit

Private Const MODULE = "MfgcustomAnnotation.DashigiriType"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dArrowVertOffset              As Double
Private m_dArrowHorOffset               As Double
Private m_dVertOffset                   As Double
Private m_dVerticalLen                  As Double
Private m_dOffset                       As Double
Private m_dHorizOffset                  As Double
Private m_dChamferHorizOffset           As Double
Private m_dChamferVertOffset            As Double
Private m_sPrefix1                      As String
Private m_sPrefix2                      As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()

    'Setting Default Values to the Attributes (Constructor Class)

    m_sControlPoint = "ll"
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dArrowVertOffset = 0.833 * m_dTextSize
    m_dArrowHorOffset = 0.6 * m_dTextSize
    m_dHorizOffset = 0.1 * m_dTextSize
    m_dChamferHorizOffset = 0.1 * m_dTextSize
    m_dChamferVertOffset = 0.1 * m_dTextSize
    m_dVertOffset = 0.333 * m_dTextSize
    m_dVerticalLen = 4.783 * m_dTextSize
    m_sPrefix1 = "$"
    m_sPrefix2 = "20"
        
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
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
                Case "VerticalLen"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalLen = Val(sAttrValue) * m_dTextSize
                    End If
                Case "Offset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ArrowVertoffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowVertOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ArrowHorOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowHorOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "VertOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVertOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "HorizOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ChamferHorizOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dChamferHorizOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "ChamferVertOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dChamferVertOffset = Val(sAttrValue) * m_dTextSize
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
    
    Dim oStartPoint         As IJDPosition                      '#1
    Dim oArrStartPoint      As IJDPosition                   '#2
    Dim oArrEndPoint        As IJDPosition                     '#3
    Dim oVertEndPoint       As IJDPosition                    '#4
    Dim oOctPoint1          As IJDPosition                       '#5
    Dim oOctPoint2          As IJDPosition                       '#6
    Dim oOctPoint3          As IJDPosition                       '#7
    Dim oOctPoint4          As IJDPosition                       '#8
    Dim oOctPoint5          As IJDPosition                       '#9
    Dim oOctPoint6          As IJDPosition                       '#10
    Dim oOctPoint7          As IJDPosition                       '#11
    Dim oOctPoint8          As IJDPosition                       '#12
    Dim oTextPoint1         As IJDPosition                      '#13
    Dim oTextPoint2         As IJDPosition                      '#14
    
    Dim strText1            As String
    Dim strText2            As String
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp

    strText1 = GetAttributeValueFromXML(sAttributeXML, "PREFIX")
    strText2 = GetAttributeValueFromXML(sAttributeXML, "DESCRIPTION")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'Normalizing The Vector
    pOrientation.length = 1
    
    'Setting All The points
    
    Set oStartPoint = New DPosition             '#1
    Set oArrStartPoint = New DPosition          '#2
    Set oArrEndPoint = New DPosition            '#3
    Set oVertEndPoint = New DPosition           '#4
    Set oOctPoint1 = New DPosition              '#5
    Set oOctPoint2 = New DPosition
    Set oOctPoint3 = New DPosition
    Set oOctPoint4 = New DPosition
    Set oOctPoint5 = New DPosition
    Set oOctPoint6 = New DPosition
    Set oOctPoint7 = New DPosition
    Set oOctPoint8 = New DPosition
    Set oTextPoint1 = New DPosition
    Set oTextPoint2 = New DPosition
    
    oStartPoint.Set 0, 0, 0
    oArrStartPoint.Set 0, m_dVertOffset, 0
    oArrEndPoint.Set -m_dArrowHorOffset, m_dArrowVertOffset, 0
    oVertEndPoint.Set 0, m_dVerticalLen, 0
    oOctPoint1.Set -(m_dHorizOffset + (m_dChamferHorizOffset)), (m_dArrowVertOffset + m_dOffset), 0
    oOctPoint2.Set -(m_dHorizOffset + (1.1 * m_dTextSize) + m_dChamferHorizOffset), (m_dArrowVertOffset + m_dOffset), 0
    oOctPoint3.Set -(m_dHorizOffset + (1.1 * m_dTextSize) + (2 * m_dChamferHorizOffset)), (m_dArrowVertOffset + m_dOffset + (m_dChamferVertOffset)), 0
    oOctPoint4.Set -(m_dHorizOffset + (1.1 * m_dTextSize) + (2 * m_dChamferHorizOffset)), (m_dArrowVertOffset + m_dOffset + ((Len(strText1 & strText2) + 0.3) * m_dTextSize) + m_dChamferVertOffset), 0  '+ Len(strText2)
    oOctPoint5.Set -(m_dHorizOffset + (1.1 * m_dTextSize) + m_dChamferHorizOffset), (m_dArrowVertOffset + m_dOffset + ((Len(strText1 & strText2) + 0.3) * m_dTextSize) + (2 * m_dChamferVertOffset)), 0 '+ Len(strText2)
    oOctPoint6.Set -(m_dHorizOffset + (m_dChamferHorizOffset)), (m_dArrowVertOffset + m_dOffset + ((Len(strText1 & strText2) + 0.3) * m_dTextSize) + (2 * m_dChamferVertOffset)), 0 ' + Len(strText2)
    oOctPoint7.Set -m_dHorizOffset, (m_dArrowVertOffset + m_dOffset + ((Len(strText1 & strText2) + 0.3) * m_dTextSize) + m_dChamferVertOffset), 0 '+ Len(strText2)
    oOctPoint8.Set -m_dHorizOffset, (m_dArrowVertOffset + m_dOffset + (m_dChamferVertOffset)), 0
    oTextPoint1.Set (1.25 * m_dHorizOffset), (m_dArrowVertOffset + m_dOffset + (0.25 * m_dTextSize)), 0
    oTextPoint2.Set (1.25 * m_dHorizOffset), (m_dArrowVertOffset + m_dOffset + ((0.25 + Len(strText1)) * m_dTextSize)), 0
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oArrStartPoint, pOrientation, pStartPoint
    TranslatePoint oArrEndPoint, pOrientation, pStartPoint
    TranslatePoint oVertEndPoint, pOrientation, pStartPoint
    TranslatePoint oOctPoint1, pOrientation, pStartPoint
    TranslatePoint oOctPoint2, pOrientation, pStartPoint
    TranslatePoint oOctPoint3, pOrientation, pStartPoint
    TranslatePoint oOctPoint4, pOrientation, pStartPoint
    TranslatePoint oOctPoint5, pOrientation, pStartPoint
    TranslatePoint oOctPoint6, pOrientation, pStartPoint
    TranslatePoint oOctPoint7, pOrientation, pStartPoint
    TranslatePoint oOctPoint8, pOrientation, pStartPoint
    TranslatePoint oTextPoint1, pOrientation, pStartPoint
    TranslatePoint oTextPoint2, pOrientation, pStartPoint

    '*** Create the lines ***'
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    'Creating Inclined Arrow Line -- 2 to 3
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrStartPoint, oArrEndPoint
    
    'Creating Vertical Line -- 2 to 4
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrStartPoint, oVertEndPoint
    
    'Creating the Octagon
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint1, oOctPoint2

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint4, oOctPoint5

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint2, oOctPoint3

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint3, oOctPoint4

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint5, oOctPoint6

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint6, oOctPoint7

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint7, oOctPoint8

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oOctPoint8, oOctPoint1
    
    '************************'

    '*** Create the XML Text ***'

    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint1, strText1, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 1.5708, "PREFIX"
                                                
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint2, strText2, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 1.5708, "DESCRIPTION"
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
        
    
    
    
CleanUp:
    Set oStartPoint = Nothing
    Set oArrStartPoint = Nothing                    '#1
    Set oArrEndPoint = Nothing                  '#2
    Set oVertEndPoint = Nothing               '#3
    Set oOctPoint1 = Nothing                  '#4
    Set oOctPoint2 = Nothing               '#5
    Set oOctPoint3 = Nothing               '#6
    Set oOctPoint4 = Nothing               '#7
    Set oTextPoint1 = Nothing               '#5
    Set oOctPoint5 = Nothing
    Set oOctPoint6 = Nothing
    Set oOctPoint7 = Nothing
    Set oOctPoint8 = Nothing
    Set oTextPoint2 = Nothing
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTempEdgeElement = Nothing
    Set oTempCurveElement = Nothing
    
    Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function
