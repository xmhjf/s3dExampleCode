VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckleLineMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\KnuckleLineMark.cls
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
' 2/10/2011    Ninad           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.KnuckleLineMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_sKnuckleName                    As String
Private m_sKNUCKLE_PREFIX               As String


Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_sControlPoint = "ul"
    m_sKNUCKLE_PREFIX = "KL"
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
    
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    
    Dim strPART_DIR                 As String
    Dim sPartDir                    As String
    Dim strKNUCKLE_PREFIX           As String
    Dim strKNUCKLE_ANGLE            As String
    Dim strKNUCKLE_DIR              As String
    
    Dim vTemp                       As Variant
    Dim oTextPoint1                 As IJDPosition
    Dim oTextPoint2                 As IJDPosition
    Dim oTextPoint3                 As IJDPosition
    
    
    strKNUCKLE_PREFIX = GetAttributeValueFromXML(sAttributeXML, "KNUCKLE_PREFIX")
    strKNUCKLE_DIR = GetAttributeValueFromXML(sAttributeXML, "KNUCKLE_DIR")
    strKNUCKLE_ANGLE = GetAttributeValueFromXML(sAttributeXML, "KNUCKLE_ANGLE")
    
    strKNUCKLE_ANGLE = Trim(Str(Round(Val(strKNUCKLE_ANGLE), 1)))
    
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")
    

    Set oTextPoint1 = New DPosition
    Set oTextPoint2 = New DPosition
    Set oTextPoint3 = New DPosition
    
    oTextPoint1.Set (Len(strKNUCKLE_ANGLE) + Len(strKNUCKLE_DIR)) * m_dTextSize / 2, 0, 0
    oTextPoint2.Set (Len(strKNUCKLE_ANGLE) - Len(strKNUCKLE_PREFIX)) * m_dTextSize / 2, 0, 0
    oTextPoint3.Set -(Len(strKNUCKLE_PREFIX) + Len(strKNUCKLE_DIR)) * m_dTextSize / 2, 0, 0
    
    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oTextPoint1, pOrientation, pStartPoint              '#1
    TranslatePoint oTextPoint2, pOrientation, pStartPoint              '#2
    TranslatePoint oTextPoint3, pOrientation, pStartPoint              '#3

    
       
    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint1, strKNUCKLE_PREFIX, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 3.1415926, "KNUCKLE_PREFIX"
                                        
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint2, strKNUCKLE_DIR, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 3.1415926, "KNUCKLE_DIR"
                                        
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint3, strKNUCKLE_ANGLE, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 3.1415926, "KNUCKLE_ANGLE"
    '****************************'

    IJDMfgOutputAnnotation_Evaluate = GetXMLDataAsString(oOutputElem)
   
        
CleanUp:
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTextPoint1 = Nothing
    Set oTextPoint2 = Nothing
    Set oTextPoint3 = Nothing

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








