VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarPlateInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\CollarPlateInfo.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\MfgCustomAnnotation.vbp
'
'
' Abstract:
'   Create Label for Collar Plate Laocation Mark
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating Label which contains the infomation
'   of Slot Type and Web Length on Collar Plate Location Mark
'
' History:
' 03/31/2011    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.CollarPlateInfo"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_sSlotTypeText                 As String
Private m_sWebLengthText                As String
Private m_dHorizOffset                  As Double
Private m_dVertOffset                   As Double
Private m_dTextGap                      As Double

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
' Initializing the Variables
    m_sControlPoint = "ll"
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_sSlotTypeText = "A"
    m_sWebLengthText = "0"
    m_dHorizOffset = 0
    m_dVertOffset = 0
    m_dTextGap = 0.5

End Sub

Private Sub IJDMfgOutputAnnotation_SetArguments(ByVal sSettingsXML As String)
Const METHOD = "IJDMfgOutputAnnotations_SetArguments"

On Error GoTo ErrorHandler

    Dim oXMLDomDoc                      As New DOMDocument
    Dim oAttributeNodeList              As IXMLDOMNodeList
    Dim oXMLNodeList                    As IXMLDOMNodeList
    Dim oXMLElement                     As IXMLDOMElement
    Dim sAttrName                       As String
    Dim sAttrValue                      As String
    Dim vTemp                           As Variant

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
                Case "HorizOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_dHorizOffset = Val(sAttrValue) * m_dTextSize
                Case "VertOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_dVertOffset = Val(sAttrValue) * m_dTextSize
                Case "TextGap"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_dTextGap = Val(sAttrValue) * m_dTextSize
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
    Dim oSlotTypeTextPos        As IJDPosition
    Dim oWebLenTextPos          As IJDPosition
    Dim sSlotTypeText           As String
    Dim sWebLenText             As String
    Dim sPART_SIDE                As String
    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim dExtraRotation          As Double
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
   
    sPART_SIDE = GetAttributeValueFromXML(sAttributeXML, "PART_SIDE")
    sSlotTypeText = GetAttributeValueFromXML(sAttributeXML, "SLOT_TYPE")
    sWebLenText = GetAttributeValueFromXML(sAttributeXML, "WEB_LENGTH")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    Set oSlotTypeTextPos = New DPosition
    Set oWebLenTextPos = New DPosition
    Set oStartPos = New DPosition

    oStartPos.Set 0, 0, 0
    
    Select Case UCase(sPART_SIDE)
        Case "R"
            oWebLenTextPos.Set m_dHorizOffset, -m_dVertOffset, 0
            oSlotTypeTextPos.Set m_dHorizOffset, -(m_dTextSize + m_dTextGap + m_dVertOffset), 0
            dExtraRotation = 3.14159265358979
        Case "L"
            oWebLenTextPos.Set m_dHorizOffset, m_dVertOffset, 0
            oSlotTypeTextPos.Set m_dHorizOffset, m_dTextSize + m_dTextGap + m_dVertOffset, 0
            dExtraRotation = 0
    End Select
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oSlotTypeTextPos, pOrientation, pStartPoint
    TranslatePoint oWebLenTextPos, pOrientation, pStartPoint
    
    'Create the Output
    
    
    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oSlotTypeTextPos, sSlotTypeText, _
                                            m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                            m_dTextSize, dExtraRotation, "SLOT_TYPE"
                                            
    
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oWebLenTextPos, sWebLenText, _
                                            m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                            m_dTextSize, dExtraRotation, "WEB_LENGTH"
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:
    Set oSlotTypeTextPos = Nothing
    Set oWebLenTextPos = Nothing
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function




