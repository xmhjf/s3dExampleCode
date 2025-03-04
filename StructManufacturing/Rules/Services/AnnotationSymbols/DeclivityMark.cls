VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeclivityMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\ArrowMark1.cls
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
' 11/03/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.DeclivityMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_sText                         As String

Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
' Initializing the Variables
    m_sControlPoint = "ll"
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_sText = "80"

End Sub

Private Sub IJDMfgOutputAnnotation_SetArguments(ByVal sSettingsXML As String)
Const METHOD = "IJDMfgOutputAnnotations_SetArguments"

On Error GoTo ErrorHandler

    Dim oXMLDomDoc As New DOMDocument
    Dim oAttributeNodeList As IXMLDOMNodeList
    Dim oXMLNodeList As IXMLDOMNodeList
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

    Dim oTextPosition1          As IJDPosition
    Dim oTextPosition2          As IJDPosition
    Dim strGUID                 As String
    Dim strUpside                As String
    Dim strAngle                As String
    
    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
        
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    
    strUpside = GetAttributeValueFromXML(sAttributeXML, "UPSIDE")
    strAngle = Val(GetAttributeValueFromXML(sAttributeXML, "ANGLE"))
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1
    
    strAngle = Trim(Str(Round(Val(strAngle), 1)))
    
    Dim sText1 As String
    
    Dim oUpsidePosition As IJDPosition
    Dim oAnglePosition As IJDPosition
    
    Set oUpsidePosition = New DPosition
    Set oAnglePosition = New DPosition
    
    oUpsidePosition.Set -Len(strAngle) * m_dTextSize / 2, 0, 0
    oAnglePosition.Set oUpsidePosition.X + ((Len(strAngle) + Len(strUpside)) * m_dTextSize / 2), 0, 0

    'Get the combined string
    sText1 = strUpside & strAngle

    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oUpsidePosition, pOrientation, pStartPoint              '#1
    TranslatePoint oAnglePosition, pOrientation, pStartPoint              '#1
    
    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oAnglePosition, strAngle, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 0, "ANGLE"
                                        
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oUpsidePosition, strUpside, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 0, "UPSIDE"
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:
    Set oTextPosition1 = Nothing
    Set oTextPosition2 = Nothing
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oUpsidePosition = Nothing
    Set oAnglePosition = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function


