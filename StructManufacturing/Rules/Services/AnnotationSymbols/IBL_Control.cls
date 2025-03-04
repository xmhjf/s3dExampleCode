VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBL_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\IBL_Control.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\MfgCustomAnnotation.vbp
'
'
' Abstract:
'
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   Inverse Bending Line output
'
' History:
' 4/4/2011    Ninad           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.IBL_Control"

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
    Dim strLANDING_STRAIGHT_LENGTH  As String
    Dim strMARK_HEIGHT              As String
    Dim vTemp                       As Variant
    Dim oTextPoint1                 As IJDPosition
    Dim oTextPoint2                 As IJDPosition
    Dim oTextPoint3                 As IJDPosition
    Dim oTextPoint4                 As IJDPosition
    
    strLANDING_STRAIGHT_LENGTH = GetAttributeValueFromXML(sAttributeXML, "LANDING_STRAIGHT_LENGTH")
    strMARK_HEIGHT = GetAttributeValueFromXML(sAttributeXML, "MARK_HEIGHT")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    Set oTextPoint1 = New DPosition
    Set oTextPoint2 = New DPosition
    Set oTextPoint3 = New DPosition
    Set oTextPoint4 = New DPosition
    
    strLANDING_STRAIGHT_LENGTH = Trim(Str(Round(Val(strLANDING_STRAIGHT_LENGTH) * 1000, 1)))
    strMARK_HEIGHT = Trim(Str(Round(Val(strMARK_HEIGHT) * 1000, 1)))
    
    oTextPoint1.Set (Len("(") + Len(strMARK_HEIGHT) + Len(")")) * m_dTextSize / 2, 0, 0
    oTextPoint2.Set (-Len(strLANDING_STRAIGHT_LENGTH) + Len(strMARK_HEIGHT) + Len(")")) * m_dTextSize / 2, 0, 0
    oTextPoint3.Set (-Len("(") - Len(strLANDING_STRAIGHT_LENGTH) + Len(")")) * m_dTextSize / 2, 0, 0
    oTextPoint4.Set -(Len("(") + Len(strMARK_HEIGHT) + Len(strLANDING_STRAIGHT_LENGTH)) * m_dTextSize / 2, 0, 0
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
   
    TranslatePoint oTextPoint1, pOrientation, pStartPoint              '#1
    TranslatePoint oTextPoint2, pOrientation, pStartPoint
    TranslatePoint oTextPoint3, pOrientation, pStartPoint
    TranslatePoint oTextPoint4, pOrientation, pStartPoint

    'normalize the vector
    pOrientation.length = 1
    
    Dim sTextString As String
    sTextString = strLANDING_STRAIGHT_LENGTH & "(" & strMARK_HEIGHT & ")"

    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint1, strLANDING_STRAIGHT_LENGTH, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 3.1415926, "LANDING_STRAIGHT_LENGTH"
                                        
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint2, "(", _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 3.1415926
                                        
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint3, strMARK_HEIGHT, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 3.1415926, "MARK_HEIGHT"
                                        
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint4, ")", _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 3.1415926
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:

    Set oOutputDom = Nothing
    Set oOutputElem = Nothing
    Set oTextPoint1 = Nothing
    Set oTextPoint2 = Nothing
    Set oTextPoint3 = Nothing
    Set oTextPoint4 = Nothing

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










