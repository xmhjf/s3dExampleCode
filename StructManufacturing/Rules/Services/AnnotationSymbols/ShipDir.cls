VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShipDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\ShipDir.cls
' Project: M:\SharedContent\Src\StructManufacturing\Rules\Services\AnnotationSymbols\AnnotationSymbols.vbp
'
'
' Abstract:
'   Create custom Ship Direction Annotation for Templates in Template Drawings
'
' Description:
'   Implements the MfgOutputAnnotation interface for creating custom output symbols and marks to create
'   a Ship Direction Annotation for Templates in Template Drawings
'
' History:
' 09/15/2010    Santosh Ram           Created
'***************************************************************************


Option Explicit
Private Const MODULE = "AnnotationSymbols.ShipDir"


'General Properties
Private m_sControlPoint                 As String
Private m_dVertOffset                   As Double
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dHorizOffset                  As Double
Private m_dTextOffset                   As Double


Implements IJDMfgOutputAnnotation

Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 30
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dTextOffset = m_dTextSize
    m_dHorizOffset = m_dTextSize * 0.4
    m_dVertOffset = m_dTextSize * 3
    m_sControlPoint = "ll"
    
End Sub

Private Sub IJDMfgOutputAnnotation_SetArguments(ByVal sSettingsXML As String)
    Const METHOD = "IJDMfgOutputAnnotation_SetArguments"
    On Error GoTo ErrorHandler

    'All the Geom Args of the Annotation are set in this SubRoutine
    
    Dim oXMLDomDoc                  As New DOMDocument
    Dim oAttributeNodeList          As IXMLDOMNodeList
    Dim oXMLElement                 As IXMLDOMElement
    Dim sAttrName                   As String
    Dim sAttrValue                  As String
    Dim vTemp                       As Variant

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
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
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
                        m_dVertOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "HorizontalOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dHorizOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_dTextOffset = Val(sAttrValue) * m_dTextSize
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
    
    'All the Text Args of the Annotation are set in this Function
             
    Dim oStartPoint                 As IJDPosition          '#1
    Dim oShipNamePoint                As IJDPosition          '#2
    Dim oPartNamePoint                 As IJDPosition          '#3
    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim strShipName                 As String
    Dim strPartName                 As String
    Dim vTemp                       As Variant
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    strShipName = GetAttributeValueFromXML(sAttributeXML, "SHIP_NAME")
    strPartName = GetAttributeValueFromXML(sAttributeXML, "PART_NAME")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")
    
    'normalize the vector
    pOrientation.length = 1
    
    Set oStartPoint = New DPosition                     '#1
    Set oShipNamePoint = New DPosition                    '#2
    Set oPartNamePoint = New DPosition                     '#3

    
    '*** Set All Points ***'
    m_dVertOffset = Round(m_dVertOffset, 1)
    m_dHorizOffset = Round(m_dHorizOffset, 1)
    m_dTextOffset = Round(m_dTextOffset, 1)

    oStartPoint.Set 0, 0, 0                                                 '#1
    oShipNamePoint.Set m_dHorizOffset, m_dVertOffset, 0                                   '#2
    oPartNamePoint.Set m_dHorizOffset, m_dVertOffset + m_dTextOffset + m_dTextSize, 0                     '#3

    
     '**********************'
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oStartPoint, pOrientation, pStartPoint
    TranslatePoint oShipNamePoint, pOrientation, pStartPoint
    TranslatePoint oPartNamePoint, pOrientation, pStartPoint
    
    

    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oShipNamePoint, strShipName, _
                                        m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                        m_dTextSize, 0, "SHIP_NAME"
                                        
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oPartNamePoint, strPartName, _
                                      m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                      m_dTextSize, 0, "PART_NAME"
    '****************************'
    
    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPoint = Nothing               '#1
    Set oShipNamePoint = Nothing              '#2
    Set oPartNamePoint = Nothing               '#3
    
    Set oOutputDom = Nothing
    Set oOutputElem = Nothing


    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Function







