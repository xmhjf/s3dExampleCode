VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TackWeldMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\TackWeldMark.cls
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
' 11/08/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.TackWeldMark"

'General Properties
Private m_sControlPoint                 As String
Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                     As String
Private m_dWeldHeight                  As Double
Private m_dWeldLength                  As Double
Private m_dLength               As Double
Private m_dOffset              As Double


Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dWeldHeight = 1.5 * m_dTextSize
    m_dWeldLength = 0.6 * m_dTextSize
    m_sControlPoint = "ll"
    m_dLength = 2 * m_dTextSize
    m_dOffset = 0.5 * m_dTextSize
    'm_SeamText = "100"
    
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
                Case "WeldHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dWeldHeight = Val(sAttrValue) * m_dTextSize
                    End If
                Case "WeldLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dWeldLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "Length"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "Offset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dOffset = Val(sAttrValue) * m_dTextSize
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
                    

    Dim m_dLengthofmark As Double
    m_dLengthofmark = pOrientation.length
            
    Dim oStartPoint             As IJDPosition
    Dim oLine1EndPoint          As IJDPosition
    Dim oLine2EndPoint          As IJDPosition
    Dim oLine3EndPoint          As IJDPosition
    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim oTempCurveElement       As IXMLDOMElement
    
    Dim vTemp                   As Variant
    Dim strTackweldCondition    As String
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
    
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")
    strTackweldCondition = GetAttributeValueFromXML(sAttributeXML, "TACKWELD_CONDITION")

    'normalize the vector
    pOrientation.length = 1

    Set oStartPoint = New DPosition
    Set oLine1EndPoint = New DPosition
    Set oLine2EndPoint = New DPosition
    Set oLine3EndPoint = New DPosition

    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
           
    '*** Create the lines ***'
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
     
     Select Case (strTackweldCondition)
     
            Case "L"
                
            oStartPoint.Set 0, 0, 0
            oLine1EndPoint.Set 0, m_dWeldHeight, 0
            oLine2EndPoint.Set -m_dOffset, m_dWeldHeight, 0
            
            
            TranslatePoint oStartPoint, pOrientation, pStartPoint
            TranslatePoint oLine1EndPoint, pOrientation, pStartPoint
            TranslatePoint oLine2EndPoint, pOrientation, pStartPoint
            
            'Line 1 -- 1 to 2
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine1EndPoint

             'Line 2 -- 2 to 3
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1EndPoint, oLine2EndPoint
            
            
            Case "C"
              
                oStartPoint.Set 0, 0, 0
                oLine1EndPoint.Set 0, m_dWeldHeight, 0
                oLine2EndPoint.Set m_dWeldLength, m_dWeldHeight, 0
                oLine3EndPoint.Set m_dWeldLength, 0, 0
            
                TranslatePoint oStartPoint, pOrientation, pStartPoint
                TranslatePoint oLine1EndPoint, pOrientation, pStartPoint
                TranslatePoint oLine2EndPoint, pOrientation, pStartPoint
                TranslatePoint oLine3EndPoint, pOrientation, pStartPoint
                'Line 1 -- 1 to 2
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine1EndPoint

                'Line 2 -- 2 to 3
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1EndPoint, oLine2EndPoint
               
                'Line 3 -- 2 to 3

                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine2EndPoint, oLine3EndPoint
     
             
            Case "R"
                 oStartPoint.Set 0, 0, 0
                 oLine1EndPoint.Set 0, m_dWeldHeight, 0
                 oLine2EndPoint.Set m_dOffset, m_dWeldHeight, 0
            
                TranslatePoint oStartPoint, pOrientation, pStartPoint
                TranslatePoint oLine1EndPoint, pOrientation, pStartPoint
                TranslatePoint oLine2EndPoint, pOrientation, pStartPoint
            
                'Line 1 -- 1 to 2
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPoint, oLine1EndPoint

                'Line 2 -- 2 to 3
                SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oLine1EndPoint, oLine2EndPoint
            
     End Select
     
    '************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:
    Set oStartPoint = Nothing
    Set oLine1EndPoint = Nothing
    Set oLine2EndPoint = Nothing
    Set oLine3EndPoint = Nothing
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














