VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpecialMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\SpecialMark.cls
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
' 11/03/2010    Karthikeyan Ravi          Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.SpecialMark"
                
            

'General Properties
Private m_dFlangelengthconst          As Double

Private m_sControlPoint               As String
Private m_dTextSize                   As Double
Private m_sTextFont                   As String
Private m_sTextStyle                  As String
Private m_dIncliAngle                 As Double
Private m_dArrowLen                   As Double
Private m_dMajorLength                As Double
Private m_dMinorLength                As Double
Private m_dWebLength                  As Double
Private m_dVerticalOffset             As Double
Private m_dFlangeThickness1           As Double
Private m_dWebThickness1              As Double
Private m_dFlangeThickness2           As Double
Private m_dWebThickness2              As Double
Private m_dTextBottomOffset           As Double
Private m_dFlangeLength1              As Double
Private m_dFlangeLength2              As Double
Private m_dFilletRadius               As Double
Private m_sPrefix                     As String
Private m_sOffset                     As String

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values

    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    
    
    m_dMajorLength = 1 * m_dTextSize
    m_dMinorLength = 1 * m_dTextSize
    
    m_dWebLength = 7 * m_dTextSize
    m_dFlangeLength1 = 3 * m_dTextSize
    m_dFlangeLength2 = 3 * m_dTextSize
    m_dFlangelengthconst = 3 * m_dTextSize
    
    m_dFlangeThickness1 = 1 * m_dTextSize
    m_dWebThickness1 = 1 * m_dTextSize
    
    m_dFlangeThickness2 = 1 * m_dTextSize
    m_dWebThickness2 = 1 * m_dTextSize
    
    m_dVerticalOffset = 1 * m_dTextSize
    
    m_dIncliAngle = 0.262 * m_dTextSize
    m_dArrowLen = 0.375 * m_dTextSize
    
    m_dTextBottomOffset = 0.333 * m_dTextSize
    m_sPrefix = "6CV"
    m_sOffset = "25"
    
    m_sControlPoint = "lm"
       
    m_dFilletRadius = 1 * m_dTextSize

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
                Case "IncliAngle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dIncliAngle = Val(sAttrValue)
                    End If
                Case "MajorLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMajorLength = Val(sAttrValue) * m_dTextSize
                    End If
                Case "MinorLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dMinorLength = Val(sAttrValue) * m_dTextSize
                    End If

                Case "WebLength"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dWebLength = Val(sAttrValue) * m_dTextSize
                    End If
                
                Case "FlangeLength1"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlangeLength1 = Val(sAttrValue) * m_dTextSize
                    End If
                    
                Case "FlangeLength2"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlangeLength2 = Val(sAttrValue) * m_dTextSize
                    End If
                
                Case "ArrowLen"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowLen = Val(sAttrValue) * m_dTextSize
                    End If

                Case "TextBottomOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextBottomOffset = Val(sAttrValue) * m_dTextSize
                    End If
                
                Case "FlangeThickness1"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlangeThickness1 = Val(sAttrValue) * m_dTextSize
                    End If

                Case "WebThickness1"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dWebThickness1 = Val(sAttrValue) * m_dTextSize
                    End If
                
                Case "FlangeThickness2"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFlangeThickness2 = Val(sAttrValue) * m_dTextSize
                    End If

                Case "WebThickness2"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dWebThickness2 = Val(sAttrValue) * m_dTextSize
                    End If
                                  
                Case "FilletRadius"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dFilletRadius = Val(sAttrValue) * m_dTextSize
                    End If
                  
                Case "VerticalOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dVerticalOffset = Val(sAttrValue) * m_dTextSize
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



    
    Dim oTextPoint As IJDPosition
    Dim oVerticalTextPoint As IJDPosition
    
    Dim oMinorlengthStartPoint As IJDPosition
    Dim oMinorlengthEndPoint As IJDPosition
    Dim oMinorlength1StartPoint As IJDPosition
    Dim oMinorlength1EndPoint As IJDPosition
    
    Dim oMajorlengthStartPoint As IJDPosition
    Dim oMajorlengthEndPoint As IJDPosition
    
    Dim oFilletCenter As IJDPosition
    Dim oFilletArcStartPoint As IJDPosition
    Dim oFilletArcEndPoint As IJDPosition
    Dim oFilletLineEndPoint As IJDPosition
    
    Dim oFlangeLength1StartPoint As IJDPosition
    Dim oFlangeLength1EndPoint As IJDPosition
    Dim oFlangeLength2StartPoint As IJDPosition
    Dim oFlangeLength2EndPoint As IJDPosition
    
    
    Dim oFlangeLength1ThickStartPoint As IJDPosition
    Dim oFlangeLength1ThickEndPoint As IJDPosition
    Dim oFlangeLength2ThickStartPoint As IJDPosition
    Dim oFlangeLength2ThickEndPoint As IJDPosition
    
    
    
    Dim oWebThick1StartPoint As IJDPosition
    Dim oWebThick1EndPoint As IJDPosition
    Dim oWebThick2StartPoint As IJDPosition
    Dim oWebThick2EndPoint As IJDPosition
    
        
    
    Dim oVerticalLinePoint1 As IJDPosition
    Dim oVerticalLinePoint2 As IJDPosition
    Dim oVerticalLinePoint3 As IJDPosition
    
    Dim oArrow1Point As IJDPosition
    Dim oArrow2Point As IJDPosition


    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    
    Dim strGrinder                  As String
    Dim strGrinder1                 As String
    Dim vTemp                       As Variant

    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If

    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0

    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp

    strGrinder = GetAttributeValueFromXML(sAttributeXML, "DESCRIPTION")
    strGrinder1 = GetAttributeValueFromXML(sAttributeXML, "OFFSET")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    'normalize the vector
    pOrientation.length = 1

    
    Set oTextPoint = New DPosition
    Set oVerticalTextPoint = New DPosition
    
    Set oMinorlengthStartPoint = New DPosition
    Set oMinorlengthEndPoint = New DPosition
    Set oMinorlength1StartPoint = New DPosition
    Set oMinorlength1EndPoint = New DPosition
    
    Set oMajorlengthStartPoint = New DPosition
    Set oMajorlengthEndPoint = New DPosition
    
    Set oFilletCenter = New DPosition
    Set oFilletArcEndPoint = New DPosition
    Set oFilletArcStartPoint = New DPosition
    Set oFilletLineEndPoint = New DPosition
    
    Set oFlangeLength1StartPoint = New DPosition
    Set oFlangeLength1EndPoint = New DPosition
    Set oFlangeLength2StartPoint = New DPosition
    Set oFlangeLength2EndPoint = New DPosition
    
    Set oFlangeLength1ThickStartPoint = New DPosition
    Set oFlangeLength1ThickEndPoint = New DPosition
    Set oFlangeLength2ThickStartPoint = New DPosition
    Set oFlangeLength2ThickEndPoint = New DPosition
    
    Set oWebThick1StartPoint = New DPosition
    Set oWebThick1EndPoint = New DPosition
    Set oWebThick2StartPoint = New DPosition
    Set oWebThick2EndPoint = New DPosition
    
    Set oVerticalLinePoint1 = New DPosition
    Set oVerticalLinePoint2 = New DPosition
    Set oVerticalLinePoint3 = New DPosition
    
    Set oArrow1Point = New DPosition
    Set oArrow2Point = New DPosition
   
    
    'condition for increasing the major length if flange length is increasing
    
     If (m_dFlangeLength2 > m_dFlangelengthconst) Then
     
     m_dMajorLength = m_dMajorLength + (0.75 * m_dFlangeLength2)
     
     End If

    '*** Set All Points ***'

        ' condition for if fillet radius is zero
        
        If (m_dFilletRadius = 0) Then
                        
            oFilletArcStartPoint.Set (m_dFlangeLength1 + m_dMinorLength), 0, 0
            oFilletArcEndPoint.Set (m_dFlangeLength1 + m_dMinorLength), (m_dFlangeThickness2 + (m_dFlangeThickness2 / 10)), 0
            oFilletLineEndPoint.Set ((m_dMinorLength + (2 * m_dFlangeLength2) + m_dTextSize) - m_dWebThickness2), (m_dFlangeThickness2 + (m_dFlangeThickness2 / 10)), 0
                  
       Else
       
         ' to increase the fillet radius according to the flange thickness
            m_dFilletRadius = 1.2 * m_dFlangeThickness2
           
          oFilletCenter.Set (m_dFlangeLength1 + m_dMinorLength + m_dFilletRadius), 0, 0
          oFilletArcStartPoint.Set (m_dFlangeLength1 + m_dMinorLength), 0, 0
          oFilletArcEndPoint.Set (m_dFlangeLength1 + m_dMinorLength + m_dFilletRadius), m_dFilletRadius, 0
          oFilletLineEndPoint.Set ((m_dMinorLength + (2 * m_dFlangeLength2) + m_dTextSize) - m_dWebThickness2), m_dFilletRadius, 0
            
       End If

        oTextPoint.Set ((0.75 * (m_dFlangeLength1 + m_dFlangeLength2)) + (0.4 * m_dMinorLength)), m_dTextBottomOffset, 0
        oVerticalTextPoint.Set m_dWebThickness1, (m_dWebLength + 3 * m_dTextSize), 0
    
        oMinorlengthStartPoint.Set m_dFlangeLength1, (m_dWebLength - m_dVerticalOffset), 0
        oMinorlengthEndPoint.Set (((2 * m_dFlangeLength2) + m_dMinorLength + m_dTextSize) - m_dWebThickness2), (m_dWebLength - m_dVerticalOffset), 0
    
        oMinorlength1StartPoint.Set -0.05 * m_dMajorLength, (m_dWebLength - m_dVerticalOffset), 0
        oMinorlength1EndPoint.Set (0.05 * m_dMajorLength), (m_dWebLength - m_dVerticalOffset), 0
    
        oMajorlengthStartPoint.Set -0.05 * m_dMajorLength, m_dWebLength, 0
        oMajorlengthEndPoint.Set 0.95 * m_dMajorLength, m_dWebLength, 0
        
        oFlangeLength1StartPoint.Set 0, 0, 0
        oFlangeLength1EndPoint.Set m_dFlangeLength1, 0, 0
        oFlangeLength2StartPoint.Set (m_dFlangeLength1 + m_dMinorLength + m_dTextSize), 0, 0
        oFlangeLength2EndPoint.Set ((2 * m_dFlangeLength2) + m_dMinorLength + m_dTextSize), 0, 0
              
        oFlangeLength1ThickStartPoint.Set 0, m_dFlangeThickness1, 0
        oFlangeLength1ThickEndPoint.Set (m_dFlangeLength1 - m_dWebThickness1), m_dFlangeThickness1, 0
        oFlangeLength2ThickStartPoint.Set (m_dFlangeLength1 + m_dMinorLength + m_dTextSize), m_dFlangeThickness2, 0
        oFlangeLength2ThickEndPoint.Set ((m_dMinorLength + (2 * m_dFlangeLength2) + m_dTextSize) - m_dWebThickness2), m_dFlangeThickness2, 0
    
        oWebThick1StartPoint.Set (m_dFlangeLength1 - m_dWebThickness1), m_dWebLength, 0
        oWebThick1EndPoint.Set m_dFlangeLength1, m_dWebLength, 0
        oWebThick2StartPoint.Set ((m_dMinorLength + (2 * m_dFlangeLength2) + m_dTextSize) - m_dWebThickness2), m_dWebLength, 0
        oWebThick2EndPoint.Set (m_dMinorLength + (2 * m_dFlangeLength2) + m_dTextSize), m_dWebLength, 0
    
        oVerticalLinePoint1.Set (0), (m_dWebLength - m_dVerticalOffset), 0
        oVerticalLinePoint2.Set (0), m_dWebLength, 0
        oVerticalLinePoint3.Set (0), (m_dWebLength + 4 * m_dTextSize), 0
    
        oArrow1Point.Set -(m_dArrowLen * Sin(m_dIncliAngle)), ((m_dWebLength - m_dVerticalOffset) + m_dArrowLen * Cos(m_dIncliAngle)), 0
        oArrow2Point.Set -(m_dArrowLen * Sin(m_dIncliAngle)), (m_dWebLength + (m_dArrowLen * Cos(m_dIncliAngle))), 0
    
    '**********************'
    'If pOrientation.x < 0 Then
        'pOrientation.x = -1 * pOrientation.x
    'End If
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""
    
    TranslatePoint oTextPoint, pOrientation, pStartPoint
    TranslatePoint oVerticalTextPoint, pOrientation, pStartPoint
    
    TranslatePoint oMajorlengthStartPoint, pOrientation, pStartPoint
    TranslatePoint oMajorlengthEndPoint, pOrientation, pStartPoint
    
    TranslatePoint oMinorlengthStartPoint, pOrientation, pStartPoint
    TranslatePoint oMinorlengthEndPoint, pOrientation, pStartPoint
    TranslatePoint oMinorlength1StartPoint, pOrientation, pStartPoint
    TranslatePoint oMinorlength1EndPoint, pOrientation, pStartPoint
    
    TranslatePoint oFilletCenter, pOrientation, pStartPoint
    TranslatePoint oFilletArcStartPoint, pOrientation, pStartPoint
    TranslatePoint oFilletArcEndPoint, pOrientation, pStartPoint
    TranslatePoint oFilletLineEndPoint, pOrientation, pStartPoint
    
    TranslatePoint oFlangeLength1StartPoint, pOrientation, pStartPoint
    TranslatePoint oFlangeLength1EndPoint, pOrientation, pStartPoint
    TranslatePoint oFlangeLength2StartPoint, pOrientation, pStartPoint
    TranslatePoint oFlangeLength2EndPoint, pOrientation, pStartPoint
    
    TranslatePoint oFlangeLength1ThickStartPoint, pOrientation, pStartPoint
    TranslatePoint oFlangeLength1ThickEndPoint, pOrientation, pStartPoint
    TranslatePoint oFlangeLength2ThickStartPoint, pOrientation, pStartPoint
    TranslatePoint oFlangeLength2ThickEndPoint, pOrientation, pStartPoint
    
    TranslatePoint oWebThick1StartPoint, pOrientation, pStartPoint
    TranslatePoint oWebThick1EndPoint, pOrientation, pStartPoint
    TranslatePoint oWebThick2StartPoint, pOrientation, pStartPoint
    TranslatePoint oWebThick2EndPoint, pOrientation, pStartPoint
    
    TranslatePoint oVerticalLinePoint1, pOrientation, pStartPoint
    TranslatePoint oVerticalLinePoint2, pOrientation, pStartPoint
    TranslatePoint oVerticalLinePoint3, pOrientation, pStartPoint
    
    TranslatePoint oArrow1Point, pOrientation, pStartPoint
    TranslatePoint oArrow2Point, pOrientation, pStartPoint
    
    '*** Create the lines ***'
    'Majorline
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oMajorlengthStartPoint, oMajorlengthEndPoint

    'Minorline1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oMinorlength1StartPoint, oMinorlength1EndPoint

    'Minor line
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oMinorlengthStartPoint, oMinorlengthEndPoint

    'flange end to flange start
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength1EndPoint, oFilletArcStartPoint

    'from arc end to web
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFilletArcEndPoint, oFilletLineEndPoint

    '*****Flange 1 and Web 1********

    'flange1 length1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength1StartPoint, oFlangeLength1EndPoint
    
    'flange1 thickness
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength1StartPoint, oFlangeLength1ThickStartPoint

    'flange1 length2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength1ThickStartPoint, oFlangeLength1ThickEndPoint

    'web1 length1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength1ThickEndPoint, oWebThick1StartPoint

    'web1 length2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength1EndPoint, oWebThick1EndPoint


    '*****Flange 2 and Web 2********

    'flange2 length1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength2StartPoint, oFlangeLength2EndPoint

    'flange2 thickness
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength2StartPoint, oFlangeLength2ThickStartPoint

    'flange2 length2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength2ThickStartPoint, oFlangeLength2ThickEndPoint

    'web2 length1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength2ThickEndPoint, oWebThick2StartPoint

    'web1 length2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFlangeLength2EndPoint, oWebThick2EndPoint

    'fillet arc
     
    'condition for fillet radius when it is zero
    
    If (m_dFilletRadius = 0) Then
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oFilletArcEndPoint, oFilletArcStartPoint
    Else
        SMS_NodeCurveArc oOutputDom, oTempEdgeElement, oFilletArcEndPoint, oFilletArcStartPoint, oFilletCenter
    End If
    
    '****Vertical line and Arrow Head **********

    'vertical line1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVerticalLinePoint1, oVerticalLinePoint2

    'vertical line2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVerticalLinePoint2, oVerticalLinePoint3

    'Arrow head 1
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVerticalLinePoint1, oArrow1Point

    'Arrow head 2
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oVerticalLinePoint2, oArrow2Point

    '************************'
    
    '*** Create the XML Text ***'
    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint, strGrinder, _
                                      m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                      m_dTextSize, 0, "DESCRIPTION"

    SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oVerticalTextPoint, strGrinder1, _
                                      m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                      m_dTextSize, -1.5707, "OFFSET"

'    ****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    
     Set oTextPoint = Nothing
    Set oVerticalTextPoint = Nothing
    
    Set oMinorlengthStartPoint = Nothing
    Set oMinorlengthEndPoint = Nothing
    Set oMinorlength1StartPoint = Nothing
    Set oMinorlength1EndPoint = Nothing
    
    Set oMajorlengthStartPoint = Nothing
    Set oMajorlengthEndPoint = Nothing
    
    Set oFilletCenter = Nothing
    Set oFilletArcEndPoint = Nothing
    Set oFilletArcStartPoint = Nothing
    Set oFilletLineEndPoint = Nothing
    
    Set oFlangeLength1StartPoint = Nothing
    Set oFlangeLength1EndPoint = Nothing
    Set oFlangeLength2StartPoint = Nothing
    Set oFlangeLength2EndPoint = Nothing
    
    Set oFlangeLength1ThickStartPoint = Nothing
    Set oFlangeLength1ThickEndPoint = Nothing
    Set oFlangeLength2ThickStartPoint = Nothing
    Set oFlangeLength2ThickEndPoint = Nothing
    
    Set oWebThick1StartPoint = Nothing
    Set oWebThick1EndPoint = Nothing
    Set oWebThick2StartPoint = Nothing
    Set oWebThick2EndPoint = Nothing
    
    Set oVerticalLinePoint1 = Nothing
    Set oVerticalLinePoint2 = Nothing
    Set oVerticalLinePoint3 = Nothing
    
    Set oArrow1Point = Nothing
    Set oArrow2Point = Nothing
    
    
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
'oTestPos.Set 0, 150, 0
'TranslatePoint oTestPos, pOrientation, pStartPoint
''MsgBox "startx: " & pStartPoint.x & ", starty: " & pStartPoint.y & vbNewLine _
''    & ", endx: " & oTestPos.x & ", endy: " & oTestPos.y
'Set oTempCurveElement = CreateSingleLineCurveNode(oOutputDom, pStartPoint, oTestPos, "plate_thickness_annotation")
'If Not oTempCurveElement Is Nothing Then oTempEdgeElement.appendChild oTempCurveElement
'
'    oOutputElem.appendChild oTempEdgeElement










