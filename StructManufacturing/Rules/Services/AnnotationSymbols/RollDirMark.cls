VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RollDirMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\RollDirMark.cls
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
' 11/04/2010    Santosh Ram           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.RollDirMark"

'General Properties
Private m_sControlPoint                 As String

Private m_dTextSize                     As Double
Private m_sTextFont                     As String
Private m_sTextStyle                    As String
Private m_dIncliAngle                   As Double
Private m_dArrowLen                     As Double
Private m_dRollRadius                   As Double
Private m_dSweepAngle                   As Double
Private m_dKnuckleAngle                 As Double
Private m_dSideOffset                   As Double
Private m_dTextBottomOffset             As Double
Private m_sUpside1                      As String
Private m_sPrefix2                      As String
Private m_sPrefix3                      As String
Private m_sRollerCondition              As String
Private m_dSecndryOffset                As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_sTextFont = "Arial"
    m_sTextStyle = "Regular"
    m_dRollRadius = 8 * m_dTextSize
    m_dIncliAngle = 0.262
    m_dSweepAngle = 2.618
    m_dKnuckleAngle = 2.618
    m_dArrowLen = 0.375 * m_dTextSize
    m_dSecndryOffset = 0.1
    
    m_dSideOffset = 200
    m_dTextBottomOffset = 0.333 * m_dTextSize
    m_sControlPoint = "ll"
    
    m_sUpside1 = "$"
    m_sPrefix2 = "1050R"
    m_sPrefix3 = "ROLLER"
    m_sRollerCondition = 1
    
    
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
                Case "ArrowLen"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dArrowLen = Val(sAttrValue) * m_dTextSize
                    End If
                Case "SecondaryOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSecndryOffset = Val(sAttrValue)
                    End If
                Case "TextBottomOffset"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextBottomOffset = Val(sAttrValue) * m_dTextSize
                    End If
                Case "TextFont"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextFont = sAttrValue
                Case "TextStyle"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sTextStyle = sAttrValue
                Case "RollerCondition"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    m_sRollerCondition = sAttrValue
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
 

'                           #2
'                           |\ #5
'                           |
'                           |#6 #8
'                           | #7
'               ____________|___________
'                 '#1       |
'                 '         |
'                 '         |
'                           | #4
'                           |/
'                           #3
'

    Dim oStartPoint         As IJDPosition          '#1
    Dim oArrow1Start        As IJDPosition          '#2
    Dim oArrow2Start        As IJDPosition          '#3
    Dim oArrow2End          As IJDPosition          '#4
    Dim oArrow1End          As IJDPosition          '#5
    Dim oTextPoint1         As IJDPosition          '#6
    Dim oTextPoint2         As IJDPosition          '#7
    Dim oTextPoint3         As IJDPosition          '#8

    
    Dim oOutputDom                  As New DOMDocument
    Dim oOutputElem                 As IXMLDOMElement
    Dim oTempEdgeElement            As IXMLDOMElement
    Dim oTempCurveElement           As IXMLDOMElement
    Dim oPlateMarkDom               As New DOMDocument
    Dim oPlateMarkElem              As IXMLDOMAttribute
    Dim oRollDirElement             As IXMLDOMElement
    
    
    Dim strUpside               As String
    Dim sKnuckleAngle           As String
    Dim RollRadius              As String
    Dim strText3                As String
    Dim vTemp                   As Variant
        
    '*** Create the XML Text ***'
    Dim oCVGTextElem1            As IXMLDOMElement
    Dim oCVGTextElem2            As IXMLDOMElement
    Dim oCVGTextElem3            As IXMLDOMElement
    Dim sText1                   As String
    Dim dTemp As Double
    
    
    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    'PART_DIR will always be L. So adjust the vector
    'pOrientation.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
    
    '*** Get GUID and PartDir ***'
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    'If sPartDir = "" Then GoTo CleanUp
   
    strUpside = GetAttributeValueFromXML(sAttributeXML, "UPSIDE")
    RollRadius = GetAttributeValueFromXML(sAttributeXML, "ROLL_RADIUS")
    m_dRollRadius = Val(RollRadius)
    RollRadius = Trim(Str(Round(m_dRollRadius, 1)))
    
    strText3 = GetAttributeValueFromXML(sAttributeXML, "DESCRIPTION2")
    m_dTextSize = Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))
    m_sTextFont = GetAttributeValueFromXML(sAttributeXML, "TextFont")
    m_sTextStyle = GetAttributeValueFromXML(sAttributeXML, "TextStyle")

    m_dSweepAngle = Val(GetAttributeValueFromXML(sAttributeXML, "SWEEP_ANGLE"))

    m_dKnuckleAngle = Val(GetAttributeValueFromXML(sAttributeXML, "KNUCKLE_ANGLE"))
    m_dKnuckleAngle = 180 / PI * m_dKnuckleAngle
    m_dKnuckleAngle = Round(m_dKnuckleAngle, 1)
    sKnuckleAngle = Trim(Str(m_dKnuckleAngle))
    
    'normalize the vector
    pOrientation.length = 1
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, ""

    Set oStartPoint = New DPosition                    '#1
    Set oArrow1Start = New DPosition                  '#2
    Set oArrow2Start = New DPosition               '#3
    Set oArrow2End = New DPosition               '#4
    Set oArrow1End = New DPosition               '#5
    Set oTextPoint1 = New DPosition                  '#7
    Set oTextPoint2 = New DPosition
    Set oTextPoint3 = New DPosition
    
    
    m_dSideOffset = ROLL_DIR_ANNOTATION_LENGTH
    '*** Set All Points ***'
    Select Case (m_sRollerCondition)
    
        Case "1"   ' With Arrow & Text
            
            'Adding the following so that the annotation symbol doesn't change
            If m_dRollRadius * m_dSweepAngle < 2 * m_dSideOffset Then
                m_dSideOffset = m_dRollRadius * m_dSweepAngle * m_dSecndryOffset
            End If
            
            'Calculate the points
            oStartPoint.Set 0, 0, 0                                                                 '#1
            oArrow1Start.Set 0, ((m_dRollRadius * m_dSweepAngle / 2) - m_dSideOffset), 0                         '#2
            oArrow2Start.Set 0, -((m_dRollRadius * m_dSweepAngle / 2) - m_dSideOffset), 0                      '#3
            oArrow2End.Set (m_dArrowLen * Sin(m_dIncliAngle)), -((m_dRollRadius * m_dSweepAngle / 2) - m_dSideOffset - (m_dArrowLen * Cos(m_dIncliAngle))), 0 '#4
            oArrow1End.Set (m_dArrowLen * Sin(m_dIncliAngle)), ((m_dRollRadius * m_dSweepAngle / 2) - m_dSideOffset - (m_dArrowLen * Cos(m_dIncliAngle))), 0       '#5
                                       '#8
            
            
            '**********************'
        '    If pOrientation.x < 0 Then
        '       pOrientation.x = -1 * pOrientation.x
        '    End If

            
            TranslatePoint oStartPoint, pOrientation, pStartPoint           '#1
            TranslatePoint oArrow1Start, pOrientation, pStartPoint          '#2
            TranslatePoint oArrow2Start, pOrientation, pStartPoint          '#3
            TranslatePoint oArrow2End, pOrientation, pStartPoint            '#4
            TranslatePoint oArrow1End, pOrientation, pStartPoint            '#5
            
            
            'Adjust the vector
            dTemp = pOrientation.X
            pOrientation.X = pOrientation.Y
            pOrientation.Y = -dTemp
            
            oTextPoint1.Set -(Len(RollRadius) + Len("R")) * m_dTextSize / 2, m_dTextSize + m_dTextBottomOffset, 0 '#6
            oTextPoint2.Set (Len(strUpside) - Len("R") - 0.5) * m_dTextSize / 2, m_dTextSize + m_dTextBottomOffset, 0
            oTextPoint3.Set (Len(strUpside) + Len(RollRadius) - 0.5) * m_dTextSize / 2, m_dTextSize + m_dTextBottomOffset, 0
            TranslatePoint oTextPoint1, pOrientation, pStartPoint     '#6
            TranslatePoint oTextPoint2, pOrientation, pStartPoint           '#7
            TranslatePoint oTextPoint3, pOrientation, pStartPoint           '#8
           
            '*** Create the lines ***'
            'From 2 to 3 -- Line 1
            Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
            oOutputElem.appendChild oTempEdgeElement
    
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrow1Start, oArrow2Start
        
            'From 2 to 5 -- Line 2
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrow1Start, oArrow1End
        
            'From 3 to 4 -- Line 3
            SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oArrow2Start, oArrow2End
            
            Set oTempCurveElement = Nothing
            Set oTempEdgeElement = Nothing
            
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint1, strUpside, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 0, "UPSIDE"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint2, RollRadius, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 0, "ROLL_RADIUS"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint3, "R", _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 0

            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, pStartPoint, strText3, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 0, "DESCRIPTION2"
            
            '************************'
        Case "2"  'Text parallel to mark
  
'            __RE______________________________
'
'                        #1 #2
'            ----------------------------------
'
'            ___________________________________
'             RE

            'Text is 90 degree extra rotated (parallel) to the mark
            oStartPoint.Set 0, 0, 0
            
            oTextPoint1.Set (Len(RollRadius) + Len("R")) * m_dTextSize / 2, 0, 0               '#6
            oTextPoint2.Set (-Len(strUpside) + Len("R") + 0.5) * m_dTextSize / 2, 0, 0
            oTextPoint3.Set -(Len(strUpside) + Len(RollRadius) - 0.5) * m_dTextSize / 2, 0, 0
            
            
            TranslatePoint oTextPoint1, pOrientation, pStartPoint           '#6
            TranslatePoint oTextPoint2, pOrientation, pStartPoint           '#7
            TranslatePoint oTextPoint3, pOrientation, pStartPoint
            
            'Make one string from MarkingSide and Angle

            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint1, strUpside, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI, "UPSIDE"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint2, RollRadius, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI, "ROLL_RADIUS"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint3, "R", _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI

        Case "3" 'Text normal to mark
'                |      |        |
'                |      |        |
'              RE|    #1| #2     |RE
'                |      |        |
'                |      |        |
'                |      |        |

            oStartPoint.Set 0, 0, 0

            oTextPoint1.Set 0, (Len(RollRadius) + Len("R")) * m_dTextSize / 2, 0 '#6
            oTextPoint2.Set 0, -(Len(strUpside) - Len("R") - 0.8) * m_dTextSize / 2, 0
            oTextPoint3.Set 0, -(Len(strUpside) + Len(RollRadius) - 0.7) * m_dTextSize / 2, 0
            
            TranslatePoint oTextPoint1, pOrientation, pStartPoint     '#6
            TranslatePoint oTextPoint2, pOrientation, pStartPoint           '#7
            TranslatePoint oTextPoint3, pOrientation, pStartPoint           '#8
            
            
            '***** USE FOLLOWING CODE TO PLACE INDIVIDUAL TEXTS ***'
'            oTextPoint1.Set ((m_dArrowLen * Sin(m_dIncliAngle)) + m_dTextSize + (2 * m_dTextBottomOffset)), Len(strText3) * m_dTextSize / 2, 0    '#6
'            oTextPoint2.Set ((m_dArrowLen * Sin(m_dIncliAngle)) + m_dTextSize + (2 * m_dTextBottomOffset)), ((Len(strText3) / 2) - Len(strUpside)) * m_dTextSize, 0                  '#7
'
'
'            TranslatePoint oTextPoint1, pOrientation, pStartPoint         '#6
'            TranslatePoint oTextPoint2, pOrientation, pStartPoint         '#7
            
'            Set oCVGTextElem1 = CreateTextNode(oOutputDom, m_sControlPoint, oTextPoint1, strUpside, _
'                                              m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
'                                               m_dTextSize, 0, "Upside")
'
'            oOutputElem.appendChild oCVGTextElem1
'
            
            
'            Set oCVGTextElem2 = CreateTextNode(oOutputDom, m_sControlPoint, oTextPoint2, RollRadius, _
'                                              m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
'                                              m_dTextSize, -1.5707, "RollRadius")
'
'            oOutputElem.appendChild oCVGTextElem2

            '*******************************************************'
            
            '***** USE FOLLOWING CODE TO PLACE BOTH TEXTS AS ONE STRING ***'
            'Make one string from MarkingSide and Angle
            
            'Rotate Vector by 90 degrees

            dTemp = pOrientation.X
            pOrientation.X = pOrientation.Y
            pOrientation.Y = -dTemp
            
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint1, strUpside, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 0, "UPSIDE"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint2, RollRadius, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 0, "ROLL_RADIUS"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint3, "R", _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, 0

            '******************************************************'
        Case "4"  'Text parallel to mark

'            __RE______________________________
'
'                        #1     #2
'            ----------------------------------
'
'            ___________________________________
'             RE
            
            'Text is 90 degree extra rotated (parallel) to the mark
            oStartPoint.Set 0, 0, 0
            oTextPoint1.Set (Len(RollRadius) + Len("R") + 1 + Len("KL") + Len(strUpside) + Len(sKnuckleAngle)) * m_dTextSize / 2, 0, 0 '#6
            oTextPoint2.Set oTextPoint1.X - ((Len(strUpside) + Len(RollRadius)) * m_dTextSize / 2), 0, 0
            oTextPoint3.Set oTextPoint2.X - ((Len("R") + Len(RollRadius) + 0.5) * m_dTextSize / 2), 0, 0
            oArrow1Start.Set oTextPoint3.X - ((Len("R") + Len("KL") + 1) * m_dTextSize / 2), 0, 0                      '#2
            oArrow2Start.Set oArrow1Start.X - ((Len(strUpside) + Len("KL")) * m_dTextSize / 2), 0, 0  '#4
            oArrow2End.Set oArrow2Start.X - ((Len(strUpside) + Len(sKnuckleAngle)) * m_dTextSize / 2), 0, 0
            
            TranslatePoint oTextPoint1, pOrientation, pStartPoint              '#1
            TranslatePoint oTextPoint2, pOrientation, pStartPoint              '#2
            TranslatePoint oTextPoint3, pOrientation, pStartPoint              '#3
            TranslatePoint oArrow1Start, pOrientation, pStartPoint              '#1
            TranslatePoint oArrow2Start, pOrientation, pStartPoint              '#2
            TranslatePoint oArrow2End, pOrientation, pStartPoint              '#3
            
            'Make one string from MarkingSide and Angle
            sText1 = strUpside & RollRadius & " " & sKnuckleAngle
            
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint1, strUpside, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI, "UPSIDE"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint2, RollRadius, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI, "ROLL_RADIUS"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oTextPoint3, "R", _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oArrow1Start, "KL", _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI
                                        
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oArrow2Start, strUpside, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI, "UPSIDE"
                                                
            SMS_NodeText oOutputDom, oOutputElem, m_sControlPoint, oArrow2End, sKnuckleAngle, _
                                                m_sTextFont, pOrientation, m_sTextStyle, "partmon", _
                                                m_dTextSize, PI, "KNUCKLE_ANGLE"

                    
    End Select
    
    Set oCVGTextElem1 = Nothing
    Set oCVGTextElem2 = Nothing
    Set oCVGTextElem3 = Nothing
    
    '****************************'

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML
    
CleanUp:
    
    Set oStartPoint = Nothing                    '#1
    Set oArrow1End = Nothing                  '#2
    Set oArrow2End = Nothing               '#3
    Set oArrow1Start = Nothing                  '#4
    Set oArrow2Start = Nothing                  '#4
    Set oTextPoint1 = Nothing               '#5
    Set oTextPoint2 = Nothing
    Set oTextPoint3 = Nothing
    
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










