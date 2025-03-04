VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateThickType1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
' Copyright(C)2000 , Intergraph Corporation. All Rights Reserved.
'
' File: M:\SharedContent\Src\StructManufacturing\Rules\Services\PartMonitorCustomAnnotation\PlateThickType1.cls
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
' 02/8/2010    Nathan Bruner           Created
'***************************************************************************

Option Explicit
Private Const MODULE = "MfgCustomAnnotation.PlateThickType1"

'General Properties
Private m_dTextSize                     As Double
Private m_dCenteredStartX               As Double
Private m_dCenteredStartY               As Double
Private m_dRightStartX                  As Double
Private m_dRightStartY                  As Double
Private m_dLeftStartX                   As Double
Private m_dLeftStartY                   As Double
Private m_dStraightHeight               As Double
Private m_dStraight2VertStart           As Double
Private m_dStraight2HorizStart          As Double
Private m_dSlant1VertStart              As Double
Private m_dSlant2VertStart              As Double

Implements IJDMfgOutputAnnotation


Private Sub Class_Initialize()
    'Set attributes to default values
    
    'General Properties
    m_dTextSize = 40
    m_dCenteredStartX = 1.625
    m_dCenteredStartY = 2.5
    m_dRightStartX = 0
    m_dRightStartY = 1.875
    m_dLeftStartX = 3.25
    m_dLeftStartY = 3.125
    m_dStraightHeight = 3.75
    m_dStraight2VertStart = 1.25
    m_dStraight2HorizStart = 3.25
    m_dSlant1VertStart = 1.25
    m_dSlant2VertStart = 2.5
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
                Case "TextSize"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dTextSize = Val(sAttrValue)
                    End If
                Case "CenteredStartX"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dCenteredStartX = Val(sAttrValue)
                    End If
                Case "CenteredStartY"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dCenteredStartY = Val(sAttrValue)
                    End If
                Case "RightStartX"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dRightStartX = Val(sAttrValue)
                    End If
                Case "RightStartY"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dRightStartY = Val(sAttrValue)
                    End If
                Case "LeftStartX"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dLeftStartX = Val(sAttrValue)
                    End If
                Case "LeftStartY"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dLeftStartY = Val(sAttrValue)
                    End If
                Case "StraightHeight"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dStraightHeight = Val(sAttrValue)
                    End If
                Case "Straight2VertStart"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dStraight2VertStart = Val(sAttrValue)
                    End If
                Case "Straight2HorizStart"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dStraight2HorizStart = Val(sAttrValue)
                    End If
                Case "Slant1VertStart"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlant1VertStart = Val(sAttrValue)
                    End If
                Case "Slant2VertStart"
                    vTemp = Trim(oXMLElement.getAttribute("VALUE"))
                    sAttrValue = IIf(VarType(vTemp) = vbString, vTemp, "")
                    If IsNumeric(sAttrValue) Then
                        m_dSlant2VertStart = Val(sAttrValue)
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
    

    Dim oStartPos               As IJDPosition
    Dim oStraight1End           As IJDPosition
    Dim oStraight2Start         As IJDPosition
    Dim oStraight2End           As IJDPosition
    Dim oSlant1Start            As IJDPosition
    Dim oSlant1End              As IJDPosition
    Dim oSlant2Start            As IJDPosition
    Dim oSlant2End              As IJDPosition

    Dim oOutputDom              As New DOMDocument
    Dim oOutputElem             As IXMLDOMElement
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim oTempCurveElement       As IXMLDOMElement

    Dim sPartDir                As String

    If pStartPoint Is Nothing Or pOrientation Is Nothing Or sAttributeXML = "" Then
        GoTo CleanUp
    End If
    
    sPartDir = GetAttributeValueFromXML(sAttributeXML, "PART_DIR")
    
    'FillPlateThickValuesFromXML sAttributeXML, strGUID, sPartDir
    If sPartDir = "" Then GoTo CleanUp


    'normalize the vector
    pOrientation.length = 1

    Set oStartPos = New DPosition
    Set oStraight1End = New DPosition
    Set oStraight2Start = New DPosition
    Set oStraight2End = New DPosition
    Set oSlant1Start = New DPosition
    Set oSlant1End = New DPosition
    Set oSlant2Start = New DPosition
    Set oSlant2End = New DPosition
    
    Select Case sPartDir
        Case "R", "r"
            oStartPos.Set m_dRightStartX * m_dTextSize, -m_dRightStartY * m_dTextSize, 0
        Case "L", "l"
            oStartPos.Set -m_dLeftStartX * m_dTextSize, -m_dLeftStartY * m_dTextSize, 0
        Case Else
            oStartPos.Set -m_dCenteredStartX * m_dTextSize, -m_dCenteredStartY * m_dTextSize, 0
    End Select
    
    oStraight1End.Set oStartPos.X, oStartPos.Y + m_dStraightHeight * m_dTextSize, 0
    oStraight2Start.Set oStartPos.X + m_dStraight2HorizStart * m_dTextSize, _
                        oStartPos.Y + m_dStraight2VertStart * m_dTextSize, 0
    oStraight2End.Set oStraight2Start.X, _
                      oStartPos.Y + (m_dStraight2VertStart + m_dStraightHeight) * m_dTextSize, 0
    oSlant1Start.Set oStartPos.X, oStartPos.Y + m_dSlant1VertStart * m_dTextSize, 0
    oSlant2Start.Set oStartPos.X, oStartPos.Y + m_dSlant2VertStart * m_dTextSize, 0
    oSlant1End.Set oStraight2Start.X, oStraight2Start.Y + oSlant2Start.Y - oSlant1Start.Y, 0
    oSlant2End.Set oStraight2Start.X, oSlant1End.Y + oSlant2Start.Y - oSlant1Start.Y, 0
    
    SMS_NodeAnnotation oOutputDom, oOutputElem, sAttributeXML, pStartPoint, pOrientation, sPartDir
    
    TranslatePoint oStartPos, pOrientation, pStartPoint
    TranslatePoint oStraight1End, pOrientation, pStartPoint
    TranslatePoint oStraight2Start, pOrientation, pStartPoint
    TranslatePoint oStraight2End, pOrientation, pStartPoint
    TranslatePoint oSlant1Start, pOrientation, pStartPoint
    TranslatePoint oSlant1End, pOrientation, pStartPoint
    TranslatePoint oSlant2Start, pOrientation, pStartPoint
    TranslatePoint oSlant2End, pOrientation, pStartPoint

    
    
    'Create the lines
    
    Set oTempEdgeElement = SMS_NodeEdge(oOutputDom)
    oOutputElem.appendChild oTempEdgeElement
    
    'Added the if condition for removing the line which coincides with the Plate/Profile Location Mark
    If UCase(sPartDir) <> "R" Then
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStartPos, oStraight1End
    End If
    
    If UCase(sPartDir) <> "L" Then
        SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oStraight2Start, oStraight2End
    End If

    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSlant1Start, oSlant1End
    SMS_NodeCurveLine oOutputDom, oTempEdgeElement, oSlant2Start, oSlant2End

    Set oTempCurveElement = Nothing
    Set oTempEdgeElement = Nothing

    Dim sOutputXML As String
    sOutputXML = GetXMLDataAsString(oOutputElem)
    IJDMfgOutputAnnotation_Evaluate = sOutputXML

CleanUp:
    Set oStartPos = Nothing
    Set oStraight1End = Nothing
    Set oStraight2Start = Nothing
    Set oStraight2End = Nothing
    Set oSlant1Start = Nothing
    Set oSlant1End = Nothing
    Set oSlant2Start = Nothing
    Set oSlant2End = Nothing

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


