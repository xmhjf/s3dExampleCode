Attribute VB_Name = "AnnotationCommon"
Option Explicit

Private Const MODULE = "MfgCustomAnnotation AnnotationCommon"

Public Const DEGREES_PER_RADIAN = 57.2957795130823
Public Const PI = 3.14159265358979

'Factors used to adjust the textbox location based on Part Monitor offsets
'public Const TEXT_ADJUST_LEFT = 0.15
Public Const TEXT_ADJUST_LEFT = 0#
Public Const TEXT_ADJUST_RIGHT = 0.075
Public Const TEXT_ADJUST_TOP = 0.25
Public Const NOSE_TEXT_ADJUST = 1.2
'Public Const TEXT_ADJUST_BOTTOM = 0.31
Public Const TEXT_ADJUST_BOTTOM = 0.3
Public Const EXTRA_TEXT_ADJUST = 0.3
Public Const MODIFIED_TEXT_ADJUST = 1
Public Const PART_TEXT_ADJUST = 1.1
Public Const ROLL_DIR_ANNOTATION_LENGTH = 200
Public Const SETTING_FILE_PATH = "M:\SharedContent\StructManufacturing\PartMonitor\Annotation\Settings.xml"

'This function creates a line with using CVG_CURVE XML and returns the XML element
Public Function CreateSingleLineCurveNode(XMLDom As DOMDocument, StartPoint As IJDPosition, _
                                          EndPoint As IJDPosition, _
                                          Optional sEdgeType As String = "annotation") _
                                          As IXMLDOMElement
                                           
    If XMLDom Is Nothing Or StartPoint Is Nothing Or EndPoint Is Nothing Then Exit Function
    
    Dim oTempCurveElement       As IXMLDOMNode
    Dim oTempVertexElement      As IXMLDOMNode
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim dAVal As Double
    Dim dBVal As Double
    Dim dCVal As Double
    
    Set oTempEdgeElement = XMLDom.createElement("SMS_EDGE")
    'oTempEdgeElement.setAttribute "TYPE", "bevel_annotation"
    If sEdgeType <> "" Then oTempEdgeElement.setAttribute "TYPE", sEdgeType

    'Create curve element
    Set oTempCurveElement = XMLDom.createElement("CVG_CURVE")
    'Determine the line attributes
    CalcLineAttributesFrom2Pts dAVal, dBVal, dCVal, StartPoint, EndPoint
    'Create the starting point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "s_point", "line", _
                                              StartPoint, dAVal, dBVal, dCVal)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    'Create the end point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "e_point", "dummy", _
                                              EndPoint, 0, 0, 0)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    Set oTempVertexElement = Nothing
    oTempEdgeElement.appendChild oTempCurveElement
    Set CreateSingleLineCurveNode = oTempEdgeElement
    Set oTempCurveElement = Nothing
End Function

'This function creates a arc with using CVG_CURVE XML and returns the XML element
Public Function CreateSingleArcCurveNode(XMLDom As DOMDocument, StartPoint As IJDPosition, _
                                          EndPoint As IJDPosition, CenterPoint As IJDPosition, _
                                          Optional sEdgeType As String = "annotation") _
                                          As IXMLDOMElement
                                           
    If XMLDom Is Nothing Or StartPoint Is Nothing Or EndPoint Is Nothing _
    Or CenterPoint Is Nothing Then Exit Function
    
    Dim oTempCurveElement       As IXMLDOMNode
    Dim oTempVertexElement      As IXMLDOMNode
    Dim oTempEdgeElement        As IXMLDOMElement
    Dim dAVal As Double
    Dim dBVal As Double
    Dim dCVal As Double
    
    Set oTempEdgeElement = XMLDom.createElement("SMS_EDGE")
    'oTempEdgeElement.setAttribute "TYPE", "bevel_annotation"
    If sEdgeType <> "" Then oTempEdgeElement.setAttribute "TYPE", sEdgeType

    'Create curve element
    Set oTempCurveElement = XMLDom.createElement("CVG_CURVE")
    'Determine the arc attributes
    dAVal = CenterPoint.X
    dBVal = CenterPoint.Y
    dCVal = StartPoint.DistPt(CenterPoint)
    If Abs(dCVal - EndPoint.DistPt(CenterPoint)) > 0.000001 Then Exit Function
    'Create the starting point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "s_point", "ccw", _
                                              StartPoint, dAVal, dBVal, dCVal)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    'Create the end point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "e_point", "dummy", _
                                              EndPoint, 0, 0, 0)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    Set oTempVertexElement = Nothing
    oTempEdgeElement.appendChild oTempCurveElement
    Set CreateSingleArcCurveNode = oTempEdgeElement
    Set oTempCurveElement = Nothing
End Function

'This function creates a CVG_TEXT XML element based on the given parameters and returns
'   the element
Public Function CreateTextNode(ByVal XMLDom As DOMDocument, ByVal Just As String, ByVal Location As IJDPosition, _
                               ByVal Text As String, ByVal Font As String, ByVal Orientation As IJDVector, _
                               ByVal Style As String, ByVal Purpose As String, ByVal dTextSize As Double, _
                               ByVal ExtraRotation As Double, Optional ByVal sType As String = "Annotation") _
                                As IXMLDOMElement
                                
    If XMLDom Is Nothing Or Location Is Nothing Or Orientation Is Nothing Then Exit Function
    Dim oCVGTextElem As IXMLDOMElement
    Dim oXUnitVector As IJDVector
    Dim oZUnitVector As IJDVector
    
    'Create reference vectors for determining the rotation angle
    Set oXUnitVector = New DVector
    oXUnitVector.X = 1
    oXUnitVector.Y = 0
    oXUnitVector.Z = 0
    Set oZUnitVector = New DVector
    oZUnitVector.X = 0
    oZUnitVector.Y = 0
    oZUnitVector.Z = 1
    
    Dim oRotMat As New DT4x4
    oRotMat.LoadIdentity
    oRotMat.Rotate ExtraRotation, oZUnitVector

    Dim oSecondaryVec As IJDVector
    Set oSecondaryVec = oRotMat.TransformVector(Orientation)
    
        '*** New Baseline Left Position ***'
    Dim oBaselinePos As IJDPosition
    Set oBaselinePos = dummyForm.GetFontMetricData(Font, dTextSize, Location, oSecondaryVec, Just, Text)
    Just = "bl"
     
    'Create the text element
    Set oCVGTextElem = XMLDom.createElement("CVG_TEXT")
    
'    If Just = "ll" Then
'        oCVGTextElem.setAttribute "TYPE", "bevel_ll"
'    ElseIf Just = "ul" Then
'        oCVGTextElem.setAttribute "TYPE", "bevel_ul"
'    Else
'        oCVGTextElem.setAttribute "TYPE", "annotation"
'    End If
    oCVGTextElem.setAttribute "TYPE", IIf(sType <> "", sType, "label")
    oCVGTextElem.setAttribute "JUST", Just
    oCVGTextElem.setAttribute "LOCX", Trim(Str(Round(oBaselinePos.X, 5)))
    oCVGTextElem.setAttribute "LOCY", Trim(Str(Round(oBaselinePos.Y, 5)))
    oCVGTextElem.setAttribute "TEXT", IIf(Text <> "", Text, "N/A")
    oCVGTextElem.setAttribute "FONT", Font
    oCVGTextElem.setAttribute "FONT_SIZE", Trim(Str(dTextSize))
    oCVGTextElem.setAttribute "FONT_STYLE", Style
    oCVGTextElem.setAttribute "PURPOSE", Purpose
    

    'Dim dAngle As Double
    'dAngle = val(CInt(Round(Orientation.Angle(oXUnitVector, oZUnitVector) * DEGREES_PER_RADIAN, 0)) Mod 360)
    'oCVGTextElem.setAttribute "ROT_ANGLE", Orientation.Angle(oXUnitVector, oZUnitVector) * DEGREES_PER_RADIAN
    oCVGTextElem.setAttribute "ROT_ANGLE", Trim(Str(Round(oXUnitVector.Angle(Orientation, oZUnitVector) + ExtraRotation, 5)))

    Set oXUnitVector = Nothing
    Set oZUnitVector = Nothing
    Set CreateTextNode = oCVGTextElem
    Set oCVGTextElem = Nothing
End Function

'This function creates a CVG_VERTEX XML element based on the given parameters and returns
'   the element
Public Function CreateVertexNode(XMLDom As DOMDocument, PntCode As String, _
                                  SegType As String, PointLoc As IJDPosition, _
                                  AVal As Double, BVal As Double, CVal As Double) _
                                  As IXMLDOMElement
    
    If XMLDom Is Nothing Or PointLoc Is Nothing Then Exit Function
    
    Dim oReturnElem As IXMLDOMElement
    
    Set oReturnElem = XMLDom.createElement("CVG_VERTEX")
    
    oReturnElem.setAttribute "POINT_CODE", PntCode
    oReturnElem.setAttribute "SEG_TYPE", SegType
    oReturnElem.setAttribute "SX", Trim(Str(Round(PointLoc.X, 5)))
    oReturnElem.setAttribute "SY", Trim(Str(Round(PointLoc.Y, 5)))
    oReturnElem.setAttribute "A", Trim(Str(Round(AVal, 5)))
    oReturnElem.setAttribute "B", Trim(Str(Round(BVal, 5)))
    oReturnElem.setAttribute "C", Trim(Str(Round(CVal, 5)))
    
    Set CreateVertexNode = oReturnElem
    Set oReturnElem = Nothing
End Function

'Given an input of 2 (x,y) points, calculate the values "a", "b", and "c" of the equation
'   of the line passing through those points in the form a*x + b*y + c = 0
Public Sub CalcLineAttributesFrom2Pts(ByRef AVal As Double, ByRef BVal As Double, _
                                       ByRef CVal As Double, ByVal Point1 As IJDPosition, _
                                       ByVal Point2 As IJDPosition)
    If Point1 Is Nothing Or Point2 Is Nothing Then Exit Sub
    
    Dim dSlope As Double
    
    'Check for infinite slope
    If Abs(Point1.X - Point2.X) < 0.000001 Then
        AVal = 1
        BVal = 0
        CVal = -Point1.X
        Exit Sub
    End If
    
    BVal = -1
    dSlope = (Point2.Y - Point1.Y) / (Point2.X - Point1.X)
    AVal = dSlope
    CVal = Point1.Y - (dSlope * Point1.X)
End Sub

Public Sub MakePerpindicular(ByRef oPoint As IJDPosition, ByVal oMidPoint As IJDPosition, _
                              ByVal dSymbolHeight As Double, ByVal dDistFromContour As Double)
    
    If oPoint Is Nothing Or oMidPoint Is Nothing Then Exit Sub
    'Make the midpoint the origin
    oPoint.X = oPoint.X - oMidPoint.X
    oPoint.Y = oPoint.Y - oMidPoint.Y
    
    Dim dTempX As Double
    dTempX = oPoint.X
    
    'Rotate the point
    oPoint.X = oPoint.Y
    oPoint.Y = -dTempX
    
    'Move the point out the correct horizontal distance
    oPoint.X = (oPoint.X + (dSymbolHeight / 2)) + dDistFromContour
    oPoint.Y = oPoint.Y + oMidPoint.X
End Sub

Public Sub TranslatePoint(ByRef oPoint As IJDPosition, _
                           ByVal OrientVector As IJDVector, _
                           ByVal oPointOnContour As IJDPosition)
    
    If oPoint Is Nothing Or OrientVector Is Nothing Or oPointOnContour Is Nothing Then Exit Sub
    
    'Make sure the vector is normalized
    OrientVector.length = 1
    
    'First rotate the point
    Dim dTempX As Double
    Dim dTempY As Double
    dTempX = oPoint.X
    dTempY = oPoint.Y
    
    '   x' =  ( x * cos(angle) ) - ( y * sin(angle) )
    oPoint.X = (dTempX * OrientVector.X) - (dTempY * OrientVector.Y)
    '   y' =  ( x * sin(angle) ) + ( y * cos(angle) )
    oPoint.Y = (dTempX * OrientVector.Y) + (dTempY * OrientVector.X)
    
    'Then Translate it
    oPoint.X = oPoint.X + oPointOnContour.X
    oPoint.Y = oPoint.Y + oPointOnContour.Y
End Sub

'Obtain all of the Bevel angle and thickness values from the XML string
Public Sub FillBevelValuesFromXML(ByVal sBevelXml As String, ByRef strGUID, ByRef dChamferAngleM As Double, _
                              ByRef dChamferDepthM As Double, ByRef dChamferAngleUM As Double, _
                              ByRef dChamferDepthUM As Double, ByRef dAngleA As Double, _
                              ByRef dAngleB As Double, ByRef dAngleN As Double, _
                              ByRef dAngleD As Double, ByRef dAngleE As Double, _
                              ByRef dDepthA As Double, ByRef dDepthB As Double, _
                              ByRef dDepthN As Double, ByRef dDepthD As Double, _
                              ByRef dDepthE As Double)
    Const METHOD = "FillBevelValuesFromXML"
    On Error GoTo ErrorHandler
    
    Dim oBevelDom As New DOMDocument
    'Dim oNodeList As IXMLDOMNodeList
    Dim oBevelElem As IXMLDOMElement
    Dim sValue As String
    Dim vTemp As Variant
        
    strGUID = ""
    dChamferAngleM = 0
    dChamferDepthM = 0
    dChamferAngleUM = 0
    dChamferDepthUM = 0
    dAngleA = 0
    dAngleB = 0
    dAngleN = 0
    dAngleD = 0
    dAngleE = 0
    dDepthA = 0
    dDepthB = 0
    dDepthN = 0
    dDepthD = 0
    dDepthE = 0
    
    If Not oBevelDom.loadXML(sBevelXml) Then GoTo CleanUp
    
    'Assumes the first node found is the desired node
    Set oBevelElem = oBevelDom.selectSingleNode("//SMS_BEVEL")
    
    If oBevelElem Is Nothing Then GoTo CleanUp
    
    vTemp = oBevelElem.getAttribute("GEOM2D_GUID")
    strGUID = IIf(VarType(vTemp) = vbString, vTemp, "")
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("CHAMFER_ANGLE_M")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dChamferAngleM = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("CHAMFER_ANGLE_UM")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dChamferAngleUM = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("CHAMFER_DEPTH_M")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dChamferDepthM = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("CHAMFER_DEPTH_UM")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dChamferDepthUM = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("ANGLE_A")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dAngleA = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("ANGLE_B")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dAngleB = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("ANGLE_N")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dAngleN = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("ANGLE_D")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dAngleD = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("ANGLE_E")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dAngleE = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("DEPTH_A")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dDepthA = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("DEPTH_B")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dDepthB = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("DEPTH_N")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dDepthN = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("DEPTH_D")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dDepthD = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.getAttribute("DEPTH_E")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If IsNumeric(sValue) Then
        dDepthE = Val(sValue)
    End If
CleanUp:
    Set oBevelDom = Nothing
    Set oBevelElem = Nothing
    'Set oNodeList = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Sub

'Obtain all of the Grind angle and thickness values from the XML string
Public Sub FillGrindValuesFromXML(ByVal sGrindXml As String, ByRef strTYPE As String, _
                              ByRef strGUID As String, ByRef dRadiusM As Double, _
                              ByRef dRadiusUM As Double, ByRef dChamferAngleM As Double, _
                              ByRef dChamferDepthM As Double, ByRef dChamferAngleUM As Double, _
                              ByRef dChamferDepthUM As Double, ByRef dAngleA As Double, _
                              ByRef dAngleB As Double, ByRef dAngleN As Double, _
                              ByRef dAngleD As Double, ByRef dAngleE As Double, _
                              ByRef dDepthA As Double, ByRef dDepthB As Double, _
                              ByRef dDepthN As Double, ByRef dDepthD As Double, _
                              ByRef dDepthE As Double)
    Const METHOD = "FillGrindValuesFromXML"
    On Error GoTo ErrorHandler
    
    Dim oGrindDom As New DOMDocument
    'Dim oNodeList As IXMLDOMNodeList
    Dim oGrindElem As IXMLDOMElement
    Dim sValue As String
    Dim vTemp As Variant
        
    strGUID = ""
    dChamferAngleM = 0
    dChamferDepthM = 0
    dChamferAngleUM = 0
    dChamferDepthUM = 0
    dAngleA = 0
    dAngleB = 0
    dAngleN = 0
    dAngleD = 0
    dAngleE = 0
    dDepthA = 0
    dDepthB = 0
    dDepthN = 0
    dDepthD = 0
    dDepthE = 0
    
    If Not oGrindDom.loadXML(sGrindXml) Then GoTo CleanUp
    
    'Assumes the first node found is the desired node
    Set oGrindElem = oGrindDom.selectSingleNode("//SMS_GRIND")
    
    If oGrindElem Is Nothing Then GoTo CleanUp
    
    vTemp = oGrindElem.getAttribute("TYPE")
    strTYPE = IIf(VarType(vTemp) = vbString, vTemp, "")
    
    vTemp = oGrindElem.getAttribute("GEOM2D_GUID")
    strGUID = IIf(VarType(vTemp) = vbString, vTemp, "")
    
    Select Case strTYPE
        Case "grind_radius"
            sValue = 0
            vTemp = oGrindElem.getAttribute("RADIUS_M")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dRadiusM = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("RADIUS_UM")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dRadiusUM = Val(sValue)
            End If
        Case "grind_flat"
            sValue = 0
            vTemp = oGrindElem.getAttribute("CHAMFER_ANGLE_M")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dChamferAngleM = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("CHAMFER_ANGLE_UM")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dChamferAngleUM = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("CHAMFER_DEPTH_M")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dChamferDepthM = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("CHAMFER_DEPTH_UM")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dChamferDepthUM = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("ANGLE_A")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dAngleA = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("ANGLE_B")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dAngleB = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("ANGLE_N")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dAngleN = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("ANGLE_D")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dAngleD = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("ANGLE_E")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dAngleE = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("DEPTH_A")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dDepthA = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("DEPTH_B")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dDepthB = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("DEPTH_N")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dDepthN = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("DEPTH_D")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dDepthD = Val(sValue)
            End If
            
            sValue = 0
            vTemp = oGrindElem.getAttribute("DEPTH_E")
            sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
            If IsNumeric(sValue) Then
                dDepthE = Val(sValue)
            End If
        Case Else
            GoTo CleanUp
    End Select
CleanUp:
    Set oGrindDom = Nothing
    Set oGrindElem = Nothing
    'Set oNodeList = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Sub

'Obtain all of the Plate Thickness Direction values from the XML string
Public Sub FillPlateThickValuesFromXML(ByVal sMarkingXml As String, ByRef strGUID, ByRef sPartDir As String)
    Const METHOD = "FillPlateThickValuesFromXML"
    On Error GoTo ErrorHandler
    
    Dim oPlateMarkDom As New DOMDocument
    Dim oPlateMarkElem As IXMLDOMAttribute
    Dim sValue As String
    Dim vTemp As Variant
        
    sPartDir = ""
    
    If Not oPlateMarkDom.loadXML(sMarkingXml) Then GoTo CleanUp
    
    Set oPlateMarkElem = oPlateMarkDom.selectSingleNode(".//@GEOM2D_GUID")
    
    If Not oPlateMarkElem Is Nothing Then
        vTemp = oPlateMarkElem.Value
        strGUID = IIf(VarType(vTemp) = vbString, vTemp, "")
    End If
    
    Set oPlateMarkElem = Nothing
    
    Set oPlateMarkElem = oPlateMarkDom.selectSingleNode(".//@PART_DIR")
       
    If oPlateMarkElem Is Nothing Then GoTo CleanUp
    
    vTemp = oPlateMarkElem.Value
    sPartDir = IIf(VarType(vTemp) = vbString, vTemp, "")
CleanUp:
    Set oPlateMarkDom = Nothing
    Set oPlateMarkElem = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Sub

'Obtain all of the flange values from the XML string
Public Sub FillFlangeValuesFromXML(ByVal sMarkingXml As String, ByRef strGUID, ByRef sXSection As String, _
                                   ByRef sThicknessDir As String, ByRef sFlangeDir As String)
    Const METHOD = "FillFlangeValuesFromXML"
    On Error GoTo ErrorHandler
    
    Dim oFlangeDom As New DOMDocument
    Dim oNodeList As IXMLDOMNodeList
    Dim oFlangeElem As IXMLDOMElement
    Dim sValue As String
    Dim vTemp As Variant
        
    sXSection = ""
    sThicknessDir = ""
    sFlangeDir = ""
    
    If Not oFlangeDom.loadXML(sMarkingXml) Then GoTo CleanUp
    
    Set oNodeList = oFlangeDom.getElementsByTagName("SMS_MARKING")
    
    If oNodeList Is Nothing Then GoTo CleanUp
    If oNodeList.length = 0 Then GoTo CleanUp
    
    oNodeList.Reset
    'Assumes the first node is the desired node
    Set oFlangeElem = oNodeList.nextNode
    If oFlangeElem Is Nothing Then GoTo CleanUp
    
    vTemp = oFlangeElem.getAttribute("GEOM2D_GUID")
    strGUID = IIf(VarType(vTemp) = vbString, vTemp, "")
    
    Set oNodeList = Nothing
    Set oFlangeElem = Nothing
    
    Set oNodeList = oFlangeDom.getElementsByTagName("SMS_MARKING_PART_LOCATION")
    
    If oNodeList Is Nothing Then Exit Sub
    
    oNodeList.Reset
    Set oFlangeElem = oNodeList.nextNode
    
    If oFlangeElem Is Nothing Then Exit Sub
    
    vTemp = oFlangeElem.getAttribute("XSECTION_NAME")
    sXSection = IIf(VarType(vTemp) = vbString, vTemp, "")
    vTemp = oFlangeElem.getAttribute("THICKNESS_DIR")
    sThicknessDir = IIf(VarType(vTemp) = vbString, vTemp, "")
    vTemp = oFlangeElem.getAttribute("FLANGE_DIR")
    sFlangeDir = IIf(VarType(vTemp) = vbString, vTemp, "")
CleanUp:
    Set oFlangeDom = Nothing
    Set oFlangeElem = Nothing
    Set oNodeList = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Sub


'Obtain all of the Margin values from the XML string
Public Sub FillDefaultMarginValuesFromXML(ByVal sMarginXML As String, ByRef strGUID, ByRef dStartVal As Double, _
                                   ByRef dEndVal As Double, ByRef dX1 As Double, ByRef dY1 As Double, _
                                   ByRef dX2 As Double, ByRef dY2 As Double)
    Const METHOD = "FillDefaultMarginValuesFromXML"
    On Error GoTo ErrorHandler
    
    Dim oMarginDom As New DOMDocument
    Dim oMarginElem As IXMLDOMElement
    Dim oNodeList As IXMLDOMNodeList
    Dim sValue As String
    Dim vTemp As Variant
        
    strGUID = ""
    dStartVal = 0
    dEndVal = 0
    dX1 = 0
    dY1 = 0
    dX2 = 0
    dY2 = 0
    
    
    If Not oMarginDom.loadXML(sMarginXML) Then GoTo CleanUp
    
    Set oMarginElem = oMarginDom.selectSingleNode("//SMS_EDGE")
    
    If oMarginElem Is Nothing Then GoTo CleanUp
    
    vTemp = oMarginElem.getAttribute("GEOM2D_GUID")
    strGUID = IIf(VarType(vTemp) = vbString, vTemp, "")
    
    Set oMarginElem = oMarginDom.selectSingleNode("//SMS_PART_MARGIN_INFO")
    vTemp = oMarginElem.getAttribute("START_VALUE")
    If IsNumeric(vTemp) Then dStartVal = Val(vTemp)
    
    vTemp = oMarginElem.getAttribute("END_VALUE")
    If IsNumeric(vTemp) Then dEndVal = Val(vTemp)
    
    Set oNodeList = oMarginElem.selectNodes("./CVG_POINT")
    
    If Not oNodeList Is Nothing Then
        If oNodeList.length > 0 Then
            Set oMarginElem = oNodeList(0)
            vTemp = oMarginElem.getAttribute("X")
            If IsNumeric(vTemp) Then dX1 = Val(vTemp)
            vTemp = oMarginElem.getAttribute("Y")
            If IsNumeric(vTemp) Then dY1 = Val(vTemp)
        End If
        If oNodeList.length > 1 Then
            Set oMarginElem = oNodeList(1)
            vTemp = oMarginElem.getAttribute("X")
            If IsNumeric(vTemp) Then dX2 = Val(vTemp)
            vTemp = oMarginElem.getAttribute("Y")
            If IsNumeric(vTemp) Then dY2 = Val(vTemp)
        End If
    End If
    
CleanUp:
    Set oMarginDom = Nothing
    Set oMarginElem = Nothing
    Set oNodeList = Nothing
    Exit Sub
ErrorHandler:
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Sub

'Obtain all of the Margin values from the XML string
Public Sub FillMargin2ValuesFromXML(ByVal sMarginXML As String, ByRef strGUID, ByRef sStage As String, _
                                   ByRef dLength As Double)
    Const METHOD = "FillMargin2ValuesFromXML"
    On Error GoTo ErrorHandler
    
    Dim oMarginDom As New DOMDocument
    Dim oMarginElem As IXMLDOMElement
    Dim sValue As String
    Dim vTemp As Variant
        
    strGUID = ""
    sStage = ""
    dLength = 0
    
    If Not oMarginDom.loadXML(sMarginXML) Then GoTo CleanUp
    
    Set oMarginElem = oMarginDom.selectSingleNode("//SMS_MARKING")
    
    If oMarginElem Is Nothing Then GoTo CleanUp
    
    vTemp = oMarginElem.getAttribute("GEOM2D_GUID")
    strGUID = IIf(VarType(vTemp) = vbString, vTemp, "")
    vTemp = oMarginElem.getAttribute("STAGE")
    sStage = IIf(VarType(vTemp) = vbString, vTemp, "")
    vTemp = oMarginElem.getAttribute("LENGTH")
    If IsNumeric(vTemp) Then dLength = Val(vTemp)
CleanUp:
    Set oMarginDom = Nothing
    Set oMarginElem = Nothing
    Exit Sub
ErrorHandler:
'MsgBox MODULE & ", " & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & ": " & METHOD
    Resume Next
End Sub

Public Function GetDirectionValue(sDirection As String) As String
    Select Case UCase(sDirection)
        
    Case "INBOARD", "IN"
        GetDirectionValue = "I"
    Case "OUTBOARD", "OUT"
        GetDirectionValue = "O"
    Case "UP", "UPPER"
        GetDirectionValue = "U"
    Case "DOWN", "LOWER"
        GetDirectionValue = "D"
    Case "FORWARD", "FORE"
        GetDirectionValue = "F"
    Case "AFT"
        GetDirectionValue = "A"
    Case "PORT"
        GetDirectionValue = "P"
    Case "STARBOARD"
        GetDirectionValue = "S"
    Case Else
        GetDirectionValue = sDirection
    End Select
End Function

Public Function GetAttributeValueFromXML(sXML As String, sAttribute As String) As String
    Dim oXMLDom As New DOMDocument
    Dim oXMLElement As IXMLDOMElement
    Dim sAttributeValue As String
    
    If Not oXMLDom.loadXML(sXML) Then GoTo CleanUp
    
    Set oXMLElement = oXMLDom.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||" & sAttribute & "']")
    If oXMLElement Is Nothing Then
        Set oXMLElement = oXMLDom.selectSingleNode(".//SMS_GEOM_ARG[@NAME='" & sAttribute & "']")
    End If
    
    If Not oXMLElement Is Nothing Then
        sAttributeValue = Trim(oXMLElement.getAttribute("VALUE"))
    End If
    'If it's an empty string, then just return an empty string.
    GetAttributeValueFromXML = sAttributeValue
CleanUp:
    Set oXMLDom = Nothing
    Set oXMLElement = Nothing
End Function

Public Function GetAttributeDblValueFromXML(sXML As String, sAttribute As String) As Double
    Dim strAttribVal As String
    strAttribVal = GetAttributeValueFromXML(sXML, sAttribute)
    If strAttribVal <> "" And IsNumeric(strAttribVal) = True Then
        GetAttributeDblValueFromXML = Val(strAttribVal)
    Else
        GetAttributeDblValueFromXML = 0
    End If
End Function

Public Function SMS_NodeEdge(XMLDom As DOMDocument, Optional sEdgeType As String = "annotation") As IXMLDOMElement
                                           
    If XMLDom Is Nothing Then Exit Function
    
    Dim oTempEdgeElement        As IXMLDOMElement
    Set oTempEdgeElement = XMLDom.createElement("SMS_EDGE")
    'oTempEdgeElement.setAttribute "TYPE", "bevel_annotation"
    If sEdgeType <> "" Then oTempEdgeElement.setAttribute "TYPE", sEdgeType
    
    Set SMS_NodeEdge = oTempEdgeElement
    
End Function


Public Sub SMS_NodeAnnotation(oOutputDom As DOMDocument, oOutputElem As IXMLDOMElement, ByVal sAttributeXML As String, _
                                    pStartPoint As IJDPosition, pOrientation As IJDVector, strPART_DIR As String)
                                    

    On Error GoTo ErrorHandler
    
    Dim oXUnitVector    As IJDVector
    Dim oZUnitVector    As IJDVector
    Dim oPlateMarkDom   As New DOMDocument
    Dim oElement        As IXMLDOMElement
    Dim strReference    As String
    Dim strFaceType     As String
    
    'Create reference vectors for determining the rotation angle
    Set oXUnitVector = New DVector
    oXUnitVector.X = 1
    oXUnitVector.Y = 0
    oXUnitVector.Z = 0
    Set oZUnitVector = New DVector
    oZUnitVector.X = 0
    oZUnitVector.Y = 0
    oZUnitVector.Z = 1
    
    oPlateMarkDom.loadXML sAttributeXML
    Set oElement = oPlateMarkDom.selectSingleNode("SMS_OUTPUT_ANNOTATION")
    
    Set oOutputElem = oOutputDom.createElement("SMS_ANNOTATION")
    oOutputElem.setAttribute "TYPE", oElement.getAttribute("TYPE")
    oOutputElem.setAttribute "MARKED_SIDE", "marking"
    oOutputElem.setAttribute "GEOM2D_GUID", GetAttributeValueFromXML(sAttributeXML, "GEOM2D_GUID")
'    If strPART_DIR <> vbNullString Then
'        oOutputElem.setAttribute "PART_DIR", strPART_DIR
'    End If
    strReference = GetAttributeValueFromXML(sAttributeXML, "REFERENCE")
    If strReference <> vbNullString Then
        oOutputElem.setAttribute "REFERENCE", strReference
    End If
    strFaceType = GetAttributeValueFromXML(sAttributeXML, "FACE_TYPE")
    If strFaceType <> vbNullString Then
        oOutputElem.setAttribute "FACE_TYPE", strFaceType
    End If
    oOutputElem.setAttribute "TEXT_SIZE", Trim(Str(Val(GetAttributeValueFromXML(sAttributeXML, "TextSize"))))
    oOutputElem.setAttribute "SX", Trim(Str(Round(pStartPoint.X, 5)))
    oOutputElem.setAttribute "SY", Trim(Str(Round(pStartPoint.Y, 5)))
    oOutputElem.setAttribute "ROT_ANGLE", Trim(Str(oXUnitVector.Angle(pOrientation, oZUnitVector)))

CleanUp:

ErrorHandler:
    
End Sub


'This function creates a line with using CVG_CURVE XML and returns the XML element
Public Sub SMS_NodeCurveLine(XMLDom As DOMDocument, oEdgeElement As IXMLDOMElement, StartPoint As IJDPosition, _
                                          EndPoint As IJDPosition)
                                           
    If XMLDom Is Nothing Or StartPoint Is Nothing Or EndPoint Is Nothing Or oEdgeElement Is Nothing Then Exit Sub
    
    Dim oTempCurveElement       As IXMLDOMElement
    Dim oTempVertexElement      As IXMLDOMNode
    Dim dAVal As Double
    Dim dBVal As Double
    Dim dCVal As Double
    
    'Create curve element
    Set oTempCurveElement = XMLDom.createElement("CVG_CURVE")
    oTempCurveElement.setAttribute "CURVE_TYPE", "line"
    'Determine the line attributes
    CalcLineAttributesFrom2Pts dAVal, dBVal, dCVal, StartPoint, EndPoint
    'Create the starting point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "s_point", "line", _
                                              StartPoint, dAVal, dBVal, dCVal)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    'Create the end point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "e_point", "dummy", _
                                              EndPoint, 0, 0, 0)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    Set oTempVertexElement = Nothing
    oEdgeElement.appendChild oTempCurveElement
    Set oTempCurveElement = Nothing
   
End Sub


'This sub creates a arc with using CVG_CURVE XML and returns the XML element
Public Sub SMS_NodeCurveArc(XMLDom As DOMDocument, oEdgeElement As IXMLDOMElement, StartPoint As IJDPosition, _
                                          EndPoint As IJDPosition, CenterPoint As IJDPosition)
                                           
    If XMLDom Is Nothing Or StartPoint Is Nothing Or EndPoint Is Nothing Or oEdgeElement Is Nothing _
    Or CenterPoint Is Nothing Then Exit Sub
    
    Dim oTempCurveElement       As IXMLDOMElement
    Dim oTempVertexElement      As IXMLDOMNode
    Dim dAVal As Double
    Dim dBVal As Double
    Dim dCVal As Double
    
    'Create curve element
    Set oTempCurveElement = XMLDom.createElement("CVG_CURVE")
    oTempCurveElement.setAttribute "CURVE_TYPE", "arc"
    'Determine the arc attributes
    dAVal = CenterPoint.X
    dBVal = CenterPoint.Y
    dCVal = StartPoint.DistPt(CenterPoint)
    If Abs(dCVal - EndPoint.DistPt(CenterPoint)) > 0.000001 Then Exit Sub
    'Create the starting point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "s_point", "ccw", _
                                              StartPoint, dAVal, dBVal, dCVal)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    'Create the end point
    Set oTempVertexElement = CreateVertexNode(XMLDom, "e_point", "dummy", _
                                              EndPoint, 0, 0, 0)
    If Not oTempVertexElement Is Nothing Then oTempCurveElement.appendChild oTempVertexElement
    Set oTempVertexElement = Nothing
    oEdgeElement.appendChild oTempCurveElement
    Set oTempCurveElement = Nothing

End Sub

'This function creates a CVG_TEXT XML element based on the given parameters and returns
'   the element
Public Sub SMS_NodeText(ByVal XMLDom As DOMDocument, oAnnotationElement As IXMLDOMElement, ByVal Just As String, ByVal Location As IJDPosition, _
                               ByVal Text As String, ByVal Font As String, ByVal Orientation As IJDVector, _
                               ByVal Style As String, ByVal Purpose As String, ByVal dTextSize As Double, _
                               ByVal ExtraRotation As Double, Optional ByVal sType As String = "Annotation")

                                
    If XMLDom Is Nothing Or Location Is Nothing Or Orientation Is Nothing Or _
    oAnnotationElement Is Nothing Then Exit Sub
    
    Dim oCVGTextElem As IXMLDOMElement
    Dim oXUnitVector As IJDVector
    Dim oZUnitVector As IJDVector
    
    'Create reference vectors for determining the rotation angle
    Set oXUnitVector = New DVector
    oXUnitVector.X = 1
    oXUnitVector.Y = 0
    oXUnitVector.Z = 0
    Set oZUnitVector = New DVector
    oZUnitVector.X = 0
    oZUnitVector.Y = 0
    oZUnitVector.Z = 1
    
    Dim oRotMat As New DT4x4
    oRotMat.LoadIdentity
    oRotMat.Rotate ExtraRotation, oZUnitVector

    Dim oSecondaryVec As IJDVector
    Set oSecondaryVec = oRotMat.TransformVector(Orientation)
    
        '*** New Baseline Left Position ***'
    Dim oBaselinePos As IJDPosition
    Set oBaselinePos = dummyForm.GetFontMetricData(Font, dTextSize, Location, oSecondaryVec, Just, Text)
    Just = "bl"
     
    'Create the text element
    Set oCVGTextElem = XMLDom.createElement("CVG_TEXT")
    
'    If Just = "ll" Then
'        oCVGTextElem.setAttribute "TYPE", "bevel_ll"
'    ElseIf Just = "ul" Then
'        oCVGTextElem.setAttribute "TYPE", "bevel_ul"
'    Else
'        oCVGTextElem.setAttribute "TYPE", "annotation"
'    End If
    oCVGTextElem.setAttribute "TYPE", IIf(sType <> "", sType, "label")
    oCVGTextElem.setAttribute "JUST", Just
    oCVGTextElem.setAttribute "LOCX", Trim(Str(Round(oBaselinePos.X, 5)))
    oCVGTextElem.setAttribute "LOCY", Trim(Str(Round(oBaselinePos.Y, 5)))
    oCVGTextElem.setAttribute "TEXT", IIf(Text <> "", Text, "N/A")
    oCVGTextElem.setAttribute "FONT", Font
    oCVGTextElem.setAttribute "FONT_SIZE", Trim(Str(dTextSize))
    oCVGTextElem.setAttribute "FONT_STYLE", Style
    oCVGTextElem.setAttribute "PURPOSE", Purpose
    

    'Dim dAngle As Double
    'dAngle = val(CInt(Round(Orientation.Angle(oXUnitVector, oZUnitVector) * DEGREES_PER_RADIAN, 0)) Mod 360)
    'oCVGTextElem.setAttribute "ROT_ANGLE", Orientation.Angle(oXUnitVector, oZUnitVector) * DEGREES_PER_RADIAN
    oCVGTextElem.setAttribute "ROT_ANGLE", Trim(Str(Round(oXUnitVector.Angle(Orientation, oZUnitVector) + ExtraRotation, 5)))
    
    oAnnotationElement.appendChild oCVGTextElem

    Set oXUnitVector = Nothing
    Set oZUnitVector = Nothing
    Set oCVGTextElem = Nothing
            
End Sub


Public Function GetXMLDataAsString(ByVal oOutputElem As IXMLDOMElement) As String
    
    If oOutputElem Is Nothing Then Exit Function
    
    Dim sOutputXML As String
    If oOutputElem.childNodes.length > 0 Then
        sOutputXML = Replace(oOutputElem.xml, "><", ">" & vbNewLine & "<")
    Else
        sOutputXML = ""
    End If
    
    GetXMLDataAsString = sOutputXML
    
    Exit Function
End Function


