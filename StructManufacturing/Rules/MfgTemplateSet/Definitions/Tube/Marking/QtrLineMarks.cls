VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QtrLineMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project  : StrMfgTemplateMarkingTube
'
'  Abstract : Creates Quarter Line Marks on unfolded tube Template
'
'  History  :
'      Siva   20th October 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingTube.QtrLineMarks"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler
    
    Dim oRuleHelper     As MfgRuleHelpers.Helper
    Set oRuleHelper = New MfgRuleHelpers.Helper
    
    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    Dim oBottomCurve2D  As IJCurve
    Dim oTopCurve2D     As IJCurve
    Dim oRefCurve       As IJCurve
        
    ' Get the bottom line, topline and ref curve used for unfold from the input reference collection
    Dim iIndex  As Long
    For iIndex = 1 To oReferenceObjColl.Count
        Dim oTempObj    As Object
        Set oTempObj = oReferenceObjColl.Item(iIndex)
        
        If TypeOf oTempObj Is IJMfgGeom2d Then
            Dim oGeom2D As IJMfgGeom2d
            Set oGeom2D = oTempObj
            If oGeom2D.GetGeometryType = STRMFG_TemplateLocationMarkLine Then
                Set oBottomCurve2D = oGeom2D.GetGeometry
            ElseIf oGeom2D.GetGeometryType = STRMFG_TopLine Then
                Set oTopCurve2D = oGeom2D.GetGeometry
            End If
        Else
            If TypeOf oTempObj Is IJCurve Then
                Set oRefCurve = oReferenceObjColl.Item(iIndex)
            End If
        End If
    Next
    
    ' Get the ref curve length
    Dim dRefCurveLen    As Double
    dRefCurveLen = oRefCurve.length
    
    ' Convert the ref curve to wire body
    Dim oRefCurveWB  As IJWireBody
    Set oRefCurveWB = oRuleHelper.ComplexStringToWireBody(oRefCurve)
    
    Dim oMfgGeomChild   As IJMfgGeomChild
    Set oMfgGeomChild = oMfgTemplate
    
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Set oMfgTemplateSet = oMfgGeomChild.GetParent
    
    ' Get the template 3D BCL curve
    Dim oBaseCtrlCurve As IJCurve
    Set oBaseCtrlCurve = oMfgTemplateSet.GetControlLine
    
    ' Get the centroid of the ref curve
    Dim dCenterX As Double, dCenterY As Double, dCenterZ As Double
    oRefCurve.Centroid dCenterX, dCenterY, dCenterZ
    
    Dim oCenPos As IJDPosition
    Set oCenPos = New DPosition
    
    oCenPos.Set dCenterX, dCenterY, dCenterZ
    
    ' Get the BCL position on ref curve
    Dim dMinDist As Double, dPar As Double
    Dim dSrcX As Double, dSrcY As Double, dSrcZ As Double, dInX As Double, dInY As Double, dInZ As Double
    Dim dDummy As Double, dTan2X As Double, dTan2Y As Double, dTan2Z As Double
    
    oBaseCtrlCurve.DistanceBetween oRefCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    
    Dim oRefBCLPos  As IJDPosition
    Set oRefBCLPos = New DPosition
    
    oRefBCLPos.Set dInX, dInY, dInZ
    
    oRefCurve.Parameter oRefBCLPos.x, oRefBCLPos.y, oRefBCLPos.z, dPar
    oRefCurve.Evaluate dPar, dDummy, dDummy, dDummy, dDummy, dDummy, dDummy, dTan2X, dTan2Y, dTan2Z
    
    Dim oCurveNormal    As IJDVector
    Set oCurveNormal = New DVector
    
    oCurveNormal.Set dTan2X, dTan2Y, dTan2Z
    oCurveNormal.length = 1
    
    Dim oCenVec As IJDVector
    Set oCenVec = oCenPos.Subtract(oRefBCLPos)
    oCenVec.length = 1
    
    If oCenVec.Dot(oCurveNormal) < 0 Then
        oCurveNormal.length = -1
    End If
    
    Dim strDirectionRef As String
    strDirectionRef = GetDirectionString(oCurveNormal, oRefBCLPos)
    
    ' Get the right pos direction (at 90 degree offset from RefBCL position)
    Dim oRightPos    As IJDPosition
    Set oRightPos = oRuleHelper.GetPointAlongCurveAtDistance(oRefCurveWB, oRefBCLPos, 0.25 * dRefCurveLen, oRefBCLPos)
    
    oRefCurve.Parameter oRightPos.x, oRightPos.y, oRightPos.z, dPar
    oRefCurve.Evaluate dPar, dDummy, dDummy, dDummy, dDummy, dDummy, dDummy, dTan2X, dTan2Y, dTan2Z
    
    oCurveNormal.Set dTan2X, dTan2Y, dTan2Z
    oCurveNormal.length = 1
    
    Set oCenVec = oCenPos.Subtract(oRightPos)
    oCenVec.length = 1
    
    If oCenVec.Dot(oCurveNormal) < 0 Then
        oCurveNormal.length = -1
    End If
    
    Dim strDirectionRight As String
    strDirectionRight = GetDirectionString(oCurveNormal, oRightPos)
    
    ' Get the other position directions on reference curve
    Dim strDirectionLeft As String
    strDirectionLeft = GetOppositeDirectionString(strDirectionRight)
    
    Dim strDirectionRefOpp As String
    strDirectionRefOpp = GetOppositeDirectionString(strDirectionRef)
    
    
    Dim oRefPos As IJDPosition
    Set oRefPos = New DPosition
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    oMarkVec.Set 0, -1, 0
    
    ' create Quarter curve at 0rigin
    Dim oQtrLine    As IJCurve
    Set oQtrLine = CreateTrimmedQtrLineAtPos(oRefPos, oBottomCurve2D, oTopCurve2D, strDirectionRef)
    oTempElems.Add oQtrLine
    Set oQtrLine = Nothing
    
    ' create Quarter curve at 90 degree from origin
    oRefPos.Set (0.25 * dRefCurveLen), 0, 0
    Set oQtrLine = CreateTrimmedQtrLineAtPos(oRefPos, oBottomCurve2D, oTopCurve2D, strDirectionLeft)
    oTempElems.Add oQtrLine
    Set oQtrLine = Nothing
    
    ' create Quarter curve at -90 degree from origin
    oRefPos.Set (-0.25 * dRefCurveLen), 0, 0
    Set oQtrLine = CreateTrimmedQtrLineAtPos(oRefPos, oBottomCurve2D, oTopCurve2D, strDirectionRight)
    oTempElems.Add oQtrLine
    Set oQtrLine = Nothing
    
    ' create Quarter curve at 180 degree from origin
    oRefPos.Set (0.5 * dRefCurveLen), 0, 0
    Set oQtrLine = CreateTrimmedQtrLineAtPos(oRefPos, oBottomCurve2D, oTopCurve2D, strDirectionRefOpp)
    oTempElems.Add oQtrLine
    Set oQtrLine = Nothing
    
    ' create Quarter curve at -180 degree from origin
    oRefPos.Set (-0.5 * dRefCurveLen), 0, 0
    Set oQtrLine = CreateTrimmedQtrLineAtPos(oRefPos, oBottomCurve2D, oTopCurve2D, strDirectionRefOpp)
    oTempElems.Add oQtrLine
    Set oQtrLine = Nothing
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oTempElems
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3020, , "RULES")
End Function

' ***********************************************************************************
' Public Function CreateTrimmedQtrLineAtPos()
'
' Description:  Create quarter lines trimmed by the boundaries(i.e., 2D bottom curve and top curve)
'
' ***********************************************************************************
Public Function CreateTrimmedQtrLineAtPos(oInputPos As IJDPosition, oBottomCurve As IJComplexString, oTopCurve As IJComplexString, strDirection As String) As Object
    Const METHOD = "CreateTrimmedQtrLineAtPos"
    On Error GoTo ErrorHandler
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    oMarkVec.Set 0, -1, 0
    
    ' Create a large line that can intersect the 2D top curve and bottom curve
    Dim oQtrMark    As IJCurve
    Set oQtrMark = CreateLineAtPosition(oInputPos, oMarkVec, 5)
    
    Dim dMinDist        As Double
    Dim dSrcX As Double, dSrcY As Double, dSrcZ As Double, dInX As Double, dInY As Double, dInZ As Double
    
    ' Get the intersection point on top curve
    oQtrMark.DistanceBetween oTopCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    
    Dim oTrimPos1 As IJDPosition
    Set oTrimPos1 = New DPosition
    oTrimPos1.Set dSrcX, dSrcY, dSrcZ
    
    ' Get the intersection point on bottom curve
    oQtrMark.DistanceBetween oBottomCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    
    Dim oTrimPos2 As IJDPosition
    Set oTrimPos2 = New DPosition
    oTrimPos2.Set dSrcX, dSrcY, dSrcZ
    
    ' Trim the curve by points on top and bottom curves
    Dim oMfgMGHelper    As GSCADMathGeom.MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    oMfgMGHelper.TrimCurveByPoints oQtrMark, oTrimPos1, oTrimPos2
    
    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    ' Create Geom2D for the mark
    Dim oMfgGeom2D As IJMfgGeom2d
    Set oMfgGeom2D = CreateGeom2D(oQtrMark, STRMFG_REF_MARK, Nothing, strDirection)
    
    Set CreateTrimmedQtrLineAtPos = oMfgGeom2D
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


