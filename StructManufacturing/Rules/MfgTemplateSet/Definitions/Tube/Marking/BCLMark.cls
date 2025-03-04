VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCLMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project  : StrMfgTemplateMarkingTube
'
'  Abstract : Create Base control line mark on unfolded tube Template
'
'  History  :
'      Siva   20th October 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingTube.BCLMark"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler
    
    Dim oBottomCurve2D  As IJCurve
    Dim oTopCurve2D     As IJCurve
    Dim oRefCurve       As IJCurve
    
    ' Get the reference objects like 2D bottom curve and top curve and ref curve used for unfold
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
    
    Dim oOrigin As IJDPosition
    Set oOrigin = New DPosition
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    
    oMarkVec.Set 0, -1, 0
    
    ' Create a large line that can intersect 2D bottom curve
    Dim oTempCurve    As IJCurve
    Set oTempCurve = CreateLineAtPosition(oOrigin, oMarkVec, 5)
    
    Dim dMinDist        As Double
    Dim dSrcX As Double, dSrcY As Double, dSrcZ As Double, dInX As Double, dInY As Double, dInZ As Double
    
    ' Get the position of BCL point on the bottom curve
    oTempCurve.DistanceBetween oBottomCurve2D, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    
    Dim oBCL2DPos As IJDPosition
    Set oBCL2DPos = New DPosition
    oBCL2DPos.Set dInX, dInY, dInZ
    
    ' Create BCL mark of fixed length
    Dim oBCLMark    As IJCurve
    Set oBCLMark = CreateLineAtPosition(oBCL2DPos, oMarkVec, TEMPLATE_BCL_MARK_LENGTH, True)
    
    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    ' Create Geom2D for BCL mark
    Dim oMfgGeom2D As IJMfgGeom2d
    Set oMfgGeom2D = CreateGeom2D(oBCLMark, STRMFG_BASELINE_MARK, Nothing, "Base Control Mark")
    
    oTempElems.Add oMfgGeom2D
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oTempElems
    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3016, , "RULES")
End Function
