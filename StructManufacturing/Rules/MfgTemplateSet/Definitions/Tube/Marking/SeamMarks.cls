VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project  : StrMfgTemplateMarkingTube
'
'  Abstract : Creates Seam marks on Template
'
'  History  :
'      Siva   20th October 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingTube.SeamMarks"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler
    
    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    Dim oBottomCurve2D  As IJCurve
    Dim oTopCurve2D     As IJCurve
    Dim oRefCurve       As IJCurve
    
    ' Get the 2D Bottom curve, top curve and ref curve
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
    
    ' Get the reference curve length
    Dim dRefCurveLen    As Double
    dRefCurveLen = oRefCurve.length
    
    ' If there is no extension then exit
    If (oTopCurve2D.length - dRefCurveLen) < 0.001 Then
        Exit Function
    End If
    
    
    Dim oRefPos As IJDPosition
    Set oRefPos = New DPosition
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    
    oMarkVec.Set 0, -1, 0
    
    ' create at 0rigin
    Dim oFittingLine    As IJCurve
    
    ' create at 180 degree from origin
    oRefPos.Set (0.5 * dRefCurveLen), 0, 0
    Set oFittingLine = CreateSeamMarkAtPos(oRefPos, oBottomCurve2D)
    oTempElems.Add oFittingLine
    Set oFittingLine = Nothing
    
    ' create at -180 degree from origin
    oRefPos.Set (-0.5 * dRefCurveLen), 0, 0
    Set oFittingLine = CreateSeamMarkAtPos(oRefPos, oBottomCurve2D)
    oTempElems.Add oFittingLine
    Set oFittingLine = Nothing
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oTempElems
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3021, , "RULES")
End Function

' ***********************************************************************************
' Public Function CreateSeamMarkAtPos()
'
' Description:  Create Seam mark at the input position
'
' ***********************************************************************************
Public Function CreateSeamMarkAtPos(oInputPos As IJDPosition, oBottomCurve As IJCurve) As Object
    Const METHOD = "CreateSeamMarkAtPos"
    On Error GoTo ErrorHandler
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    
    oMarkVec.Set 0, -1, 0
    
    Dim oTempCurve    As IJCurve
    Set oTempCurve = CreateLineAtPosition(oInputPos, oMarkVec, 5)
    
    ' Get the seam mark position on bottom curve
    Dim dMinDist        As Double
    Dim dSrcX As Double, dSrcY As Double, dSrcZ As Double, dInX As Double, dInY As Double, dInZ As Double
    
    oTempCurve.DistanceBetween oBottomCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dInZ
    
    Dim oMarkPos As IJDPosition
    Set oMarkPos = New DPosition
    oMarkPos.Set dInX, dInY, dInZ
    
    ' Get the curve normal
    Dim dTanX As Double, dTanY As Double, dTanZ As Double, dPar As Double, dDummy As Double
    
    oBottomCurve.Parameter dInX, dInY, dInZ, dPar
    oBottomCurve.Evaluate dPar, dDummy, dDummy, dDummy, dTanX, dTanY, dTanZ, dDummy, dDummy, dDummy
    
    Dim oTanVec   As IJDVector
    Set oTanVec = New DVector
    
    oTanVec.Set dTanX, dTanY, dTanZ
    
    Dim oZVec   As IJDVector
    Set oZVec = New DVector
    
    oZVec.Set 0, 0, 1
    
    ' Create mark vector for seam mark
    Set oMarkVec = oTanVec.Cross(oZVec)
            
    Set oTempCurve = Nothing
    Set oTempCurve = CreateLineAtPosition(oMarkPos, oMarkVec, 0.05, True)
    
    ' Create Geom2D for seam mark
    Dim oMfgGeom2D As IJMfgGeom2d
    Set oMfgGeom2D = CreateGeom2D(oTempCurve, STRMFG_SEAM_MARK, Nothing, "Seam Mark")
    
    Set CreateSeamMarkAtPos = oMfgGeom2D
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
