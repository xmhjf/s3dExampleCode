VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project  : StrMfgTemplateMarkingEdge
'
'  Abstract : Create custom marks
'
'  History  :
'      Siva   20th April 2012    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingEdge.CustomMark"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
    Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
    On Error GoTo ErrorHandler
    
    ' Below implementation is done for thickness direction annotation
    
    Dim oUVLine  As IJLine
    Dim oBottomCurve  As IJCurve
    
    ' Get the reference objects like 2D bottom curve and top curve and ref curve used for unfold
    Dim iIndex  As Long
    For iIndex = 1 To oReferenceObjColl.Count
        Dim oTempObj    As Object
        Set oTempObj = oReferenceObjColl.Item(iIndex)
        
        If TypeOf oTempObj Is IJMfgGeom2d Then
            Dim oTempGeom2d As IJMfgGeom2d
            Set oTempGeom2d = oTempObj
            
            If oTempGeom2d.GetGeometryType = STRMFG_UV_MARK Then
                Set oUVLine = oTempGeom2d.GetGeometry
            ElseIf oTempGeom2d.GetGeometryType = STRMFG_TEMPLATELOCATIONMARK Then
                Set oBottomCurve = oTempGeom2d.GetGeometry
            End If
        End If
    Next
    
    ' Exit when inputs are not available
    If oUVLine Is Nothing Then
        Exit Function
    End If
    
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
    
    '*** Get the wirebody from complexstring ***'
    Dim oBottomWB As IJWireBody
    Set oBottomWB = oMfgRuleHelper.ComplexStringToWireBody(oBottomCurve)

    '*** Get the middle point of the curve ***'
    Dim oMidPos As IJDPosition
    Set oMidPos = oMfgRuleHelper.GetMiddlePoint(oBottomWB)
    
    ' Place thickness direction annotation at 0.2 m from center of bottom curve
    Dim oRefPos     As IJDPosition
    Set oRefPos = oMfgRuleHelper.GetPointAlongCurveAtDistance(oBottomWB, oMidPos, 0.2, Nothing)
    
    ' Get the curve vector
    Dim dTanX As Double, dTanY As Double, dTanZ As Double
    
    Dim oMarkVec    As IJDVector
    Set oMarkVec = New DVector
    
    oUVLine.GetDirection dTanX, dTanY, dTanZ
    oMarkVec.Set dTanX, dTanY, dTanZ
    oMarkVec.length = -1    ' Thickness direction is opposite to UV mark direction
    
    ' Create BCL mark of fixed length
    Dim oBCLMark    As IJCurve
    Set oBCLMark = CreateLineAtPosition(oRefPos, oMarkVec, 0.05)
    
    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    ' Create Geom2D for BCL mark
    Dim oMfgGeom2D As IJMfgGeom2d
    Set oMfgGeom2D = CreateGeom2D(oBCLMark, STRMFG_REF_MARK, Nothing, "Base Control Mark")
    
    oMfgGeom2D.IsSupportOnly = True ' It is meant for thickness direction
    
    oTempElems.Add oMfgGeom2D
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oTempElems
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3017, , "RULES")
End Function
