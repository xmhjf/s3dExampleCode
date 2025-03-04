VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidateMarking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2010, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileProcess
' Module: ValidateMarking
'
' Description:  In this function, user can control Validate marks.
'  History:
'       StructMfg         October 6th, 2010   created
'
'*******************************************************************************
Implements IJMfgMarkingValidate
Const Module = "MfgProfileProcess.ValidateMarking"
Private Sub IJMfgMarkingValidate_ValidateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal MfgGeomCol3d As GSCADMfgRulesDefinitions.IJMfgGeomCol3d)
    Const METHOD = "ValidateMarking: IJMfgMarkingValidate_ValidateBeforeUnfold"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 2058, , "RULES")
    GoTo CleanUp
End Sub
Private Sub IJMfgMarkingValidate_ValidateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "ValidateMarking: IJMfgMarkingValidate_ValidateAfterUnfold"
    On Error GoTo ErrorHandler
    
    Dim i As Long, j As Long
    Dim oKnuckleMark As IJMfgGeom2d
    Dim oRollBoundaryMark As IJMfgGeom2d
    Dim oMfgGeomHelper As New MfgGeomHelper
        
    For i = MfgGeomCol2d.Getcount To 1 Step -1
        Set oKnuckleMark = MfgGeomCol2d.GetGeometry(i)
         'Checking for the Knuckle Mark
        If oKnuckleMark.GetGeometryType = STRMFG_KNUCKLE_MARK Then
            For j = MfgGeomCol2d.Getcount To 1 Step -1
                If Not MfgGeomCol2d.GetGeometry(j) Is oKnuckleMark Then
                    Set oRollBoundaryMark = MfgGeomCol2d.GetGeometry(j)
                    'Checking for the Roll Boundary Mark
                    If oRollBoundaryMark.GetGeometryType = STRMFG_ROLL_BOUNDARIES_MARK Then
                        'Checking whether there is a overlapping between Roll Boundary and Knuckle Mark
                        If oMfgGeomHelper.CheckOverlapBetweenTwoCurvesWithinTol(oKnuckleMark.GetGeometry, oRollBoundaryMark.GetGeometry, 0.005) = True Then
                            'oRollBoundaryMark.IsSupportOnly = True
                            Dim oObjToDel As IJDObject
                            Set oObjToDel = oRollBoundaryMark
                            oObjToDel.Remove
                            Set oObjToDel = Nothing
                        End If
                    End If
                End If
            Next
            Set oRollBoundaryMark = Nothing
        End If
        Set oKnuckleMark = Nothing
    Next
    
CleanUp:
    Exit Sub
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD, , "SMCustomWarningMessages", 2058, , "RULES")
    GoTo CleanUp
End Sub

