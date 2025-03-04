VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldCompNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2005, Intergraph Corp.  All rights reserved.
'
' Module:       WeldCompNo
'
' Description:  Remove Weld Compensation from Profile
'
' Author:       hkarumug@ingr.com
'
' Comments:
'
'   2005.04.13  hkarumug@ingr.com     Initial Creation.
'
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkRule

Const MODULE = "MfgProfileProcess.WeldCompNo"

Private Sub Execute(ByVal Part As Object)
    Const METHOD = "Execute"
    On Error GoTo ErrorHandler
           
    Dim oWCFactory As IJMfgWeldCompensationFactory
    Set oWCFactory = New MfgWeldCompensationFactory
    
    ' Does a compensation already exist ?
    On Error Resume Next
    
    Dim oWC As IJMfgWeldCompensation
    Set oWC = oWCFactory.GetWeldCompensationOnProfile(Part)
    
    On Error GoTo ErrorHandler
    If Not oWC Is Nothing Then
        ' Delete the WC object
        Dim oWCObject As IJDObject
        Set oWCObject = oWC
        
        oWCObject.Remove
    End If
            
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IJDMfgSystemMarkRule_CreateAfterUnfold(ByVal Part As Object, ByVal Upside As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
'
End Function

Private Function IJDMfgSystemMarkRule_CreateBeforeUnfold(ByVal Part As Object, ByVal Upside As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgSystemMarkRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    ' Remove WeldCompensation
    Call Execute(Part)
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2036, , "RULES")
End Function

 
