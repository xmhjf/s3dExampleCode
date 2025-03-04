VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shrinkage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgFrameSysRule
' Module: Shrinkage
'
' Author: Anand Hariharan
'
' Comments:
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgFrameSysRule.Shrinkage"

Implements IJMfgRuleForFrameSystem

Private Function GetResourceManager() As Object
Const METHOD = "GetPOM"
On Error GoTo ErrHandler
    
    Dim oContext As IJContext
    Set oContext = GetJContext()
    
    Dim oAccessMiddle As IJDAccessMiddle
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    
    Set GetResourceManager = oAccessMiddle.GetResourceManagerFromType("Model")
    
    Exit Function
    
ErrHandler:
        Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IJMfgRuleForFrameSystem_GetFrameSystemForMfgEntity(ByVal pEntityDsp As Object) As IJElements
    Const METHOD = "IJMfgRuleForFrameSystem_GetFrameSystemForMfgEntity"
    On Error GoTo ErrorHandler
    
    Dim oFSglobal As IJFrameSystemGlobals
    Set oFSglobal = New JFrameSystemGlobals
    
    ' Get the unique FrameSystemFactory object
    Dim oFrameSysFactory As IJFrameSystemFactory
    Set oFrameSysFactory = oFSglobal.FrameSystemFactory(GetResourceManager(), True)

    Dim RetElements As IJElements
    Set RetElements = New JObjectCollection
    
    Dim j As Long
    For j = 1 To oFrameSysFactory.Count
        Dim oFrameSystem As IJFrameSystem
        Set oFrameSystem = oFrameSysFactory.Item(j)
        
        RetElements.Add oFrameSystem
        Set oFrameSystem = Nothing
    Next j
    
    Set IJMfgRuleForFrameSystem_GetFrameSystemForMfgEntity = RetElements
    
    Set oFSglobal = Nothing
    Set oFrameSysFactory = Nothing
    Set RetElements = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3002, , "RULES")
End Function


 
