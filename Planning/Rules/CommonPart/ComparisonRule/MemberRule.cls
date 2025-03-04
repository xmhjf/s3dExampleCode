VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonPartRules\ComparisonRule\
'  File        : MemberRule.cls
'
'  Description : Common Part member comparison rule
'
'  History     :
'   6th July 2011      Apparao.N     Initial creation
'**************************************************************************************

Option Explicit

Implements IJCmnPartComparison

Private Const sSOURCEFILE As String = "MemberRule.cls"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Private m_oError As IJEditError     ' Defined here for convenience

Private Sub Class_Initialize()
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Function IJCmnPartComparison_IsCommon(ByVal oMemberPart1 As Object, ByVal oMemberPart2 As Object, _
                                              ByVal oCmnPartMgr As Object) As Boolean
Const METHOD = "IJCmnPartComparison_IsCommon"
On Error GoTo ErrorHandler
    
    Dim bIsCommon           As Boolean
    Dim oPlnCompareHelper   As IJPlnCompareHelper
    
    bIsCommon = False
    Set oPlnCompareHelper = New CPlnCommonPartHelper
    
    'Set Candidate and target
    oPlnCompareHelper.Candidate = oMemberPart1
    oPlnCompareHelper.Target = oMemberPart2
        
    'Set distance tolerance
    oPlnCompareHelper.DistanceTolerance = 0.001
    
    'Set the comparision criteria
    oPlnCompareHelper.Checks MaterialTypeAndGrade
    oPlnCompareHelper.Checks Definition
'    oPlnCompareHelper.Checks Features
    
    'Check whether objects are common
    bIsCommon = oPlnCompareHelper.AreCommon
    
    IJCmnPartComparison_IsCommon = bIsCommon

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function
