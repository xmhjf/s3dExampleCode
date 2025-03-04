VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'  Project     : SharedContent\Src\Planning\Rules\CommonPartRules\ComparisonRule\
'  File        : PlateRule.cls
'
'  Description : Common Part Plate comparison rule
'
'  History     :
'   14th June 2010      Triveni CH     Initial creation
'**************************************************************************************

Option Explicit

Private Const sSOURCEFILE As String = "CollarRule.cls"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Private m_oError As IJEditError     ' Defined here for convenience

Implements IJCmnPartComparison


Private Sub Class_Initialize()
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Function IJCmnPartComparison_IsCommon(ByVal oCollarSrcObj As Object, ByVal oCollarDestObj As Object, _
                                              ByVal oCmnPartMgr As Object) As Boolean
Const METHOD = "IJCmnPartComparison_IsCommon"
On Error GoTo ErrorHandler

    IJCmnPartComparison_IsCommon = False
    
    Dim bIsCommon           As Boolean
    Dim oPlnCompareHelper   As IJPlnCompareHelper
    
    bIsCommon = False
    Set oPlnCompareHelper = New CPlnCommonPartHelper
    
    oPlnCompareHelper.Candidate = oCollarSrcObj
    oPlnCompareHelper.Target = oCollarDestObj
        
    oPlnCompareHelper.DistanceTolerance = 0.001
    
    oPlnCompareHelper.Checks MaterialTypeAndGrade
    oPlnCompareHelper.Checks Definition
    oPlnCompareHelper.Checks OuterContours
    oPlnCompareHelper.Checks InnerContours
            
    bIsCommon = oPlnCompareHelper.AreCommon
    IJCmnPartComparison_IsCommon = bIsCommon
        
    Exit Function
    
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function
