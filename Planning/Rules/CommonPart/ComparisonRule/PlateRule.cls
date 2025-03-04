VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateRule"
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

Private m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Private m_oError As IJEditError     ' Defined here for convenience

Private m_strLogFile As String
Private m_oStream As TextStream

Private Const sSOURCEFILE As String = "PlateRule.cls"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"

Implements IJCmnPartComparison

Private Sub Class_Initialize()
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject(ERRORPROGID)
    End If
    
'    OpenLogFile m_strLogFile, m_oStream
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
'    CloseLogFile m_oStream
End Sub

Private Function IJCmnPartComparison_IsCommon(ByVal oPlateSrcObj As Object, ByVal oPlateDestObj As Object, _
                                              ByVal oCmnPartMgr As Object) As Boolean
Const METHOD = "IJCmnPartComparison_IsCommon"
On Error GoTo ErrorHandler

    Dim bIsCommon           As Boolean
    Dim oPlnCompareHelper   As IJPlnCompareHelper
    
    bIsCommon = False
    Set oPlnCompareHelper = New CPlnCommonPartHelper
    
    oPlnCompareHelper.Candidate = oPlateSrcObj
    oPlnCompareHelper.Target = oPlateDestObj
    
    'Base & Secondary Plate Parts should compare
    'the following information :
        '1. Outer contour & Inner contour ( Geometry )
        '2. Thickness - CPlatePartUtils::GetParentOfPlatePart
        '3. Grade
        '4. Short primer
        '5. Grinder
    
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
