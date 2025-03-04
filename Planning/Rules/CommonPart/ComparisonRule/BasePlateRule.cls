VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasePlateRule"
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

Implements IJCmnPartComparison

Private Const sSOURCEFILE As String = "BasePlateRule.cls"
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors   ' To collect and propagate the errors.
Private m_oError As IJEditError     ' Defined here for convenience

Private m_strLogFile As String
Private m_oStream As TextStream

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
    Dim oTransMatrix        As IJDT4x4
    
    Dim oMatchingObjects    As IJElements
'    Dim oMatchingFeatures   As IJElements
    
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

    oPlnCompareHelper.RelationChecks PhysicalConnections
    oPlnCompareHelper.RelationChecks PlateParts
    oPlnCompareHelper.RelationChecks ProfileParts
        
    bIsCommon = oPlnCompareHelper.AreCommon(oMatchingObjects)        ', oMatchingFeatures)
    
    Set oTransMatrix = oPlnCompareHelper.TransMatrix
    
    Dim eMatch              As PhysConnMatch
    Dim oMatchingpair       As IJPlnMatchingPair
    Dim oPhysConnPair       As IJPlnPhysConnPair
    Dim oCndtConnObject     As Object
    Dim oTgtConnObject      As Object
    Dim lIsProfile          As Long
    Dim lIsPlate            As Long
    Dim strObjType          As String
    
    For Each oMatchingpair In oMatchingObjects
    
        If oMatchingpair.TypeOfMatch <> PlnMatch_Full Then
            bIsCommon = False
            Exit For
        End If
    
        lIsPlate = oMatchingpair.TypeOfObject And PlateParts
        lIsProfile = oMatchingpair.TypeOfObject And ProfileParts
        
        Set oCndtConnObject = oMatchingpair.CandidateConnObject
        Set oTgtConnObject = oMatchingpair.TargetConnObject
        
        If (lIsPlate = PlateParts) Then
            If TypeOf oCndtConnObject Is IJCollarPart And TypeOf oTgtConnObject Is IJCollarPart Then
                strObjType = "Collars"
            ElseIf IsBracket(oCndtConnObject) And IsBracket(oTgtConnObject) Then
                strObjType = "Brackets"
            Else
                strObjType = "Plates"
            End If
        ElseIf (lIsProfile = ProfileParts) Then
                strObjType = "Profiles"
        End If
        
        If AreObjectsCommon(strObjType, oTransMatrix, oCndtConnObject, oTgtConnObject) = False Then
            bIsCommon = False
        Else
            bIsCommon = True
        End If

ContinueLoop:
    Next

    IJCmnPartComparison_IsCommon = bIsCommon

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function

Private Function AreObjectsCommon(strObjType As String, oTransMatrix As IJDT4x4, _
                                  oCndtConnObject As Object, oTgtConnObject As Object) As Boolean
Const METHOD = "AreObjectsCommon"
On Error GoTo ErrorHandler

    Dim bIsCommon           As Boolean
    Dim oPlnCompareHelper   As IJPlnCompareHelper
    
    Dim oNamedItem          As IJNamedItem
    Dim strName1            As String
    Dim strName2            As String
    
    Set oNamedItem = oCndtConnObject
    strName1 = oNamedItem.Name
    
    Set oNamedItem = oTgtConnObject
    strName2 = oNamedItem.Name
    
    bIsCommon = False
    Set oPlnCompareHelper = New CPlnCommonPartHelper
    
    oPlnCompareHelper.Candidate = oCndtConnObject
    oPlnCompareHelper.Target = oTgtConnObject
    
    oPlnCompareHelper.TransMatrix = oTransMatrix
        
    oPlnCompareHelper.DistanceTolerance = 0.001
    
    Select Case strObjType
        
        Case "Collars"
            oPlnCompareHelper.Checks MaterialTypeAndGrade
            oPlnCompareHelper.Checks Definition
            oPlnCompareHelper.Checks OuterContours
            oPlnCompareHelper.Checks InnerContours
        
        Case "Brackets"
            oPlnCompareHelper.Checks MaterialTypeAndGrade
            oPlnCompareHelper.Checks Definition
            oPlnCompareHelper.Checks OuterContours
            oPlnCompareHelper.Checks InnerContours
        
        Case "Plates"
            oPlnCompareHelper.Checks MaterialTypeAndGrade
            oPlnCompareHelper.Checks Definition
            oPlnCompareHelper.Checks OuterContours
            oPlnCompareHelper.Checks InnerContours
        
        Case "Profiles"
            oPlnCompareHelper.Checks MaterialTypeAndGrade
            oPlnCompareHelper.Checks Definition
            oPlnCompareHelper.Checks Features
    End Select
    
    bIsCommon = oPlnCompareHelper.AreCommon
    
    If bIsCommon = False Then
        WriteToLog (strObjType + " " + strName1 + ", " + strName2 + " are not common"), m_oStream
    Else
        WriteToLog (strObjType + " " + strName1 + ", " + strName2 + " are common"), m_oStream
    End If
    
    AreObjectsCommon = bIsCommon
                                                    
Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function
