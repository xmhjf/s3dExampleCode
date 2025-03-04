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
' Project: MfgPlateProcess
' Module: ValidateMarking
'
' Description:  In this function, user can control Validate marks.
'  History:
'       StructMfg         October 6th, 2010   created
'
'*******************************************************************************
Option Explicit

Implements IJMfgMarkingValidate

Const MODULE = "MfgPlateProcess.ValidateMarking"

Private Sub IJMfgMarkingValidate_ValidateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal MfgGeomCol2d As GSCADMfgRulesDefinitions.IJMfgGeomCol2d)
    Const METHOD = "ValidateMarking: IJMfgMarkingValidate_ValidateAfterUnfold"
    On Error GoTo ErrorHandler

CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1076, , "RULES")
    GoTo CleanUp
End Sub

Private Function IdxOfGeomTypeInPrecedenceList(CurGeomType As StrMfgGeometryType, PrecedenceList() As StrMfgGeometryType) As Long
    IdxOfGeomTypeInPrecedenceList = -1
    Dim i As Long
    For i = LBound(PrecedenceList) To UBound(PrecedenceList)
        If CurGeomType = PrecedenceList(i) Then
            IdxOfGeomTypeInPrecedenceList = i
            Exit Function
        End If
    Next
End Function

Private Function IsLocationMark(MarkGeomType As StrMfgGeometryType) As Boolean
    IsLocationMark = ((MarkGeomType = STRMFG_PLATELOCATION_MARK) Or _
                      (MarkGeomType = STRMFG_PROFILELOCATION_MARK) Or _
                      (MarkGeomType = STRMFG_BRACKETLOCATION_MARK))
End Function

Private Function IsReferenceMark(MarkGeomType As StrMfgGeometryType) As Boolean
    IsReferenceMark = ((MarkGeomType = STRMFG_FRAMELINE_MARK) Or _
                      (MarkGeomType = STRMFG_WATERLINE_MARK) Or _
                      (MarkGeomType = STRMFG_BUTTOCKLINE_MARK))
End Function

Private Function ComparePriorities(ThisGeomType As StrMfgGeometryType, ThatGeomType As StrMfgGeometryType, PrecedenceList() As StrMfgGeometryType) As Integer
    Dim ThisPriority As Long
    ThisPriority = IdxOfGeomTypeInPrecedenceList(ThisGeomType, PrecedenceList)
    
    Dim ThatPriority As Long
    ThatPriority = IdxOfGeomTypeInPrecedenceList(ThatGeomType, PrecedenceList)
    
    ' Check if either of the geom types are missing from the list
    If ThisPriority < 0 And ThatPriority < 0 Then
        ComparePriorities = 0
        Exit Function
    ElseIf ThisPriority >= 0 And ThatPriority < 0 Then
        ComparePriorities = -1
        Exit Function
    ElseIf ThisPriority < 0 And ThatPriority >= 0 Then
        ComparePriorities = 1
        Exit Function
    End If
    
    ' Both geom types are in the precedence list.
    If IsLocationMark(ThisGeomType) And IsLocationMark(ThatGeomType) Then
        ComparePriorities = 0
        Exit Function
    ElseIf ThisPriority < ThatPriority Then
        ComparePriorities = -1
        Exit Function
    ElseIf ThisPriority > ThatPriority Then
        ComparePriorities = 1
        Exit Function
    ElseIf ThisPriority = ThatPriority Then
        ComparePriorities = 0
        Exit Function
    End If
End Function

Private Function PickWhichOneToOffset(Xgo As IJMfgGeom3d, Xsm As IJMfgSystemMark, _
                                      Ygo As IJMfgGeom3d, Ysm As IJMfgSystemMark, _
                                      Xside As Long, Yside As Long, UpSide As Long) As IJMfgGeom3d
    Const METHOD = "PickWhichOneToOffset"
    On Error GoTo ErrorHandler

    If Xside = UpSide And Yside <> UpSide Then
        Set PickWhichOneToOffset = Ygo
    ElseIf Xside <> UpSide And Yside = UpSide Then
        Set PickWhichOneToOffset = Xgo
    End If
    
    ' TODO: Logic based on relative thickness direction vectors
    
CleanUp:
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1076, , "RULES")
    GoTo CleanUp
End Function

Private Sub OffsetAndResetAttribution(GeomOutput As IJMfgGeom3d)
    Const METHOD = "OffsetAndResetAttribution"
    On Error GoTo ErrorHandler

    Dim MarkInfo As MarkingInfo
    Set MarkInfo = GeomOutput.SystemMark
    
    ' First find how much to offset along what direction
    Dim OffsetVec As IJDVector
    Set OffsetVec = MarkInfo.ThicknessDirection
    
    Dim OffsetAmount As Double
    OffsetAmount = OffsetVec.Length
    
    Dim oCrv As IJCurve
    Set oCrv = GeomOutput.GetGeometry
    
    Dim x As Double, y As Double, z As Double, dummy As Double
    oCrv.EndPoints x, y, z, dummy, dummy, dummy
    
    x = x + OffsetVec.x
    y = y + OffsetVec.y
    z = z + OffsetVec.z
    
    ' Now do offset
    Dim GeomService As IGeometryServices
    Set GeomService = New GeometryFactory
    
    Dim OffsetCS As IJComplexString
    Set OffsetCS = GeomService.CreateByOffset(Nothing, oCrv, x, y, z, OffsetAmount, 1)
                                              
'    Dim oGH As New MfgGeomHelper
'    oGH.DumpGeometryToFile GeomOutput.GetGeometry, False, False, Environ("TEMP"), "OrigCS"
'    oGH.DumpGeometryToFile OffsetCS, False, False, Environ("Temp"), "OffsetCS"
'    Set oGH = Nothing

    ' Now modify properties of Geom output
    
    GeomOutput.PutGeometry OffsetCS
    
    OffsetVec.Length = -1 * OffsetAmount
    MarkInfo.ThicknessDirection = OffsetVec
    
    MarkInfo.FittingAngle = 3.14159265358979 - MarkInfo.FittingAngle
    
    Select Case LCase(MarkInfo.Direction)
        Case "aft": MarkInfo.Direction = "fore"
        Case "fore": MarkInfo.Direction = "aft"
        Case "down": MarkInfo.Direction = "up"
        Case "up": MarkInfo.Direction = "down"
        Case "in": MarkInfo.Direction = "out"
        Case "out": MarkInfo.Direction = "in"
        Case "port": MarkInfo.Direction = "starboard"
        Case "starboard": MarkInfo.Direction = "port"
    End Select
    
CleanUp:
    Set GeomService = Nothing
    Set MarkInfo = Nothing
    Set OffsetCS = Nothing
    Set OffsetVec = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1076, , "RULES")
    GoTo CleanUp
End Sub

Private Sub IJMfgMarkingValidate_ValidateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal oMfgGeomCol3d As IJMfgGeomCol3d)
    Const METHOD = "ValidateMarking: IJMfgMarkingValidate_ValidateBeforeUnfold"
    On Error GoTo ErrorHandler

    Dim OverlapPrecedence(1 To 9) As StrMfgGeometryType
    OverlapPrecedence(1) = STRMFG_PLATELOCATION_MARK
    OverlapPrecedence(2) = STRMFG_PROFILELOCATION_MARK
    OverlapPrecedence(3) = STRMFG_BRACKETLOCATION_MARK
    OverlapPrecedence(4) = STRMFG_FRAMELINE_MARK
    OverlapPrecedence(5) = STRMFG_WATERLINE_MARK
    OverlapPrecedence(6) = STRMFG_BUTTOCKLINE_MARK
    OverlapPrecedence(7) = STRMFG_TemplateLocationMarkLine
    OverlapPrecedence(8) = STRMFG_KNUCKLE_MARK
    OverlapPrecedence(9) = STRMFG_ROLL_BOUNDARIES_MARK
    
    Dim oMfgGeomHelper As New MfgGeomHelper

    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
    Dim oElements As IJElements
    
    Dim oSOElements As IJElements
    Set oSOElements = New JObjectCollection
    
    Dim i As Long, j As Long
    For i = 1 To oMfgGeomCol3d.GetCount ' Evaluate count
        Dim ThisOne As IJMfgGeom3d
        Set ThisOne = oMfgGeomCol3d.GetGeometry(i)
        'For UV Mark and Contours, there is no system mark. Hence proceed with next one
        If ThisOne.GetGeometryType = 6 Or ThisOne.GetGeometryType = 0 Then GoTo NextI
        
        Dim oMarkingInfo As MarkingInfo
        Set oMarkingInfo = ThisOne.SystemMark
        
        If Not ThisOne.IsSupportOnly Then
            For j = 1 To oMfgGeomCol3d.GetCount
                'Geometry should not compare with itself
                If j = i Then GoTo NextJ
                Dim ThatOne As IJMfgGeom3d
                Set ThatOne = oMfgGeomCol3d.GetGeometry(j)
                'For UV Mark and Contours, there is no system mark. Hence proceed with next one
                If ThatOne.GetGeometryType = 6 Or ThatOne.GetGeometryType = 0 Then GoTo NextJ
                
                If Not ThatOne.IsSupportOnly Then
                    Dim oCrv1 As IJCurve, oCrv2 As IJCurve
                    Set oCrv1 = ThisOne.GetGeometry
                    Set oCrv2 = ThatOne.GetGeometry
                    
                    Dim ThisSystemMark As IJMfgSystemMark, ThatSystemMark As IJMfgSystemMark
                    Set ThisSystemMark = ThisOne.SystemMark
                    Set ThatSystemMark = ThatOne.SystemMark
                    
                    Dim TempOne As IJMfgGeom3d
                    Set TempOne = oMfgGeomCol3d.GetGeometry(1)
                    
                    Dim oTempCrv As IJCurve
                    Set oTempCrv = TempOne.GetGeometry
                    
                    Dim ThisSide As Long
                    If ThisSystemMark Is Nothing Then
                        ThisSide = UpSide
                    Else
                    ThisSide = ThisSystemMark.GetMarkingSide
                    End If
                    
                    Dim ThatSide As Long
                    If ThatSystemMark Is Nothing Then
                        ThatSide = UpSide
                    Else
                    ThatSide = ThatSystemMark.GetMarkingSide
                    End If
                    
                    Dim ThisGeomType As StrMfgGeometryType, ThatGeomType As StrMfgGeometryType
                    ThisGeomType = ThisOne.GetGeometryType
                    ThatGeomType = ThatOne.GetGeometryType
                    
                    Dim PrCmp As Integer
                    PrCmp = ComparePriorities(ThisGeomType, ThatGeomType, OverlapPrecedence)
                        
                    ' Check for (any kind of) overlap within 1 mm.
                    If oMfgGeomHelper.CheckOverlapBetweenTwoCurvesWithinTol(oCrv1, oCrv2, 0.001) Then
                        Dim InputGeomCol As IJElements
                        Set InputGeomCol = New JObjectCollection
                        
                        InputGeomCol.Add oCrv1
                        InputGeomCol.Add oCrv2
                        
                        Dim MergeCurveCol As IJElements
                        Set MergeCurveCol = oMfgGeomHelper.OptimizedMergingOfInputCurvesEx(InputGeomCol, 0.001, 0.001)
                        Set InputGeomCol = Nothing

                        Dim MergedCS As IJComplexString
                        If MergeCurveCol.Count = 1 Then
                            Set MergedCS = MergeCurveCol.Item(1)
                        Else
                            StrMfgLogError Err, MODULE, METHOD, _
                                           "Review merge between " & i & " and " & j, _
                                           "SMCustomWarningMessages", 1076, , "RULES"
                        End If
                        
                        If ThisSide = UpSide And ThatSide <> UpSide Then
                            ' This is on Marking side, whereas That is on AntiMarking side -- Remove That One
                            ThatOne.IsSupportOnly = True
                            oSOElements.Add ThatOne
                            If Not MergedCS Is Nothing Then
                                ThisOne.PutGeometry MergedCS
                                'ThisSystemMark.SetMarkingSide ThatSide
                            End If
                        ElseIf ThatSide = UpSide And ThisSide <> UpSide Then
                            ' That is on Marking side, whereas This is on AntiMarking side
                            ThisOne.IsSupportOnly = True
                            oSOElements.Add ThisOne
                            If Not MergedCS Is Nothing Then
                                ThatOne.PutGeometry MergedCS
                                'ThatSystemMark.SetMarkingSide ThisSide
                            End If
                        ElseIf PrCmp < 0 Or (PrCmp = 0 And _
                               (oCrv1.Length > oCrv2.Length)) _
                        Then
                            ' This has either higher priority or longer than That -- Remove That One
                            ThatOne.IsSupportOnly = True
                            oSOElements.Add ThatOne
                            If Not MergedCS Is Nothing Then
                                ThisOne.PutGeometry MergedCS
                                'ThisSystemMark.SetMarkingSide ThatSide
                            End If
                        ElseIf PrCmp > 0 Or (PrCmp = 0 And _
                               (oCrv1.Length < oCrv2.Length)) _
                        Then
                            ' That has either higher priority or longer than This -- Remove This One
                            ThisOne.IsSupportOnly = True
                            oSOElements.Add ThisOne
                            If Not MergedCS Is Nothing Then
                                ThatOne.PutGeometry MergedCS
                                'ThatSystemMark.SetMarkingSide ThisSide
                            End If
                        Else ' They have equal priority, identical in length and on same side!
                            'Add exceptions to fitting marks
                            If Not (ThisOne.GetGeometryType = STRMFG_FITTING_MARK Or ThatOne.GetGeometryType = STRMFG_FITTING_MARK) Then
                            StrMfgLogError Err, MODULE, METHOD, _
                                           "Two marking cuves of same type on the same side of plate completely overlap!", _
                                           "SMCustomWarningMessages", 1076, , "RULES"
                            End If
                        End If ' Priority comparison
                        
                        ModifyFittingMarksForOverlap oMfgGeomCol3d, UpSide, ThisOne, ThatOne
                        Set MergedCS = Nothing
                        
                        
                    ElseIf oMfgGeomHelper.CheckOverlapBetweenTwoCurvesWithinTol(oCrv1, oCrv2, 0.005) Then
                        ' Overlap exists between 1 and 5mm
                        ' If one of them is not a location mark, it should not be part of output
                        If Not IsLocationMark(ThisGeomType) And Not IsLocationMark(ThatGeomType) Then
                            ' Neither are location marks
                            If PrCmp < 0 Or (PrCmp = 0 And _
                               (oCrv1.Length > oCrv2.Length Or _
                               (ThisSide = UpSide And ThatSide <> UpSide))) _
                            Then
                                ThatOne.IsSupportOnly = True
                                oSOElements.Add ThatOne
                                ModifyFittingMarksForOverlap oMfgGeomCol3d, UpSide, ThisOne, ThatOne
                                
                            ElseIf PrCmp > 0 Or (PrCmp = 0 And _
                                   (oCrv1.Length < oCrv2.Length Or _
                                   (ThatSide = UpSide And ThisSide <> UpSide))) _
                            Then
                                ThisOne.IsSupportOnly = True
                                oSOElements.Add ThisOne
                                ModifyFittingMarksForOverlap oMfgGeomCol3d, UpSide, ThisOne, ThatOne
                                
                            End If
                        ElseIf Not IsLocationMark(ThisGeomType) Then
                            ' This is not a location mark
                            ThisOne.IsSupportOnly = True
                            oSOElements.Add ThisOne
                        ElseIf Not IsLocationMark(ThatGeomType) Then
                            ' That is not a location mark
                            ThatOne.IsSupportOnly = True
                            oSOElements.Add ThatOne
                        Else ' Both are location marks
                            ' One of them has to be offset by the thickness amount
                            Dim ObjToOffsetGeom As IJMfgGeom3d
                            Set ObjToOffsetGeom = PickWhichOneToOffset(ThisOne, ThisSystemMark, _
                                                                       ThatOne, ThatSystemMark, _
                                                                       ThisSide, ThatSide, UpSide)
                            
                            OffsetAndResetAttribution ObjToOffsetGeom
                        End If
                    End If
                    
                    Set oCrv1 = Nothing
                    Set oCrv2 = Nothing
                End If ' That one is considered for overlapping
                
NextJ:
                Set ThatOne = Nothing
            Next ' Inner loop
        End If ' This one is considered for overlapping
        
NextI:
        Set ThisOne = Nothing
    Next ' Outer loop
    
CleanUp:
    Set oMfgGeomHelper = Nothing
    Set MergedCS = Nothing
    Set oCrv1 = Nothing
    Set oCrv2 = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1076, , "RULES")
    GoTo CleanUp
End Sub


Private Sub ModifyFittingMarksForOverlap(ByRef oMfgGeomCol3d As IJMfgGeomCol3d, ByVal UpSide As Long, _
                                            ByVal ThisOne As IJMfgGeom3d, ByVal ThatOne As IJMfgGeom3d)

    Const METHOD = "ValidateMarking: ModifyFittingMarksForOverlap"
    On Error GoTo ErrorHandler

    Dim oRelatedGeom3d As IJMfgGeom3d
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oMB As IJDModelBody
    Dim oMarkingInfo As MarkingInfo
    Dim oRollRegionMonikers As IJElements
    Dim oMarkingLineMoniker As IMoniker
    Dim kk As Long, jj As Long
    Dim oNewCurve As IJCurve
    Dim oNewCS As IJComplexString
    Dim oNewLine As IJLine
    Dim o3dVec As New DVector
    Dim oRelatedCurve As IJCurve
    Dim dSX As Double, dSY As Double, dSZ As Double, dEX As Double, dEY As Double, dEZ As Double
    Dim oMyMark As IJMfgGeom3d, oMyOtherMark As IJMfgGeom3d
    
    Dim ThisSystemMark As IJMfgSystemMark
    Set ThisSystemMark = ThisOne.SystemMark
    Dim ThatSystemMark As IJMfgSystemMark
    Set ThatSystemMark = ThatOne.SystemMark
    
    If ThisSystemMark Is Nothing Or ThatSystemMark Is Nothing Then Exit Sub
    
    '*** LOGIC OF OVERLAPPING:
    '... 1. Add SubgeometryType STRMFG_ANNOTATION_MARK to the other (support only) mark in order to flip the name of that mark
    '... 2. Flip the End Fitting Marks of LocationMark on Marking Side
    '... 3. Add SubgeometryType STRMFG_FITTING_MARK to that location mark's all End Fitting Marks to get the Annotation
    '... 4. Add SubgeometryType STRMFG_FITTING_MARK to the other(support only) location marks' End Fitting Marks to get the Annotation
    
    'Check AntiMarkingSide
    If ThisSystemMark.GetMarkingSide <> UpSide Then
        Set oMyMark = ThisOne
        Set oMyOtherMark = ThatOne
        ThisOne.PutSubGeometryType STRMFG_ANNOTATION_MARK
        Set oMarkingInfo = oMyMark.SystemMark
        oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "OverlapFlip"
    End If

    'Else
    If ThatSystemMark.GetMarkingSide <> UpSide Then
        Set oMyMark = ThatOne
        Set oMyOtherMark = ThisOne
        ThatOne.PutSubGeometryType STRMFG_ANNOTATION_MARK
    Set oMarkingInfo = oMyMark.SystemMark
    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "OverlapFlip"
    End If
    
    If oMyOtherMark Is Nothing Then Exit Sub
                
    For kk = 1 To oMfgGeomCol3d.GetCount
        Set oRelatedGeom3d = oMfgGeomCol3d.GetGeometry(kk)
        If oRelatedGeom3d.GetGeometryType = STRMFG_FITTING_MARK Then
        
            If oRelatedGeom3d.GetMoniker Is oMyOtherMark.GetMoniker And IsLocationMark(oMyOtherMark.GetGeometryType) Then
            
                Set oMarkingInfo = oRelatedGeom3d.SystemMark
                oRelatedGeom3d.PutSubGeometryType STRMFG_FITTING_MARK
                oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "Overlap"
                
                Dim om As MfgRuleHelpers.Helper
                Dim oMfgGeomHelper As New MfgGeomHelper

                Set oCS = oRelatedGeom3d
                oCS.GetCurve 1, oRelatedCurve
                oRelatedCurve.EndPoints dSX, dSY, dSZ, dEX, dEY, dEZ
                
                o3dVec.Set dEX - dSX, dEY - dSY, dEZ - dSZ
                o3dVec.Length = o3dVec.Length * -1
                
                Set oNewLine = New Line3d
                oNewLine.DefineBy2Points dSX, dSY, dSZ, dSX + o3dVec.x, dSY + o3dVec.y, dSZ + o3dVec.z
                
                Set oNewCS = New ComplexString3d
                oNewCS.AddCurve oNewLine, True
                oRelatedGeom3d.PutGeometry oNewCS
                
            End If
        End If
    Next kk
    
    'For the other side set subgeometrytype of End Fitting Marks in order to generate the annotation
    For jj = 1 To oMfgGeomCol3d.GetCount
        Set oRelatedGeom3d = oMfgGeomCol3d.GetGeometry(jj)
        If oRelatedGeom3d.GetGeometryType = STRMFG_FITTING_MARK Then
            If oRelatedGeom3d.GetMoniker Is oMyMark.GetMoniker And IsLocationMark(oMyMark.GetGeometryType) Then
            
                Set oMarkingInfo = oRelatedGeom3d.SystemMark
                oRelatedGeom3d.PutSubGeometryType STRMFG_FITTING_MARK
                oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "Overlap"
                
            End If
        End If
    Next jj
    
CleanUp:
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1076, , "RULES")
    GoTo CleanUp
'End If
End Sub
