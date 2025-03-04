Attribute VB_Name = "BevelCommon"
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPlateProcess/MfgProfileProcess
' Module: BevelCustom
'
' Description:  Determines the process settings for the mfg plate
'
' Author:
'
' Comments:
'
'*******************************************************************************
Option Explicit

Public Const MODULE = "BevelCommon"

Public Const CONNECTION_LENGTH = 0.003
Public Const CONNECTION_MIN_DIFF = 0.00005
Public Const CONNECTION_MAX_DIFF = 0.003
Public Const CONNECTION_ANGLE = 189.5

Public Const CONNECTION_VERTICAL_LINE = 1
Public Const CONNECTION_SLOPED_LINE = 2

Public Const UNDEFINED_BEVEL = 0
Public Const SQUARE_CUT = 1
Public Const Y_CUT = 2
Public Const X_CUT = 3
Public Const V_CUT = 4

Public Const AsRadians As Double = 1.74532925199433E-02
Public Const AsDegrees As Double = 57.2957795130823

Public Const M_PI As Double = 3.14159265358979
Public Sub InitializeBevelDepthsAndAngles(ByRef BevelProps As IJMfgBevelDetailProperties)
    Const METHOD = "InitializeBevelDepthsAndAngles"
    On Error GoTo ErrorHandler

    BevelProps.Angle1_M = 0#
    BevelProps.Depth1_M = 0#
    BevelProps.Method1_M = 0
    
    BevelProps.Angle1_UM = 0#
    BevelProps.Depth1_UM = 0#
    BevelProps.Method1_UM = 0
    
    BevelProps.Angle2_M = 0#
    BevelProps.Depth2_M = 0#
    BevelProps.Method2_M = 0
    
    BevelProps.Angle2_UM = 0#
    BevelProps.Depth2_UM = 0#
    BevelProps.Method2_UM = 0

    BevelProps.Nose = 0#
    BevelProps.NoseAngle = 0#
    BevelProps.NoseMethod = 0
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Sub

Public Function CheckIfBothBevelsBelongsToSameBoundingElement(ByRef oInputPart As Object, ByRef oMfgBevelProps1 As IJMfgBevelDetailProperties, ByRef oMfgBevelProps2 As IJMfgBevelDetailProperties) As Long
    Const METHOD = "CheckIfBothBevelsBelongsToSameBoundingElement"
    
    Dim oPCWrapper As StructDetailObjects.PhysicalConn
    Dim oMfgBevel1 As IJMfgBevel
    Dim oMfgBevel2 As IJMfgBevel
    Dim oOTherPart1 As Object
    Dim oOTherPart2 As Object
  
    Set oPCWrapper = New PhysicalConn
    
    Set oMfgBevel1 = oMfgBevelProps1
    Set oMfgBevel2 = oMfgBevelProps2
    
    If ((TypeOf oMfgBevel1.PhysicalConnection Is IJStructPhysicalConnection) And (TypeOf oMfgBevel2.PhysicalConnection Is IJStructPhysicalConnection)) Then
        Set oPCWrapper.object = oMfgBevel1.PhysicalConnection
        
        If oPCWrapper.ConnectedObject1 Is oInputPart Then
            Set oOTherPart1 = oPCWrapper.ConnectedObject2
        End If
        
        If oPCWrapper.ConnectedObject2 Is oInputPart Then
            Set oOTherPart1 = oPCWrapper.ConnectedObject1
        End If
        
        Set oPCWrapper.object = oMfgBevel2.PhysicalConnection
        
        If oPCWrapper.ConnectedObject1 Is oInputPart Then
            Set oOTherPart2 = oPCWrapper.ConnectedObject2
        End If
        
        If oPCWrapper.ConnectedObject2 Is oInputPart Then
            Set oOTherPart2 = oPCWrapper.ConnectedObject1
        End If
        
        If oOTherPart1 Is oOTherPart2 Then
            CheckIfBothBevelsBelongsToSameBoundingElement = True
        Else
            CheckIfBothBevelsBelongsToSameBoundingElement = False
        End If
    Else
        CheckIfBothBevelsBelongsToSameBoundingElement = False
    End If

    Set oPCWrapper = Nothing
    Set oMfgBevel1 = Nothing
    Set oMfgBevel2 = Nothing
    Set oOTherPart1 = Nothing
    Set oOTherPart2 = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
  
End Function
Public Function GetBevelType(ByVal oMfgBevel As IJMfgBevelDetailProperties, ByRef dVAngle As Double) As Long
    Const METHOD = "GetBevelType"
    On Error GoTo ErrorHandler

    Dim bChamferDepth_M As Boolean
    Dim bChamferDepth_UM As Boolean
    Dim bDepth1_M As Boolean
    Dim bDepth1_UM As Boolean
    Dim bDepth2_M As Boolean
    Dim bDepth2_UM As Boolean
    Dim bNoseDepth As Boolean
    Dim bSqaureCut As Boolean
    Dim bVCut1 As Boolean
    Dim bVCut2 As Boolean
    Dim dAngle As Double
    
    bVCut1 = False
    bVCut2 = False
    bChamferDepth_M = (Abs(oMfgBevel.ChamferDepth_M) > 0.0001)
    bChamferDepth_UM = (Abs(oMfgBevel.ChamferDepth_UM) > 0.0001)
    bDepth1_M = (Abs(oMfgBevel.Depth1_M) > 0.0001)
    bDepth1_UM = (Abs(oMfgBevel.Depth1_UM) > 0.0001)
    bDepth2_M = (Abs(oMfgBevel.Depth2_M) > 0.0001)
    bDepth2_UM = (Abs(oMfgBevel.Depth2_UM) > 0.0001)
    bNoseDepth = (Abs(oMfgBevel.Nose) > 0.0001)
    
    dVAngle = 0#
    
    If bChamferDepth_M Then
        If (Abs(oMfgBevel.ChamferAngle_M) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            bVCut1 = True
        End If
    End If
    
    If bChamferDepth_UM Then
        If (Abs(oMfgBevel.ChamferAngle_UM) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            bVCut1 = True
        End If
    End If
    
    If bDepth1_M Then
        dAngle = oMfgBevel.Angle1_M
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            dVAngle = dAngle
            bVCut1 = True
        End If
    End If
    
    If bDepth1_UM Then
        dAngle = oMfgBevel.Angle1_UM
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
                dVAngle = -1# * dAngle
            bVCut1 = True
        End If
    End If
    
    If bDepth2_M Then
        dAngle = oMfgBevel.Angle2_M
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
            dVAngle = dAngle
            bVCut1 = True
        End If
    End If
    
    If bDepth2_UM Then
        dAngle = oMfgBevel.Angle2_UM
        
        If (Abs(dAngle) < 0.000001) Then
            bSqaureCut = True
        ElseIf bVCut1 Then
            bVCut2 = True
        Else
                dVAngle = -1# * dAngle
            bVCut1 = True
        End If
    End If
    
    If bNoseDepth Then
        dAngle = oMfgBevel.NoseAngle
            If (Abs(dAngle) < 0.01) Then
                bSqaureCut = True
            ElseIf bVCut1 Then
                bVCut2 = True
            Else
                dVAngle = dAngle
                bVCut1 = True
            End If
    End If
    
    If bSqaureCut Then
        If bVCut1 Or bVCut2 Then
            GetBevelType = Y_CUT
        Else
            GetBevelType = SQUARE_CUT
        End If
    ElseIf bVCut1 And bVCut2 Then
        GetBevelType = X_CUT
    ElseIf bVCut1 Or bVCut2 Then
        GetBevelType = V_CUT
    Else
        GetBevelType = UNDEFINED_BEVEL
    End If
        
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1040, , "RULES")
End Function
Public Function MergeGeom2dsWithSameBevels(ByRef GeomCollection As IMSCoreCollections.IJElements)

    Dim i As Integer
    Dim lCollStartIndex As Integer, lPrevIndex As Integer
    Dim oPrevPC As IUnknown
    Dim oPrevMfgBevel As IJMfgBevel
    Dim oGeomElems As IJElements
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Dim oGeomCol2d As IJMfgGeomCol2d
    
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oGeomCol2d = Nothing
    
StartAgain:
    Set oGeomElems = New JObjectCollection
    
    lCollStartIndex = -1
    
    For i = 1 To GeomCollection.Count
        Dim oGeom2d As IJMfgGeom2d
        Dim oMfgBevel As IJMfgBevel
        Dim oPCUnk As IUnknown
        Dim bAddedToColl As Boolean
        
        bAddedToColl = False
        Set oGeom2d = GeomCollection.Item(i)
        
        If oGeomCol2d Is Nothing Then
            Dim oMfgGeomChild As IJMfgGeomChild
            Set oMfgGeomChild = oGeom2d
            Set oGeomCol2d = oMfgGeomChild.GetParent
            Set oMfgGeomChild = Nothing
        End If
        
        If (Not ((oGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR) Or _
         (oGeom2d.GetGeometryType = STRMFG_INNER_CONTOUR))) Then GoTo NextContourEdge
         
        Set oMfgBevel = oGeom2d.GetBevel
        If oMfgBevel Is Nothing Then
            Set oPrevPC = Nothing
            Set oPrevMfgBevel = Nothing
            GoTo NextContourEdge
        End If
        
        Set oPCUnk = oMfgBevel.PhysicalConnection
        If oPCUnk Is Nothing Then
            Set oPrevPC = Nothing
            Set oPrevMfgBevel = Nothing
            GoTo NextContourEdge
        End If
        
        If Not oPrevPC Is Nothing Then
          If (oPrevPC Is oPCUnk) And (Abs(oMfgBevel.AttachmentAngle - oPrevMfgBevel.AttachmentAngle) > 0.002) And _
                (BevelsValuesAreDifferent(oMfgBevel, oPrevMfgBevel) = False) Then
            If lCollStartIndex < 0 Then
                Dim oPrevGeom2d As IJMfgGeom2d
                Set oPrevGeom2d = GeomCollection.Item(lPrevIndex)
                oGeomElems.Add oPrevGeom2d
                lCollStartIndex = lPrevIndex
            End If
            oGeomElems.Add oGeom2d
            bAddedToColl = True
          End If
        End If
        
        Set oPrevPC = oPCUnk
        Set oPrevMfgBevel = oMfgBevel
        Set oPCUnk = Nothing
        Set oMfgBevel = Nothing
        
NextContourEdge:
        If (oGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR Or oGeom2d.GetGeometryType = STRMFG_INNER_CONTOUR) Then
            If bAddedToColl = False Then
               lPrevIndex = i
               If (oGeomElems.Count > 1) Then
                    Dim ind As Integer
                    Dim oAllCSElems As IJElements
                    Dim oFittedCS As IJComplexString
                    Dim oGeom2dToUpdate As IJMfgGeom2d
                    Dim oFittedBspCurve As Object
                    
                    Set oAllCSElems = New JObjectCollection
                    
                    For ind = 1 To oGeomElems.Count
                        Dim oThisGeom2d As IJMfgGeom2d
                        Dim oGeomCS As IJComplexString
                        Dim oCSElems As IJElements
                        Dim oIJDObject As IJDObject
                        
                        Set oThisGeom2d = oGeomElems.Item(ind)
                        
                        If ind = 1 Then
                            Set oGeom2dToUpdate = oThisGeom2d
                            Dim oLastMfgGeomGeom2d As IJMfgGeom2d
                            Set oLastMfgGeomGeom2d = oGeomElems.Item(oGeomElems.Count)
                            oGeom2dToUpdate.PutEOC oLastMfgGeomGeom2d.GetEOC
                            Set oLastMfgGeomGeom2d = Nothing
                        Else
                            oGeomCol2d.RemoveGeometry oThisGeom2d
                            Set oIJDObject = oThisGeom2d
                            oIJDObject.Remove
                            
                            GeomCollection.Remove (GeomCollection.GetIndex(oThisGeom2d))
                        End If
                        Set oGeomCS = oThisGeom2d.GetGeometry
                        'oGeomCS.GetCurves oCSElems
                        oAllCSElems.Add oGeomCS
                    Next ind
                    
                    
                    'Set oFittedBspCurve = oMfgGeomHelper.ApproximateComplexStringByStrokeAndFit(oAllCSElems, False, 3, 0.1, 0.0001)
                    Set oFittedCS = oMfgGeomHelper.OptimizedMergingOfInputCurves(oAllCSElems).Item(1)
                    'oFittedCS.SetCurves oAllCSElems
                    'oFittedCS.AddCurve oFittedBspCurve, True
                    
                    oGeom2dToUpdate.PutGeometry oFittedCS
                    oGeomElems.Clear
                    Set oPrevPC = Nothing
                    Set oPrevMfgBevel = Nothing
                    Set oPCUnk = Nothing
                    Set oMfgBevel = Nothing
                    GoTo StartAgain
               End If
            End If
        End If
    Next i
    
    If (oGeomElems.Count > 1) Then
         
         Set oAllCSElems = New JObjectCollection
         
         For ind = 1 To oGeomElems.Count
            
             Set oThisGeom2d = oGeomElems.Item(ind)
             
             If ind = 1 Then
                 Set oGeom2dToUpdate = oThisGeom2d
                 Set oLastMfgGeomGeom2d = oGeomElems.Item(oGeomElems.Count)
                 oGeom2dToUpdate.PutEOC oLastMfgGeomGeom2d.GetEOC
                 Set oLastMfgGeomGeom2d = Nothing
             Else
                 oGeomCol2d.RemoveGeometry oThisGeom2d
                 Set oIJDObject = oThisGeom2d
                 oIJDObject.Remove
                 GeomCollection.Remove (GeomCollection.GetIndex(oThisGeom2d))
             End If
             Set oGeomCS = oThisGeom2d.GetGeometry
             'oGeomCS.GetCurves oCSElems
             oAllCSElems.Add oGeomCS
         Next ind

         'Set oFittedBspCurve = oMfgGeomHelper.ApproximateComplexStringByStrokeAndFit(oAllCSElems, False, 3, 0.1, 0.0001)
         
         Set oFittedCS = oMfgGeomHelper.OptimizedMergingOfInputCurves(oAllCSElems).Item(1)
         'oFittedCS.AddCurve oFittedBspCurve, True
         'oFittedCS.SetCurves oAllCSElems
         
         oGeom2dToUpdate.PutGeometry oFittedCS
         oGeomElems.Clear
    End If
    
End Function
Public Function BevelsValuesAreDifferent(ByRef oMfgBevel1 As IJMfgBevel, ByRef oMfgBevel2 As IJMfgBevel) As Boolean
    BevelsValuesAreDifferent = True
    Dim oMfgBevelParams1 As IJMfgBevelDetailProperties
    Dim oMfgBevelParams2 As IJMfgBevelDetailProperties
    
    Set oMfgBevelParams1 = oMfgBevel1
    Set oMfgBevelParams2 = oMfgBevel2
    
    If (Abs(oMfgBevelParams1.Angle1_M - oMfgBevelParams2.Angle1_M) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Angle1_UM - oMfgBevelParams2.Angle1_UM) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Angle2_M - oMfgBevelParams2.Angle2_M) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Angle2_UM - oMfgBevelParams2.Angle2_UM) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.ChamferAngle_M - oMfgBevelParams2.ChamferAngle_M) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.ChamferAngle_UM - oMfgBevelParams2.ChamferAngle_UM) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.ChamferDepth_M - oMfgBevelParams2.ChamferDepth_M) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.ChamferDepth_UM - oMfgBevelParams2.ChamferDepth_UM) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Depth1_M - oMfgBevelParams2.Depth1_M) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Depth1_UM - oMfgBevelParams2.Depth1_UM) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Depth2_M - oMfgBevelParams2.Depth2_M) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Depth2_UM - oMfgBevelParams2.Depth2_UM) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.Nose - oMfgBevelParams2.Nose) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.NoseAngle - oMfgBevelParams2.NoseAngle) > 0.001) Then
        Exit Function
    End If
    
    If (Abs(oMfgBevelParams1.RootGap - oMfgBevelParams2.RootGap) > 0.001) Then
        Exit Function
    End If
    
    BevelsValuesAreDifferent = False
    
End Function
Public Function IsValueWholeNumber(ByRef dDblValue, ByRef dValDifference) As Boolean

    Dim dFloor As Double
    
    dValDifference = 0#
    
    dFloor = CDbl(CLng(dDblValue * 1000#))
    dFloor = dFloor / 1000#
    
    If (Abs(dDblValue - dFloor) < 0.00001) Then
        IsValueWholeNumber = True
        Exit Function
    End If
    
    dFloor = CDbl(CLng(dDblValue * 1000# - 0.5))
    dFloor = dFloor / 1000#
    
    If (Abs(dDblValue - dFloor) > 0.00001) Then
        dValDifference = dDblValue - dFloor
        IsValueWholeNumber = False
    Else
        IsValueWholeNumber = True
    End If
    
End Function
Public Function RoundUpBevel(ByRef T1 As Double, ByRef T2 As Double)
    Dim dDiff1 As Double, dDiff2 As Double
    If (IsValueWholeNumber(T1, dDiff1) = False) Then
        If (IsValueWholeNumber(T2, dDiff2)) Then
            T1 = T1 - dDiff1 + 0.001
        Else
            T1 = T1 - dDiff1
            T2 = T2 - dDiff2 + 0.001
        End If
    End If
    
End Function
Public Function AdjustBevelParams(ByRef GeomCollection As IMSCoreCollections.IJElements)

    Dim i As Integer
    Dim T1 As Double, T2 As Double, T3 As Double, T4 As Double, T5 As Double
    Dim dVerySmallVal As Double
    
    dVerySmallVal = 0.00001
    
    For i = 1 To GeomCollection.Count
        Dim oGeom2d As IJMfgGeom2d
        Dim oMfgBevel As IJMfgBevelDetailProperties

        Set oGeom2d = GeomCollection.Item(i)

        If (Not ((oGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR) Or _
         (oGeom2d.GetGeometryType = STRMFG_INNER_CONTOUR))) Then GoTo NextContourEdge
         
        Set oMfgBevel = oGeom2d.GetBevel
        If oMfgBevel Is Nothing Then
            GoTo NextContourEdge
        End If
        
        If oMfgBevel.ChamferDepth_M > dVerySmallVal Then
            T1 = oMfgBevel.ChamferDepth_M
        Else
            T1 = 0#
        End If
        
        If oMfgBevel.Depth1_M > dVerySmallVal Then
            T2 = oMfgBevel.Depth1_M
        Else
            T2 = 0#
        End If
        
        If oMfgBevel.Nose > dVerySmallVal Then
            T3 = oMfgBevel.Nose
        Else
            T3 = 0#
        End If
        
        If oMfgBevel.Depth1_UM > dVerySmallVal Then
            T4 = oMfgBevel.Depth1_UM
        Else
            T4 = 0#
        End If
        
        If oMfgBevel.ChamferDepth_UM > dVerySmallVal Then
            T5 = oMfgBevel.ChamferDepth_UM
        Else
            T5 = 0#
        End If
        
        If (T1 > dVerySmallVal) Then
            If (T2 > dVerySmallVal) Then
                RoundUpBevel T1, T2
            ElseIf (T3 > dVerySmallVal) Then
                RoundUpBevel T1, T3
            ElseIf (T4 > dVerySmallVal) Then
                RoundUpBevel T1, T4
            End If
        End If
                
        If (T5 > dVerySmallVal) Then
            If (T4 > dVerySmallVal) Then
                RoundUpBevel T5, T4
            ElseIf (T3 > dVerySmallVal) Then
                RoundUpBevel T5, T3
            ElseIf (T2 > dVerySmallVal) Then
                RoundUpBevel T5, T2
            End If
        End If
        
        If (T2 > dVerySmallVal) Then
            If (T3 > dVerySmallVal) Then
                RoundUpBevel T2, T3
            ElseIf (T4 > dVerySmallVal) Then
                RoundUpBevel T2, T4
            End If
        End If
        
        If (T4 > dVerySmallVal) Then
            If (T3 > dVerySmallVal) Then
                RoundUpBevel T4, T3
            ElseIf (T2 > dVerySmallVal) Then
                RoundUpBevel T4, T2
            End If
        End If
        
        oMfgBevel.ChamferDepth_M = T1
        oMfgBevel.Depth1_M = T2
        oMfgBevel.Nose = T3
        oMfgBevel.Depth1_UM = T4
        oMfgBevel.ChamferDepth_UM = T5
        
NextContourEdge:
    Next i
    
End Function
