VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BevelChangeMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Buttock Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule2
Private Const PI As Double = 3.14159265358979
Private Const MODULE = "SKDY_MfgProfileMarking.BevelChangeMark"

Private Sub Class_Initialize()
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule2_CreateAtTheEnd(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "IJDMfgSystemMarkingRule2_CreateAtTheEnd"
    On Error GoTo ErrorHandler
    
    If bSelectiveRecompute Then
        Exit Function
    End If
    
    Dim dMarkLength As Double
    dMarkLength = BEVEL_MARK_TRANS_LENGTH

    '*** Create the Plate Wrapper Object ***'
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Dim oMfgPart As Object
    Dim oMfgProfilePart As IJMfgProfilePart
    
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    '*** Get the Manufactured Plate Part ***'
    If Not oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Exit Function
    Else
        Set oMfgProfilePart = oMfgPart
    End If

    Dim oGeomCol2d As IJMfgGeomCol2d
    Dim oPrevPC As IUnknown
    Dim oPrevMfgBevel As IUnknown
    Dim oPrevMfgGeom2d As IJMfgGeom2d
    Dim i As Long
    Dim oZVec As New DVector
    Dim oGeomTanVec As New DVector
    Dim oLines As ILines3d
    Dim oComplexStrings As IComplexStrings3d
    Dim oGeom2dColFactory As New MfgGeomCol2dFactory
    Dim oGeom2dCol As IJMfgGeomCol2d
    Dim lGeomCount As Long
    
    lGeomCount = 0
    Set oGeom2dCol = oGeom2dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
      
    Set oLines = New GeometryFactory
    Set oComplexStrings = New GeometryFactory
    
    oZVec.Set 0#, 0#, 1#
    
    Set oGeomCol2d = oMfgProfilePart.FinalGeometriesAfterProcess2D
    
    Dim pResMgr As IUnknown
    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If
    
    Dim oGeom2dFactory As New GSCADMfgGeometry.MfgGeom2dFactory
    
    For i = 1 To oGeomCol2d.Getcount
        Dim oGeom2d As IJMfgGeom2d
        Dim oMfgBevel As IJMfgBevel
        Dim oPCUnk As IUnknown
        Dim oLineVec As IJDVector
        
        Set oGeom2d = oGeomCol2d.GetGeometry(i)
        
        If (Not (oGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR)) Then GoTo NextContourEdge
         
        Set oMfgBevel = oGeom2d.GetBevel
        If oMfgBevel Is Nothing Then
            Set oPrevPC = Nothing
            Set oPrevMfgBevel = Nothing
            Set oPrevMfgGeom2d = Nothing
            GoTo NextContourEdge
        End If
        
        Set oPCUnk = oMfgBevel.PhysicalConnection
        If oPCUnk Is Nothing Then
            Set oPrevPC = Nothing
            Set oPrevMfgBevel = Nothing
            Set oPrevMfgGeom2d = Nothing
            GoTo NextContourEdge
        End If
        
        If Not oPrevMfgBevel Is Nothing Then
          If ((CheckIfBothBevelsBelongsToSameBoundingElement(Part, oMfgBevel, oPrevMfgBevel) = True) _
                And (BevelsValuesAreDifferent(oMfgBevel, oPrevMfgBevel))) Then
            Dim oGeomCS As IJComplexString
            Dim oGeomCurve As IJCurve
            Dim dStartParam As Double, dEndParam As Double, dLineStartX As Double, dLineStartY As Double, dLineStartZ As Double
            Dim dGeomVecX  As Double, dGeomVecY As Double, dGeomVecZ As Double, dDummy As Double
            Dim oCurveColl As IJElements
            Dim oLine As IJLine
            Dim oNewGeom2d As IJMfgGeom2d
            Dim oLineCS As IJComplexString
            Dim oCurveElems As IJElements
            Dim oCSCurve As IJCurve
            
            Set oGeomCS = oGeom2d.GetGeometry
            
            oGeomCS.GetCurves oCurveElems
            
            Set oGeomCurve = oGeomCS
            oGeomCurve.EndPoints dLineStartX, dLineStartY, dLineStartZ, dDummy, dDummy, dDummy
            
            For Each oCSCurve In oCurveElems
                If oCSCurve.length > 0.004 Then
                    Set oGeomCurve = oCSCurve
                    GoTo GotCurve
                End If
            Next
            
GotCurve:
            If oGeomCurve Is Nothing Then
                Set oGeomCurve = oGeomCS
            End If
            
            oGeomCurve.ParamRange dStartParam, dEndParam
            oGeomCurve.Evaluate dStartParam, dDummy, dDummy, dDummy, dGeomVecX, dGeomVecY, dGeomVecZ, dDummy, dDummy, dDummy
            oGeomTanVec.Set dGeomVecX, dGeomVecY, dGeomVecZ
            
            ' If the curve length is still small, take tagent at the end point of previous geometry
            If oGeomCurve.length < 0.004 Then
                If Not oPrevMfgGeom2d Is Nothing Then
                    Set oGeomCurve = oPrevMfgGeom2d.GetGeometry
                    oGeomCurve.ParamRange dStartParam, dEndParam
                    oGeomCurve.Evaluate dEndParam, dDummy, dDummy, dDummy, dGeomVecX, dGeomVecY, dGeomVecZ, dDummy, dDummy, dDummy
                    oGeomTanVec.Set dGeomVecX, dGeomVecY, dGeomVecZ
                End If
            End If
                        
            oGeomTanVec.length = 1#
            
            Set oLineVec = oZVec.Cross(oGeomTanVec)
            
            Set oCurveColl = New JObjectCollection
            Set oLine = oLines.CreateByPtVectLength(Nothing, dLineStartX, dLineStartY, dLineStartZ, _
                                    oLineVec.x, oLineVec.y, oLineVec.z, dMarkLength)
            oCurveColl.Add oLine
            
            Set oLineCS = oComplexStrings.CreateByCurves(Nothing, oCurveColl)
            
            Set oNewGeom2d = oGeom2dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            oNewGeom2d.PutGeometry oLineCS
            oNewGeom2d.PutGeometrytype STRMFG_BEVEL_MARK
            
            lGeomCount = lGeomCount + 1
            oGeom2dCol.AddGeometry lGeomCount, oNewGeom2d
            Set oGeomCS = Nothing
            Set oCurveElems = Nothing
          End If
        End If
        
        Set oPrevPC = oPCUnk
        Set oPrevMfgBevel = oMfgBevel
        Set oPrevMfgGeom2d = oGeom2d
        Set oLineVec = Nothing
        Set oPCUnk = Nothing
        Set oMfgBevel = Nothing
NextContourEdge:
    Next i
    
CleanUp:
    Set IJDMfgSystemMarkingRule2_CreateAtTheEnd = oGeom2dCol
    Set oLines = Nothing
    Set oComplexStrings = Nothing
    Set oGeom2dCol = Nothing
    Set pResMgr = Nothing
    Set oMfgProfilePart = Nothing
    Set oPrevPC = Nothing
    Set oPrevMfgGeom2d = Nothing
    Set oMfgBevel = Nothing
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1010, , "RULES")
    GoTo CleanUp
End Function
Private Function BevelsValuesAreDifferent(ByRef oMfgBevel1 As IJMfgBevel, ByRef oMfgBevel2 As IJMfgBevel) As Boolean
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


