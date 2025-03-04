VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamCntrlMark2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Seam Control Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      StructMfg        09.09.2010      Creation
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule
Private Const dSeamOffset = 0.15
Private Const dOverlapDist = 0.015
Private Const dSeamOffsetIfOverlaps = 0.1
Private Const PI As Double = 3.14159265358979

Private Const MODULE = "MfgPlateMarking.SeamCntrlMark2D"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub
Private Sub ConstructSeamMark(ByRef oCSElems As IJElements, _
                                ByVal UpSide As Long, _
                                ByRef pResMgr As IUnknown, _
                                ByVal oPC As IUnknown, _
                                ByVal oGeomCol2d As IJMfgGeomCol2d, _
                                ByVal dTotalMargin As Double, _
                                ByRef oOutputGeomCol2d As IJMfgGeomCol2d)
    If ((oCSElems.Count = 0) Or (oPC Is Nothing)) Then
        Exit Sub
    End If
    Dim dMinDist As Double, dPar As Double, dDummy As Double
    Dim dEdgeX As Double, dEdgeY As Double, dEdgeZ As Double
    Dim dMarkX As Double, dMarkY As Double, dMarkZ As Double
    Dim dSeamTanX As Double, dSeamTanY As Double, dSeamTanZ As Double
    Dim dMarkTanX As Double, dMarkTanY As Double, dMarkTanZ As Double
    Dim dMarkWithinOverlapDist As Boolean
    Dim oComplexStr As New ComplexString3d
    
    oComplexStr.SetCurves oCSElems
    
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
    Set oSDPhysicalConn.object = oPC

    Dim dRootGap As Double

    oSDPhysicalConn.GetBevelParameterValue "RootGap", dRootGap, DoubleType
     
    Dim oSeamEdgeCurve As IJCurve
    Set oSeamEdgeCurve = oComplexStr
    
    ' If the length of the CS is less than 1 mm then skip that curve
    If oSeamEdgeCurve.Length < 0.001 Then
        Exit Sub
    End If
    
    Dim indGeom As Long, lGeomCount As Long
    lGeomCount = oGeomCol2d.GetCount
    dMarkWithinOverlapDist = False
    For indGeom = 1 To lGeomCount
        Dim oMfgGeom2d As IJMfgGeom2d
        Dim oMarkCS As IJComplexString
        Dim oMarkCurve As IJCurve
        Set oMfgGeom2d = oGeomCol2d.GetGeometry(indGeom)
        If ((oMfgGeom2d.GetGeometryType <> STRMFG_OUTER_CONTOUR) And _
            (oMfgGeom2d.GetGeometryType <> STRMFG_INNER_CONTOUR)) Then
            Set oMarkCS = oMfgGeom2d.GetGeometry
            Set oMarkCurve = oMarkCS
            oSeamEdgeCurve.DistanceBetween oMarkCurve, dMinDist, dEdgeX, dEdgeY, dEdgeZ, dMarkX, dMarkY, dMarkZ
            If ((dMinDist >= (dSeamOffset - dOverlapDist)) And (dMinDist <= (dSeamOffset + dOverlapDist))) Then
                oSeamEdgeCurve.Parameter dEdgeX, dEdgeY, dEdgeZ, dPar
                oSeamEdgeCurve.Evaluate dPar, dDummy, dDummy, dDummy, dSeamTanX, dSeamTanY, dSeamTanZ, dDummy, dDummy, dDummy
                
                oMarkCurve.Parameter dMarkX, dMarkY, dMarkZ, dPar
                oMarkCurve.Evaluate dPar, dDummy, dDummy, dDummy, dMarkTanX, dMarkTanY, dMarkTanZ, dDummy, dDummy, dDummy
                
                If (Abs(dSeamTanX * dMarkTanX + dSeamTanY * dMarkTanY + dSeamTanZ * dMarkTanZ) > Sin(PI / 4#)) Then
                    dMarkWithinOverlapDist = True
                End If

            End If
        End If
    Next
    
    Dim oMfgMathGeom As New MfgMathGeom
    
    If dMarkWithinOverlapDist Then
        oMfgMathGeom.ApplyOffset oComplexStr, (dTotalMargin - dRootGap - dSeamOffsetIfOverlaps), (dTotalMargin - dRootGap - dSeamOffsetIfOverlaps)
    Else
        oMfgMathGeom.ApplyOffset oComplexStr, (dTotalMargin - dRootGap - dSeamOffset), (dTotalMargin - dRootGap - dSeamOffset)
    End If
       
    'Create in systemmark object and add to output-collection
    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom2d As IJMfgGeom2d
    Set oGeom2d = m_oGeom2dFactory.Create(pResMgr)
                    
    'Create a SystemMark object to store additional information
    Set oSystemMark = m_oSystemMarkFactory.Create(pResMgr)
    'Set the marking side
    oSystemMark.SetMarkingSide UpSide '(should be connecting side)

    'QI for the MarkingInfo object on the SystemMark
    Set oMarkingInfo = oSystemMark

    oMarkingInfo.Name = "SEAM CONTROL"

    oSystemMark.Set2dGeometry oGeom2d
    oGeom2d.PutGeometry oComplexStr
    
    oGeom2d.TrimToBoundaries = True
    
    Set oComplexStr = Nothing
    oGeom2d.PutGeometrytype STRMFG_SEAM_MARK
    
    If dMarkWithinOverlapDist Then
        'oGeom2d.PutSubGeometryType STRMFG_SEAM_MARK
        oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "SeamControl"
    End If

    oGeom2d.PutMoniker m_oMfgRuleHelper.GetMoniker(oPC)
    If oOutputGeomCol2d Is Nothing Then
        Set oOutputGeomCol2d = m_oGeom2dColFactory.Create(pResMgr)
    End If
    oOutputGeomCol2d.AddGeometry 1, oGeom2d

    
End Sub
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
Const METHOD = "SeamControlMark: IJDMfgSystemMarkingRule_CreateAfterUnfold"
On Error GoTo ErrorHandler
        
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part
    
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    Dim oGeomCol2d As IJMfgGeomCol2d
    Dim lGeomCount As Long, indGeom As Long
    Dim oPrevPC As IUnknown
    Dim oNewCSElems As IJElements
    Dim dTotalMargin As Double
    
    Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries

    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If
    
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    
    
    Dim pResMgr As IUnknown
    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    lGeomCount = oGeomCol2d.GetCount
    
    Set oNewCSElems = New JObjectCollection
    
    For indGeom = 1 To lGeomCount
        Dim oMfgGeom2d As IJMfgGeom2d
        Dim oGeomCS As IJComplexString
        Dim oCSElems As IJElements
        Dim oMfgBevel As IJMfgBevel
        Dim oMoniker As IMoniker
        Dim oPortObject As IUnknown
        Dim oPortOperator As IUnknown
        Dim oSeamType As IJDSeamType
            
        Set oMfgGeom2d = oGeomCol2d.GetGeometry(indGeom)
        If oMfgGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR Then
        
            Set oMoniker = oMfgGeom2d.GetMoniker
            If oMoniker Is Nothing Then GoTo NextContourEdge
            
            Set oPortObject = m_oMfgRuleHelper.BindToObject(pResMgr, oMoniker)
            
            Dim oPort As IJPort
            Set oPort = oPortObject
            
            Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
            Set oSDPartSupport.Part = oPort.Connectable
            oSDPartSupport.GetPortOperator oPortObject, oPortOperator
            If Not oPortOperator Is Nothing Then
                If (TypeOf oPortOperator Is IJDSeamType) Then
                    Set oSeamType = oPortOperator
                    
                    'the seam type we are interested in is a planning seam
                    If oSeamType.SeamType = sstPlanningSeam Then
                        ConstructSeamMark oNewCSElems, UpSide, pResMgr, oPrevPC, oGeomCol2d, dTotalMargin, IJDMfgSystemMarkingRule_CreateAfterUnfold
                        oNewCSElems.Clear
                        GoTo NextContourEdge
                    End If
                ElseIf (TypeOf oPortOperator Is IJStructFeature) Then
                    ConstructSeamMark oNewCSElems, UpSide, pResMgr, oPrevPC, oGeomCol2d, dTotalMargin, IJDMfgSystemMarkingRule_CreateAfterUnfold
                    oNewCSElems.Clear
                    GoTo NextContourEdge
                End If
            Else
                ConstructSeamMark oNewCSElems, UpSide, pResMgr, oPrevPC, oGeomCol2d, dTotalMargin, IJDMfgSystemMarkingRule_CreateAfterUnfold
                oNewCSElems.Clear
                GoTo NextContourEdge
            End If
            
            Set oMfgBevel = oMfgGeom2d.GetBevel
            If oMfgBevel Is Nothing Then
                If oNewCSElems.Count > 0 Then
                    ConstructSeamMark oNewCSElems, UpSide, pResMgr, oPrevPC, oGeomCol2d, dTotalMargin, IJDMfgSystemMarkingRule_CreateAfterUnfold
                    oNewCSElems.Clear
                    dTotalMargin = 0#
                End If
                Set oPrevPC = Nothing
            Else
                If Not oMfgBevel.PhysicalConnection Is oPrevPC Then
                    ConstructSeamMark oNewCSElems, UpSide, pResMgr, oPrevPC, oGeomCol2d, dTotalMargin, IJDMfgSystemMarkingRule_CreateAfterUnfold
                    oNewCSElems.Clear
                    dTotalMargin = 0#
                End If
                Set oGeomCS = oMfgGeom2d.GetGeometry
                oGeomCS.GetCurves oCSElems
                oNewCSElems.AddElements oCSElems
                Set oCSElems = Nothing
                            
                Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double
                lFabMargin = 0
                lAssyMargin = 0
                lCustomMargin = 0
                                
                Dim i As Integer
                Dim oConstMargin As IJConstMargin
                Dim oObliqueMargin As IJObliqueMargin
                
                Dim oMfgDefport As IJPort
                Dim pMfgDef As IJMfgDefinition
                For i = 1 To ReferenceObjColl.Count
                    If TypeOf ReferenceObjColl.Item(i) Is IJMfgDefinition Then
                        Set pMfgDef = ReferenceObjColl.Item(i)
                        Set oMfgDefport = pMfgDef.GetPort
                        'check if Margin exist on the same port where SeamControl mark is being created
                        If oPortObject Is oMfgDefport Then
                            If TypeOf pMfgDef Is IJDFabMargin Then
                                Dim oFabMargin As IJDFabMargin
                                Set oFabMargin = pMfgDef
                            
                                If oFabMargin.GeometryChange = AsMargin Then ''1 = As Margin, 2 = As Shrinkage, 3 = As Reference
                                    If TypeOf pMfgDef Is IJAssyMarginChild Then
                                        Set oConstMargin = pMfgDef
                                        lAssyMargin = lAssyMargin + oConstMargin.Value
                                    ElseIf TypeOf pMfgDef Is IJObliqueMargin Then
                                        Set oObliqueMargin = pMfgDef
                                        If oObliqueMargin.EndValue > oObliqueMargin.StartValue Then
                                            lFabMargin = lFabMargin + oObliqueMargin.EndValue
                                        Else
                                            lFabMargin = lFabMargin + oObliqueMargin.StartValue
                                        End If
                                    ElseIf TypeOf pMfgDef Is IJConstMargin Then
                                        Set oConstMargin = pMfgDef
                                        lFabMargin = lFabMargin + oConstMargin.Value
                                    'ElseIf TypeOf oMfgDefCol.Item(j) Is ??? Then
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
                If lAssyMargin <> 0 Or lFabMargin <> 0 Or lCustomMargin <> 0 Then
                     dTotalMargin = lAssyMargin + lFabMargin + lCustomMargin
                End If
                
                Set oPrevPC = oMfgBevel.PhysicalConnection
            End If
				Set oSDPartSupport = Nothing
        End If
NextContourEdge:
    Next indGeom
    
    ConstructSeamMark oNewCSElems, UpSide, pResMgr, oPrevPC, oGeomCol2d, dTotalMargin, IJDMfgSystemMarkingRule_CreateAfterUnfold
    oNewCSElems.Clear
    dTotalMargin = 0#
    
CleanUp:
    Set oSDPartSupport = Nothing
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1026, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
End Function

