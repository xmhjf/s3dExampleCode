VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LapConnectionMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Lap Connection Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'       TBH         feb. 6. 2002    created
'       KONI        june 22 2002    Tested OK
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.LapConnectionMark"
Private Const REMOVEINNERCONTOUR As Boolean = False
Private Const NEEDCROSSMARKS As Boolean = False

Private Const dLengthCondition = 0.15
Private Const dExtendLength = 0.05

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "LapConnectionMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_LAP_MARK, ReferenceObjColl, oGeomCol3d

    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    'Get the Plate Part Physically Connected Objects
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    If oConObjsCol Is Nothing Then
        'Since there is no connecting structure you can leave the marking rule
        GoTo CleanUp
    End If

    Dim oMfgPart As IJMfgPlatePart
    If (oPlateWrapper.PlateHasMfgPart(oMfgPart) = False) Then
        Exit Function
    End If
    
    Dim oPlateCreation_AE As IJMfgPlateCreation_AE
    Dim oNeutralSurface As IJDModelBody
    
    Set oPlateCreation_AE = oMfgPart.ActiveEntity
    Set oNeutralSurface = oPlateCreation_AE.NeutralSurface
    Dim dDummyLength As Double
    Dim dEstRelAccyAchieved As Double
    Dim dNSurfaceArea As Double
    Dim dDummyVolume As Double

    oNeutralSurface.GetDimMetrics 0.001, dEstRelAccyAchieved, dDummyLength, dNSurfaceArea, dDummyVolume
    
    Dim Item As Object
    Dim oConnectionData As ConnectionData
    
    Dim nIndex As Long, nWBIndex As Long
    'Initialize the profile wrapper and the Physical Connection wrapper
    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Set oSDConPlateWrapper = New StructDetailObjects.PlatePart
        
    ' Loop thru each Physical Connections
    Dim bContourLap As Boolean
    Dim oWBColl As Collection
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D
    Dim lGeomCount As Long
    Dim oMfgGeomHelper As New MfgGeomHelper
    
    lGeomCount = 1
    
    Dim oConnPartColl As IJElements
    Set oConnPartColl = New JObjectCollection
                
    For nIndex = 1 To oConObjsCol.Count
        
        oConnectionData = oConObjsCol.Item(nIndex)
        
        ' When the physical connection geometry has multiple lumps or disjointed geometries, child PCs are created
        ' Example: Lapped plate with sketch feature at center
        ' In the above case, Make sure parent PC is not processed
        Dim oPhyConnSys As IJSystem
        Set oPhyConnSys = oConnectionData.AppConnection
        
        If Not oPhyConnSys Is Nothing Then
            Dim oChildPCColl As IJDTargetObjectCol
            Set oChildPCColl = oPhyConnSys.GetChildren
            
            If Not oChildPCColl Is Nothing Then
                If oChildPCColl.Count > 0 Then
                    GoTo NextItem
                End If
            End If
        End If
        
        'For avoiding duplication of Lapped marks in the Mfg Plate Output
        If oConnPartColl.Contains(oConnectionData.ToConnectable) Then
            GoTo NextItem
        Else
            oConnPartColl.Add oConnectionData.ToConnectable
        End If
        
        ' As both Bracket and Collar are implementing IJPlatePart we need to check if the
        ' conneected item is not either of those
        Dim oRootPartSystem         As IJSystem
        Dim oStructDetailHelper     As GSCADStructDetailUtil.StructDetailHelper
        Dim oPlateUtils             As IJPlateAttributes
        
        Set oPlateUtils = New PlateUtils
        Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
        oStructDetailHelper.IsPartDerivedFromSystem Part, oRootPartSystem, True
        
        If Not oRootPartSystem Is Nothing Then
            If oPlateUtils.IsBracketByPlane(oRootPartSystem) Then
                GoTo NextItem
            End If
        End If
        
        If TypeOf oConnectionData.ToConnectable Is IJCollarPart Then
             GoTo NextItem
        End If
        
        If Not TypeOf oConnectionData.ToConnectable Is IJPlatePart And _
            Not TypeOf oConnectionData.ToConnectable Is ISPSMemberPartPrismatic Then
             GoTo NextItem
        End If
        
        If TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
            Set oSDConPlateWrapper.object = oConnectionData.ToConnectable
        End If
        
        bContourLap = oSDPlateWrapper.Connection_ContourLap(oConnectionData.AppConnection, oWBColl)
        
        If ((bContourLap = True) And Not (oWBColl Is Nothing)) Then
            If oWBColl.Count = 0 Then
                GoTo NextItem
            End If
            
            'Convert the IJWireBody to a IJComplexString
            'For nWBIndex = 1 To 1
            'Fix for TR 54242 Lapped Plate is incorrect in xml file
            'Loop through all Wirebodies.
            'Suresh S.V.  25 March 2004
            
            If NEEDCROSSMARKS = True Then
                Dim oMfgMGHelper As MfgMGHelper
                Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
                
                Dim oCurveElements As IJElements
                Set oCurveElements = New JObjectCollection
                
                'Logic for Converting all the Items in the WireBody Collection to Complex Strings
                Dim lCount As Long
                For lCount = 1 To oWBColl.Count
                    Dim oTempElements As IJElements
                    oMfgMGHelper.WireBodyToComplexStrings oWBColl.Item(lCount), oTempElements
                    oCurveElements.AddElements oTempElements
                Next
                
                'Logic For Getting Minimum Bounding Box on a 3D object
            
                '                       ________________
                '                      /.              /'
                '                     /_._____________/2'
                '                     ' .             ' '
                '                     ' .3............'.'
                '                     '/______________'/
                '                     0               1
                '                    Minimum Boynding Box
    
                Dim oCOG As IJDPosition
                Dim oProjectionPoint As IJDPosition
                Dim oProjectionVector As IJDVector
                Dim oConnectionPartPort As IJPort
                
                ' Get the lapped plate port and it's normal at projected COG
                Set oConnectionPartPort = oConnectionData.ToConnectedPort
                
                Dim oConnSurfaceBody As IJSurfaceBody
                Set oConnSurfaceBody = oConnectionPartPort.Geometry
                oConnSurfaceBody.GetCenterOfGravity oCOG
                oMfgMGHelper.ProjectPointOnSurfaceBody oConnSurfaceBody, oCOG, oProjectionPoint, oProjectionVector
                
                Dim oSurfaceNormal As IJDVector
                oConnSurfaceBody.GetNormalFromPosition oProjectionPoint, oSurfaceNormal
                
                Dim oBoxPoints As IJElements
                Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVector(oCurveElements, oSurfaceNormal)
                
                Dim oPoints(1 To 4) As IJDPosition
                Set oPoints(1) = oBoxPoints.Item(1)
                Set oPoints(2) = oBoxPoints.Item(2)
                Set oPoints(3) = oBoxPoints.Item(3)
                Set oPoints(4) = New DPosition
    
                oPoints(4).x = oPoints(1).x + oPoints(3).x - oPoints(2).x
                oPoints(4).y = oPoints(1).y + oPoints(3).y - oPoints(2).y
                oPoints(4).z = oPoints(1).z + oPoints(3).z - oPoints(2).z
    
                Dim dLength(1 To 2) As Double
                dLength(1) = oPoints(1).DistPt(oPoints(2))
                dLength(2) = oPoints(2).DistPt(oPoints(3))
    
                '************** For Creating Mark if the length is less than 150 mm *************
    
                If (dLength(1) <= dLengthCondition) And (dLength(2) <= dLengthCondition) Then
                    If (dLength(1) - dLength(2)) < 0.001 Then
                        If IsALine(oCurveElements.Item(1)) = False Then
                            
                            Dim oBoxVector As IJDVector
                            Set oBoxVector = oPoints(1).Subtract(oPoints(2))
                            oBoxVector.Length = 1
    
                            Dim oXVector As IJDVector, oYVector As IJDVector, oZVector As IJDVector
                            Set oXVector = New DVector
                            Set oYVector = New DVector
                            Set oZVector = New DVector
    
                            oXVector.Set 1, 0, 0
                            oYVector.Set 0, 1, 0
                            oZVector.Set 0, 0, 1
    
                            If Not (Abs(oBoxVector.Dot(oXVector)) > 0.999 Or Abs(oBoxVector.Dot(oYVector)) > 0.999 Or Abs(oBoxVector.Dot(oZVector)) > 0.999) Then
                                Dim oVecElems As IJElements
                                Set oVecElems = New JObjectCollection
    
                                oVecElems.Add oXVector
                                oVecElems.Add oYVector
                                oVecElems.Add oZVector
    
                                Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVectors(oCurveElements, oVecElems)
                                Set oPoints(1) = oBoxPoints.Item(1)
                                Set oPoints(2) = oBoxPoints.Item(2)
                                Set oPoints(3) = oBoxPoints.Item(3)
    
                                oPoints(4).x = oPoints(1).x + oPoints(3).x - oPoints(2).x
                                oPoints(4).y = oPoints(1).y + oPoints(3).y - oPoints(2).y
                                oPoints(4).z = oPoints(1).z + oPoints(3).z - oPoints(2).z
                           End If
                        End If
                    End If
    
                    Dim oPartPort As IJPort
                    Set oPartPort = oConnectionData.ConnectingPort
    
                    Dim oSurfaceBody As IJSurfaceBody
                    Set oSurfaceBody = oPartPort.Geometry
    
                    Dim oCrossLineColl As IJElements
                    Set oCrossLineColl = GetCrossLinesCollection(oPoints, oSurfaceBody)
                    
                    Dim oCS1 As IJComplexString
                    Dim oCS2 As IJComplexString
     
                    Set oCS1 = oCrossLineColl.Item(1)
                    Set oCS2 = oCrossLineColl.Item(2)
                    
                    Dim oMfgGeomUtilWrapper As New MfgGeomUtilWrapper
                    oMfgGeomUtilWrapper.ExtendWire oCS1, dExtendLength
                    oMfgGeomUtilWrapper.ExtendWire oCS2, dExtendLength
                    
                    Dim oCrossLine1 As IJCurve
                    Set oCrossLine1 = oCS1
                    Dim oCrossLine2 As IJCurve
                    Set oCrossLine2 = oCS2
                    
                    
                    Dim jCount As Long
                    For jCount = 1 To 2
                        'Create a SystemMark object to store additional information
                        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                        
                        'Set the marking side
                        oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
                        
                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfo = oSystemMark
                        
                        oMarkingInfo.Thickness = oSDConPlateWrapper.PlateThickness
                        
                        Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                        
                        If jCount = 1 Then
                            oGeom3d.PutGeometry oCS1
                            
                           If oCrossLine1.Length >= oCrossLine2.Length Then
                                oMarkingInfo.Name = oSDConPlateWrapper.Name
                            End If
                        
                        Else
                            oGeom3d.PutGeometry oCS2
                            If oCrossLine2.Length > oCrossLine1.Length Then
                                oMarkingInfo.Name = oSDConPlateWrapper.Name
                            End If
                            
                        End If
                        
                        oGeom3d.PutGeometrytype STRMFG_LAP_REF_MARK
                        oSystemMark.Set3dGeometry oGeom3d
                        
                        Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                        oGeom3d.PutMoniker oMoniker
                        oGeomCol3d.AddGeometry 1, oGeom3d
                        
                        Set oGeom3d = Nothing
                    Next
                    
                    GoTo NextItem
                    
                Else
                    'the below code is written for cases where inner contour of the Lapped Plate is to be removed(special cases like U, J)
                    If REMOVEINNERCONTOUR = True Then
                        Dim oOuterWBColl As Collection
                        Set oOuterWBColl = GetOuterWBCollection(oPoints, oConnSurfaceBody, oWBColl)
                        
                        If oOuterWBColl.Count = 0 Then
                            GoTo CleanUp
                        Else
                            Set oWBColl = Nothing
                            Set oWBColl = oOuterWBColl
                        End If
                    End If
                   
                End If
                
                Erase oPoints
                
            End If
            
            Dim oWireBodyColl As IJElements
            Set oWireBodyColl = New JObjectCollection
            
            For nWBIndex = 1 To oWBColl.Count
                Set oWB = oWBColl.Item(nWBIndex)
                oWireBodyColl.Add oWB
            Next nWBIndex
            
            Dim oCSColl As IJElements
            Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWireBodyColl)
            
            If oCSColl Is Nothing Then
                GoTo NextItem
            Else
                If oCSColl.Count = 0 Then
                    Set oCSColl = Nothing
                    GoTo NextItem
                End If
            End If
            
            For Each oCS In oCSColl
                Dim oCSCrv As IJCurve
                Set oCSCrv = oCS
                If (Abs(oCSCrv.Area - dNSurfaceArea) > 0.0001) Then ' TR#261394
                
                    'Create a SystemMark object to store additional information
                    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    
                    'Set the marking side
                    oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)
                    
                    'QI for the MarkingInfo object on the SystemMark
                    Set oMarkingInfo = oSystemMark
                    
                    If TypeOf oConnectionData.ToConnectable Is IJPlatePart Then
                        oMarkingInfo.Thickness = oSDConPlateWrapper.PlateThickness
                        oMarkingInfo.Name = oSDConPlateWrapper.Name
                    End If
                    
                    Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    oGeom3d.PutGeometry oCS
                    
                    If REMOVEINNERCONTOUR = True Then
                        ' here we need toshi mark so give geom type as STRMFG_LAP_TRACE_MARK
                        oGeom3d.PutGeometrytype STRMFG_LAP_TRACE_MARK
                    Else
                        oGeom3d.PutGeometrytype STRMFG_LAP_MARK
                    End If
                    
                    oGeomCol3d.AddGeometry 1, oGeom3d
                    oSystemMark.Set3dGeometry oGeom3d
                    
                    Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                    oGeom3d.PutMoniker oMoniker
                    
                    Set oSystemMark = Nothing
                    Set oMarkingInfo = Nothing
                    Set oGeom3d = Nothing
                End If
            Next
        End If
              
NextItem:
        Set oWB = Nothing
        Set oWireBodyColl = Nothing
        Set oCSColl = Nothing
        Set oCS = Nothing
        Set oSDConPlateWrapper = Nothing
    Next nIndex
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Set oCS = Nothing
    Set oConnPartColl = Nothing
    
CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1012, , "RULES")
    GoTo CleanUp
End Function

''**************************************************************************************
'' Routine      : GetCrossLinesCollection
'' Abstract     : This Function gets the Crosslines of the minimumbounding box points being
''                the input
''
''**************************************************************************************
Private Function GetCrossLinesCollection(oPoints() As IJDPosition, oSurfaceBody As IJSurfaceBody) As IJElements
    Const METHOD = "GetCrosslinesCollection"
    On Error GoTo ErrorHandler
    Dim oCurveElements As IJElements
    Set oCurveElements = New JObjectCollection

        '  For creating the Marking Line
        
        '                    _______.3______
        '                   |               |
        '                   |               |
        '                   .4              .2          1,2,3,4 are cross points
        '                   |               |
        '                   |______.1_______|
        
        Dim oCrossPoint1 As IJDPosition
        Set oCrossPoint1 = New DPosition
        oCrossPoint1.x = (oPoints(1).x + oPoints(2).x) / 2
        oCrossPoint1.y = (oPoints(1).y + oPoints(2).y) / 2
        oCrossPoint1.z = (oPoints(1).z + oPoints(2).z) / 2
    
    
        Dim oCrossPoint2 As IJDPosition
        Set oCrossPoint2 = New DPosition
        oCrossPoint2.x = (oPoints(2).x + oPoints(3).x) / 2
        oCrossPoint2.y = (oPoints(2).y + oPoints(3).y) / 2
        oCrossPoint2.z = (oPoints(2).z + oPoints(3).z) / 2
    
    
        Dim oCrossPoint3 As IJDPosition
        Set oCrossPoint3 = New DPosition
        oCrossPoint3.x = (oPoints(3).x + oPoints(4).x) / 2
        oCrossPoint3.y = (oPoints(3).y + oPoints(4).y) / 2
        oCrossPoint3.z = (oPoints(3).z + oPoints(4).z) / 2
    
        Dim oCrossPoint4 As IJDPosition
        Set oCrossPoint4 = New DPosition
        oCrossPoint4.x = (oPoints(1).x + oPoints(4).x) / 2
        oCrossPoint4.y = (oPoints(1).y + oPoints(4).y) / 2
        oCrossPoint4.z = (oPoints(1).z + oPoints(4).z) / 2
    
        Dim oCrossLine1 As IJLine
        Dim oCrossLine2 As IJLine
        Set oCrossLine1 = New Line3d
        Set oCrossLine2 = New Line3d
    
        oCrossLine1.DefineBy2Points oCrossPoint1.x, oCrossPoint1.y, oCrossPoint1.z, oCrossPoint3.x, oCrossPoint3.y, oCrossPoint3.z
        oCrossLine2.DefineBy2Points oCrossPoint2.x, oCrossPoint2.y, oCrossPoint2.z, oCrossPoint4.x, oCrossPoint4.y, oCrossPoint4.z
                            
        Dim oCS1 As IJComplexString
        Dim oCS2 As IJComplexString
        Set oCS1 = New ComplexString3d
        Set oCS2 = New ComplexString3d
    
        oCS1.AddCurve oCrossLine1, True
        oCS2.AddCurve oCrossLine2, True
    
        Dim oWireBody1 As IJWireBody
        Set oWireBody1 = m_oMfgRuleHelper.ComplexStringToWireBody(oCS1)
    
        Dim oWireBody2 As IJWireBody
        Set oWireBody2 = m_oMfgRuleHelper.ComplexStringToWireBody(oCS2)
    
        Set oCS1 = Nothing
        Set oCS2 = Nothing
        
        Dim oProjectionVector As IJDVector
        
        Set oCS1 = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurfaceBody, oWireBody1, oProjectionVector)
        Set oCS2 = m_oMfgRuleHelper.CurveAlongVectorOnToSurface(oSurfaceBody, oWireBody2, oProjectionVector)
        
        Dim oCScoll1 As IJElements
        Set oCScoll1 = New JObjectCollection
        oCScoll1.Add oCS1
        oCScoll1.Add oCS2
        
        Set GetCrossLinesCollection = oCScoll1
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

''**************************************************************************************
'' Routine      : GetOuterWBCollection
'' Abstract     : This Function gets the Crosslines of the minimumbounding box points being
''                the input
''
''**************************************************************************************
Private Function GetOuterWBCollection(oPoints() As IJDPosition, oConnSurfaceBody As IJSurfaceBody, oWBColl As Collection) As Collection
    Const METHOD = "GetOuterWBCollection"
    On Error GoTo ErrorHandler
    
    Dim oMfgMGHelper As MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oCurveElements As IJElements
    Set oCurveElements = New JObjectCollection
    
    Dim oTempWBColl As Collection
    Set oTempWBColl = New Collection
    
    Dim lCount As Long
    For lCount = 1 To oWBColl.Count
        Dim oTempElements As IJElements
        oMfgMGHelper.WireBodyToComplexStrings oWBColl.Item(lCount), oTempElements
        oCurveElements.AddElements oTempElements
        oTempWBColl.Add oWBColl.Item(lCount)
    Next
    
    Dim dMinDist As Double
                
    ' Get the Box face(formed by points 1,2,3) mid point (ie., point between 1 and 3 or 2 and 4)
    Dim oMidPoint As IJDPosition
    Set oMidPoint = New DPosition
        
    oMidPoint.x = (oPoints(1).x + oPoints(3).x) / 2
    oMidPoint.y = (oPoints(1).y + oPoints(3).y) / 2
    oMidPoint.z = (oPoints(1).z + oPoints(3).z) / 2
    
    Dim oVector1 As IJDVector
    Dim oVector2 As IJDVector
    
    Set oVector1 = oPoints(2).Subtract(oPoints(1))
    Set oVector2 = oPoints(3).Subtract(oPoints(2))
    
    oVector1.Length = 1
    oVector2.Length = 1

    Dim oPlane1 As IJPlane
    Dim oPlane2 As IJPlane
                
    Set oPlane1 = CreatePlane(oMidPoint, oVector2)
    Set oPlane2 = CreatePlane(oMidPoint, oVector1)

    Dim oMfgRuleHelpers As MfgRuleHelpers.Helper
    Set oMfgRuleHelpers = New MfgRuleHelpers.Helper

    Dim oIntersectingWB As IJWireBody
    Set oIntersectingWB = GetIntersection(oConnSurfaceBody, oPlane1)

    Dim oIntCSColl As IJElements
    Set oIntCSColl = New JObjectCollection

    Dim oTempColl As IJElements
    Set oTempColl = oMfgRuleHelpers.WireBodyToComplexStrings(oIntersectingWB)
    oIntCSColl.AddElements oTempColl
    
    Set oIntersectingWB = Nothing
    Set oIntersectingWB = GetIntersection(oConnSurfaceBody, oPlane2)

    Set oTempColl = Nothing
    Set oTempColl = oMfgRuleHelpers.WireBodyToComplexStrings(oIntersectingWB)
    oIntCSColl.AddElements oTempColl
    
    Dim bSpecialShape  As Boolean
    bSpecialShape = False
                
    If (oIntCSColl.Count = 2) Then
        
        Dim oIntMB1 As IJDModelBody
        Set oIntMB1 = m_oMfgRuleHelper.ComplexStringToWireBody(oIntCSColl.Item(1))
        
        Dim oIntMB2 As IJDModelBody
        Set oIntMB2 = m_oMfgRuleHelper.ComplexStringToWireBody(oIntCSColl.Item(2))
        
        Dim oClosePos1 As IJDPosition, oClosePos2 As IJDPosition
        oIntMB1.GetMinimumDistance oIntMB2, oClosePos1, oClosePos2, dMinDist
        
        ' In this case if the distance between the wires are greater than 1 mm, treat it as spectal shape(J shape)
        If (dMinDist > 0.001) Then
            bSpecialShape = True
        End If
    ElseIf oIntCSColl.Count = 3 Then
        ' In this case(intersection CS count = 3), treat it as spectal shape(J shape)
        bSpecialShape = True
    End If
    
    If bSpecialShape = True Then
        
        Dim oInnerPointColl() As IJDPosition
        ReDim oInnerPointColl(1 To oIntCSColl.Count)
        
        For lCount = 1 To oIntCSColl.Count
        
            Dim oStartPoint As IJDPosition
            Dim oEndPoint As IJDPosition
    
            Set oStartPoint = New DPosition
            Set oEndPoint = New DPosition
            Set oInnerPointColl(lCount) = New DPosition
    
            Dim oTempCurve As IJCurve
            Set oTempCurve = oIntCSColl.Item(lCount)
    
            Dim dStartX As Double, dStartY As Double, dStartZ As Double
            Dim dEndX As Double, dEndY As Double, dEndZ As Double
            
            oTempCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
            oStartPoint.Set dStartX, dStartY, dStartZ
            oEndPoint.Set dEndX, dEndY, dEndZ
            
            ' Get the inerior point by checking the disance from box face mid point to a curve extreme points
            If oMidPoint.DistPt(oStartPoint) < oMidPoint.DistPt(oEndPoint) Then
                Set oInnerPointColl(lCount) = oStartPoint
            Else
                Set oInnerPointColl(lCount) = oEndPoint
            End If
        Next
                    
        For lCount = 1 To oIntCSColl.Count
            Dim lWBCount As Long
            Dim dTempDist As Double
            dTempDist = 100000# ' initialize with large number
            Dim lCurveToRemove As Long
            lCurveToRemove = 0
            
            For lWBCount = 1 To oTempWBColl.Count
                Dim oModelBody As IJDModelBody
                Set oModelBody = oTempWBColl.Item(lWBCount)
                Dim oClosestPos As IJDPosition
                oModelBody.GetMinimumDistanceFromPosition oInnerPointColl(lCount), oClosestPos, dMinDist
                
                If dMinDist < dTempDist And dMinDist < 0.001 Then
                    dTempDist = dMinDist
                    lCurveToRemove = lWBCount
                End If
            Next
            
            If lCurveToRemove <> 0 Then
                oTempWBColl.Remove lCurveToRemove
            End If
        Next
        
        Erase oInnerPointColl
    End If
    
    ' if bSpecialShape = True then return the oTempWBColl
    If bSpecialShape = True Then
        Set GetOuterWBCollection = oTempWBColl
    Else
        Set GetOuterWBCollection = oWBColl
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
