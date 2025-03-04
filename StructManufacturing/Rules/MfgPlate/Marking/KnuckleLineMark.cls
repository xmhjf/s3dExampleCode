VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckleLineMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Bending Control Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      KONI     June 12 2002    created
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule
Private Const PI As Double = 3.14159265358979
Private Const MODULE = "MfgPlateMarking.KnuckleLineMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As JCmnShp_CollectionAlias) As IJMfgGeomCol2d
    Const METHOD = "IJDMfgSystemMarkingRule_CreateAfterUnfold"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1011, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As JCmnShp_CollectionAlias) As IJMfgGeomCol3d
    Const METHOD = "KnuckleLineMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_KNUCKLE_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    If bSelectiveRecompute Then
        Exit Function
    End If
    
    Dim oMfgPart As IJMfgPlatePart
    Dim oPlateWrapper As New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part
    Dim oNeutralSurface As IJSurfaceBody

    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Dim oMfgPlateCreation_AE As IJMfgPlateCreation_AE
        Set oMfgPlateCreation_AE = oMfgPart.ActiveEntity
        Set oNeutralSurface = oMfgPlateCreation_AE.NeutralSurface
        Set oMfgPlateCreation_AE = Nothing
    Else
        Exit Function
    End If
    Set oPlateWrapper = Nothing
        
    Dim oPlateSystem As IJPlateSystem
    Dim oPlateUtils As IJPlateAttributes
    Dim oKnuckleColl As IJElements
    Dim nIndex As Long
    Dim oGeom3d As IJMfgGeom3D
    Dim oCS As IJComplexString
    Dim oMoniker As IMoniker
    
    On Error Resume Next
    Set oPlateSystem = m_oMfgRuleHelper.GetTopMostParentSystem(Part)
    On Error GoTo ErrorHandler
    
    Set oPlateUtils = New PlateUtils
    
    If Not oPlateSystem Is Nothing Then
        Set oKnuckleColl = oPlateUtils.GetKnucklesOnPlate(oPlateSystem)
        
        For nIndex = 1 To oKnuckleColl.Count
            Dim oKnuckleWire As IJWireBody
            Set oKnuckleWire = oKnuckleColl.Item(nIndex)
            'Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oKnuckleWire)
            
            Dim oKNElem As IJElements
            Set oKNElem = m_oMfgRuleHelper.WireBodyToComplexStrings(oKnuckleWire)
            
            For Each oCS In oKNElem
                Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            oGeom3d.PutGeometry oCS
            oGeom3d.PutGeometrytype STRMFG_KNUCKLE_MARK
            oGeom3d.PutSubGeometryType STRMFG_NAME_MARK
            oGeom3d.IsSupportOnly = True
            
            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oKnuckleWire)
            oGeom3d.PutMoniker oMoniker
                    
            oGeomCol3d.AddGeometry 1, oGeom3d
            Next
            
        Next nIndex
    
    End If
    
    ' Get the collection of Knuckle Curves
    Dim oKnuckleCol As Collection
    Dim oSurfaceSideACol As Collection
    Dim oSurfaceSideBCol As Collection
    
    Dim oCleanedSurf As IJSurfaceBody
    Dim oSGOModelBodyUtilities As IJSGOModelBodyUtilities
    Set oSGOModelBodyUtilities = New SGOModelBodyUtilities
    oSGOModelBodyUtilities.ReplaceSliverFacesWithTolerantEdges oNeutralSurface, 0.0035, oCleanedSurf
    
    Dim oMfgMGHelper As IJMfgMGHelper
    'Create an instance of the StrMfg math geom helper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    ' We have to use the neutral surface to get the knuckle lines.
    Dim oSurfaceBodyUtilities As IJSGOSurfaceBodyUtilities
    Set oSurfaceBodyUtilities = New SGOSurfaceBodyUtilities
    oSurfaceBodyUtilities.HasKnuckles oCleanedSurf, oKnuckleCol, _
                                            oSurfaceSideACol, oSurfaceSideBCol
                                            
    If oKnuckleCol Is Nothing Then
        'Since there are no knuckles we can leave the marking rule
        GoTo CleanUp
    End If
	
    If oKnuckleCol.Count = 0 Then
        'Since there are no knuckles we can leave the marking rule
        GoTo CleanUp
    End If
	
	Dim oMfgGeomHelper As New MfgGeomHelper
	
    For nIndex = 1 To oKnuckleCol.Count
        Dim oKnuckle As IJWireBody
        Dim oSurfaceA As IJSurfaceBody
        Dim oSurfaceB As IJSurfaceBody
        Dim oSystemMark As IJMfgSystemMark
        Dim oMarkingInfo As MarkingInfo
        Dim oObjSystemMark As IUnknown
        
        Set oKnuckle = oKnuckleCol.Item(nIndex)
        Set oSurfaceA = oSurfaceSideACol.Item(nIndex)
        Set oSurfaceB = oSurfaceSideBCol.Item(nIndex)
        
        ' If the knuckle angle is not between planar surfaces, just ignore it.
        If (Not (TypeOf oSurfaceA Is IJPlane)) Then
            GoTo NextKnuckle
        End If
                
        If (Not (TypeOf oSurfaceB Is IJPlane)) Then
            GoTo NextKnuckle
        End If
		
        ' compute Knuckle angle
        Dim oGeomOps As IJDTopologyToolBox
        Dim oMidPoint As IJDPosition
        Dim oPointA As IJDPosition
        Dim oPointB As IJDPosition
        Dim oVectorA As IJDVector
        Dim oVectorB As IJDVector
        Dim dAngle As Double
        
        Set oGeomOps = New DGeomOpsToolBox
        oGeomOps.GetMiddlePointOfCompositeCurve oKnuckle, oMidPoint
        
        oMfgMGHelper.ProjectPointOnSurfaceBody oSurfaceA, oMidPoint, oPointA, oVectorA
        oMfgMGHelper.ProjectPointOnSurfaceBody oSurfaceB, oMidPoint, oPointB, oVectorB
        
        If Not (oVectorA Is Nothing Or oVectorB Is Nothing) Then
            oVectorA.Length = 1
            oVectorB.Length = 1
            dAngle = PI - ArcCos(oVectorA.Dot(oVectorB))
        End If
        
        ' If the knuckle angle is less than 0.08 degree, just ignore it.
        If (Abs(dAngle - PI) < 0.08 * PI / 180#) Then
            GoTo NextKnuckle
        End If
        
        'compute Knuckle direction
        Dim oVectorDir As IJDVector
        Dim strDir As String
        strDir = "failed"
                
        oSurfaceA.GetCenterOfGravity oPointA
        oSurfaceB.GetCenterOfGravity oPointB
        
        Dim oProjPointA As IJDPosition
        Dim oprojPointB As IJDPosition
        
        oMfgMGHelper.ProjectPointOnSurfaceBody oSurfaceA, oPointA, oProjPointA, oVectorA
        oMfgMGHelper.ProjectPointOnSurfaceBody oSurfaceB, oPointB, oprojPointB, oVectorB
        
        Dim oLines3d As ILines3d
        Dim oCurve As IJCurve
        Dim dMinDist As Double, dDummy As Double, dX As Double, dY As Double, dZ As Double
        Dim oCOGVec As New DVector
        
        Set oLines3d = New GeometryFactory
        
        Set oCurve = oLines3d.CreateBy2Points(Nothing, oPointA.x, oPointA.y, oPointA.z, oPointB.x, oPointB.y, oPointB.z)
        
        oCurve.DistanceBetween oMidPoint, dMinDist, dX, dY, dZ, dDummy, dDummy, dDummy
        
        oCOGVec.Set (dX - oMidPoint.x), (dY - oMidPoint.y), (dZ - oMidPoint.z)
        
        oCOGVec.Length = 1#
        
        
        Set oCurve = Nothing
        Set oLines3d = Nothing
                
        'check if surfaceB is bent towards base side or offset side w.r.t surfaceA
        If Not (oVectorA Is Nothing Or oVectorB Is Nothing) Then
            Set oVectorDir = oPointB.Subtract(oPointA)
            oVectorA.Length = 1
            oVectorB.Length = 1
            oVectorDir.Length = 1
            If oVectorDir.Dot(oVectorB) < 0 Then
                strDir = "base"
            Else
                strDir = "offset"
            End If
            strDir = Direction(oVectorA, strDir)
        End If
                         
        'Convert the IJWireBody to a IJComplexString
        Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oKnuckle)
        
        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
        Dim NameString As String
        
        'Set the marking side
        If oCOGVec.Dot(oVectorA) > 0# Then
            oSystemMark.SetMarkingSide UpSide
        Else
            If UpSide = BaseSide Then
                oSystemMark.SetMarkingSide OffsetSide
            Else
                oSystemMark.SetMarkingSide BaseSide
            End If
        End If
        
        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark
        
        'Set the angle
        oMarkingInfo.FittingAngle = dAngle
        
        'set the direction
        oMarkingInfo.Direction = strDir
        
        ' Commenting this name because it is now handled by Annotation.
        ' Uncomment to see how the name should appear as.
'        Dim AngleInDegrees As Double
'        AngleInDegrees = dAngle * 180 / PI
'        If Abs(AngleInDegrees - 90) > 1 Then
'            oMarkingInfo.Name = NameString & CStr(Round(AngleInDegrees, 0))
'        Else
'            oMarkingInfo.Name = NameString
'        End If
        
        Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oGeom3d.PutGeometry oCS
        oGeom3d.PutGeometrytype STRMFG_KNUCKLE_MARK
        Set oObjSystemMark = oSystemMark

        oSystemMark.Set3dGeometry oGeom3d
            
        oGeomCol3d.AddGeometry 1, oGeom3d
        
NextKnuckle:
        Set oKnuckle = Nothing
        Set oSurfaceA = Nothing
        Set oSurfaceB = Nothing
        Set oCS = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oObjSystemMark = Nothing
        Set oMoniker = Nothing
        
        Set oGeomOps = Nothing
        Set oMidPoint = Nothing
        Set oPointA = Nothing
        Set oPointB = Nothing
        Set oVectorA = Nothing
        Set oVectorB = Nothing
        Set oVectorDir = Nothing
        Set oCOGVec = Nothing
    Next nIndex

    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d


CleanUp:
    Set oKnuckleCol = Nothing
    Set oSurfaceSideACol = Nothing
    Set oSurfaceSideBCol = Nothing
    Set oMfgMGHelper = Nothing
    Set oNeutralSurface = Nothing
	Set oMfgGeomHelper = Nothing
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1011, , "RULES")
    GoTo CleanUp
End Function


Private Function ArcCos(x As Double) As Double
   If Abs(x) < 1 Then
      ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
   Else
      ArcCos = 0
   End If
End Function

' Check Vector with Major CS axis and find which is closer, X or Y or Z axis
' Then determine if knuckle is also same dir or opposite dir and return string accordingly
Private Function Direction(oVec As IJDVector, oKnuckleDir As String) As String
    If Abs(oVec.x) >= Abs(oVec.y) Then
        If Abs(oVec.x) >= Abs(oVec.z) Then
            If oVec.x >= 0 Then
                If oKnuckleDir = "base" Then
                    Direction = "fore"
                Else
                    Direction = "aft"
                End If
            Else
                If oKnuckleDir = "base" Then
                    Direction = "aft"
                Else
                    Direction = "fore"
                End If
            End If
        Else
            If oVec.z >= 0 Then
                If oKnuckleDir = "base" Then
                    Direction = "upper"
                Else
                    Direction = "lower"
                End If
            Else
                If oKnuckleDir = "base" Then
                    Direction = "lower"
                Else
                    Direction = "upper"
                End If
            End If
        End If
    Else
        If Abs(oVec.y) >= Abs(oVec.z) Then
            If oVec.y >= 0 Then
                If oKnuckleDir = "base" Then
                    Direction = "port"
                Else
                    Direction = "starboard"
                End If
            Else
                If oKnuckleDir = "base" Then
                    Direction = "starboard"
                Else
                    Direction = "port"
                End If
            End If
        Else
            If oVec.z >= 0 Then
                If oKnuckleDir = "base" Then
                    Direction = "upper"
                Else
                    Direction = "lower"
                End If
            Else
                If oKnuckleDir = "base" Then
                    Direction = "lower"
                Else
                    Direction = "upper"
                End If
            End If
        End If
    End If
End Function




 
