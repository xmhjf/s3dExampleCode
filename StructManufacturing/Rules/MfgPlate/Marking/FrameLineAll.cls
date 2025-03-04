VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameLineAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2010 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Frame Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'           09/24/2010   Suma Mallena      Created
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.FrameLineAll"

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    PlMrkHelpers.Initialize
    Exit Sub

ErrorHandler:
Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1010, , "RULES")
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "FrameLineAll: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error Resume Next
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(oResourceManager)
    
    CreateAPSMarkings STRMFG_FRAMELINE_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Dim dDotP           As Double
    Dim oPlateNormal    As IJDVector
    Dim oXAxis          As New DVector
    
    oXAxis.Set 1#, 0#, 0#
    Set oPlateNormal = GetPlateNeutralSurfaceNormal(Part)
    
    If Not oPlateNormal Is Nothing Then
        dDotP = oPlateNormal.Dot(oXAxis)
        If Abs(1# - Abs(dDotP)) < 0.0001 Then
            Exit Function
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    'Initialize the PlateWrapper with the PlatePart
    Dim oPlate As StructDetailObjects.PlatePart
    Set oPlate = New StructDetailObjects.PlatePart
    Set oPlate.object = Part
    
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part
   
    'Get the Coordinate System object
    Dim oMfgFrameSystem As IJDCoordinateSystem
    Dim nIndex As Long
    For nIndex = 1 To ReferenceObjColl.Count
        If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
            Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
        End If
    Next nIndex
    
    ' Within V7, the Mfg object inherit the Frame System. When not found we do
    ' leave by the default rule. Customers can activate the Main CS.
    ' TR-CP·121466  manufacturing Plate profile being marked even though Frame system not defined

    If oMfgFrameSystem Is Nothing Then
        Exit Function
    End If
    
    Dim oFrameSysMoniker As IMoniker
    Set oFrameSysMoniker = m_oMfgRuleHelper.GetMoniker(oMfgFrameSystem)
    
    Dim oMfgHFrameSys As IHFrameSystem
    Set oMfgHFrameSys = oMfgFrameSystem
    
    'Get the frame line intersections using the assumption that there is a coordinate system called "Main CS"
    'This is mandatory for all the Hicadec unwrapping algorithms therefore this seems a safe bet.
    'Since it is framelines it automatically means the "X" direction
    'If this one works, the buttock and waterlines could be made as well based on the different directions
    'untill we have the wrappers available.
    
    'With radial CS being supported, a call to the same function with "R" as axis parameter
    'is made and the obtained planes, curves are added to collection and processed further.
    '   "R" - Stands for Radial planes in Cylindrical CS.
    
    Dim oConObjsCol     As Collection
    Dim oRefplanes      As Collection
    
    If oSDPlateWrapper.MoldedSide = "Base" Or oSDPlateWrapper.MoldedSide = "Offset" Then
        oPlate.ReferencePlaneRange oMfgHFrameSys.Name, "X", oRefplanes, oConObjsCol, BPT_Offset
        
        'Below line needs to uncommented for the system to demonstrate the capability to
        'support Radial Co-ordinate System.
        'oPlate.ReferencePlaneRange "", "R", oRefplanes, oConObjsCol, BPT_Offset
    Else
        Dim lErrNumber As Long
        lErrNumber = LogMessage(Err, MODULE, METHOD, " Unexpected Molded side ")
    End If
    
    Dim oIntersectHelper As New StructGenericUtilities.IntersectHelper
    Dim xLow As Double, yLow As Double, zLow As Double
    Dim xHigh As Double, yHigh As Double, zHigh As Double
    Dim oOffsetPort As IJPort
    Dim oPlateSurfGeom As Object
    
    Set oOffsetPort = oPlate.BasePort(BPT_Offset)
    
    Set oPlateSurfGeom = oOffsetPort.Geometry
    Set oOffsetPort = Nothing
        
    oIntersectHelper.RangeOfElement oPlateSurfGeom, xLow, yLow, zLow, xHigh, yHigh, zHigh
    
    Set oIntersectHelper = Nothing
    
    If yLow < 0# And yHigh > 0# Then
        If Not oMfgHFrameSys Is Nothing Then
            Dim oGridMiddleHelper As New SPGMiddleHelper
            Dim oCenterFrame As IHFrame
            
            oGridMiddleHelper.IsPlaneAtLocation oMfgHFrameSys, y, 0#, oCenterFrame

            If Not oCenterFrame Is Nothing Then
                Dim oIntersectedUnknown As IUnknown
                
                Dim oGenericIntersector As IMSModelGeomOps.DGeomOpsIntersect
                Set oGenericIntersector = New IMSModelGeomOps.DGeomOpsIntersect
                
                oGenericIntersector.PlaceIntersectionObject Nothing, _
                                                            oPlateSurfGeom, oCenterFrame, _
                                                            Nothing, _
                                                            oIntersectedUnknown
                    
                If Not oIntersectedUnknown Is Nothing Then
                    oRefplanes.Add oCenterFrame.Name
                    oConObjsCol.Add oIntersectedUnknown
                End If
                Set oGenericIntersector = Nothing
                Set oIntersectedUnknown = Nothing
            End If
            Set oGridMiddleHelper = Nothing
            Set oCenterFrame = Nothing
        End If
        
        Set oMfgHFrameSys = Nothing
    End If
        
    If Not (oConObjsCol Is Nothing) Then
        
        On Error Resume Next
        For nIndex = 1 To oConObjsCol.Count
          
            Dim oCS             As IJComplexString
            Dim oSystemMark     As IJMfgSystemMark
            Dim oMarkingInfo    As MarkingInfo
            Dim oGeom3d         As IJMfgGeom3D
            Dim oFrameline      As IUnknown
            Dim oCurves         As IJElements
            
            Set oFrameline = oConObjsCol.Item(nIndex)
            
    '       Convert the IJWireBody to a collection of IJComplexString's
            Set oCurves = m_oMfgRuleHelper.WireBodyToComplexStrings(oFrameline)
            
            ' Get type of Frame and set as custom attribute on Marking Info
            Dim oGridData As ISPGGridData
            Set oGridData = oMfgHFrameSys.PrincipalXAxis.Frame(oRefplanes.Item(nIndex))
       
            Dim oCodeListMetaData As IJDCodeListMetaData
            Set oCodeListMetaData = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

            Dim strPlaneType As String
            strPlaneType = oCodeListMetaData.LongStringValue("GridPlaneType", oGridData.Type)
                        
            For Each oCS In oCurves
            
    '           Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                    
    '           Set the marking side
                oSystemMark.SetMarkingSide UpSide
    
    '           QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
               
    '           Get the Frame Name
                oMarkingInfo.Name = oRefplanes(nIndex)
                
                'Set the Frame Type
                If Not strPlaneType = vbNullString Then
                    oMarkingInfo.SetAttributeNameAndValue "REFERENCE", strPlaneType
                End If
                
                Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_FRAMELINE_MARK
                
                oGeom3d.PutMoniker oFrameSysMoniker
                oGeom3d.Put3dModelMoniker m_oMfgRuleHelper.GetMoniker(oMfgHFrameSys.PrincipalXAxis.Frame(oRefplanes.Item(nIndex)))
                
                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry 1, oGeom3d
                
                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oGeom3d = Nothing
               
            Next oCS
            
            Set oGridData = Nothing
            Set oCodeListMetaData = Nothing
            
        Next nIndex
    End If
    
'    Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

CleanUp:

    Set oPlate = Nothing
    Set oCS = Nothing
    Set oFrameline = Nothing
    Set oConObjsCol = Nothing
    Set oSystemMark = Nothing
    Set oGeom3d = Nothing
    Set oResourceManager = Nothing
    Set oMfgFrameSystem = Nothing
    Set oMfgHFrameSys = Nothing
    Set oFrameSysMoniker = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1010, , "RULES")
    GoTo CleanUp
End Function
