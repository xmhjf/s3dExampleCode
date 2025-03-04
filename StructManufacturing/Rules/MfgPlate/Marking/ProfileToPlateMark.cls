VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileToPlateMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Profile To Plate Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.ProfileToPlateMark"
Private Const TEEJOINTFITTINGMARK As Boolean = True
Private Const LAPJOINTFITTINGMARK As Boolean = False

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

' ***********************************************************************************
' Public Function IJDMfgSystemMarkingRule_CreateBeforeUnfold
'
' Description:  function creates fitting marks on the this plate for each T - connected profile.
'               Input arguments: this part, marking side.
'               Output argument: collection of marking lines.
'
' ***********************************************************************************

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d

    Const METHOD = "ProfileToPlateMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
       
    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3D As IJMfgGeomCol3d
    Set oGeomCol3D = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_PROFILE_TO_PLATE_MARK, ReferenceObjColl, oGeomCol3D
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3D
    
    'Prepare collections
    Dim oSDPlateWrapper As New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part
    
    Dim oMfgRuleHlprPlatePart As MfgRuleHelpers.PlatePartHlpr
    Set oMfgRuleHlprPlatePart = New MfgRuleHelpers.PlatePartHlpr
    Set oMfgRuleHlprPlatePart.object = Part
    
    'Retrieve collection of ConnectionData objects
    Dim oConnectedObjectsColl As Collection
    'Set oConnectedObjectsColl = oSDPlateWrapper.ConnectedObjects
    Set oConnectedObjectsColl = GetPhysicalConnectionData(Part, ReferenceObjColl, False)
    
    If oConnectedObjectsColl Is Nothing Then
        'Nothing to be marked so therefore you can go to cleanup and exit
        GoTo CleanUp
    End If
    
    Dim oConnectionData As ConnectionData
    Dim iConnObjIndex As Integer 'iteration object
    
    ' The below code is customer specific
    If TEEJOINTFITTINGMARK = True Then
        'iterate through the collection of profile parts connected to this plate
        For iConnObjIndex = 1 To oConnectedObjectsColl.Count
            oConnectionData = oConnectedObjectsColl.Item(iConnObjIndex)
            If Not TypeOf oConnectionData.ToConnectable Is IJProfilePart Then GoTo NextItem
            
            Dim oSDConnProfileWrapper As New StructDetailObjects.ProfilePart
            Set oSDConnProfileWrapper.object = oConnectionData.ToConnectable
            
            'Get collection of mark positions
            Dim oMarkPosColl As Collection
            Set oMarkPosColl = GetFittingMarkPositions(oSDPlateWrapper.object, oConnectionData, UpSide)
            
            'exit method if no mark point has been found
            If oMarkPosColl Is Nothing Then GoTo NextItem
            
            'Iterate through the collection of mark points in oMarkPosColl
            Dim oMarkPosObj As Object 'iteration object
            For Each oMarkPosObj In oMarkPosColl
                
                'Get connected ports on both parts
                Dim oThisPartPort As IJPort
                Dim oConnPartPort As IJPort
                Dim oWebPort As IJPort
        
                'get connected port on this plate
                Set oThisPartPort = oConnectionData.ConnectingPort
                'get mounting port on connected profile
                Set oConnPartPort = oConnectionData.ToConnectedPort
        
                'get web adjacent to mounting face
                'oWebPort = oConnProfileWrapper.ProfileMountFaceAdjacentPort oConnPartPort
        
                '****REMOVE AFTER ADJACENT WEB IS AVAILABLE ON PROFILE WRAPPER----------->
                Dim oPartSupport As IJPartSupport
                Dim oProfilePartSupport As IJProfilePartSupport
                Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
                Set oPartSupport.Part = oSDConnProfileWrapper.object
                Set oProfilePartSupport = oPartSupport
                On Error Resume Next
                Set oWebPort = oProfilePartSupport.GetFacePortAdjacentToMountingFace(1)
                '****REMOVE AFTER ADJACENT WEB IS AVAILABLE ON PROFILE WRAPPER<-----------
                If oWebPort Is Nothing Then
                    GoTo NextMarkPosObj
                End If
                On Error GoTo ErrorHandler
                       
                'Get surfaces of ports
                Dim oThisPartSuface As Object 'IJSurface
                Dim oConnPartSurface As Object 'IJSurface
                Set oThisPartSuface = oThisPartPort.Geometry
                Set oConnPartSurface = oWebPort.Geometry
        
        
                'Check If oMarkPosObj is on the surface of this plate
                Dim oIJDModelBody As IJDModelBody
                Set oIJDModelBody = New MODELGEOMETRYLib.DWireBody2
                Set oIJDModelBody = oThisPartSuface
                Dim oPositionStatus As JS_TOPOLOGY_POINT_BODY
                oIJDModelBody.GetPositionStatus oMarkPosObj, oPositionStatus
                'Note: JS_TOPOLOGY_POINT_BODY_INSIDE = 1 or JS_TOPOLOGY_POINT_BODY_BOUNDARY = 2 means position is on surface
                
                If Not oPositionStatus = 1 Then
                ElseIf Not oPositionStatus = 2 Then
                    GoTo NextMarkPosObj
                End If
        
        
                'Create mark line through oMarkPosObj
                Dim oCS As IJComplexString
                Set oCS = New ComplexString3d
                Set oCS = CreateLocationFittingMark(oMarkPosObj, oThisPartSuface, oConnPartSurface)
        
                If oCS Is Nothing Then
                    GoTo NextMarkPosObj
                End If
        
                'Prepare some different assisting objects
                Dim oMarkingInfo As MarkingInfo
                Dim oGeom3d As IJMfgGeom3D
                Dim oMoniker As IMoniker
                Dim oObjSystemMark As IUnknown
                Dim oSystemMark As IJMfgSystemMark
        
                'Dim oSystemMarkFactory As GSCADMfgSystemMark.MfgSystemMarkFactory
                Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                Dim ePlateSide As enumPlateSide
                ePlateSide = oMfgRuleHlprPlatePart.GetSide(oThisPartPort)
                If (ePlateSide = UnDefinedSide) Then
                    oSystemMark.SetMarkingSide UpSide
                Else
                    oSystemMark.SetMarkingSide ePlateSide
                End If
        
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
        
                'Set marking information, if any
                oMarkingInfo.Name = "PROFILE FITTING MARK"
        
                'Add ML's geometry to collection
                Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                oGeom3d.PutGeometry oCS
                oGeom3d.PutGeometrytype STRMFG_PROFILE_TO_PLATE_MARK
        
                Set oObjSystemMark = oSystemMark
                
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                oGeom3d.PutMoniker oMoniker
                        
                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry 1, oGeom3d
                
NextMarkPosObj:
                Set oGeom3d = Nothing
                Set oMoniker = Nothing
                Set oObjSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oSystemMark = Nothing
                Set oCS = Nothing
                Set oThisPartSuface = Nothing
                Set oConnPartSurface = Nothing
                Set oIJDModelBody = Nothing
                Set oWebPort = Nothing
                Set oConnPartPort = Nothing
                Set oThisPartPort = Nothing
                Set oMarkPosObj = Nothing
            Next oMarkPosObj
        
NextItem:
            Set oSDConnProfileWrapper = Nothing
            Set oMarkPosColl = Nothing
        Next iConnObjIndex
    End If
    
    ' The below code is customer specific
    If LAPJOINTFITTINGMARK = True Then
        ' If the part is bracket create special fitting mark
        If CheckIfPartIsBracket(Part) = True Then
            Dim oAssemblySequence As GSCADPlanningInterfaces.IJAssemblySequence
            Set oAssemblySequence = GetAssemblySequenceObject(Part)
            
            Dim oRefLapCurve       As IJCurve
            Dim lFittingIndex       As Long
            Dim lConnObjSeqIndex    As Long
            Dim nWBIndex            As Long
            Dim oLandingCurve       As IJWireBody
            Dim oLandCurveVec       As IJDVector
            Dim oRefLandCurveVec    As IJDVector
            Dim oTopologyLocate     As New TopologyLocate
                                
            lConnObjSeqIndex = 100000   ' intialize with large number
                            
            ' Lap Fitting mark for brackets
            For iConnObjIndex = 1 To oConnectedObjectsColl.Count
                
                oConnectionData = oConnectedObjectsColl.Item(iConnObjIndex)
                If Not TypeOf oConnectionData.ToConnectable Is IJProfilePart Then GoTo NextItem
                
                ' Get the current index
                Dim lChildSeqIndex  As Long
                lChildSeqIndex = oAssemblySequence.ChildIndex(oConnectionData.ToConnectable)
                
                ' Get the lap connection contours
                Dim oWBColl As Collection
                Dim bContourLap As Boolean
                bContourLap = oSDPlateWrapper.Connection_ContourLap(oConnectionData.AppConnection, oWBColl)
                
                If ((bContourLap = True) And Not (oWBColl Is Nothing)) Then
                    
                    If oWBColl.Count = 0 Then
                        GoTo NextItem
                    End If
                    
                    For nWBIndex = 1 To oWBColl.Count
                    
                        Dim oWireBody As IJWireBody
                        Set oWireBody = oWBColl.Item(nWBIndex)
                        
                        Dim oCurveStartPos      As IJDPosition
                        Dim oCurveEndPos        As IJDPosition
                        
                        Dim oPCCurveVec         As IJDVector
                        Dim dDotProdVal         As Double
                        
                        If nWBIndex = 1 Then
                            Set oLandingCurve = Nothing
                            Set oLandCurveVec = Nothing
                            
                            ' Get the profile landing curve and its vector
                            Set oLandingCurve = oTopologyLocate.GetProfileParentWireBody(oConnectionData.ToConnectable)
                            
                            oLandingCurve.GetEndPoints oCurveStartPos, oCurveEndPos
                            Set oLandCurveVec = oCurveStartPos.Subtract(oCurveEndPos)
                            oLandCurveVec.Length = 1
                        End If
                        
                        Set oCurveStartPos = Nothing
                        Set oCurveEndPos = Nothing
                        oWireBody.GetEndPoints oCurveStartPos, oCurveEndPos
                        
                        Set oPCCurveVec = oCurveStartPos.Subtract(oCurveEndPos)
                        oPCCurveVec.Length = 1
                        
                        dDotProdVal = oLandCurveVec.Dot(oPCCurveVec)
                        
                        ' Check if lap contour and landing curve are parallel
                        If Abs(dDotProdVal) > 0.99 Then
                        
                            Dim oPCCSColl   As IJElements
                            Set oPCCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oWireBody)
                            
                            If Not oPCCSColl Is Nothing Then
                                For Each oCS In oPCCSColl
                                    Dim oPCCurve As IJCurve
                                    Set oPCCurve = oCS
                                    
                                    ' Get the connected part which is first in assembly sequence
                                    If lChildSeqIndex <= lConnObjSeqIndex And Not (lChildSeqIndex = -1 And Not oRefLapCurve Is Nothing) Then
                                        Set oRefLapCurve = Nothing
                                        Set oRefLapCurve = oPCCurve
                                        lFittingIndex = iConnObjIndex
                                        lConnObjSeqIndex = lChildSeqIndex
                                        Set oRefLandCurveVec = oLandCurveVec
                                    End If
                                    
                                Next
                            End If
                        End If
                        
                    Next nWBIndex
                End If
            Next iConnObjIndex
            
            If Not oRefLapCurve Is Nothing Then
            
                Dim dLength     As Double
                dLength = oRefLapCurve.Length
                
                Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double, dAlong As Double
                oRefLapCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
                
                Dim oStartPos   As IJDPosition
                Set oStartPos = New DPosition
                    
                oStartPos.Set dStartX, dStartY, dStartZ
                                
                Dim oMfgGeomHelper As IJMfgGeomHelper
                Set oMfgGeomHelper = New MfgGeomHelper
                
                dAlong = dLength * (2 / 3)
                Dim oPosition       As IJDPosition
                Dim oNewPosition    As IJDPosition
                Dim oFittingVec     As IJDVector
                Set oPosition = oMfgGeomHelper.GetPointAtDistAlongCurve(oRefLapCurve, oStartPos, dAlong)
                
                oConnectionData = oConnectedObjectsColl.Item(lFittingIndex)
                
                Set oSDConnProfileWrapper = New StructDetailObjects.ProfilePart
                Set oSDConnProfileWrapper.object = oConnectionData.ToConnectable
                
                bContourLap = oSDConnProfileWrapper.Connection_ContourLap(oConnectionData.AppConnection, oWBColl)
                
                ' Get the reference point on the lap curve that overlaps with plate outer contour
                For nWBIndex = 1 To oWBColl.Count
                
                    Set oWireBody = oWBColl.Item(nWBIndex)
                    
                    Set oCurveStartPos = Nothing
                    Set oCurveEndPos = Nothing
                    oWireBody.GetEndPoints oCurveStartPos, oCurveEndPos
                    
                    Set oPCCurveVec = oCurveStartPos.Subtract(oCurveEndPos)
                    oPCCurveVec.Length = 1
                    
                    dDotProdVal = oRefLandCurveVec.Dot(oPCCurveVec)
                    
                    If Abs(dDotProdVal) > 0.99 Then
                        Set oNewPosition = m_oMfgRuleHelper.ProjectPointOnWire(oWireBody, oPosition)
                        
                        Set oFittingVec = oNewPosition.Subtract(oPosition)
                        oFittingVec.Length = 1
                        Exit For
                    End If
                       
                Next nWBIndex
                            
                If oNewPosition Is Nothing Then
                    Exit Function
                End If
                
                Set oThisPartPort = oConnectionData.ConnectingPort
                Set oConnPartPort = oConnectionData.ToConnectedPort
                
                ' Create lap fitting mark
                Dim oFittingCS          As IJComplexString
                Set oFittingCS = CreateLapConnFittingMark(oNewPosition, oThisPartPort.Geometry, oConnPartPort.Geometry, 0.08, oFittingVec, True)
                
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                
                oSystemMark.SetMarkingSide oMfgRuleHlprPlatePart.GetSide(oConnectionData.ConnectingPort)
                
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
                
                oMarkingInfo.Thickness = oSDConnProfileWrapper.WebThickness
                oMarkingInfo.Name = "PROFILE FITTING"
                
                Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                oGeom3d.PutGeometry oFittingCS
                oGeom3d.PutGeometrytype STRMFG_PROFILE_TO_PLATE_MARK
                'oGeom3d.PutSubGeometryType STRMFG_LAP_MARK
                oMarkingInfo.SetAttributeNameAndValue "REFERENCE", "Lap"
                
                oGeomCol3d.AddGeometry 1, oGeom3d
                oSystemMark.Set3dGeometry oGeom3d
                
                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                oGeom3d.PutMoniker oMoniker
                
                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oGeom3d = Nothing
                
            End If
        End If
    End If
    
    'Return collection of ML's geometry
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oMfgRuleHlprPlatePart = Nothing
    Set oConnectedObjectsColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1023, , "RULES")
    GoTo CleanUp
End Function
