VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndConnMarkLogical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 '*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the (Profile) Logical End Connection Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      Siva        04 Feb 2011      Created
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.EndConnMarkLogical"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "EndConnectionMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_END_MARK, ReferenceObjColl, oGeomCol3d

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
        'Since there is no connecting structure we can leave the rule
        GoTo CleanUp
    End If
    
    Dim oConnectionData As ConnectionData

    Dim nIndex As Long, jIndex As Long
    Dim oSDProfileWrapper As New StructDetailObjects.ProfilePart
    Dim oSDMemberWrapper As New StructDetailObjects.MemberPart
    Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn

    ' Loop thru each Physical Connections
    Dim bContourEnd As Boolean
    Dim oWBColl As Collection
    Dim oSystemMark As IJMfgSystemMark
    Dim oMoniker As IMoniker
    Dim oMarkingInfo As MarkingInfo
    Dim oGeom3d As IJMfgGeom3D
    Dim lGeomCount As Long
    Dim oPlatePort As IJPort
    
    lGeomCount = 1
    
    Dim oConnPartColl As IJElements
    Set oConnPartColl = New JObjectCollection
    
    Dim oUpsidePort As IJPort
    Set oUpsidePort = oPlateWrapper.GetSurfacePort(UpSide)

    Dim oRefPort As IJPort
    Dim oSideAColl  As Collection
    Set oSideAColl = New Collection
    Dim oSideBColl  As Collection
    Set oSideBColl = New Collection
    
    'Grouping the profiles on either side of plate to two different Collections
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        If nIndex = 1 Then
            Set oRefPort = oConnectionData.ConnectingPort
        End If
        
        'Check if the connected object is a profile
        If Not TypeOf oConnectionData.ToConnectable Is IJStructProfilePart Then
             GoTo NextIndex
        Else
            If oConnectionData.ConnectingPort Is oRefPort Then
                oSideAColl.Add oConObjsCol.Item(nIndex)
            Else
                oSideBColl.Add oConObjsCol.Item(nIndex)
            End If
        End If
        
NextIndex:
    Next
    
    Dim oCollToRemove As Collection
    Set oCollToRemove = New Collection
     
    'To check the overlap of the profiles B/N two collections
    For nIndex = 1 To oSideAColl.Count
        
        oConnectionData = oSideAColl.Item(nIndex)
        
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            Set oSDProfileWrapper.object = oConnectionData.ToConnectable
        Else
            Set oSDMemberWrapper.object = oConnectionData.ToConnectable
        End If
        
        Dim dWebLengthA     As Double
        
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            dWebLengthA = oSDProfileWrapper.WebLength
        Else
            dWebLengthA = oSDMemberWrapper.WebLength
        End If
        
        Dim oSideAPort As IJPort
        Set oSideAPort = oConnectionData.ConnectingPort
        
        'Getting the Landing curve for profiles in oSideAColl
        Dim oLandingCurveA   As IJWireBody
        Dim oThicknessDir As IJDVector, bCentered As Boolean
        
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            oSDProfileWrapper.LandingCurve oLandingCurveA, oThicknessDir, bCentered
        Else
            oSDMemberWrapper.LandingCurve oLandingCurveA, oThicknessDir, bCentered
        End If
        
        For jIndex = 1 To oSideBColl.Count
           
            oConnectionData = oSideBColl.Item(jIndex)
            
            Dim oLandingCurveB   As IJWireBody
            
            If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
                Set oSDProfileWrapper.object = oConnectionData.ToConnectable
                'Getting the Landing curve for profiles in oSideBColl
                oSDProfileWrapper.LandingCurve oLandingCurveB, oThicknessDir, bCentered
            Else
                Set oSDMemberWrapper.object = oConnectionData.ToConnectable
                oSDMemberWrapper.LandingCurve oLandingCurveB, oThicknessDir, bCentered
            End If

            Dim oModelLandingCurve As IJDModelBody
            Set oModelLandingCurve = oLandingCurveB
            
            Dim oClosestPos1  As IJDPosition, oClosestPos2   As IJDPosition, dMinDist As Double
            
            'Finiding minimun distance B/N two landing curves of the profiles
            oModelLandingCurve.GetMinimumDistance oLandingCurveA, oClosestPos1, oClosestPos2, dMinDist
            
            'Checking the minimun distance with the plate thickness
            If (Abs(dMinDist - oSDPlateWrapper.PlateThickness) <= 0.005) Then
                'Removing one of the Overlapping Profile is based on UpSide
                If oSideAPort Is oUpsidePort Then
                    oCollToRemove.Add oSideBColl.Item(jIndex)
                Else
                    oCollToRemove.Add oSideAColl.Item(nIndex)
                End If
            End If
        Next
        Set oSideAPort = Nothing
    Next
    
    'To remove the Overlapped Profile from the oConObjsCol
    For nIndex = 1 To oCollToRemove.Count
    
        oConnectionData = oCollToRemove.Item(nIndex)
        
        Dim oProfilePartToRemove As Object
        Set oProfilePartToRemove = oConnectionData.ToConnectable
         
        For jIndex = 1 To oConObjsCol.Count
        
            oConnectionData = oConObjsCol.Item(jIndex)
            If oConnectionData.ToConnectable Is oProfilePartToRemove Then
                oConObjsCol.Remove (jIndex)
                Exit For
            End If
            
        Next
    Next
    
    ' Loop thru each Physical Connections
    For nIndex = 1 To oConObjsCol.Count
        oConnectionData = oConObjsCol.Item(nIndex)
        
        'Check if the connected object is a profile
        If Not TypeOf oConnectionData.ToConnectable Is IJStructProfilePart Then
             GoTo NextItem
        End If
        
        Dim oProfilePart As IJStructProfilePart
        Set oProfilePart = oConnectionData.ToConnectable
       
        'Initialize the profile wrapper and the Physical Connection wrapper
        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            Set oSDProfileWrapper.object = oProfilePart
        Else
            Set oSDMemberWrapper.object = oProfilePart
        End If
       
        Set oPlatePort = oConnectionData.ConnectingPort
        
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = oConnectionData.AppConnection

        If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
            bContourEnd = oSDPlateWrapper.Connection_ContourProfileEnd(oConnectionData.AppConnection, oWBColl)
        Else
            bContourEnd = oSDMemberWrapper.Connection_ContourProfileEnd(oConnectionData.AppConnection, oWBColl)
        End If

        If ((bContourEnd = True) And Not (oWBColl Is Nothing)) Then
            If oWBColl.Count = 0 Then
                GoTo NextItem
            End If
            
            If oConnPartColl.Contains(oProfilePart) Then
                GoTo NextItem
            Else
                oConnPartColl.Add oProfilePart
            End If
            
            Dim oCSMarkColl As IJElements
            Set oCSMarkColl = CreateSpecialEndConnMarks(oProfilePart, oPlatePort)
            
            If oCSMarkColl Is Nothing Then
                GoTo NextItem
            ElseIf oCSMarkColl.Count = 0 Then
                GoTo NextItem
            End If
            
            Dim oMarkCS1 As IJComplexString
            Dim oMarkCS2 As IJComplexString

            Set oMarkCS1 = oCSMarkColl.Item(1)
            Set oMarkCS2 = oCSMarkColl.Item(2)

            Dim jCount As Long
            For jCount = 1 To 2
                'Create a SystemMark object to store additional information
                 Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

                'Set the marking side
                oSystemMark.SetMarkingSide oPlateWrapper.GetSide(oConnectionData.ConnectingPort)

                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark

                If TypeOf oConnectionData.ToConnectable Is IJProfilePart Then
                    oMarkingInfo.Thickness = oSDProfileWrapper.WebThickness
                Else
                    oMarkingInfo.Thickness = oSDMemberWrapper.WebThickness
                End If

                oMarkingInfo.FittingAngle = oSDPhysicalConn.MountingAngle

                Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                If jCount = 1 Then
                    oGeom3d.PutGeometry oMarkCS1
                Else
                    oGeom3d.PutGeometry oMarkCS2
                End If
                oGeom3d.PutGeometrytype STRMFG_END_MARK

                Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                oGeom3d.PutMoniker oMoniker

                oSystemMark.Set3dGeometry oGeom3d
                oGeomCol3d.AddGeometry lGeomCount, oGeom3d
                lGeomCount = lGeomCount + 1
            Next
            
            Set oMarkCS1 = Nothing
            Set oMarkCS2 = Nothing
            Set oCSMarkColl = Nothing
            Set oGeom3d = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
        End If
NextItem:

        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oSDProfileWrapper = Nothing
        Set oSDMemberWrapper = Nothing
        Set oSDPhysicalConn = Nothing
        Set oPlatePort = Nothing
    Next nIndex

    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oConnPartColl = Nothing
    Set oUpsidePort = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1009, , "RULES")
    GoTo CleanUp
End Function


