VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnsupportedInputHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : UnsupportedInputHelper.cls
'
'Author : Struct Dev
'
'Description :
'    SmartPlant Structural unsupported Frame Connection Definition file
'
'History:
'
' 05/21/03   Struct Dev    Changed error handling in GetRelatedObjects
'                          and SetRelatedObjects to raise the error;
'                          also replaced the Msgbox with HandleError
' 08/29/03   Struct Dev    TR 47396.  Establish, but do not maintain, relations to grids.
'                          This allows commands to set Unsupported to clear relations with grids.
' 04/18/06   Struct Dev    TR 96246 - Added a check in SetRelatedObjects to verify if
'                                the reference collection exists to the Frame Connection prior to calling
'                                GetRelatedObjects. This is because GetRelatedObjects creates a reference
'                                collection if one does not exist. Typically this is okay, however not for a
'                                Frame Connection that is tagged for delete. Adding this reference collection
'                                will cause the delete transaction to abort. Creating one of these conditions
'                                requires: placing a member axis-end to a supporting member, change permission
'                                group of placed supported member, modify access of supported members
'                                permission group to no write access, delete the supporting member and
'                                change permission group of supported member back to full-access.
' 05/10/06  Struct Dev    Support PointOn to curved objects, not just IJLine
' 05/25/06  Struct Dev    Follow on for TR 96246. If FC not being deleted, go ahead and create the RefColl and
'                                avoid setting bDisconnectShared to true to allow a shared joint connection to remain
'                                even though the joint's PointOn is set to nothing.  Place AxisAlong and AxisEnd shared-joint.
'                                Delete the AxisAlong.  The AxisEnd-shared joint should remain.
' 09/26/06  Struct Dev    TR 104591  ReadPathPointOnObjectsForFC
'
' 10/04/06  Struct Dev    TR 107067 ISPSFCInputHelper_SetRelatedObjects() clears the offsets only
'                                   if the current FC is not unsupported type.
' 12/07/06  Struct Dev    TR 111350 Modified SetRelatedObjects() to call getRelatedObjects on current fc
'                                   only if the current fc is not nothing.
' 08/17/07  Struct Dev    TR 125888 In SetRelatedObjects, set PointOnPositionRule to ByIntersect for grid lines etc
' 06/24/09  Sdas          TR 165583 Standard member endcut is deleted when fram connection modified
'
' 07/1/09   RP            CR165819 - Added code in SetRelatedObjects() to remove association to CanRule
'                         if the FC had a canrule in its RefColl
' 09/23/09  MH            TR 171361 SetRelatedObjects - clear Joint's PO relation to another joint.
' 11/02/09  MH            TR 118121 reset PointOnEndRule to Auto when setting to PointOnPositionRule to Intersect
' 03/04/11  MH            TR 187673 remove ACs
'*****************************************************************************************************

Const MODULE = "UnsupportedInputHelper"
Private m_oLocalizer As IJLocalizer

Implements SPSMembers.ISPSFCInputHelper
Implements SPSMembers.ISPSFCInputHelperEx



Private Property Get ISPSFCInputHelperEx_AMPOption() As SPSMembers.SPSFCAMPOptions
    ISPSFCInputHelperEx_AMPOption = SPSFCAMPOption_NONE
End Property

'*************************************************************************
'Function
'ISPSFCInputHelper_ExecuteSelectionRule
'
'Abstract
'Selects a particular type of FrameConnection to be used
'
'Arguments
'FC as FrameConnection
'name as String
'
'Return
'Returns the catalog smartitem name as string and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************

Private Function ISPSFCInputHelper_ExecuteSelectionRule(ByVal FC As SPSMembers.ISPSFrameConnection, ByRef selection As String) As SPSMembers.SPSFCInputHelperStatus

    'Could an Unsupported do something else here ??
    selection = "Unsupported"
    ISPSFCInputHelper_ExecuteSelectionRule = SPSFCInputHelper_Ok

End Function

'*************************************************************************
'Function
'ISPSFCInputHelper_GetRelatedObjects
'
'Abstract
'Returns the objects in the ReferenceCollection used   as input by the FrameConnection.
'
'Arguments
'FC As ISPSFrameConnection
'RelatedObject1 As Object
'RelatedObject2 As Object
'
'Return
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_GetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, RelatedObject1 As Object, RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus
    Const METHOD = "ISPSFCInputHelper_GetRelatedObjects"
    ' The unsupported connection type can only be related via pointOn such as grids.
    ' By definition, this connection type cannot be related to other members using the ReferenceCollection.
    ' But we look into the RefCol anyway for debug purposes.
    
    On Error GoTo ErrorHandler
    Dim IHStatus As SPSFCInputHelperStatus
    Dim oRefColl As IMSSymbolEntities.IJDReferencesCollection
    Dim oObject As Object
    Dim count As Long
    Dim oRel1 As Object, oRel2 As Object
    Dim oCR As ISPSCanRule
    
    IHStatus = SPSFCInputHelper_UnexpectedError

    Set oRefColl = GetRefColl(FC)
    If oRefColl Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNSUPP_MISSING_REQDDATA, "Unsupported frame connection is missing required data. Delete this frame connection.")
    End If
    
    count = oRefColl.IJDEditJDArgument.GetCount
    
    If count > 1 Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNSUPP_LINKEDTO_MORETHANONE, "Unsupported frame connection is linked to more than one member system. Delete this frame connection.")
    ElseIf count > 0 Then
        If TypeOf oRefColl.IJDEditJDArgument.GetEntityByIndex(1) Is ISPSCanRule Then
            Set oCR = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
            oCR.GetPrimaryMemberSystem oObject
            Set oCR = Nothing
        Else
            Set oObject = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        End If
        If Not TypeOf oObject Is ISPSMemberSystem Then
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNSUPP_MISSING_PARENT, "Unsupported frame connection is missing its parent member system information. Delete this frame connection.")
        End If
        If Not FC.MemberSystem Is oObject Then
            IHStatus = SPSFCInputHelper_InconsistentRelations
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNSUPP_NOTLINKED_MEMBS, "Unsupported frame connection is linked to something other than a member system. Delete this frame connection.")
        End If
    End If
    
    FC.Joint.GetPointOn RelatedObject1, RelatedObject2
    
    IHStatus = SPSFCInputHelper_Ok

    If ObjectIsSPSMemberObject(RelatedObject1) Or ObjectIsSPSMemberObject(RelatedObject2) Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject

    ElseIf ReadPathPointOnObjectsForFC(FC, oRel1, oRel2) Then
        If oRel1 Is RelatedObject1 And oRel2 Is RelatedObject2 Then
            IHStatus = SPSFCInputHelper_Ok
        ElseIf oRel1 Is RelatedObject2 And oRel2 Is RelatedObject1 Then
            IHStatus = SPSFCInputHelper_Ok
        Else
            IHStatus = SPSFCInputHelper_InconsistentRelations
        End If
    End If
    Set oCR = Nothing
    ISPSFCInputHelper_GetRelatedObjects = IHStatus
    
    Exit Function

ErrorHandler:
    If ObjectAssocFlags(FC, &H200, 0) Then      ' object is not RELATION_INSERTED_IN_TRANSACTION
         HandleError MODULE, METHOD             ' log error
    End If
    Err.Clear
    ISPSFCInputHelper_GetRelatedObjects = IHStatus
End Function

'*************************************************************************
'Function
'ISPSFCInputHelper_SetRelatedObjects
'
'Abstract
'Sets the objects in the ReferenceCollection used as input by the FrameConnection.
'SetRelatedObjects is called by the command to enable this code to set connections to the FrameConnection's ReferenceCollection.
'
'Different connections will establish different relations according to what needs to be watched.
'This connection is the RootSelector connection, and serves as a hub for the "ByRule".
'
'See documentation and metadata for more details.
'
'Arguments
'FC As ISPSFrameConnection
'RelatedObject1 As Object
'RelatedObject2 As Object
'
'Return
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_SetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal RelatedObject1 As Object, ByVal RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus
    Const METHOD = "ISPSFCInputHelper_SetRelatedObjects"
    On Error GoTo ErrorHandler

    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim ii As Long, eleCount As Long
    Dim ijEles As IJElements
    Dim oRefColl As IMSSymbolEntities.IJDReferencesCollection
    Dim oR1 As Object, oR2 As Object
    Dim oOtherR1 As Object, oOtherR2 As Object
    Dim oCanRule As ISPSCanRule
    Dim oOtherMS As ISPSMemberSystem
    Dim oOtherFC As ISPSFrameConnection
    
    Dim oFC As ISPSFrameConnection
    Dim iJoint As ISPSAxisJoint
    Dim iMemberPart As ISPSMemberPartPrismatic
    Dim oIJDObject As IJDObject
    Dim oMeMS As ISPSMemberSystem
    Dim IH As ISPSFCInputHelper
    
    Dim bKeepNonMemberPO As Boolean
    Dim bDisconnectShared As Boolean
    Dim iSPSPort As ISPSSplitAxisPort
    Dim portIndex As SPSMemberAxisPortIndex
    Dim iIJPort As IJPort
    Dim iWPO As ISPSAxisWPO
    Dim iSmartOcc As IJSmartOccurrence
    Dim oAC As Object

    IHStatus = SPSFCInputHelper_UnexpectedError

    ' if only one obj, then make it be RelatedObject1
    If RelatedObject1 Is Nothing Then
        If Not RelatedObject2 Is Nothing Then
            Set RelatedObject1 = RelatedObject2
            Set RelatedObject2 = Nothing
        End If
    End If

    'Unsupported FC cannot be PO to a MemberSystem, MemberPart, Joint, FrameConnection, or SplitConnection.
    If ObjectIsSPSMemberObject(RelatedObject1) Then
        Set RelatedObject1 = Nothing
    End If
    If ObjectIsSPSMemberObject(RelatedObject2) Then
        Set RelatedObject2 = Nothing
    End If
    
   'this should be caught by the command, but we fix it here.  Cannot be PO to the same object twice.
    If RelatedObject1 Is RelatedObject2 Then
        Set RelatedObject2 = Nothing
    End If
    
    ' If only one obj, it must be a curve.
    If Not RelatedObject1 Is Nothing Then
        If RelatedObject2 Is Nothing Then
            If TypeOf RelatedObject1 Is IJCurve Then
                Set RelatedObject2 = Nothing    ' dummy statement

            ElseIf IsOkPointType(RelatedObject1) Then
                Set RelatedObject2 = Nothing    ' dummy statement
    
            ElseIf IsOkSurfaceType(RelatedObject1) Then
                Set RelatedObject2 = Nothing    ' dummy statement
    
            Else
                Set RelatedObject1 = Nothing    ' don't know how to be PointOn.  So we ignore the object.

            End If  'If TypeOf RelatedObject1 is IJCurve
        End If  'If RelatedObject2 Is Nothing
    End If  'If Not RelatedObject1 Is Nothing

    ' If it's not a surface, and it supports IJKeypoint, and has more than one keypoint,
    ' and does NOT support IJPoint or IJDPosition ( via IJPort )
    ' then it's a keypoint object that we cannot support, because we do not support the keypoint index.
    ' I set it to Nothing here so it gets ignored.

    If Not RelatedObject1 Is Nothing Then
        If Not TypeOf RelatedObject1 Is IJSurface Then
            If TypeOf RelatedObject1 Is IJKeyPoint Then
                Dim keypointCount As Long
                Dim iKeypoint As IJKeyPoint
                Set iKeypoint = RelatedObject1
                keypointCount = iKeypoint.keypointCount
                If keypointCount > 1 Then
                    Dim tmpKeypointObj As Object
                    If TypeOf RelatedObject1 Is IJPort Then
                        Dim pPort As IJPort
                        Set pPort = RelatedObject1
                        Set tmpKeypointObj = pPort.Geometry
                    Else
                        Set tmpKeypointObj = RelatedObject1
                    End If
                    If TypeOf tmpKeypointObj Is IJPoint Then
                        Set tmpKeypointObj = Nothing    ' noop = ok
                    ElseIf TypeOf tmpKeypointObj Is IJDPosition Then
                        Set tmpKeypointObj = Nothing    ' noop = ok
                    Else
                        Set RelatedObject1 = Nothing    ' cannot support it.  Cancel this guy.
                    End If
                End If
            End If
        End If
    End If
    
    Set oMeMS = FC.MemberSystem

    ' disconnectShared = false.
    ' Test case:  AxisAlong that is sharing joint with AxisEnd.
    ' Set the AxisAlong to unsupported should leave the AxisEnd okay.
    ' But if this was the AxisEnd, disconnectShared is set to true, with an exception.
    ' If the other side is AxisEnd now watching my MemberSystem, then bDisconnectShared is again set to false.
    ' Test case: Place AxisEnd onto an Unsupported.  Then swap them so the Unsupported is made AxisEnd.
    '   That AxisEnd is now becoming Unsupported, but leave the joint alone.
    
    bDisconnectShared = False
    bKeepNonMemberPO = False    ' make me PO to the given objects, which might be Nothing or grids.

    On Error Resume Next    'moh.  02/17/2004.  May not have write access to other FC's, or AC's.

    'Check that the reference collection exists because GetRelatedObjects will create a new ref coll
    'if one does not exist and this cannot be allowed if the Frame Connection is being deleted.
    
    If FC.IsMarkedForDelete Then
        Set oRefColl = GetRefCollNoCreate(FC)    'get ref coll iff it already exists.
    Else
        Set oRefColl = GetRefColl(FC)                   'if one does not exist, then make a new one.
    End If

    If Not oRefColl Is Nothing Then
        'Get the existing related objects, to determine how to disconnect at the end.
        'If we are sharing a joint with a supporting Member, then TransferMember to Nothing to create a new joint.
        
        If Not FC.definition Is Nothing Then

            eleCount = oRefColl.IJDEditJDArgument.GetCount

            If eleCount = 3 Then        ' could be VCB or AxisEnd - sharedJoint
                Set oR1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(3)
                If TypeOf oR1 Is ISPSAxisWPO Then   ' Yes, this FC was an AxisEnd - sharedJoint
                    
                    bDisconnectShared = True        ' unsually, we want to disconnect
                    
                    ' check for the exception case.  is the FC on the other side now watching my MemberSystem ?
                    Set iWPO = oR1
                    portIndex = iWPO.portIndex

                    Set oOtherMS = iWPO.MemberSystem
                    Set oR1 = Nothing
                    Set oR2 = Nothing
                    Set oOtherFC = oOtherMS.FrameConnectionAtEnd(portIndex)
            
                    If Not oOtherFC Is Nothing Then
                        IHStatus = oOtherFC.InputHelper.GetRelatedObjects(oOtherFC, oR1, oR2)
                        If oR1 Is oMeMS Or oR2 Is oMeMS Then
                            bDisconnectShared = False
                            bKeepNonMemberPO = True
                        End If
                    End If
                End If  ' if AxisEnd
            End If
            
            'set dependent FC's to Unsupported, if they depend on my connectivity.
            'Eg:  VerticalCornerBrace
            If eleCount = 2 Or eleCount = 3 Then
                On Error Resume Next
                Dim oPOJoint As ISPSAxisJoint
                Dim elesPOJoints As IJElements, iiPOJoint As Long, nPOJoints As Long
                Dim elesPOFCs As IJElements, iiPOFC As Long, nPOFCs As Long
                
                Set oR1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)    'current watched member-system
                If Not oR1 Is Nothing Then
                    If TypeOf oR1 Is ISPSCanRule Then
                        Set oCanRule = oR1
                        Set oR1 = Nothing
                        oCanRule.GetPrimaryMemberSystem oR1
                    End If
                End If

                Set elesPOJoints = FC.Joint.PointOnJoints
                nPOJoints = elesPOJoints.count
                For iiPOJoint = 1 To nPOJoints
                    Set oPOJoint = elesPOJoints(iiPOJoint)
                    Set elesPOFCs = oPOJoint.FrameConnections
                    nPOFCs = elesPOFCs.count
                    For iiPOFC = 1 To nPOFCs
                        Set oOtherFC = elesPOFCs(iiPOFC)
                        If Not oOtherFC.IsMarkedForDelete Then
                            IHStatus = oOtherFC.InputHelper.GetRelatedObjects(oOtherFC, oOtherR1, oOtherR2)
                            'if this PO FC is watching my parent MemberSystem and the MemberSystem I am watching
                            If oOtherR2 Is oMeMS And oOtherR1 Is oR1 Then
                                Set oOtherFC.definition = Nothing
                            End If
                        End If
                    Next iiPOFC
                Next iiPOJoint
                On Error GoTo ErrorHandler
            End If
        End If ' if current definition is not Unsupported.
    
        'if this end of the member is input to can, then remove the member's relation to the Can
        RemoveSecondaryMembFromCanRules FC
        
    Else
        bDisconnectShared = True
    End If
    
    Set oR1 = Nothing
    Set oR2 = Nothing

    Set iWPO = FC.WPO
    portIndex = iWPO.portIndex

    'Delete any AssemblyConnections that are not a child of the memberPart.
    Dim oMemberPart As ISPSMemberPartCommon
    Dim iAppConnection As IJAppConnection
    Dim elesPorts As IJElements
    Dim iIJDesignChild As IJDesignChild
    
    Set oMemberPart = oMeMS.MemberPartAtEnd(portIndex)
    Set iSPSPort = oMemberPart.AxisPort(portIndex)
    Set iIJPort = iSPSPort.Port
    
    'Get the collection of connections on ConnHasPorts relation.   These are ass'y connections.
    
    Call iIJPort.enumConnections(ijEles, 0, 0)
    If Not ijEles Is Nothing Then
        eleCount = ijEles.count
        For ii = eleCount To 1 Step -1    '( decrement, in case Remove affects the list )

            Set oAC = ijEles.Item(ii)
            If Not CheckForReadOnlyAccess(oAC) Then
                If TypeOf oAC Is IJAppConnection Then
                    Set iAppConnection = oAC
                    
                    iAppConnection.enumPorts elesPorts
                    
                    ' there is a case for which this will delete the AC when it should not.
                    ' Place a split, then a splice AC on that split.
                    ' Convert to MemberSystems.  Setting the Unsupported FC on one side of the AxisEnd results in the AC getting deleted.
                    ' This indicates a need to apply this fix logic during compute instead of during connect/disconnect of FCs
                    ' so that all changes are present at that time.  Maybe set a flag in the AC object that tells it to check its topology.
                    
                    ' delete the AC if it has more than one port, unless it is a child of the member part.  ( Generic AC is such a child of MemberPart )
                    If elesPorts.count > 1 Then
                        If (Not TypeOf elesPorts.Item(1) Is IJPlane) And (Not TypeOf elesPorts.Item(2) Is IJPlane) Then
                            If HasNoPortsOnRefColl(oAC) Then
                                Set oIJDObject = oAC
                                oIJDObject.Remove
                                Set oIJDObject = Nothing
                            End If
                        End If
                    End If
                    Set elesPorts = Nothing
                    Set iAppConnection = Nothing
                End If
            End If
            Set oAC = Nothing
        Next ii
    End If
    Set ijEles = Nothing

    'Clear WPO only if the current FC is not Unsupported type. Otherwise
    'we may loose the offsets propagated from the other end due to
    'the align property (TR#107067)
    If Not FC.definition Is Nothing Then
        ' Clear Work Point Offsets and Work Point CP
        iWPO.SetWPO 0#, 0#, 0#, 0#, 0#, 0#
        iWPO.WPOCardinalPoint = 0
    End If
    'Remove previous associations
    If Not FC.IsMarkedForDelete Then
        Dim oCanRuleSupported As ISPSCanRule
        Set oRefColl = GetRefColl(FC)
        Set oCanRuleSupported = FC.GetCrossSectionObject(SPSFCSupported)
        oRefColl.IJDEditJDArgument.RemoveAll
        
        If oCanRuleSupported Is Nothing Then 'no can on the end, so establish relation with memb sys
            If portIndex = SPSMemberAxisStart Then
                oRefColl.IJDEditJDArgument.SetEntity 1, oMeMS, ISPSMemberSystemStartEndNotify, "MemberSysStartNotifyRC_DEST"
            Else
                oRefColl.IJDEditJDArgument.SetEntity 1, oMeMS, ISPSMemberSystemEndEndNotify, "MemberSysEndNotifyRC_DEST"
            End If
        Else
            FC.SetCrossSectionObject SPSFCSupported, oCanRuleSupported
        End If
        
    End If
    
    'If we are setting to Unsupported while another FC is not Unsupported, then
    'get the current PointOn objects so they can be kept, as in PO to a grid line.
    'We check though to make sure it is not a SPS Object Type.
    If bKeepNonMemberPO And RelatedObject1 Is Nothing Then
        FC.Joint.GetPointOn RelatedObject1, RelatedObject2
        If ObjectIsSPSMemberObject(RelatedObject1) Then
            Set RelatedObject1 = Nothing
        End If
        If ObjectIsSPSMemberObject(RelatedObject2) Then
            Set RelatedObject2 = Nothing
        End If
    End If

    'Clear the logical relation with the supporting member.
    'TransferMember to Nothing also calls:  SetPointOn Nothing, Nothing
    If bDisconnectShared Then
        
        FC.Joint.TransferMember oMeMS, Nothing

    End If
    
    SetPathPointOnObjectsForFC FC, RelatedObject1, RelatedObject2
    
    'The above code has insured that RelatedObjects are ok non-Member entities.
    If Not FC.IsMarkedForDelete Then
        FC.Joint.SetPointOn RelatedObject1, RelatedObject2
        FC.Joint.PointOnDimension.PointOnPositionRule = SPSPODimension_Position_Intersection
        FC.Joint.PointOnDimension.PointOnEndRule = SPSPODimension_EndSelection_Auto
    End If

    'clear the relation setup for GAP2 frame connection
    RemoveGAPRelations FC

    ISPSFCInputHelper_SetRelatedObjects = SPSFCInputHelper_Ok
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    ISPSFCInputHelper_SetRelatedObjects = IHStatus
    Exit Function
End Function


'*************************************************************************
'Function
'ISPSFCInputHelper_UserAttributeMgmt
'
'Abstract
'Function on ISPSFACInputHelper to return the UserAttributeMgmt interface
'
'Arguments
'
'Return
'The interface is returned.
'
'Exceptions
'
'***************************************************************************
Private Property Get ISPSFCInputHelper_UserAttributeMgmt() As SPSMembers.IJUserAttributeMgmt
    Set ISPSFCInputHelper_UserAttributeMgmt = Nothing
End Property

'*************************************************************************
'Function
'ISPSFCInputHelper_ValidateLocatedObjects
'
'Abstract
'Validates the objects to be set as input to the ReferenceCollection for the FrameConnection.
'ValidateLocatedObjects is called by the command during Mouse Move events of PlaceMemberSystem command
'Depending on the type of FrameConnection selected in the ribbon bar its repective ValidateLocatedObjects
'function will be called to determine if that object is valid for placing the selected type of FrameConnection
'the located objects are validated and returned as relatedobjects in the argument
'
'Arguments
'FC As ISPSFrameConnection
'LocatedObject1 As Object
'LocatedObject2 As Object
'RelatedObject1 As Object
'RelatedObject2 As Object
'RelatedPositions X,Y,Z as Doubles
'LocatedPositions X,Y,Z as Doubles
'The other arguments are not being used and is for future enhancements
'
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, _
        ByVal options As Long, ByVal snapDistance As Double, _
        ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, _
        ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, _
        RelatedObject1 As Object, RelatedObject2 As Object, _
        RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus

' FC is optional and can be Nothing
' what we want to do here is if snap is enabled, then
' based on LocatedObject1's type we can return the x,y,z of nearby points of interest such as
' grid-intersections or pointOn joints, endpoints

    On Error GoTo ErrorHandler
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    IHStatus = SPSFCInputHelper_UnexpectedError
    ISPSFCInputHelper_ValidateLocatedObjects = SPSFCInputHelper_UnexpectedError
    
    If FC Is Nothing Then ' during smartsketch locate we pass FC as nothing
        IHStatus = SPSFCInputHelper_Ok
        GoTo wrapup
    End If
     
'    Dim oRel1 As Object, oRel2 As Object
'    ReadPathPointOnObjectsForFC FC, oRel1, oRel2
'    If Not oRel1 Is Nothing Then
'        Set LocatedObject1 = oRel1
'        Set LocatedObject2 = oRel2
'    End If

    If ObjectIsSPSMemberObject(LocatedObject1) Then
        Set LocatedObject1 = Nothing
    End If
    If ObjectIsSPSMemberObject(LocatedObject2) Then
        Set LocatedObject2 = Nothing
    End If
    
    If LocatedObject1 Is Nothing And Not LocatedObject2 Is Nothing Then
        Set LocatedObject1 = LocatedObject2
        Set LocatedObject2 = Nothing
    End If
        
    'TODO: should exclude spsmembersystem !
'commenting the code below due to TR#101626 (Middle mousing and then left mousing on the same grid line)
'the code down below already takes care of this situation by setting relateobj2 to nothing
'    If Not LocatedObject1 Is Nothing And Not LocatedObject2 Is Nothing Then
'        If LocatedObject1 Is LocatedObject2 Then
'            IHStatus = SPSFCInputHelper_DuplicateObject
'            GoTo wrapup
'        End If
'    End If

    If Not LocatedObject1 Is Nothing Then
        Set RelatedObject1 = LocatedObject1
    Else
        Set RelatedObject1 = Nothing
    End If
    
    If Not LocatedObject2 Is Nothing Then
        Set RelatedObject2 = LocatedObject2
    Else
        Set RelatedObject2 = Nothing
    End If
    
    'Regardless of above considerations, do not make the FC be related to the same obj twice.
    If Not RelatedObject2 Is Nothing Then
        If Not RelatedObject1 Is Nothing Then
            If RelatedObject1 Is RelatedObject2 Then
                Set RelatedObject2 = Nothing
            End If
        End If
    End If
     
'    Membersystems and Frame connections not allowed for Unsupported type FC
    If ObjectIsSPSMemberObject(LocatedObject1) Then
        Set RelatedObject1 = Nothing
    End If
    If ObjectIsSPSMemberObject(LocatedObject2) Then
        Set RelatedObject2 = Nothing
    End If
    
    RelatedObjectX = LocateX
    RelatedObjectY = LocateY
    RelatedObjectZ = LocateZ
    
    IHStatus = SPSFCInputHelper_Ok

wrapup:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
    Exit Function

ErrorHandler:

    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus

End Function
 
Private Function IsOkPointType(RelatedObject As Object) As Boolean

    Const METHOD = "ISOkPointType"
    On Error GoTo ErrorHandler
    
    Dim geomObj As Object

    IsOkPointType = False
        
    If RelatedObject Is Nothing Then
        GoTo wrapup
    End If

    Set geomObj = RelatedObject

    If TypeOf geomObj Is ISPSSplitMemberConnection Then
        GoTo wrapup

    ElseIf TypeOf geomObj Is IJPoint Then
        IsOkPointType = True
        GoTo wrapup
    
    ElseIf TypeOf RelatedObject Is IJPort Then
        Dim iPort As IJPort
        Set iPort = geomObj
        Set geomObj = Nothing
        Set geomObj = iPort.Geometry
        
    End If

    If Not geomObj Is Nothing Then
    
        If TypeOf geomObj Is IJPoint Then
            IsOkPointType = True
        
        ElseIf TypeOf geomObj Is IJDPosition Then
            IsOkPointType = True
        
        ElseIf TypeOf geomObj Is IJKeyPoint Then
            Dim keypointCount As Long
            Dim iKeypoint As IJKeyPoint
            Set iKeypoint = geomObj
            keypointCount = iKeypoint.keypointCount
            If keypointCount = 1 Then
                IsOkPointType = True
            End If
    
        End If
    End If

wrapup:
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear

End Function

Private Function IsOkSurfaceType(RelatedObject As Object) As Boolean

    Const METHOD = "ISOkSurfaceType"
    On Error GoTo ErrorHandler
    
    Dim geomObj As Object

    IsOkSurfaceType = False
        
    If RelatedObject Is Nothing Then
        GoTo wrapup
    End If
    Set geomObj = RelatedObject

    If TypeOf geomObj Is IJPort Then
        Dim iPort As IJPort
        Set iPort = geomObj
        Set geomObj = Nothing
        Set geomObj = iPort.Geometry
    End If
    
    If Not geomObj Is Nothing Then
        If TypeOf geomObj Is IJSurface Then
            IsOkSurfaceType = True
        ElseIf TypeOf geomObj Is IJPlane Then
            IsOkSurfaceType = True
        End If
    End If

wrapup:
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear

End Function

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

Private Function HasNoPortsOnRefColl(oConnection As Object) As Boolean

    On Error GoTo ErrorHandler
    Const METHOD = "HasNoPortsOnRefColl"

    Dim oObj As Object
    Dim nCount As Long, ii As Long
    Dim oRefColl As IMSSymbolEntities.IJDReferencesCollection
    
    HasNoPortsOnRefColl = True

    Set oRefColl = GetRefCollNoCreate(oConnection)
    If oRefColl Is Nothing Then
        Exit Function
    End If
    
    nCount = oRefColl.IJDEditJDArgument.GetCount
    For ii = 1 To nCount
        Set oObj = oRefColl.IJDEditJDArgument.GetEntityByIndex(ii)
        If Not oObj Is Nothing Then
            If TypeOf oObj Is IJPort Then
                HasNoPortsOnRefColl = False
                Exit For
            End If
        End If
    Next ii

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function
