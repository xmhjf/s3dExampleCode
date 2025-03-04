Attribute VB_Name = "Common"
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : Common.bas
'
'Author : RP
'
'Description :
'    Module for common connection constants/utilities
'
'History:
'   14-April-2004   RP   Added method to check coplanarity of members for vertical corner brace
'   18-April-2006   RP   Added a method to propagate WPO to the other end another to clear
'                                           the offset on the othen end
'   07-July -2006   RP   If alignment is false and the other end is unsupported clear offset on
'                        the other end.
'   07-Sept -2006   RP   Added GetBetaAngle() from axis and orientation vector
'   15-Sept-2006    RP   Moved IsTwoMemberSystemsOkforVCB() to this module from the selector
'                        now checks if all three members are in the same quadrant (TR#104603)
'   25-Sept-2006    MH   TR 104591  Added GetPathPointOnObjectsForFC and SetPathPointOnObjectsForFC
'   21-Nov-2006     JMS  Propagate WPO method was erroring for curved members because it does not support
'                        the ISPSMemberSystemAlignment interface
'   09-May-2007     MOH  TR 120069  added ObjectAssocFlags
'
'   02-Mar-2009     RP   Resolved CR153139 - Axial Gap Both - doesn't work when end is constrained to elevation plane
'   01-Jul-2009     RP   CR165819 - Added RemoveSecondaryMembFromCanRules() to remove the member's association to Cans                                       association to CanRules
'*********************************************************************************************************************
Option Explicit
Private Const MODULE = "Common"
Public Const IJPlane = "{4317C6B3-D265-11D1-9558-0060973D4824}"
Public Const strISPSAxisEndPort = "{D1136F2C-51A0-4FD2-A6FE-E0E7DD59EC8A}"
Public Const strISPSFrameConnection = "{09C7178A-A735-4692-B713-E4253FBF6397}"
Public Const strAxisEndAlong1Dest = "AxisEndAxisAlong1_DEST"
Public Const strAxisEndAlong2Dest = "AxisEndAxisAlong2_DEST"
Public Const strAxisEndAxisEndDest = "AxisEndAxisEnd_DEST"
Public Const strAxisEndAxisEndOrig = "AxisEndAxisEnd_ORIG"
Public Const strFCToGapOutputOrig = "FCToGapOutput_ORIG"
Public Const lFCAxisAlong = 1
Public Const lFCAxisEnd = 2
Public Const lFCAxisColinear = 3
Public Const SPSvbError = vbObjectError + 512
Public Const E_FAIL As Long = -2147467259
'*************************************************************************
'Function
'
'<SetRefColl>
'
'Abstract
'
'<Adds the object to the ReferencesCollection Object>
'
'Arguments
'
'<FrameConnection businessobject as Object, refcoll object As IJDReferencesCollection>
'
'Return

'Exceptions
'***************************************************************************
Private Sub SetRefColl(pFC As Object, pRefColl As IJDReferencesCollection)
Const MT = "SetRefColl"
 On Error GoTo ErrorHandler
  
   'connect the reference collection to the smart occurrence
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pFC
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    pCollectionHelper.Add pRefColl, "RC", pRelationshipHelper
    Set pRevision = New JRevision
    pRevision.AddRelationship pRelationshipHelper
  
   Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'
'<GetRefColl>
'
'Abstract
'
'<Gets the object from the ReferencesCollection Object>
'
'Arguments
'
'<FrameConnection businessobject as Object>
'
'Return
'
'<refcoll object As IJDReferencesCollection>
'
'Exceptions
'***************************************************************************
Public Function GetRefColl(pFC As Object) As IJDReferencesCollection
Const MT = "GetRefColl"
 On Error GoTo ErrorHandler
  
   'traverse the relation from the SO to the RefColl
   'if none exists, make one and connect it
   
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim oObject As IJDObject
    Dim oSymbolEntitiesFactory  As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim count As Long
    
    Set pRelationHelper = pFC
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    If Not pCollectionHelper Is Nothing Then
        count = pCollectionHelper.count
        If count >= 1 Then
            Set GetRefColl = pCollectionHelper.Item(count)
        End If
    End If
  
    If GetRefColl Is Nothing Then
        Set oObject = pFC
        Set GetRefColl = oSymbolEntitiesFactory.CreateEntity(referencesCollection, oObject.ResourceManager)
        SetRefColl pFC, GetRefColl
    End If

 Exit Function
 
ErrorHandler:
    HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'AreMembersInOnePlane
'
'Abstract
'
'Checks axes of three member systems are in the same plane in the
'context of CornerBrace Frame connection
'
'Arguments
'
'Supported Member system, first Supporting member system, second supporting member system
'
'Return
'   True - If they are coplanar
'   False - If they are not coplanar
'***************************************************************************
Public Function AreMembersInOnePlane(oSupped As ISPSMemberSystemLinear, oSupping1 As ISPSMemberSystemLinear, _
oSupping2 As ISPSMemberSystemLinear) As Boolean
    Const MT = "AreMembersInOnePlane"
    On Error GoTo ErrorHandler
    
    Dim oLine0 As IJLine, oLine1 As IJLine, oLine2 As IJLine
    Dim oVec0 As New DVector, oVec1 As New DVector, oVec2 As New DVector
    Dim oVec1xVec2 As IJDVector, oVec0xVec2 As IJDVector
    Dim oAxis As ISPSLogicalAxis
    Dim uX#, uY#, uZ#, X0#, Y0#, Z0#, x1#, y1#, z1#
    Dim cos As Double
    
    AreMembersInOnePlane = False
    'get the physical axis of supping1 and supping2
    Set oLine1 = oSupping1
    Set oLine2 = oSupping2
    
    'get the logical axis of the supported, because the logical axis of the supported
    'is offsetted relative to the physical axis of the supporting
    Set oAxis = oSupped
    
    oLine1.GetDirection uX, uY, uZ
    oVec1.Set uX, uY, uZ
    'normalize
    oVec1.Length = 1
    
    oLine2.GetDirection uX, uY, uZ
    oVec2.Set uX, uY, uZ
    'normalize
    oVec2.Length = 1
    
    Set oVec1xVec2 = oVec1.Cross(oVec2)
    'normalize
    oVec1xVec2.Length = 1
    
    oAxis.GetLogicalStartPoint X0, Y0, Z0
    oAxis.GetLogicalEndPoint x1, y1, z1
    
    oVec0.Set x1 - X0, y1 - Y0, z1 - Z0
    'normalize
    oVec0.Length = 1
    
    Set oVec0xVec2 = oVec0.Cross(oVec2)
    oVec0xVec2.Length = 1
    cos = oVec1xVec2.Dot(oVec0xVec2)
    'if oVec1xVec2 and oVec0xVec2 are parallel then all three members are in one plane
    If Abs(Abs(cos) - 1) < angleTol Then
        AreMembersInOnePlane = True
    End If
    
    Exit Function
ErrorHandler:
    AreMembersInOnePlane = False
End Function

Public Function HasMeAsSupportingMember(oMeMS As ISPSMemberSystem, oSingFC As ISPSFrameConnection) As Boolean

    Const MT = "HasMeAsSupportingMember"
    On Error GoTo ErrorHandler

''    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim oSingSingMS As ISPSMemberSystem
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim count As Long
    Dim oRelatedObject As Object

    HasMeAsSupportingMember = False

'' if supporting FC is marked for delete, then no need to report that oMeMS is its supping MS
''
    If Not oSingFC.IsMarkedForDelete Then
        Set oRC = GetRefColl(oSingFC)
        count = oRC.IJDEditJDArgument.GetCount
        
        If count > 1 Then
    
            Set oRelatedObject = oRC.IJDEditJDArgument.GetEntityByIndex(2)     ' supporting member system
            
            If oRelatedObject Is oMeMS Then
                HasMeAsSupportingMember = True
    
            ElseIf count > 2 Then
                Set oRelatedObject = oRC.IJDEditJDArgument.GetEntityByIndex(3)
                
                If oRelatedObject Is oMeMS Then
                    HasMeAsSupportingMember = True
                End If
            End If
        End If
    End If

''  09/25/09. This is called by SetRelatedObjects, after joints and refColl of the supported FC is set.
''  Commented out usage of GetRelatedObjects since GetRelatedObjects checks integrity of the supping FC.
''  In the case of AxisEnd, we might be in an in-between state, so it is more robust to use the supping refColl directly.
''
''    IHStatus = oSingFC.InputHelper.GetRelatedObjects(oSingFC, oR1, oR2)
''
''    If IHStatus <> SPSFCInputHelper_Ok Then
''        HasMeAsSupportingMember = False
''
''    ElseIf oR1 Is oMeMS Or oR2 Is oMeMS Then
''        HasMeAsSupportingMember = True
''
''----------------------
    'If either of this FC's supporting objects is a MemberSystem, then we need to make sure
    'that none of it's supporting objects is the given member system.
    'HOWEVER, it is commented out for several reasons:
    '(1) When modifying a FC that has many ancestors ( on a tertiary member ) the search
    '    will traverse a whole structure before it confirms that the user is not trying to
    '    attach a column onto that tertiary member.
    '(2) With this in place, then one of the FC's of the supporting member is set to Unsupported -
    '    which is unseen and unexpected.
    '(3) Such a loop will be detected by assoc, which will abort the transaction.
    'To test this scenario:
    'Place several members in "route-mode."  Then take the first FC and attempt to make it dependent
    'on the last member placed.  With this code active, that will succeed, but the last member will
    'have it's FC by which it had a relationship with the first member set to Unsupported.
    '
'
'    Else
'        If Not oR1 Is Nothing Then
'            If TypeOf oR1 Is ISPSMemberSystem Then
'                Set oSingSingMS = oR1
'                If HasMeAsSupportingMember(oMeMS, oSingSingMS.FrameConnectionAtEnd(SPSMemberAxisStart)) Then
'                    HasMeAsSupportingMember = True
'                ElseIf HasMeAsSupportingMember(oMeMS, oSingSingMS.FrameConnectionAtEnd(SPSMemberAxisEnd)) Then
'                    HasMeAsSupportingMember = True
'                End If
'            End If
'        End If
'        If Not oR2 Is Nothing Then
'            If TypeOf oR2 Is ISPSMemberSystem Then
'                Set oSingSingMS = oR2
'                If HasMeAsSupportingMember(oMeMS, oSingSingMS.FrameConnectionAtEnd(SPSMemberAxisStart)) Then
'                    HasMeAsSupportingMember = True
'                ElseIf HasMeAsSupportingMember(oMeMS, oSingSingMS.FrameConnectionAtEnd(SPSMemberAxisEnd)) Then
'                    HasMeAsSupportingMember = True
'                End If
'            End If
'        End If
'    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function


Public Sub RemoveCommonSplitConnections(oMS1 As ISPSMemberSystem, oMS2 As Object)
    Const MT = "RemoveCommonSplitConnections"
    On Error GoTo ErrorHandler
    
    'remove any SplitConnections that are common between these MemberSystems

    Dim oIJDObject As IJDObject
    Dim ii As Long, count As Long
    Dim SCparents As IJElements, SCs As IJElements
    Dim SC As ISPSSplitMemberConnection
    
    If oMS1 Is Nothing Then
        Exit Sub
    End If
    If oMS2 Is Nothing Then
        Exit Sub
    End If
    
    Set SCs = oMS1.SplitConnections
    count = SCs.count
    For ii = count To 1 Step -1
        Set SC = SCs(ii)
        Set SCparents = SC.InputObjects
        If SCparents.Contains(oMS2) Then
            Set oIJDObject = SC
            oIJDObject.Remove
        End If
    Next ii

    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Function ObjectIsOkSurface(oInputObj As Object) As Boolean

    Const METHOD = "ObjectIsOkSurface"
    On Error GoTo ErrorHandler

    Dim oGeom As Object
    Dim oPort As IJPort
    Dim bOk As Boolean

    bOk = False

    If Not oInputObj Is Nothing Then

        If TypeOf oInputObj Is IJDynamicSurfaceFind Then
            Set oGeom = oInputObj
        ElseIf TypeOf oInputObj Is IJPort Then
            Set oPort = oInputObj
            Set oGeom = oPort.Geometry
        Else
            Set oGeom = oInputObj
        End If
        
        If oGeom Is Nothing Then
            bOk = False
        ElseIf TypeOf oGeom Is IJSurface Then
            bOk = True
        ElseIf TypeOf oGeom Is IJPlane Then
            bOk = True
        ElseIf TypeOf oGeom Is IJDynamicSurfaceFind Then
            bOk = True
        End If
    End If

    ' any surface port whose connectable is a MemberPart is not a valid Surface for FC's
    If Not oPort Is Nothing Then
        Set oGeom = oPort.Connectable
        If Not oGeom Is Nothing Then
            If TypeOf oGeom Is ISPSMemberPartPrismatic Then
                bOk = False
            End If
        End If
    End If

    ObjectIsOkSurface = bOk
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    ObjectIsOkSurface = False
    Exit Function
End Function

' "Transfer" joint to Nothing will create a new joint so its not shared with another MS.
' Heres the deal ...
' we are about to set the joint PointOn to a MemberSystem.  If the joint is shared,
' we must create a new joint at the FC's end, unless all non-self FC's have
' related objects which include the MemberSystems at the joint.
'
' Example: two members systems share a joint.  One FC is Unsupported, other is AxisEnd.
' Setting the Unsupported to AxisAlong does not require transfer.
' Setting the AxisEnd to AxisAlong does require transfer.

Public Sub ConditionallyTransferJoint(oMeFC As ISPSFrameConnection, Optional oSuppingMember As Object)

    Const METHOD = "ConditionallyTransferJoint"
    On Error GoTo ErrorHandler

    Dim oMeJoint As ISPSAxisJoint
    Dim oMeMemberSystem As ISPSMemberSystem
    Dim oOtherFC As ISPSFrameConnection
    Dim elesEndMemberSystems As IJElements
    Dim elesFCs As IJElements
    Dim bDoTransfer As Boolean
    Dim IHStatus As SPSFCInputHelperStatus
    Dim oR1 As Object, oR2 As Object
    Dim ii As Long, count As Long
    Dim jj As Long, countFCs As Long
    
    Set oMeJoint = oMeFC.Joint
    Set elesFCs = oMeJoint.FrameConnections

    Set elesEndMemberSystems = oMeJoint.EndMemberSystems
    
    bDoTransfer = False
    
    count = elesFCs.count

    For ii = 1 To count
    
        Set oOtherFC = elesFCs(ii)
        
        If Not oOtherFC Is oMeFC Then
            
            IHStatus = oOtherFC.InputHelper.GetRelatedObjects(oOtherFC, oR1, oR2)
            
            If IHStatus <> SPSFCInputHelper_Ok Then
                bDoTransfer = True
            ElseIf oR1 Is Nothing Then
                bDoTransfer = True

            ElseIf Not elesEndMemberSystems.Contains(oR1) Then  ' otherFC's supping member not in set of shared MS's
                bDoTransfer = True

            ElseIf Not oSuppingMember Is Nothing Then           ' AxisEnd sends in suppingMember.  cannot share with that MS.
                If elesEndMemberSystems.Contains(oSuppingMember) Then
                    bDoTransfer = True
                End If
                
            End If
            
            If bDoTransfer Then
                Exit For
            End If
        End If

        Set oR1 = Nothing
        Set oR2 = Nothing
    
    Next ii
    
    If bDoTransfer Then
        
''        Dim elesMePOJointsToMove As IJElements
''        Dim elesMePOJoints As IJElements
''        Dim oJoint As ISPSAxisJoint
''        Dim oFC As ISPSFrameConnection
''
''        Set elesMePOJointsToMove = New JObjectCollection
''        Set oMeMemberSystem = oMeFC.MemberSystem
''
''        Set elesMePOJoints = oMeJoint.PointOnJoints
''        count = elesMePOJoints.count
''
''        'Before doing the TransferJoint to Nothing - which creates a new joint - obtain a list of joints
''        'that are PO to the FC's current Joint, and need to be made PO to the newly created joint.  It
''        'needs to be transferred when that joints' FCs use this MemberSystem as a supporting member.
''
''        For ii = 1 To count
''
''            Set oJoint = elesMePOJoints(ii)
''            Set elesFCs = oJoint.FrameConnections
''
''            countFCs = elesFCs.count
''
''            For jj = 1 To countFCs
''
''                Set oOtherFC = elesFCs(jj)
''                Set oR1 = Nothing
''                Set oR2 = Nothing
''                IHStatus = oOtherFC.InputHelper.GetRelatedObjects(oOtherFC, oR1, oR2)
''                If IHStatus = SPSFCInputHelper_Ok Then
''                    If oR1 Is oMeMemberSystem Or oR2 Is oMeMemberSystem Then
''                        elesMePOJointsToMove.Add oJoint
''                        GoTo nextPOJoint
''                    End If
''                End If
''            Next jj
''nextPOJoint:
''        Next ii
            
        'make sure that oMeMemberSystem has its own joint since prev def was sharing the joint with another MS
        oMeJoint.TransferMember oMeFC.MemberSystem, Nothing
    
''        'transfer joints found above to the newly created joint
''        SetJointsPointOn elesMePOJointsToMove, oMeFC.Joint, Nothing
        
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise Err.Number
    Exit Sub
End Sub

Public Function GetStablePort(objPort As Object) As Object
    Const MT = "GetStablePort"
    On Error GoTo ErrorHandler

    Dim iIJPort As IJPort
    Dim oConnectable As Object
    Dim iIJStablePort As IJStablePort
    Dim oOutput As Object, oStableOutput As Object
    
    Set oOutput = Nothing
    If Not objPort Is Nothing Then
        Set oOutput = objPort
        If TypeOf objPort Is IJPort Then
            Set iIJPort = objPort
            Set oConnectable = iIJPort.Connectable
            If TypeOf oConnectable Is IJStablePort Then 'true for slab,wall,member part (SP3d structure objects)
                Set iIJStablePort = oConnectable
                On Error Resume Next
                Set oStableOutput = iIJStablePort.StablePort(objPort)
                Err.Clear
                On Error GoTo ErrorHandler
                If oStableOutput Is Nothing Then
                    Set oStableOutput = objPort
                End If
                Set oOutput = oStableOutput
            'remove the else block once PlateSystem implements IJStablePort(DI#237012)
            Else ' for instance, a marine plate system
                Dim oCommonConnectable As IJCommonStructEntity
                Dim oBoundingPort As IJPort
                On Error Resume Next ' guard against the QI or the FindEquivalentPort call failure
                Set oCommonConnectable = oConnectable
                If Not oCommonConnectable Is Nothing Then
                    oCommonConnectable.FindEquivalentPort oBoundingPort, iIJPort, "PlateEntity.PlateBound_AE.1", False
                End If
                On Error GoTo ErrorHandler
                
                If Not oBoundingPort Is Nothing Then
                    Set oOutput = oBoundingPort
                End If
            End If
            
        End If
    End If

    Set GetStablePort = oOutput
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<ClearWPO>
'
'Abstract
'
'Clears the WPO associated to the FC if the other end is unsupported and align is false. If align
'is true and other end is not unsupported then copy the WPO from the other end
'
'Arguments
'
'<FrameConnection businessobject as Object>
'
'Return
'
'
'
'Exceptions

Public Sub ClearWPO(iFC As ISPSFrameConnection)
    Const MT = "ClearWPO"
    On Error GoTo ErrorHandler
    
    If Not iFC Is Nothing Then
        Dim iWPO As ISPSAxisWPO
        ' Clear Work Point Offsets and Work Point CP
        Set iWPO = iFC.WPO
        iWPO.SetWPO 0#, 0#, 0#, 0#, 0#, 0#
        iWPO.WPOCardinalPoint = 0

        Dim oAlign As ISPSMemberSystemAlignment
        Dim oSupportedMem As ISPSMemberSystem
        
        Set oSupportedMem = iFC.MemberSystem
        If TypeOf oSupportedMem Is ISPSMemberSystemAlignment Then
            Set oAlign = oSupportedMem
            'based on the align flag and FC on the other end, either copy   offsets from
            'that end to this end or clear that end's offset
            If oAlign.Align = True Then
                Dim gX As Double, gY As Double, gZ As Double, lX As Double, lY As Double, lZ As Double
                Dim offsetCP As Long
                'get FC on the other end
                Dim oFCotherEnd As ISPSFrameConnection
                Dim oWPOotherEnd As ISPSAxisWPO
                If iFC.WPO.portIndex = SPSMemberAxisStart Then
                    Set oFCotherEnd = oSupportedMem.FrameConnectionAtEnd(SPSMemberAxisEnd)
                    Set oWPOotherEnd = oSupportedMem.WPOAtEnd(SPSMemberAxisEnd)
                Else
                    Set oFCotherEnd = oSupportedMem.FrameConnectionAtEnd(SPSMemberAxisStart)
                    Set oWPOotherEnd = oSupportedMem.WPOAtEnd(SPSMemberAxisStart)
                End If
                If Not oFCotherEnd Is Nothing Then
                    If Not oFCotherEnd.definition Is Nothing Then
                        'copy the offset from the other end
                        oWPOotherEnd.GetWPO gX, gY, gZ, lX, lY, lZ
                        oWPOotherEnd.WPOCardinalPoint = offsetCP
                        iWPO.SetWPO gX, gY, gZ, lX, lY, lZ
                        iWPO.WPOCardinalPoint = offsetCP
                    Else
                        'clear the other end's offset
                        oWPOotherEnd.SetWPO 0, 0, 0, 0, 0, 0
                        oWPOotherEnd.WPOCardinalPoint = 0
                    End If
                Else
                    'probably the FC on the other end got deleted
                    'clear the other end's offset if there is any
                    oWPOotherEnd.SetWPO 0, 0, 0, 0, 0, 0
                    oWPOotherEnd.WPOCardinalPoint = 0
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT

End Sub

'*********************************************************************************************
'Function
'
'<PropagateWPOToOtherEnd>
'
'Abstract
'
'If the align flag is true and on the other end if there is no FC or if the FC unsupported
'then the offset from this end is copied to the other end. If align is false and the other
'end is unsupported then the offset on othe other end is cleared.
'
'Arguments
'
'<FrameConnection businessobject as Object>
'
'Return
'
'
'
'Exceptions
'*********************************************************************************************
Public Sub PropagateWPOToOtherEnd(iFC As ISPSFrameConnection)
    Const MT = "PropagateWPOToOtherEnd"
    On Error GoTo ErrorHandler
    
    If Not iFC Is Nothing Then
        Dim WhichEnd As SPSMemberAxisPortIndex
        Dim oAlign As ISPSMemberSystemAlignment
        Dim oSupportedMem As ISPSMemberSystem
        Dim globalOffsetX As Double, globalOffsetY As Double, globalOffsetZ As Double
        Dim localOffsetX As Double, localOffsetY As Double, localOffsetZ As Double
        Dim oWPO As ISPSAxisWPO
        
        Set oSupportedMem = iFC.MemberSystem
               
        ' Curved members do not support the alignment interface so do nothing if this
        '   interface is not supported
        If TypeOf oSupportedMem Is ISPSMemberSystemAlignment Then
            WhichEnd = iFC.WPO.portIndex
            Set oWPO = iFC.WPO
            Set oAlign = oSupportedMem
            
            'get FC and wpo on the other end
            Dim oFCotherEnd As ISPSFrameConnection
            Dim oWPOotherEnd As ISPSAxisWPO
            If WhichEnd = SPSMemberAxisStart Then
                Set oFCotherEnd = oSupportedMem.FrameConnectionAtEnd(SPSMemberAxisEnd)
                Set oWPOotherEnd = oSupportedMem.WPOAtEnd(SPSMemberAxisEnd)
            Else
                Set oFCotherEnd = oSupportedMem.FrameConnectionAtEnd(SPSMemberAxisStart)
                Set oWPOotherEnd = oSupportedMem.WPOAtEnd(SPSMemberAxisStart)
            End If
            
            oWPO.GetWPO globalOffsetX, globalOffsetY, globalOffsetZ, localOffsetX, localOffsetY, localOffsetZ
            
            If oAlign.Align = True Then
                
                If Not oFCotherEnd Is Nothing Then
                    If oFCotherEnd.definition Is Nothing Then
                        'unsupported FC so propagate the offset to the other end
                        oWPOotherEnd.SetWPO globalOffsetX, globalOffsetY, globalOffsetZ, localOffsetX, localOffsetY, localOffsetZ
                        oWPOotherEnd.WPOCardinalPoint = oWPO.WPOCardinalPoint
                    End If
                Else
                    'no FC on the other end,so Propagate the offset to the other end
                    oWPOotherEnd.SetWPO globalOffsetX, globalOffsetY, globalOffsetZ, localOffsetX, localOffsetY, localOffsetZ
                    oWPOotherEnd.WPOCardinalPoint = oWPO.WPOCardinalPoint
                End If
            Else ' set offset to zero on the other end if it is unsupported
                If Not oFCotherEnd Is Nothing Then
                    If oFCotherEnd.definition Is Nothing Then
                        'unsupported FC so propagate the offset to the other end
                        oWPOotherEnd.SetWPO 0, 0, 0, 0, 0, 0
                        oWPOotherEnd.WPOCardinalPoint = 0
                    End If
                Else
                    'no FC on the other end,so Propagate the offset to the other end
                    oWPOotherEnd.SetWPO 0, 0, 0, 0, 0, 0
                    oWPOotherEnd.WPOCardinalPoint = 0
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT

End Sub

' user can select a memberPart.  if it is a part, then return it's parent member system, if exists.
' else return the given LocatedObject.

Public Function SwitchLocatedPartToMemberSystem(LocatedObject As Object) As Object
    Const MT = "SwitchLocatedPartToMemberSystem"
    On Error GoTo ErrorHandler

    Dim oPartPris As ISPSMemberPartPrismatic

    Set SwitchLocatedPartToMemberSystem = LocatedObject

    If LocatedObject Is Nothing Then Exit Function

    If TypeOf LocatedObject Is ISPSMemberSystem Then Exit Function
        
    If TypeOf LocatedObject Is ISPSMemberPartPrismatic Then
        Set oPartPris = LocatedObject
        Set SwitchLocatedPartToMemberSystem = oPartPris.MemberSystem

    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, MT
    Exit Function
End Function

Public Function SwitchLocatedFCToMemberSystemForPointOn(LocatedObject As Object) As Object
    Const MT = "SwitchLocatedFCToMemberSystemForPointOn"
    On Error GoTo ErrorHandler

    Dim oFC As ISPSFrameConnection

    Set SwitchLocatedFCToMemberSystemForPointOn = LocatedObject

    If LocatedObject Is Nothing Then Exit Function

    If TypeOf LocatedObject Is ISPSFrameConnection Then
        Set oFC = LocatedObject
        Set SwitchLocatedFCToMemberSystemForPointOn = oFC.MemberSystem
        'this can be changed to this:
        'oFC.InputHelper.GetRelatedObjects(oFC, oObj1, oObj2 )
        'if not oObj1 is nothing then
        '   if typeof oObj1 is ISPSMemberSytsem then
        '       SwitchLocatedFCToMemberSystemForPointOn = oObj1
        '
    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, MT
    Exit Function
End Function

Public Function ExtractValidRelatedObjectFromLocatedObject(FC As ISPSFrameConnection, linearOnly As Boolean, LocatedObject As Object, ByRef RelatedObject As Object) As SPSMembers.SPSFCInputHelperStatus

    On Error GoTo ErrorHandler
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus

    Dim oFC As ISPSFrameConnection
    Dim oFCMembSys As ISPSMemberSystem
    Dim oPart As ISPSMemberPartPrismatic

    IHStatus = SPSFCInputHelper_UnexpectedError

    Set RelatedObject = Nothing

    If LocatedObject Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        GoTo wrapup
    End If

    If TypeOf LocatedObject Is ISPSMemberPartPrismatic Then
        Set oPart = LocatedObject
        Set LocatedObject = oPart.MemberSystem
    ElseIf TypeOf LocatedObject Is ISPSFrameConnection Then
        Set oFC = LocatedObject
        Set LocatedObject = oFC.MemberSystem
    ElseIf TypeOf LocatedObject Is ISPSSplitMemberConnection Then
        Dim ii As Long
        Dim eles As IJElements
        Dim oSC As ISPSSplitMemberConnection
        Dim oSCResult As IJStructILCConnectionResult
        Dim iSplitStatus As SplitStatus
        
        Set oSC = LocatedObject
        Set eles = oSC.InputObjects
        If eles.count = 2 Then
            Set oSCResult = oSC
            iSplitStatus = oSCResult.SplitParentStatusResult
            If iSplitStatus = ssSplitSecond Then
                Set LocatedObject = eles(1)
            Else
                Set LocatedObject = eles(2)
            End If
        End If
    End If

    'Always return the "located" object even if not a MembSys
    Set RelatedObject = LocatedObject

    If linearOnly And Not TypeOf LocatedObject Is ISPSMemberSystemLinear Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
    
    ElseIf Not TypeOf LocatedObject Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
    
    ElseIf FC Is Nothing Then
        IHStatus = SPSFCInputHelper_Ok
    
    ElseIf LocatedObject Is FC.MemberSystem Then        'LocatedObject is supporting and cannot also be supported
        IHStatus = SPSFCInputHelper_DuplicateObject
        
    Else
        IHStatus = SPSFCInputHelper_Ok
    
    End If

wrapup:
    ExtractValidRelatedObjectFromLocatedObject = IHStatus
    Exit Function

ErrorHandler:
    Err.Clear
    ExtractValidRelatedObjectFromLocatedObject = IHStatus
End Function

'*********************************************************************************************
'Function
'
'<GetBetaAnagle>
'
'Abstract
'
'Computes beta angle from the axis vector  and orientation vector
'
'Arguments
'
'<axisVector as IJDVector, orientationVector as IJDVector>
'
'Return
'<betaAngle as double>
'
'
'
'Exceptions
'*********************************************************************************************

Public Function GetBetaAngle(oAxis As IJDVector, zVec As IJDVector) As Double
    Dim bVertical As Boolean
    Dim oMat As iJDT4x4
    Dim vecRefAxis As IJDVector
    Dim xLoc As IJDVector, yLoc As IJDVector, zLoc As IJDVector


    'interpret vertical using the logical axis. It is unaffected by offsets.
    'if vertical, use the global x-axis as reference, otherwise use Z-axis.
    If ((Abs(oAxis.x) < distTol) And (Abs(oAxis.y) < distTol)) Then
        If (Abs(oAxis.z) < distTol) Then
            Err.Raise E_FAIL
        End If
        bVertical = True
    End If

    'get betaAngle as the angle of the reference-axis in the member's coordinate system.
    Set xLoc = oAxis.Clone
    xLoc.Length = 1
    
    Set zLoc = zVec.Clone
    zLoc.Length = 1
    'compute y from x and z
    Set yLoc = zLoc.Cross(xLoc)
    yLoc.Length = 1
    
    'compute z again as previous z may not be at right angle with x
    Set zLoc = xLoc.Cross(yLoc)
    zLoc.Length = 1
        
    Set oMat = New DT4x4
    
    oMat.LoadIdentity
    oMat.IndexValue(0) = xLoc.x
    oMat.IndexValue(1) = xLoc.y
    oMat.IndexValue(2) = xLoc.z
    
    oMat.IndexValue(4) = yLoc.x
    oMat.IndexValue(5) = yLoc.y
    oMat.IndexValue(6) = yLoc.z
    
    oMat.IndexValue(8) = zLoc.x
    oMat.IndexValue(9) = zLoc.y
    oMat.IndexValue(10) = zLoc.z
    
    Set vecRefAxis = New DVector
    If (bVertical) Then
        vecRefAxis.Set 1, 0, 0
    Else
        vecRefAxis.Set 0, 0, 1
    End If
    
    oMat.Invert 'global to local
    Set vecRefAxis = oMat.TransformVector(vecRefAxis)
    'what is the angle of the reference axis in local coord's and in yz plane.
    GetBetaAngle = Atn(vecRefAxis.y / vecRefAxis.z)
End Function

'determine if the two given objects are ok to use for Vertical Corner Brace
                
Public Function IsTwoMemberSystemsOkforVCB(oSuppedFC As ISPSFrameConnection, oCv1 As IJCurve, oCv2 As IJCurve) As Boolean

    Const METHOD = "IsTwoMemberSystemsOkforVCB"
    On Error GoTo ErrorHandler
    Dim dotp As Double
    Dim parAlongCv As Double, x As Double, y As Double, z As Double
    Dim posx As Double, posy As Double, posz As Double
    Dim vecTanX As Double, vecTanY As Double, vecTanZ As Double
    Dim vecTan2X As Double, vecTan2Y As Double, vecTan2Z As Double
    Dim vecTan1 As IJDVector, vecTan2 As IJDVector
    Dim sX As Double, sY As Double, sZ As Double, eX As Double, eY As Double, eZ As Double
    Dim oPos1 As New DPosition, oPos2 As New DPosition, oPos3 As New DPosition
    Dim oVec1 As IJDVector, oVec2 As IJDVector
    Dim idxPort As SPSMemberAxisPortIndex
    Dim iAxis As ISPSLogicalAxis
    
    IsTwoMemberSystemsOkforVCB = False
    
    'The supporting objects not ok for VCB unless...
    '   Both supporting system's are Linear
    '   The supported system is linear
    '   The supporting member systems are not parallel
    '   Supported and Both suporting are in one plane
    
    If TypeOf oCv1 Is ISPSMemberSystemLinear And TypeOf oCv2 Is ISPSMemberSystemLinear And _
            TypeOf oSuppedFC.MemberSystem Is ISPSMemberSystemLinear Then
    
        oSuppedFC.Joint.Point.GetPoint x, y, z
        
        oCv1.Parameter x, y, z, parAlongCv
        oCv1.Evaluate parAlongCv, posx, posy, posz, vecTanX, vecTanY, vecTanZ, vecTan2X, vecTan2Y, vecTan2Z
        Set vecTan1 = New DVector
        vecTan1.Set vecTanX, vecTanY, vecTanZ
        vecTan1.Length = 1#
    
        oCv2.Parameter x, y, z, parAlongCv
        oCv2.Evaluate parAlongCv, posx, posy, posz, vecTanX, vecTanY, vecTanZ, vecTan2X, vecTan2Y, vecTan2Z
        Set vecTan2 = New DVector
        vecTan2.Set vecTanX, vecTanY, vecTanZ
        vecTan2.Length = 1#
    
        dotp = vecTan2.Dot(vecTan1)
        If Abs(dotp) < 0.8 Then                     'the two supping's are not parallel
            If AreMembersInOnePlane(oSuppedFC.MemberSystem, oCv1, oCv2) Then
                'let us check if all 3 members are in the same quadrant
  
                'get FC point
                oSuppedFC.Joint.Point.GetPoint x, y, z
                oPos1.Set x, y, z
                
                idxPort = oSuppedFC.WPO.portIndex
                
                Set iAxis = oSuppedFC.MemberSystem.LogicalAxis
                iAxis.GetLogicalStartPoint sX, sY, sZ
                oPos2.Set sX, sY, sZ
                iAxis.GetLogicalEndPoint eX, eY, eZ
                oPos3.Set eX, eY, eZ
                
                'get vector from the connected end to the other end for supported
                If idxPort = SPSMemberAxisStart Then
                    Set oVec1 = oPos3.Subtract(oPos2)
                Else
                    Set oVec1 = oPos2.Subtract(oPos3)
                End If
                
                'check if FC  is at the end of the first supporting
                oCv1.EndPoints sX, sY, sZ, eX, eY, eZ
                
                oPos2.Set sX, sY, sZ
                oPos3.Set eX, eY, eZ
                
                If oPos1.DistPt(oPos2) < distTol Then
                    'supporting1 has its start end at the connection
                    Set oVec2 = oPos3.Subtract(oPos2)
                ElseIf oPos1.DistPt(oPos3) < distTol Then
                    'supporting1 has its end end at the connection
                    Set oVec2 = oPos2.Subtract(oPos3)
                End If
                
                If Not oVec2 Is Nothing Then
                    If oVec2.Dot(oVec1) < 0 Then
                        'supported and supporting1 are not in same quadrant
                        Exit Function
                    End If
                    Set oVec2 = Nothing
                End If
                
                'check if FC is at the end of the second supporting
                oCv2.EndPoints sX, sY, sZ, eX, eY, eZ
                
                oPos2.Set sX, sY, sZ
                oPos3.Set eX, eY, eZ
                
                If oPos1.DistPt(oPos2) < distTol Then
                    'supporting2 has its start end at the connection
                    Set oVec2 = oPos3.Subtract(oPos2)
                ElseIf oPos1.DistPt(oPos3) < distTol Then
                    'supporting2 has its end end at the connection
                    Set oVec2 = oPos2.Subtract(oPos3)
                End If
                If Not oVec2 Is Nothing Then
                    If oVec2.Dot(oVec1) < 0 Then
                        'supported and supporting2 are not in same quadrant
                        Exit Function
                    End If
                    Set oVec2 = Nothing
                End If
                
                IsTwoMemberSystemsOkforVCB = True
            End If
        End If

    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    IsTwoMemberSystemsOkforVCB = False
End Function


Public Sub SetPathPointOnObjectsForFC(oFC As ISPSFrameConnection, oRel1 As Object, oRel2 As Object)

    Const METHOD = "SetPathPointOnObjectsForFC"
    On Error GoTo ErrorHandler

    Dim oMF As New SPSMembers.SPSMemberFactory
    Dim oConnServices As ISPSMemberConnectionServices
        
    Set oConnServices = oMF.CreateConnectionServices

    oConnServices.SetPathPointOnObjectsForFC oFC, oRel1, oRel2

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
End Sub

Public Function ReadPathPointOnObjectsForFC(oFC As ISPSFrameConnection, ByRef oRel1 As Object, ByRef oRel2 As Object) As Boolean

    Const METHOD = "ReadPathPointOnObjectsForFC"
    On Error GoTo ErrorHandler

    Dim bMemberSystemHasPath As Boolean
    Dim oMF As New SPSMembers.SPSMemberFactory
    Dim oConnServices As ISPSMemberConnectionServices
        
    Set oConnServices = oMF.CreateConnectionServices

    oConnServices.ReadPathPointOnObjectsForFC oFC, bMemberSystemHasPath, oRel1, oRel2

    ReadPathPointOnObjectsForFC = bMemberSystemHasPath
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
End Function

Public Function ObjectIsSPSMemberObject(oTest As Object) As Boolean

    Const METHOD = "ObjectIsSPSMemberObject"
    On Error GoTo ErrorHandler
    
    If oTest Is Nothing Then
        ObjectIsSPSMemberObject = False
    ElseIf TypeOf oTest Is ISPSMemberSystem Then
        ObjectIsSPSMemberObject = True
    ElseIf TypeOf oTest Is ISPSMemberPartPrismatic Then
        ObjectIsSPSMemberObject = True
    ElseIf TypeOf oTest Is ISPSFrameConnection Then
        ObjectIsSPSMemberObject = True
    ElseIf TypeOf oTest Is ISPSSplitMemberConnection Then
        ObjectIsSPSMemberObject = True
    ElseIf TypeOf oTest Is ISPSAxisJoint Then
        ObjectIsSPSMemberObject = True
    Else
        ObjectIsSPSMemberObject = False
    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
End Function

' this simply sets a list of joints to be PO to the same objects
Public Sub SetJointsPointOn(elesJoints As IJElements, oTarget1 As Object, oTarget2 As Object)

    Const METHOD = "SetJointsPointOn"
    On Error GoTo ErrorHandler
    
    Dim ii As Long, count As Long
    Dim oJoint As ISPSAxisJoint
    
    If elesJoints Is Nothing Then
        Exit Sub
    End If

    count = elesJoints.count
    
    For ii = 1 To count
        Set oJoint = elesJoints(ii)
        If Not oJoint Is oTarget1 Then
            oJoint.SetPointOn oTarget1, oTarget2
        End If
    Next ii

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise Err.Number
End Sub

Public Function ObjectAssocFlags(obj As Object, flagMask As Long, flagState As Long) As Boolean

    Const METHOD = "ObjectAssocFlags"
    On Error GoTo ErrHandler
    
    Dim assocFlags As Long
    Dim iCompute As IJStructAssocCompute

    If obj Is Nothing Then
        ObjectAssocFlags = False
    Else
        Set iCompute = New StructAssocTools
        iCompute.GetAssocFlags obj, assocFlags

        assocFlags = assocFlags And flagMask

        If assocFlags = flagState Then
            ObjectAssocFlags = True
        Else
            ObjectAssocFlags = False
        End If
    End If
    
    Exit Function
    
ErrHandler:
     HandleError MODULE, METHOD
End Function




'*************************************************************************
'method
'
'<SetGAPRelations>
'
'Abstract
'
'<sets relations specific to the GAP frame connection>
'
'Arguments
'
'<FrameConnection, type of gap connection >
'
'Return
'
'<none>
'
'Exceptions
'***************************************************************************
Public Sub SetGAPRelations(oFC As ISPSFrameConnection, eGapType As SPSGapConnType)
Const MT = "SetGAPRelations"
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    Dim oSupportingMembSys As ISPSMemberSystem, oSecondaryMembSys As ISPSMemberSystem, oSupportedMemberSystem As ISPSMemberSystem
    Dim IHStatus As SPSFCInputHelperStatus
    Dim lngRelated As Long
    Dim oOtherFC As ISPSFrameConnection
    Dim oPointConn As IJPoint
    
    If oFC Is Nothing Then Exit Sub
    
    If eGapType <> SPSGap1 And eGapType <> SPSGap2 Then Exit Sub
        
    'get primary and secondary members
    IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oSupportingMembSys, oSecondaryMembSys)
    
    If IHStatus <> SPSFCInputHelper_Ok Then Err.Raise E_FAIL
    
    If oSupportingMembSys Is Nothing Or oSecondaryMembSys Is Nothing Then Err.Raise E_FAIL
    
    'add relation between port and  supporting member
    Set pRevision = New JRevision
    Set pRelationHelper = oFC.WPO
    Set pCollectionHelper = pRelationHelper.CollectionRelations(strISPSAxisEndPort, strAxisEndAlong1Dest)
    pCollectionHelper.Add oSupportingMembSys, vbNullString, pRelationshipHelper
    pRevision.AddRelationship pRelationshipHelper
    

    If eGapType = SPSGap1 Then
        'add relation between port and secondary member
        Set pCollectionHelper = pRelationHelper.CollectionRelations(strISPSAxisEndPort, strAxisEndAlong2Dest)
        pCollectionHelper.Add oSecondaryMembSys, vbNullString, pRelationshipHelper
    Else
        'add relation between port and secondary member port at connected end
        Set pCollectionHelper = pRelationHelper.CollectionRelations(strISPSAxisEndPort, strAxisEndAxisEndDest)
        
        'get the FC between supporting and secondary member
        GetCommonConnection oSupportingMembSys, oSecondaryMembSys, oPointConn, lngRelated
        
        Set oOtherFC = oPointConn
        'add relation between ports
        pCollectionHelper.Add oOtherFC.WPO, vbNullString, pRelationshipHelper
    End If
    
    pRevision.AddRelationship pRelationshipHelper
    
    'establish extend relation between FC and port, so that the port's propeties show up in the filter dialogs
    Set pRelationHelper = oFC
    Set pCollectionHelper = pRelationHelper.CollectionRelations(strISPSFrameConnection, strFCToGapOutputOrig)
    pCollectionHelper.Add oFC.WPO, vbNullString, pRelationshipHelper
    pRevision.AddRelationship pRelationshipHelper
    
    Exit Sub
End Sub

'*************************************************************************
'method
'
'<RemoveGAPRelations>
'
'Abstract
'
'<removes relations specific to the GAP frame connection>
'
'Arguments
'
'<FrameConnection, type of gap connection >
'
'Return
'
'<none>
'
'Exceptions
'***************************************************************************
Public Sub RemoveGAPRelations(oFC As ISPSFrameConnection)
Const MT = "RemoveGAPRelations"
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    Dim oAssocRel As IJDAssocRelation
    Dim pRelCol As IJDRelationshipCol
    Dim ii As Long, count As Long
   
    If oFC Is Nothing Then Exit Sub
    
       
    Set pRevision = New JRevision
    Set oAssocRel = oFC.WPO
   
    Set pRelCol = oAssocRel.CollectionRelations(strISPSAxisEndPort, strAxisEndAlong1Dest)
    'remove relation between port and  supporting member
  

    If Not pRelCol Is Nothing Then
        count = pRelCol.count
        If count > 0 Then
            'the count should be 1. But there is a chance that we see more than 1 as the previous relation
            'was not deleted due to some reason. So clear all of them now
            For ii = 1 To count
                Set pRelationshipHelper = pRelCol.Item(ii)
                     
                pRevision.RemoveRelationship pRelationshipHelper
            Next ii
        End If
    End If
    
    'remove relation between port and secondary member
    Set pRelCol = oAssocRel.CollectionRelations(strISPSAxisEndPort, strAxisEndAlong2Dest)
    If Not pRelCol Is Nothing Then
        count = pRelCol.count
        If count > 0 Then
            'the count should be 1. But there is a chance that we see more than 1 as the previous relation
            'was not deleted due to some reason. So clear all of them now
            For ii = 1 To count
                Set pRelationshipHelper = pRelCol.Item(ii)
                pRevision.RemoveRelationship pRelationshipHelper
            Next ii
        End If
    End If

    'remove relation between port and secondary member port at connected end
    Set pRelCol = oAssocRel.CollectionRelations(strISPSAxisEndPort, strAxisEndAxisEndDest)
    If Not pRelCol Is Nothing Then
        count = pRelCol.count
        If count > 0 Then
            'the count should be 1. But there is a chance that we see more than 1 as the previous relation
            'was not deleted due to some reason. So clear all of them now
            For ii = 1 To count
                Set pRelationshipHelper = pRelCol.Item(ii)
                pRevision.RemoveRelationship pRelationshipHelper
            Next ii
        End If
    End If
    
    'remove relation between a port which is controoling this port
    Set pRelCol = oAssocRel.CollectionRelations(strISPSAxisEndPort, strAxisEndAxisEndOrig)
    If Not pRelCol Is Nothing Then
        count = pRelCol.count
        If count > 0 Then
            'the count should be 1. But there is a chance that we see more than 1 as the previous relation
            'was not deleted due to some reason. So clear all of them now
            For ii = 1 To count
                Set pRelationshipHelper = pRelCol.Item(ii)
                pRevision.RemoveRelationship pRelationshipHelper
            Next ii
        End If
    End If
    
    
    'remove extend relation between FC and port, so that the port's propeties show up in the filter dialogs
    Set oAssocRel = oFC

    Set pRelCol = oAssocRel.CollectionRelations(strISPSFrameConnection, strFCToGapOutputOrig)
    If Not pRelCol Is Nothing Then
        If pRelCol.count > 0 Then
            Set pRelationshipHelper = pRelCol.Item(1)
            pRevision.RemoveRelationship pRelationshipHelper
        End If
    End If
    
    Exit Sub
End Sub

'*************************************************************************
'method
'
'<IsGAP2Enabled>
'
'Abstract
'
'<checks if the FC has relations specific to the GAPBoth frame connection>
'
'Arguments
'
'<FrameConnection >
'
'Return
'
'<True/false>
'
'Exceptions
'***************************************************************************
Public Function IsGAP2Enabled(oFC As ISPSFrameConnection) As Boolean
Const MT = "IsGAP2Enabled"
    
    Dim pRevision As IJRevision
    Dim oAssocRel As IJDAssocRelation
    Dim pRelCol As IJDRelationshipCol
   
    IsGAP2Enabled = False
    
    If oFC Is Nothing Then Exit Function
    
       
    Set pRevision = New JRevision
    Set oAssocRel = oFC.WPO
   
    'check if this FC    has a support member
    Set pRelCol = oAssocRel.CollectionRelations(strISPSAxisEndPort, strAxisEndAlong1Dest)
  
    If Not pRelCol Is Nothing Then
        If pRelCol.count > 0 Then
            GoTo wrapup ' this FC has an associated support member
        End If
    End If
    
    'check if this FC is controlling the port of another FC
    Set pRelCol = oAssocRel.CollectionRelations(strISPSAxisEndPort, strAxisEndAxisEndDest)
    If Not pRelCol Is Nothing Then
        If pRelCol.count > 0 Then
            GoTo wrapup ' this FC is controlling another FC's port
        End If
    End If
   
    'check if this FC's port is controlled by another FC
    Set pRelCol = oAssocRel.CollectionRelations(strISPSAxisEndPort, strAxisEndAxisEndOrig)
    If Not pRelCol Is Nothing Then
        If pRelCol.count > 0 Then
            GoTo wrapup ' another FC is controlling this Fc's port
        End If
    End If
    'It is still possible that the FC is another of type. so the
    'caller should make sure that the type is Gap2 before calling this function
    Exit Function
wrapup:
    IsGAP2Enabled = True 'if we get here one or more of the above relations exist, so the connection is
    'paired up with another gap2 connection.
    
End Function

'*************************************************************************
'method
'
'<GetMembWithPassiveGapBothConn>
'
'Abstract
'
'<Gets the member with passive Gap Both conn at the location of the input FC>
'
'Arguments
'
'<FrameConnection, supporting membersystem, string - progid of gapboth conection >
'
'Return
'
'<nothing | member system>
'
'Exceptions
'***************************************************************************

Public Function GetMembWithPassiveGapBothConn(oFC As ISPSFrameConnection, oSing As ISPSMemberSystem, strProgID As String) As ISPSMemberSystem

    Dim oCatalogPOM As IJDPOM
    Dim oNamingContext As IJDNamingContextObject
    Dim oSecMemb As Object
    Dim colJoints As IJElements, colFCs As IJElements
    Dim posx As Double, posy As Double, posz As Double
    Dim oJoint As ISPSAxisJoint
    Dim oExistingFC As ISPSFrameConnection
    Dim ii As Long, jj As Long
    Dim oSmartItem As IJSmartItem

    
    Set oCatalogPOM = GetCatalogResourceManager()
    Set oNamingContext = New NamingContextObject
    
        
    'get the location of the FC
    oFC.Joint.Point.GetPoint posx, posy, posz
    'get all joints at this location on the primary member
    Set colJoints = oSing.PointOnJointsAtPosition(posx, posy, posz)


    If Not colJoints Is Nothing Then
        For ii = 1 To colJoints.count
            Set oJoint = colJoints.Item(ii)
            Set colFCs = oJoint.FrameConnections
            If Not colFCs Is Nothing Then
                If colFCs.count = 1 Then ' if count is more than 1 then it cannot be a passive gap2
                    For jj = 1 To colFCs.count
                        Set oExistingFC = colFCs.Item(jj)
                        If Not oFC Is oExistingFC Then
                            Set oSmartItem = oNamingContext.ObjectMoniker(oCatalogPOM, oExistingFC.DefinitionName)
                            If Not oSmartItem Is Nothing Then
                                If oSmartItem.definition = strProgID Then 'it is gap2
                                    If IsGAP2Enabled(oExistingFC) = False Then
                                        'found a passive gap2 FC at this location
                                        'so pair up with this one
                                         Set oSecMemb = oExistingFC.MemberSystem
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next jj
                End If
            End If
            If Not oSecMemb Is Nothing Then
                Set GetMembWithPassiveGapBothConn = oSecMemb
                Exit For
            End If
        Next
    End If


End Function


'*************************************************************************
'method
'
'<RemoveSecondaryMembFromCanRules>
'
'Abstract
'
'This method removes the association between the parent memb of the FC
'and any CanRule the FC is related to.
'
'Arguments
'
'<FrameConnection >
'
'Return
'
'<nothing >
'
'Exceptions
'***************************************************************************

Public Sub RemoveSecondaryMembFromCanRules(oFC As ISPSFrameConnection)

    'no error handler. Caller is expected to handle error thrown from this method
    
    Dim oCanRule As ISPSCanRule
    Dim colSecMembSys As IJElements
    Dim oSecMembSys As ISPSMemberSystem
    Dim ii As Long, count As Long
    
    'the FC may have 1 supporting member (Axis, flush, centerline, etc,...)
    'or 2 supporting members (VerticalCornerBrace, Gap, etc..)
    
    'the parent member FC may be a secondary to either of the 2 supporting member
    'secondary to both is not expected but is not  ruled out
    Set oCanRule = oFC.GetCrossSectionObject(SPSFCPrimary)
    
    If Not oCanRule Is Nothing Then
        oCanRule.GetSecondaryMemberSystems colSecMembSys
        If Not colSecMembSys Is Nothing Then
            For ii = 1 To colSecMembSys.count
                Set oSecMembSys = colSecMembSys.Item(ii)
                If oSecMembSys Is oFC.MemberSystem Then
                    colSecMembSys.Remove (ii)
                    Exit For
                End If
            Next ii
            oCanRule.SetSecondaryMemberSystems colSecMembSys
            Set colSecMembSys = Nothing
        End If
    End If
    
    Set oCanRule = oFC.GetCrossSectionObject(SPSFCSecondary)
    
    If Not oCanRule Is Nothing Then
        oCanRule.GetSecondaryMemberSystems colSecMembSys
        If Not colSecMembSys Is Nothing Then
            For ii = 1 To colSecMembSys.count
                Set oSecMembSys = colSecMembSys.Item(ii)
                If oSecMembSys Is oFC.MemberSystem Then
                    colSecMembSys.Remove (ii)
                    Exit For
                End If
            Next ii
            oCanRule.SetSecondaryMemberSystems colSecMembSys
        End If
    End If
    
End Sub

' Try to fix FC relations using PointOn relations
' Return True if fix was successful allowing caller to continue,
' otherwise return false telling caller the FC has been set to Unsupported.

Public Function FixFC(iFC As ISPSFrameConnection) As Boolean

On Error Resume Next

Dim bFixed As Boolean
Dim oRefCollFixer As ISPSFixFCRefColl

bFixed = False

Set oRefCollFixer = New SPSDBIRepairCmd.FixFCRefColl

If Not oRefCollFixer Is Nothing Then
    
    ' DBIRepairCmd.FixFCRefColl attempts to repair RefColl relations
    ' using PointOn relations.   If they also do not exist, then FC is set to Unsupported.
    
    ' if it cannot fix it, RefCollFixer sets to Unsupported, and returns bFixed = false
    ' so that this caller knows to stop compute of FC.
    bFixed = oRefCollFixer.FixFC(iFC)
    
Else  ' problem creating this component.   Cannot fix.

   Set iFC.definition = Nothing

End If

FixFC = bFixed

End Function

Public Sub SetRecatchOfWPO(oObject As Object)
    On Error GoTo ErrHandler

    ' TR 235218.   Recatch the WPO to force recompute of SPSAxisNotify semantic.
    ' This should only be called from the CMConstruct methods
    
    Dim iControlFlags As IJControlFlags
    Dim iStructAssocTools As SP3DStructGenericTools.StructAssocTools
    Dim iIJStructAssocCompute As SP3DStructGenericTools.IJStructAssocCompute
    Dim lAssocFlags As Long

    Set iIJStructAssocCompute = New SP3DStructGenericTools.StructAssocTools
    iIJStructAssocCompute.GetAssocFlags oObject, lAssocFlags

    If lAssocFlags And &H100000 = &H100000 Then    ' RELATION_INSERTED (in this compute cycle)
        Set iStructAssocTools = iIJStructAssocCompute
        Set iControlFlags = oObject
        iControlFlags.ControlFlags(CTL_FLAG_GRAPH_RECATCH) = CTL_FLAG_GRAPH_RECATCH
        iStructAssocTools.UpdateObject oObject, "{34DF6BED-41D5-4d19-B090-D58D93A1CF64}"        'ISPSAxisWPO
        iControlFlags.ControlFlags(CTL_FLAG_GRAPH_RECATCH) = 0
        Set iControlFlags = Nothing
        Set iStructAssocTools = Nothing
    End If
    Set iIJStructAssocCompute = Nothing

    Exit Sub

ErrHandler:
    HandleError MODULE, "SetRecatchOfWPO"
    Err.Raise E_FAIL
End Sub

Public Sub AddOutputWPO(oFC As ISPSFrameConnection, dispid As Long)
    On Error GoTo ErrHandler
    Dim iIJDMemberObjects As IJDMemberObjects

    Set iIJDMemberObjects = oFC
    If iIJDMemberObjects.ItemByDispid(dispid) Is Nothing Then
        iIJDMemberObjects.Add oFC.WPO, dispid, -1
    End If

    Set iIJDMemberObjects = Nothing
    
    Exit Sub
    
ErrHandler:
    HandleError MODULE, "AddOutputWPO"
    Err.Raise E_FAIL
End Sub

'*************************************************************************
'Function
'UpdateAxisOffsetWPO
'
'Abstract
'Evaluates the member property
'
'Arguments
'IJDPropertyDescription interface describing the property to be evaluated
'pObject is the object whose property is being computed
'FCAxisType: lFCAxisAlong, lFCAxisEnd,lFCAxisColinear
'Return
'
'Exceptions
'
'***************************************************************************

Public Sub UpdateAxisOffsetWPO(FC As ISPSFrameConnection, fcInputHelper As ISPSFCInputHelper, pOutObject As Object, FCAxisType As Long)
    Const MT = "UpdateAxisOffsetWPO"
    On Error GoTo ErrorHandler
    
    Dim geomCondition As Long

    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim obj As Object
    Dim oWPO As ISPSAxisWPO
    Dim iFCPoint As IJPoint
            
    Dim suppingEnd As SPSMemberAxisPortIndex
    Dim oRotation  As ISPSAxisRotation
    Dim oSuppingLocToGlobal As iJDT4x4
    Dim oMemberPart As ISPSMemberPartCommon
    Dim oMemberPart2 As ISPSMemberPartCommon
    Dim oCrossSection As ISPSCrossSection
    Dim oSuppingSys As ISPSMemberSystem
    Dim dLogicalx As Double, dLogicaly As Double, dLogicalz As Double
    Dim dLogicalx2 As Double, dLogicaly2 As Double, dLogicalz2 As Double
    Dim dPhysicalx#, dPhysicaly#, dPhysicalz#
    Dim suppingCP As Long, offsetCP As Long
    Dim lCoordSys As Long
    
    Dim oAttributes As CollectionProxy
    Dim pIJAttrbs As IJDAttributes
    
    Dim oVec As DVector, oVecSupping As DVector, oVecSupping2 As DVector

    Set pIJAttrbs = FC
    
    'the CP of the supporting where the supported needs be positioned
    offsetCP = pIJAttrbs.CollectionOfAttributes("IJUASPSFCAxis").Item("SupportingCP").Value
    
    Set oAttributes = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset")

    lCoordSys = oAttributes.Item("CoordinateSystem").Value
    
    Set oWPO = pOutObject   ' the WPO object of the supported Member System

    Set oVec = New DVector
    Set oVecSupping = New DVector
    Set oVecSupping2 = New DVector
    
    'only get supporting member info if needed...
    If FC.IsImporting Or lCoordSys = 2 Or offsetCP <> 0 Or FCAxisType <> lFCAxisAlong Then

        IHStatus = fcInputHelper.GetRelatedObjects(FC, oSuppingSys, obj)
    
        ' if error occured at reading supporting objects, it may be fixable so call FixFC to see.
        ' if FixFC fails to fix it, it returns false, the FC has been set to Unsupported and we must exit.
        ' if FixFC returns okay, try again to get_RelatedObjects.  If it fails now, we set to unsupported here and exit.
        ' Otherwise, we have been healed and are okay to continue
        
        If IHStatus <> SPSFCInputHelper_Ok Then
            If FixFC(FC) Then
                IHStatus = fcInputHelper.GetRelatedObjects(FC, oSuppingSys, obj)
                If IHStatus <> SPSFCInputHelper_Ok Then
                    Set FC.definition = Nothing
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    
        FC.Joint.Point.GetPoint dLogicalx, dLogicaly, dLogicalz     ' supported end

        ' if joint and end of logical axis not at same location, write message to error log
        If oWPO.portIndex = SPSMemberAxisStart Then
            FC.MemberSystem.LogicalAxis.GetLogicalStartPoint dLogicalx2, dLogicaly2, dLogicalz2
        Else
            FC.MemberSystem.LogicalAxis.GetLogicalEndPoint dLogicalx2, dLogicaly2, dLogicalz2
        End If
        If FCMacroDistance(dLogicalx, dLogicaly, dLogicalz, dLogicalx2, dLogicaly2, dLogicalz2) > 0.001 Then
            ReportError Err, "FCMacros.Common.bas", "UpdateAxisOffsetWPO", "logical positions are not in agreement"
        End If
            
        If FCAxisType = lFCAxisAlong Then
            oSuppingSys.GetMemberPartsAtPosition dLogicalx, dLogicaly, dLogicalz, oMemberPart, oMemberPart2
    
            If oMemberPart Is Nothing Then
                If oMemberPart2 Is Nothing Then  'should not happen.   resolve using closest end.
                    Dim dSuppingPhysicalStartx#, dSuppingPhysicalStarty#, dSuppingPhysicalStartz#
                    Dim dSuppingPhysicalEndx#, dSuppingPhysicalEndy#, dSuppingPhysicalEndz#
                    
                    oSuppingSys.PhysicalAxis.EndPoints dSuppingPhysicalStartx, dSuppingPhysicalStarty, dSuppingPhysicalStartz, dSuppingPhysicalEndx, dSuppingPhysicalEndy, dSuppingPhysicalEndz
                    
                    If FCMacroDistance(dLogicalx, dLogicaly, dLogicalz, dSuppingPhysicalStartx, dSuppingPhysicalStarty, dSuppingPhysicalStartz) < _
                       FCMacroDistance(dLogicalx, dLogicaly, dLogicalz, dSuppingPhysicalEndx, dSuppingPhysicalEndy, dSuppingPhysicalEndz) Then
                        suppingEnd = SPSMemberAxisStart
                    Else
                        suppingEnd = SPSMemberAxisEnd
                    End If
                    Set oMemberPart = oSuppingSys.MemberPartAtEnd(suppingEnd)
                Else
                    Set oMemberPart = oMemberPart2
                End If
            End If
    
        Else
            suppingEnd = oSuppingSys.ResolveEnd(FC.Joint)
            Set oMemberPart = oSuppingSys.MemberPartAtEnd(suppingEnd)
        End If
    
        'if there is a can at the connection point, get section info from the can
        Set oCrossSection = FC.GetCrossSectionObject(SPSFCPrimary)
    
        'if no can, get directly from the part
        If oCrossSection Is Nothing Then
            Set oCrossSection = oMemberPart.CrossSection
        End If
    
        suppingCP = oCrossSection.CardinalPoint ' the CP with which supporting is placed
    
        CheckForUndefinedValueAndRaiseError FC, offsetCP, "StructFCSupportingCPs", 9
        CheckForUndefinedValueAndRaiseError FC, lCoordSys, "StructCoordSysReference", 10
    
        If offsetCP = 0 Then
            offsetCP = suppingCP
        End If
        
        'read transform if needed..
        If offsetCP <> suppingCP Or lCoordSys = 2 Then
            Set oRotation = oMemberPart.Rotation
            oRotation.GetTransform oSuppingLocToGlobal
            If oRotation.Mirror Then
                oSuppingLocToGlobal.IndexValue(4) = -oSuppingLocToGlobal.IndexValue(4)
                oSuppingLocToGlobal.IndexValue(5) = -oSuppingLocToGlobal.IndexValue(5)
                oSuppingLocToGlobal.IndexValue(6) = -oSuppingLocToGlobal.IndexValue(6)
            End If
        End If
          
        If offsetCP <> suppingCP Then   'compute an offset for selected different CP
        
            Dim x As Double, y As Double
            Dim oJointPos As New DPosition
            
            oJointPos.Set dLogicalx, dLogicaly, dLogicalz
            
            oCrossSection.GetCardinalPointDelta oJointPos, suppingCP, offsetCP, x, y
       
            oVecSupping.Set 0#, -x, y    'cross-section x = member's -y, y = member's z
            Set oVecSupping = oSuppingLocToGlobal.TransformVector(oVecSupping)
        
        End If
        
        'Now add the offset of the supporting Member Part so the apparent offset is relative to the supporting's physical.
        If FCAxisType = lFCAxisEnd Or FCAxisType = lFCAxisColinear Then
            Dim xPs#, yPs#, zPs#, xPe#, yPe#, zPe#, xL#, yL#, zL#

            oSuppingSys.PhysicalAxis.EndPoints xPs, yPs, zPs, xPe, yPe, zPe
            
            If suppingEnd = SPSMemberAxisStart Then
            
                oSuppingSys.LogicalAxis.GetLogicalStartPoint xL, yL, zL
                oVecSupping2.Set (xPs - xL), (yPs - yL), (zPs - zL)
            
            Else
                
                oSuppingSys.LogicalAxis.GetLogicalEndPoint xL, yL, zL
                oVecSupping2.Set (xPe - xL), (yPe - yL), (zPe - zL)
            
            End If
        End If
    End If   'If FC.IsImporting Or lCoordSys = 2 Or offsetCP <> 0 Or FCAxisType <> lFCAxisAlong
    
    If FC.IsImporting Then
        'Adjust the offset to keep the physical unchanged
        Set iFCPoint = FC       'IJPoint for FC returns MemberSystem PhysicalAxis end point
        iFCPoint.GetPoint dPhysicalx, dPhysicaly, dPhysicalz

        oWPO.SetWPO dPhysicalx - dLogicalx, dPhysicaly - dLogicaly, dPhysicalz - dLogicalz, 0#, 0#, 0#
        'save back offsets
        oVec.Set dPhysicalx - dLogicalx - oVecSupping.x - oVecSupping2.x, dPhysicaly - dLogicaly - oVecSupping.y - oVecSupping2.z, dPhysicalz - dLogicalz - oVecSupping.z - oVecSupping2.z
        If lCoordSys = 2 Then
            oSuppingLocToGlobal.Invert
            'Convert Offsets back to local coordinates
            Set oVec = oSuppingLocToGlobal.TransformVector(oVec)
         End If
        oAttributes.Item("XOffset").Value = oVec.x
        oAttributes.Item("YOffset").Value = oVec.y
        oAttributes.Item("ZOffset").Value = oVec.z
    Else
        oVec.x = oAttributes.Item("XOffset").Value
        oVec.y = oAttributes.Item("YOffset").Value
        oVec.z = oAttributes.Item("ZOffset").Value
        If lCoordSys = 2 Then 'Offsets were given in supping local coordinates
            Set oVec = oSuppingLocToGlobal.TransformVector(oVec)
        End If
        Set oVec = oVec.Add(oVecSupping)
        Set oVec = oVec.Add(oVecSupping2)
        oWPO.SetWPO oVec.x, oVec.y, oVec.z, 0#, 0#, 0#
    End If
   
    oWPO.WPOCardinalPoint = 0
    If FC.IsImporting Then
        SetWPOAtOtherEndIfUnsupported FC    ' if other end is unsupported, set its offsets to keep physical where it is.
        FC.IsImporting = False
    Else
        PropagateWPOToOtherEnd FC   'if the other end is unsupported, then propagate offset to the other end based on the align flag
    End If
        
    Set iFCPoint = Nothing
    
    If FCAxisType = lFCAxisColinear Then
        geomCondition = GeometricalCondition(FC, oSuppingSys)
        
        If (geomCondition And 4) <> 4 Then
        
            SPSToDoErrorNotify FCToDoMsgCodelist, TDL_FCMACROS_AXISCOLI_MEMS_COLINEAR, FC, Nothing
            Err.Raise SPS_MACRO_WARNING
        
        End If
    End If

    Exit Sub

ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise E_FAIL
    End If
End Sub

Private Function FCMacroDistance(d1x As Double, d1y As Double, d1z As Double, d2x As Double, d2y As Double, d2z As Double) As Double

    Dim deltax As Double, deltay As Double, deltaz As Double
    
    deltax = d2x - d1x
    deltay = d2y - d1y
    deltaz = d2z - d1z

    FCMacroDistance = Sqr(deltax * deltax + deltay * deltay + deltaz * deltaz)
    
End Function

' if other end FC is unsupported, set its offset values based on that ends difference in logical and physical.
Public Sub SetWPOAtOtherEndIfUnsupported(FC As ISPSFrameConnection)

On Error GoTo ErrorHandler

    Dim otherFC As ISPSFrameConnection
    Dim otherFCEnd As SPSMemberAxisPortIndex
    Dim iCurve As IJCurve
    Dim physStartX As Double, physStartY As Double, physStartZ As Double, physEndX As Double, physEndY As Double, physEndZ As Double
    Dim logicalX As Double, logicalY As Double, logicalZ As Double
    
    If FC Is Nothing Then
        Exit Sub
    End If
    If TypeOf FC.MemberSystem Is ISPSMemberSystemCurve Then
        Exit Sub
    End If
    
    If FC.WPO.portIndex = SPSMemberAxisStart Then
        otherFCEnd = SPSMemberAxisEnd
    Else
        otherFCEnd = SPSMemberAxisStart
    End If

    Set otherFC = FC.MemberSystem.FrameConnectionAtEnd(otherFCEnd)
    If otherFC Is Nothing Then
        Exit Sub
    End If
    If Not otherFC.definition Is Nothing Then
        Exit Sub
    End If

    Set iCurve = otherFC.MemberSystem
    iCurve.EndPoints physStartX, physStartY, physStartZ, physEndX, physEndY, physEndZ
    
    otherFC.WPO.WPOCardinalPoint = 0
    If otherFCEnd = SPSMemberAxisStart Then
        otherFC.MemberSystem.LogicalAxis.GetLogicalStartPoint logicalX, logicalY, logicalZ
        otherFC.WPO.SetWPO physStartX - logicalX, physStartY - logicalY, physStartZ - logicalZ, 0#, 0#, 0#
    Else
        otherFC.MemberSystem.LogicalAxis.GetLogicalEndPoint logicalX, logicalY, logicalZ
        otherFC.WPO.SetWPO physEndX - logicalX, physEndY - logicalY, physEndZ - logicalZ, 0#, 0#, 0#
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, "SetWPOAtOtherEndIfUnsupported"
    Err.Raise E_FAIL
End Sub


