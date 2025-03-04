Attribute VB_Name = "Common"
'*******************************************************************
'
'Copyright (C) 2004 Intergraph Corporation. All rights reserved.
'
'File : Common.bas
'
'Author : M. Holderer
'
'Description :
'    Module for common connection constants/utilities
'
'History:
'   08-Jul-2004 M. Holderer     started
'   14-NOv-2006 A. Singh    TR#109960- Change the PG of the split point to the Parent SC. Also retrunr error
'                           in evaluatesplitrule if the split parent is not writable.

'**********************************************************************************************************
Option Explicit
Private Const MODULE = "Common"
Public Const SPSvbError = vbObjectError + 512
Public Const E_FAIL As Long = -2147467259

Public Const ConstIJStructContinuity2 = "{0FBF4139-92B4-43e8-9D08-B2917FCACD05}"
Public Const ConstISPSMemberSystemSuppingNotify3 = "{1DABF8D1-5534-44d2-9896-063D5F62DA79}"

Public Const ConstErrMsg_UnexpectedError = "Split connection encountered an unexpected error. See To Do List messages in the Troubleshooting Guide for more information."
Public Const ConstErrMsg_NoValidIntersection = "The member system to split and the object selected to define the split location do not intersect."

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
 On Error GoTo errorHandler
  
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

errorHandler:
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
 On Error GoTo errorHandler
  
   'traverse the relation from the SO to the RefColl
   'if none exists, make one and connect it
   
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim oObject As IJDObject
    Dim oSymbolEntitiesFactory  As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set pRelationHelper = pFC
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    If Not pCollectionHelper Is Nothing Then
        If pCollectionHelper.count = 1 Then
            Set GetRefColl = pCollectionHelper.Item("RC")
        End If
    End If
  
    If GetRefColl Is Nothing Then
        Set oObject = pFC
        Set GetRefColl = oSymbolEntitiesFactory.CreateEntity(referencesCollection, oObject.ResourceManager)
        SetRefColl pFC, GetRefColl
    End If

 Exit Function
 
errorHandler:
    HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'
'<GetRefCollNoCreate>
'
'Abstract
'
'<Gets the related RefColl object if one exists>
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
Public Function GetRefCollNoCreate(pFC As Object) As IJDReferencesCollection
Const MT = "GetRefCollNoCreate"
 On Error GoTo errorHandler
  
   'traverse the relation from the SO to the RefColl
   
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim oObject As IJDObject
    Dim oSymbolEntitiesFactory  As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set GetRefCollNoCreate = Nothing

    Set pRelationHelper = pFC
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "toArgs_O")
    If Not pCollectionHelper Is Nothing Then
        If pCollectionHelper.count = 1 Then
            Set GetRefCollNoCreate = pCollectionHelper.Item("RC")
        End If
    End If
  
    Exit Function

errorHandler:
    HandleError MODULE, MT
End Function

Public Sub GetDesignParent(parent As Object, iIJDesignParent As IJDesignParent)

    Dim pParent As Object
    Dim iChild As IJDesignChild

    If TypeOf parent Is IJDesignChild Then

        Set iChild = parent
        Set pParent = iChild.GetParent
        
        If pParent Is Nothing Then
            Set iIJDesignParent = Nothing
        ElseIf TypeOf pParent Is IJDesignParent Then
            Set iIJDesignParent = pParent
        Else
            GetDesignParent pParent, iIJDesignParent
        End If
    End If
        
End Sub


'*************************************************************************
'Function
'
'<GetRefCollObjects>
'
'Abstract
'
'<Gets the objects from the ReferencesCollection>
'
'Arguments
'
'<ReferenceCollection as input, IJElements as output with unique objects>
'
'Return
'
'Exceptions
'***************************************************************************
Public Function GetRefCollObjects(oConnObj As Object) As IJElements
    
    Const MT = "GetRefCollObjects"
    On Error GoTo errorHandler

    Dim oRC As IJDReferencesCollection
    Dim uniqueObjs As IJElements
    Dim obj As Object
    Dim ii As Long
    
    Set GetRefCollObjects = New JObjectCollection
    
    If oConnObj Is Nothing Then
        Exit Function
    End If

    If TypeOf oConnObj Is IJDReferencesCollection Then
        Set oRC = oConnObj
    Else
        Set oRC = GetRefCollNoCreate(oConnObj)
    End If

    If Not oRC Is Nothing Then
        For ii = 1 To oRC.IJDEditJDArgument.GetCount
            Set obj = oRC.IJDEditJDArgument.GetEntityByIndex(ii)
            If Not GetRefCollObjects.Contains(obj) Then
                GetRefCollObjects.Add obj
            End If
        Next ii
    End If

    Set obj = Nothing
    Exit Function

errorHandler:
    HandleError MODULE, MT
End Function

Public Sub ComputeIntersection(parent1 As Object, parent2 As Object, splitStatus As splitStatus, _
                distanceAlong As Double, _
                ByRef parent1x As Double, ByRef parent1y As Double, ByRef parent1z As Double, _
                ByRef parent2x As Double, ByRef parent2y As Double, ByRef parent2z As Double, _
                ByRef ok As Boolean)

Const MT = "ComputeIntersection"
On Error GoTo errorHandler

Dim MF As New SPSMemberFactory
Dim iConnServices As SPSMembers.ISPSMemberConnectionServices

Set iConnServices = MF.CreateConnectionServices

iConnServices.ComputeIntersectionPointForSplit parent1, parent2, splitStatus, distanceAlong, _
        parent1x, parent1y, parent1z, parent2x, parent2y, parent2z, ok

Exit Sub

errorHandler:
   HandleError MODULE, MT
End Sub

Public Function EvaluateSplitRule(ByVal iSmartOccurrence As Object, useStatus As Boolean, strSelectedItem As String) As SPSMembers.StructSOCInputHelperStatus

    On Error GoTo errorHandler

    Dim splitStat As splitStatus
    Dim iIJStructILCConnection As IJStructILCConnection

    Dim socStatus As SPSMembers.StructSOCInputHelperStatus
    Dim iContinuity2 As IJStructContinuity2
    Dim continuityType1 As StructContinuityType2, continuityType2 As StructContinuityType2
    Dim priorityNumber1 As Long, priorityNumber2 As Long

    Dim parents As IJElements
    Dim parent1 As Object, parent2 As Object
 
    socStatus = StructSOCInputHelper_Ok
    strSelectedItem = ""
    
    Set parents = GetRefCollObjects(iSmartOccurrence)

    If parents.count = 1 Then
        Set parent1 = parents(1)
    ElseIf parents.count = 2 Then
        Set parent1 = parents(1)
        Set parent2 = parents(2)
    Else
        socStatus = StructSOCInputHelper_BadNumberOfObjects
        GoTo wrapup
    End If
    Set parents = Nothing

    If parent2 Is Nothing Then
        If TypeOf parent1 Is IJStructSystemParent Then
            strSelectedItem = "SplitByPoint-1"
        Else
            socStatus = StructSOCInputHelper_InvalidTypeOfObject
        End If

    'if both parents are potentially splittable, then we get the splitStatus to determine which connection to use.
    ElseIf TypeOf parent1 Is IJStructSystemParent And TypeOf parent2 Is IJStructSystemParent Then

        If useStatus Then       ' called by selector
    
            Set iIJStructILCConnection = iSmartOccurrence
            splitStat = iIJStructILCConnection.SplitParentStatus
            
            If splitStat = ssRule Then
                strSelectedItem = "SplitByRule"
            ElseIf splitStat = ssSplitNone Then
                strSelectedItem = "SplitNone"
            ElseIf splitStat = ssSplitBoth Then
                strSelectedItem = "SplitBoth"
            ElseIf splitStat = ssSplitFirst Then
                strSelectedItem = "SplitFirst"
            ElseIf splitStat = ssSplitSecond Then
                strSelectedItem = "SplitSecond"
            Else
                strSelectedItem = ""        ' error
            End If
        
        Else        ' called ByRule.  Note that IJStructContinuity2 is an input interface.
    
            Set iContinuity2 = parent1
            continuityType1 = iContinuity2.ContinuityType
            priorityNumber1 = iContinuity2.PriorityNumber
            Set iContinuity2 = parent2
            continuityType2 = iContinuity2.ContinuityType
            priorityNumber2 = iContinuity2.PriorityNumber
            
            If continuityType1 = continuityType2 Then       ' they are the same.
                If continuityType1 = ContinuousType Then    ' both are continuous.
                    strSelectedItem = "SplitNone"
                Else
                    If priorityNumber1 = priorityNumber2 Then
                        strSelectedItem = "SplitBoth"
                    ElseIf priorityNumber1 > priorityNumber2 Then  ' split parent with higher priorityNumber
                        strSelectedItem = "SplitFirst"
                    Else
                        strSelectedItem = "SplitSecond"
                    End If
                End If
            ElseIf continuityType1 = IntercostalType Then   ' not the same.
                strSelectedItem = "SplitFirst"
            Else
                strSelectedItem = "SplitSecond"
            End If
            
            'TR109960 Check the approval status of the parent and return error if case of no acess
            Dim bPermOK As Boolean
            Dim bPermParent1OK As Boolean
            Dim bPermParent2OK As Boolean
            
            bPermOK = False
            bPermParent1OK = False
            bPermParent2OK = False
           
            Dim oObjMember1 As IJDObject
            Dim oObjMember2 As IJDObject
            Dim acConst As AccessControlConstant
            
            Set oObjMember1 = parent1
            Set oObjMember2 = parent2
            
            acConst = oObjMember1.AccessControl
            If acConst And acUpdate Then
                bPermParent1OK = True
            End If
            
            acConst = oObjMember2.AccessControl
            If acConst And acUpdate Then
                bPermParent2OK = True
            End If
            
            'MsgBox "AccessControl1=" + Str(oObjMember1.AccessControl) + " AccessControl1=" + Str(oObjMember2.AccessControl)
            If strSelectedItem = "SplitFirst" Then
                 If bPermParent1OK = True Then
                    bPermOK = True
                 End If
            ElseIf strSelectedItem = "SplitSecond" Then
                 If bPermParent2OK = True Then
                    bPermOK = True
                 End If
            ElseIf strSelectedItem = "SplitBoth" Then
                 If ((bPermParent1OK = True) And (bPermParent2OK = True)) Then
                    bPermOK = True
                 End If
            ElseIf strSelectedItem = "SplitNone" Then
                'check the permission of the parent system
                Dim oObjParentSys As IJDObject
                Dim oDesignPar As IJDesignParent
                                
                GetDesignParent parent1, oDesignPar
                
                If Not oDesignPar Is Nothing Then
                    Set oObjParentSys = oDesignPar
                    acConst = oObjParentSys.AccessControl
                    If acConst And acWrite Then
                        bPermOK = True
                    End If
                End If
            End If
            
            If bPermOK = False Then
                socStatus = StructSOCInputHelper_NoWriteAccess
            End If
            
        End If

    'else one is not splittable.  select based on element type.
    ' may want to support IJDynamicSurfaceFind here too
    Else
        If useStatus Then       ' called by selector
            Set iIJStructILCConnection = iSmartOccurrence
            splitStat = iIJStructILCConnection.SplitParentStatus
        End If
        If TypeOf parent1 Is IJStructSystemParent Then
            If splitStat = ssPenetration Then
                strSelectedItem = "SplitByPenetration-1"
            ElseIf ObjectIsOkSurface(parent2) Then
                strSelectedItem = "SplitBySurface-1"
            End If
            
        ElseIf TypeOf parent2 Is IJStructSystemParent Then
            If splitStat = ssPenetration Then
                strSelectedItem = "SplitByPenetration-1"
            ElseIf ObjectIsOkSurface(parent1) Then
                strSelectedItem = "SplitBySurface-1"
            End If
        End If
    End If
    
wrapup:
    EvaluateSplitRule = socStatus
    Exit Function

errorHandler:
    EvaluateSplitRule = StructSOCInputHelper_UnexpectedError
    Err.Raise E_FAIL
End Function



Public Sub InstallSplitPoint(doWhen As Long, doSysChild As Long, splitConn As ISPSSplitConnectionComputeHelper, parent As IJStructSystemParent, point As IJPoint)

    Const MT = "InstallSplitPoint"
    On Error GoTo errorHandler

    ' doWhen can be:
    '   0 = always  (called at construct)
    '   1 = only on Paste  (called in compute)
    '   2 = only on Paste and clear the Paste flag   (called in compute)
    
    'TR#109960- Change the PG of the split point to the Parent SC
    If Not point Is Nothing Then
        Dim oObjSC As IJDObject
        Dim oObjPoint As IJDObject
        Set oObjSC = splitConn
        Set oObjPoint = point
        oObjPoint.PermissionGroup = oObjSC.PermissionGroup
    End If

    If 0 = doWhen Or splitConn.IsPasted Then

        If Not point Is Nothing Then
            parent.AddSplit point
        End If
        
        If doSysChild > 0 Then
            AddSysChild doSysChild, parent, splitConn
        End If
    
        If 2 = doWhen Then
            splitConn.IsPasted = False
        End If
    
    End If

    Exit Sub

errorHandler:
   HandleError MODULE, MT
End Sub

Public Sub AddSysChild(whichParent As Long, parent As Object, ISOC As IJDesignChild)

    Const MT = "AddSysChild"
    On Error GoTo errorHandler

    Dim oldParent As IJDesignParent
    Dim newParent As IJDesignParent

    'whichParent can be:
    '   1 = make SplitConn be a child of parent
    '   2 = make SplitConn be a child of parents parent
    
    Set oldParent = ISOC.GetParent
    
    If whichParent = 1 Then
        Set newParent = parent
    ElseIf whichParent = 2 Then
        GetDesignParent parent, newParent
    End If

    If Not oldParent Is newParent Then
        If Not oldParent Is Nothing Then
            'First detach split from existing parent, before relating it to new parent
            oldParent.RemoveChild ISOC
        End If
        newParent.AddChild ISOC
    End If

    Exit Sub

errorHandler:
   HandleError MODULE, MT
End Sub

Public Sub UpdateSplitPoint(splitConn As ISPSSplitConnectionComputeHelper, point As IJPoint, x As Double, y As Double, z As Double)
    Const MT = "UpdateSplitPoint"
    On Error GoTo errorHandler

    splitConn.UpdatePointOperator point, x, y, z

    Exit Sub

errorHandler:
   HandleError MODULE, MT
End Sub

Private Sub AddJointAndUp(inMS As ISPSMemberSystem, inJoint As ISPSAxisJoint, connectedList As IJElements)
    Const METHOD = "AddJointAndUp"
    On Error GoTo ErrHandler
    
    'Add joint's end MS's that are not the given MS
    'Do the same for any joint that is PO to this joint.  ( Recurse "up" )

    Dim oMS As ISPSMemberSystem
    Dim oJoint As ISPSAxisJoint
    Dim coll As IJElements
    Dim ii As Long, count As Long
    
    ' add end member systems that are not the same as the one given.
    Set oJoint = inJoint
    Set coll = oJoint.EndMemberSystems
    count = coll.count
    For ii = 1 To count
        Set oMS = coll(ii)
        If Not inMS Is oMS Then
            connectedList.Add oMS
        End If
    Next ii
    Set oMS = Nothing
    Set coll = Nothing
    
    ' recurse with any joints that are PointOn to this joint.
    Set coll = oJoint.PointOnJoints
    Set oJoint = Nothing
    count = coll.count
    For ii = 1 To count
        Set oJoint = coll(ii)
        AddJointAndUp inMS, oJoint, connectedList
    Next ii
    
    Exit Sub

ErrHandler:
    HandleError MODULE, METHOD
End Sub
Private Sub AddJointDown(inJoint As ISPSAxisJoint, connectedList As IJElements)
    Const METHOD = "AddJointDown"
    On Error GoTo ErrHandler
    
    'Add objects that this joint is PO to.
    'But if that object is a joint, add that Joint's End MSs and recurse in case it is also PO to something.

    Dim p1 As Object, p2 As Object
    Dim oJoint As ISPSAxisJoint
    Dim coll As IJElements
    Dim ii As Long, count As Long
    
    Set oJoint = inJoint
    oJoint.GetPointOn p1, p2
    Set oJoint = Nothing

    If p1 Is Nothing Then
        Exit Sub
    
    ElseIf TypeOf p1 Is ISPSAxisJoint Then
        
        Set oJoint = p1
        Set coll = oJoint.EndMemberSystems
        count = coll.count
        For ii = 1 To count
            connectedList.Add coll(ii)
        Next ii
        Set coll = Nothing
        AddJointDown oJoint, connectedList
    
    Else

        connectedList.Add p1
        
        If Not p2 Is Nothing Then
            connectedList.Add p2
        End If
    End If

    Exit Sub

ErrHandler:
    HandleError MODULE, METHOD
End Sub
Private Sub GetAllRelatedObjects(inMS As ISPSMemberSystem, inISOC As SPSMembers.IJStructSmartOccurrenceConnection, connectedList As IJElements)

    Const METHOD = "GetAllRelatedObjects"
    On Error GoTo ErrHandler
    
    Dim oJoint As ISPSAxisJoint
    Dim ii As Long, count As Long, jj As Long, parentCount As Long
    Dim coll As IJElements, parentColl As IJElements
    Dim p1 As Object

    Dim ISOC As IJStructSmartOccurrenceConnection
    Dim socStatus As StructSOCInputHelperStatus
    
    'Add objects at start we are dependent on.
    Set oJoint = inMS.JointAtEnd(SPSMemberAxisStart)
    AddJointDown oJoint, connectedList
    'Add objects at start that are dependent on me.
    AddJointAndUp inMS, oJoint, connectedList

    'Add objects at end we are dependent on, or are depending on me.
    Set oJoint = inMS.JointAtEnd(SPSMemberAxisEnd)
    AddJointAndUp inMS, oJoint, connectedList
    AddJointDown oJoint, connectedList
    
    'Add objects at pointOn joints that are dependent on inMS
    Set coll = inMS.PointOnJoints
    count = coll.count
    For ii = 1 To count
        Set oJoint = coll(ii)
        AddJointAndUp inMS, oJoint, connectedList
    Next ii
    Set oJoint = Nothing
    
    'Add object that are related through the split connection
    Set coll = inMS.SplitConnections
    count = coll.count
    
    For ii = 1 To count
        Set ISOC = coll(ii)
        'don't add objects related to this split conn itself.
        If Not ISOC Is inISOC Then

            Set parentColl = ISOC.InputObjects

            parentCount = parentColl.count
            For jj = 1 To parentCount
                Set p1 = parentColl(jj)
                If Not p1 Is inMS Then
                    connectedList.Add p1
                End If
            Next jj
        End If
    
        Set p1 = Nothing
        Set ISOC = Nothing
        Set parentColl = Nothing
    Next ii
    
    Set parentColl = Nothing
    Set coll = Nothing
    
    Exit Sub

ErrHandler:
    HandleError MODULE, METHOD
End Sub

Public Function RedundantConnectionExists(ISOC As SPSMembers.IJStructSmartOccurrenceConnection, _
                            InObj1 As Object, InObj2 As Object) As Boolean

    'Determine whether a connection exists between the given MemberSystem and the other object.
    'If the objects are connected ONLY with the input connection, then the answer is false.
    'This permits ValidateParents to call this at placement time or paste to check for redundant connection.
    '
    'Even if the members are connected indirectly through PointOn's, here we say that a Redundant Connection Exists.
    '

    Const MT = "RedundantConnectionExists"
    On Error GoTo errorHandler

    Dim bConnected As Boolean
    Dim connectedList As IJElements
    Dim inMS As ISPSMemberSystem
    Dim InObj As Object

    Set connectedList = New JObjectCollection

    ' get a list of every object that is related to the input member-system, which is
    ' all objects that the MS depends on, or are dependent on the MS through FC's and SC's.

    If TypeOf InObj1 Is ISPSMemberSystem Then
        Set inMS = InObj1
        Set InObj = InObj2
    Else
        Set inMS = InObj2
        Set InObj = InObj1
    End If

    GetAllRelatedObjects inMS, ISOC, connectedList
      
    If connectedList.Contains(InObj) Then
        RedundantConnectionExists = True
    Else
        RedundantConnectionExists = False
    End If
    
    Exit Function

errorHandler:
   HandleError MODULE, MT
End Function

Public Function IsNOTWritable(InObj As IJDObject) As Boolean
    Const MT = "IsNOTWritable"
    On Error GoTo errorHandler

    If InObj.AccessControl And acUpdate Then
        IsNOTWritable = False
    Else
        IsNOTWritable = True
    End If

    Exit Function

errorHandler:
   HandleError MODULE, MT
End Function
    
Public Sub RemoveCommonFCs(parent1 As Object, parent2 As Object)
    Const MT = "RemoveCommonFCs"
    On Error GoTo errorHandler

    Dim oMS As ISPSMemberSystem
    Dim oFC As ISPSFrameConnection
    Dim oR1 As Object, oR2 As Object
    Dim IHStatus As SPSFCInputHelperStatus

    On Error Resume Next        ' write access failure will result in transaction abort

    If Not parent1 Is Nothing Then
        If TypeOf parent1 Is ISPSMemberSystem Then
            Set oMS = parent1
            Set oFC = oMS.FrameConnectionAtEnd(SPSMemberAxisStart)
            IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oR1, oR2)
            If IHStatus <> SPSFCInputHelper_Ok Then
                Set oFC.definition = Nothing
            ElseIf oR1 Is parent2 Or oR2 Is parent2 Then
                Set oFC.definition = Nothing
            End If
            Set oR1 = Nothing
            Set oR2 = Nothing
            Set oFC = oMS.FrameConnectionAtEnd(SPSMemberAxisEnd)
            IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oR1, oR2)
            If IHStatus <> SPSFCInputHelper_Ok Then
                Set oFC.definition = Nothing
            ElseIf oR1 Is parent2 Or oR2 Is parent2 Then
                Set oFC.definition = Nothing
            End If
            Set oR1 = Nothing
            Set oR2 = Nothing
        End If
    End If
    If Not parent2 Is Nothing Then
        If TypeOf parent2 Is ISPSMemberSystem Then
            Set oMS = parent2
            Set oFC = oMS.FrameConnectionAtEnd(SPSMemberAxisStart)
            IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oR1, oR2)
            If IHStatus <> SPSFCInputHelper_Ok Then
                Set oFC.definition = Nothing
            ElseIf oR1 Is parent1 Or oR2 Is parent1 Then
                Set oFC.definition = Nothing
            End If
            Set oR1 = Nothing
            Set oR2 = Nothing
            Set oFC = oMS.FrameConnectionAtEnd(SPSMemberAxisEnd)
            IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oR1, oR2)
            If IHStatus <> SPSFCInputHelper_Ok Then
                Set oFC.definition = Nothing
            ElseIf oR1 Is parent1 Or oR2 Is parent1 Then
                Set oFC.definition = Nothing
            End If
            Set oR1 = Nothing
            Set oR2 = Nothing
        End If
    End If
    
    Exit Sub

errorHandler:
   HandleError MODULE, MT
End Sub

Public Sub SetTDLErrorFlag(oSplitConnection As Object, bOk As Boolean)
    Const MT = "SetTDLErrorFlag"
    On Error GoTo errorHandler

    ' this method sets or clears the TDL list flag to help indicate whether the split Connection is in an error state.
    ' bit zero is used to indicate whether the SplitConn compute found an error and is on TDL.

    Dim oSPSPersistFlagsAccess As ISPSPersistFlagsAccess

    Set oSPSPersistFlagsAccess = oSplitConnection

    If bOk Then     ' not in error state.   make sure the error flag is cleared.
        oSPSPersistFlagsAccess.PersistFlags = oSPSPersistFlagsAccess.PersistFlags And &HFFFFFFFE
    Else            ' set the error state flag.
        oSPSPersistFlagsAccess.PersistFlags = oSPSPersistFlagsAccess.PersistFlags Or 1
    End If
    
    Exit Sub

errorHandler:
   HandleError MODULE, MT
End Sub

Public Sub SetIsUnaryFlag(oSplitConnection As Object, bIsUnary As Boolean)
    Const MT = "SetIsUnaryFlag"
    On Error GoTo errorHandler

    ' this method sets or clears the IsUnary split connection type
    ' it is actually persisted using the same mechanism as ISPSPersistFlagsAccess uses for the error state.

    Dim iSPSSplitConnection As ISPSSplitMemberConnection
    
    Set iSPSSplitConnection = oSplitConnection
    
    iSPSSplitConnection.IsUnary = bIsUnary

    Exit Sub

errorHandler:
   HandleError MODULE, MT
End Sub

Public Function CheckAccess(whichParent As Long, parent As Object, ISOC As IJDesignChild) As StructSOCInputHelperStatus

    Const MT = "CheckAccess"
    On Error GoTo errorHandler

    Dim oldParent As IJDesignParent
    Dim newParent As IJDesignParent

    'whichParent can be:
    '   1 = make SplitConn be a child of parent
    '   2 = make SplitConn be a child of parents parent
    
    Set oldParent = ISOC.GetParent
    
    CheckAccess = StructSOCInputHelper_Ok
    
    If whichParent = 1 Then
        Set newParent = parent
    ElseIf whichParent = 2 Then
        GetDesignParent parent, newParent
    End If

    If Not oldParent Is newParent Then
        Dim oObject As IJDObject
        
        Set oObject = ISOC
        
        'Check the access control of the split conn being modiifed
        If oObject.AccessControl And acUpdate Then
            CheckAccess = StructSOCInputHelper_Ok
        Else
            CheckAccess = StructSOCInputHelper_NoWriteAccess
            Set oObject = Nothing
            Exit Function
        End If
        
        'Check the access control of the parent of the split conn being modiifed
        Set oObject = ISOC.GetParent
        If Not oObject Is Nothing Then
            If oObject.AccessControl And acUpdate Then
                CheckAccess = StructSOCInputHelper_Ok
            Else
                CheckAccess = StructSOCInputHelper_NoWriteAccess
                Set oObject = Nothing
                Exit Function
            End If
        End If
        
        'Check the access control of the parent to which the split conn being added.
        Set oObject = newParent
        
        If oObject.AccessControl And acUpdate Then
            CheckAccess = StructSOCInputHelper_Ok
        Else
            CheckAccess = StructSOCInputHelper_NoWriteAccess
            Set oObject = Nothing
            Exit Function
        End If
        
        Set oObject = Nothing
        
        
    End If


    Exit Function

errorHandler:
   HandleError MODULE, MT
End Function

