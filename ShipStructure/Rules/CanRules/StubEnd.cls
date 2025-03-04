VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2009 Intergraph Corporation. All rights reserved.
'
'File : InLine.cls
'
'Author : MH
'
'Description :
'    class to drive can rule object for the inline placement
'
'History:
'   some time in 2009   MH   creation
'   May 5 2009          RP   Added migration, disconnection and evaluate logic
'   June 05, 2009       MH   165811 detect overlapping can conditions
'   June 26, 2009       GG   163724 NameRules are required for DM's driven by CanRules
'   June 26, 2009       RP   166645 Modified input/output interface definition  for secondaries, plane and can member
'                            also added code to call UpdatePosition during evaluate.
'   July 02, 2009       RP   TR167563 Removed spliconnection as an input for the stubendcan
'   July 17, 2009       GG   166187 Attribute management code needs to validate user keyed in values and update related values
'   July 23, 2009       MH   163732 update parameters during compute
'   Mar  04, 2010       MH   175947 watch secondary on physical axis to force my compute to precede an update of related split-connection.
'                            also go TDL if split location is off the member system
'*****************************************************************************************************************

Option Explicit

Private Const MODULE = "StubEndCan"

Implements IJGeometricConstructionDefinitionService
Implements IJGCMigrate
Implements ISPSCanRuleHelper
Implements IJGCSemanticConnection
Implements IJUserAttributeMgmt

Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal pDispatchOfRelationship As Object)
    'Any code that needs to be called when a relation is established to the canrule goes here
End Sub

Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal pDispatchOfRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Const METHOD = "IJGCSemanticConnection_PreConnectionRemoved"
    On Error GoTo ErrorHandler

    Dim oRelationShip As IJDRelationship
    Dim oIJDObject As IJDObject
    Dim strRlnName As String
    
    'This code is called for disconnect of InputRel or OutputRel
    'CanRule is Origin for InputRel and OutputRel
    '
    'Use cases:
    '   delete canrule with no secondary
    '   delete canrule with AxisAlong as secondary
    '   delete canrule with VerticalCornerBrace as secondary
    '   delete primary member-system
    '   delete secondary member-system
    '   set secondary member-system FC to unsupported
    '   install neighbor can
    '   delete neighbor can
    '   delete stubEnd can placed at a split
    
    Set oRelationShip = pDispatchOfRelationship
    strRlnName = oRelationShip.Name
    ''MsgBox strRlnName & ", bIsOriginDeleted=" & bIsOriginDeleted & ", bIsDestDeleted=" & bIsDestinationDeleted

    'The CanRule is being deleted.
    If bIsOriginDeleted = True Then
    
        'Delete the planes that are input to the split connections.
        'That deletes the split connectons, which will delete the BUCan in the split semantic.
        If Left(strRlnName, Len(StructCanRuleCollectionNames.StructCanRule_Planes)) = StructCanRuleCollectionNames.StructCanRule_Planes Then
            Set oIJDObject = oRelationShip.Destination
            oIJDObject.Remove

        'The CanRule is being deleted and not the primaryMS
        ElseIf Left(strRlnName, Len(StructCanRuleCollectionNames.StructCanRule_Primary)) = StructCanRuleCollectionNames.StructCanRule_Primary Then

            'Reset secondary member FC's to be in relation to their memberSystem again.
            'and also reset the related FC's "Supported" object.
            ResetMemberFCs oRelationShip.Origin, bIsDestinationDeleted
        
        End If
    
    ' A StubEnd Can at a split has PointOperator.1 relation, and when the Split is deleted, the PointOperator is also deleted.
    ' If the CanRule is not being deleted yet, then we need to delete the CanRule.  That will trigger more deletes..
    ElseIf bIsDestinationDeleted = True And Left(strRlnName, Len(StructCanRuleCollectionNames.StructCanRule_PointOperator)) = StructCanRuleCollectionNames.StructCanRule_PointOperator Then

        Set oIJDObject = oRelationShip.Origin
        If Not oIJDObject Is Nothing Then
            oIJDObject.Remove
        End If
    
    End If
    
    Exit Sub

ErrorHandler:
    ' no specific code for here right now other than clearing
    ' the error so that delete operation completes
    WriteToErrorLog Err.Number, MODULE, METHOD, "Unexpected error"
    Err.Clear

End Sub


Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal MyDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Const METHOD = "IJGeometricConstructionDefinitionService_Initialize"
    On Error GoTo ErrorHandler
    '
    ' Order of listing the inputs is important !! The order is  used by client in fitering inputs
    '
    ' input 1 - frame connection or split connection where the end can needs to be placed
    
    Call MyDefinition.AddInput(StructCanRuleCollectionNames.StructCanRule_EndPort, "Select a Frame Connection or Split Connection", "ISPSFrameConnection AND [SM3DCanRules.CanFilter,IsTube] OR ISPSSplitMemberConnection AND [SM3DCanRules.CanFilter,IsTube] OR ISPSAxisEndPort AND [SM3DCanRules.CanFilter,IsTube]", 1, 1, "ISPSAxisEndPort") '' only for endcans
    
    ' input 2               : the member system of the secondary member systems contributing to my length
    ' compute triggering    : SuppingNotify2 is post-part notification
    Call MyDefinition.AddInput(StructCanRuleCollectionNames.StructCanRule_Secondary, "Select Secondary Member Systems", "ISPSMemberSystemLinear", 0, 100, "ISPSMemberSystemEndEndNotify ISPSMemberSystemLinear")
    
    ' obsolete input collection.  this input collection is no longer created or used.
    Call MyDefinition.AddInput(StructCanRuleCollectionNames.StructCanRule_Target, "Select Member Part", "ISPSDesignedMember", 0, 1, "ISPSDesignedMember")
        
    ' compute triggering    : SuppingNotify3 is pre-part notification
    Call MyDefinition.AddControlledInput(StructCanRuleCollectionNames.StructCanRule_Primary, "ISPSMemberSystemPhysicalAxis")
    
    ' retrieved inputs      : the member parts along the Column
    ' compute triggering    : NotifyInput triggers me to change split locations, and/or tube-diameter
    Call MyDefinition.AddControlledInput(StructCanRuleCollectionNames.StructCanRule_MemberToSplit, "ISPSCanRule")
    
    Call MyDefinition.AddControlledInput(StructCanRuleCollectionNames.StructCanRule_Neighbors, "IStructCrossSectionDesignProperties")
    Call MyDefinition.AddControlledInput(StructCanRuleCollectionNames.StructCanRule_SupportingSecondary, "ISPSMemberSystemXSectionNotify ISPSMemberSystemPhysicalAxis")
    Call MyDefinition.AddControlledInput(StructCanRuleCollectionNames.StructCanRule_PointOperator, "IJPoint")
''

    Call MyDefinition.AddControlledOutput(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan, "IUABuiltUpCan IJUASMCanRuleResult")
    '
    ' outputs related to the Splitters
    '
    Call MyDefinition.AddControlledOutput(StructCanRuleCollectionNames.StructCanRule_Planes, "IJPlane", GCOverridePropagateFromMacro)
    Call MyDefinition.AddControlledOutput(StructCanRuleCollectionNames.StructCanRule_SplitConnections, "ISPSSplitMemberConnection")

    '
    ' to avoid creating a GCSharedParameters root clsid at bulkload time unstead of a GCMacro
    Call MyDefinition.AddOutput(GCGTypePoint3d, "Dummy")

    Call MyDefinition.AddSelfOutput("IJPoint")
    Call MyDefinition.AddSelfOutput("IJUASMCanRuleStubEnd")     ' declare this as self-output to keep params in synch

    ' parameters
    Call MyDefinition.AddParameter(attrDiameterRule, "DiameterRule", GCCodeList, UNIT_SCALAR, Nothing, Nothing, Nothing, DiameterRule_User)
    Call MyDefinition.AddParameter("CanOD", "CanOD", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("CanID", "CanID", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("CanThickness", "CanThickness", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("CanMaterial", "CanMaterial", GCChar, 0, 0, 0, 0)
    Call MyDefinition.AddParameter("CanGrade", "CanGrade", GCChar, 0, 0, 0, 0)

    Call MyDefinition.AddParameter("MinExtensionDistance", "MinExtensionDistance", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("L2Method", "L2Method", GCCodeList)
    Call MyDefinition.AddParameter("L2Factor", "L2Factor", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("L2Length", "L2Length", GCDouble, 1, 0, 0, 0)

    Call MyDefinition.AddParameter("ConeMethod", "ConeMethod", GCLong)
    Call MyDefinition.AddParameter("ConeSlope", "ConeSlope", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("ConeAngle", "ConeAngle", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("ConeLength", "ConeLength", GCDouble, 1, 0, 0, 0)

    Call MyDefinition.AddParameter("ConeThickness", "ConeThickness", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("ConeMaterial", "ConeMaterial", GCChar, 0, 0, 0, 0)
    Call MyDefinition.AddParameter("ConeGrade", "ConeGrade", GCChar, 0, 0, 0, 0)

    Call MyDefinition.AddParameter("ChamferSlope", "ChamferSlope", GCDouble, 1, 0, 0, 0)
    Call MyDefinition.AddParameter("RoundoffDistance", "RoundoffDistance", GCDouble, 1, 0, 0, 0)
    
    MyDefinition.UseErrorValue strCodeListTablename

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal My As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    
    Call Evaluate(My, False)
    
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Sub Evaluate(ByVal MyGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal bIsAPSDeleting As Boolean)
    Const METHOD = "Evaluate"
    On Error GoTo ErrorHandler
    
    ' use the GC as a GCMacro
    Dim iCanRule As ISPSCanRule
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim crStatus As SPSCanRuleStatus
    Dim iAxisPort As ISPSAxisEndPort
    Dim iPort As IJPort
    Dim pMemberSystem As iSPSMemberSystem
    Dim distCanLen As Double        ' initial can length based on secondary member sizes, but not coneLengths or L2/L3
    Dim posL2Hull As IJDPosition, posL2Centerline As IJDPosition
    Dim posL3Hull As IJDPosition, posL3Centerline As IJDPosition

    Set iCanRule = MyGC
    Set MyGCMacro = MyGC

    If MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_EndPort).count = 0 Then
        MyGC.PostError MISSING_MANDATORY_INPUT, True, strCodeListTablename
        Err.Raise E_FAIL
    End If
    'update the position of the canrule to be consistent with that of the endport
    crStatus = iCanRule.UpdatePosition
    If crStatus <> SPSCanRule_Ok Then
        MyGC.PostError GetCodeListErrorNumber(crStatus), True, strCodeListTablename
        Err.Raise E_FAIL
    End If
    
    
    Set iAxisPort = MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_EndPort).Item("1")
   
    ' check inputs, set pMemberSystem
    If MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Primary).count = 0 Then

        Set pMemberSystem = iAxisPort.MemberSystem
        MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Primary).Add pMemberSystem, "1"

    Else
        Set pMemberSystem = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Primary).Item("1")
        If Not pMemberSystem Is iAxisPort.MemberSystem Then
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename, iAxisPort, pMemberSystem
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Inconsistency: MemberSystem of AxisPort is not same as Primary input"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Inconsistency: MemberSystem of AxisPort is not same as Primary input"
        End If

        CheckAndAddSupportingSecondary iCanRule
    End If

    If MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors).count = 0 Then               ' have not migrated yet

        iCanRule.Services.ComputeMinMaxPoints 0, posL2Hull, posL2Centerline, posL3Centerline, posL3Hull, crStatus

        ' if I have no output yet, then find the part to split and set it as input.
        If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).count = 0 And MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_MemberToSplit).count = 0 Then

            Dim iPartCommon As ISPSMemberPartCommon
            Dim elesParts As IJElements, elesDummy As IJElements
            Dim iMF As New SPSMemberFactory
            Dim iMemberFeatureServices As ISPSMemberFeatureServices
            
            Set iMemberFeatureServices = iMF.CreateMemberFeatureServices
            iMemberFeatureServices.FindPartsAlongMemberSystem pMemberSystem, posL2Hull, posL3Hull, elesParts, elesDummy

            If (elesParts.count > 1) And (Not TypeOf iAxisPort.ILC Is ISPSSplitMemberConnection) Then              ' attempting to place across an existing split. it is ok
            'for a stubend placed at an existing split connection to return more than one part during the call above
                RemoveOutputs MyGCMacro
                MyGC.PostError SPLIT_ALREADY_EXISTS, True, strCodeListTablename
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Cannot place a Can across existing splits."
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Cannot place a Can across existing splits."
            Else
                Set iPort = iAxisPort
                Set iPartCommon = iPort.Connectable
            End If

            If iPartCommon.IsPrismatic Then
                RemoveOutputs MyGCMacro
                MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Cannot place a Can on prismatic part"
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Cannot place a Can on prismatic part"
            End If

            If Not iPartCommon.Rule Is Nothing Then
                RemoveOutputs MyGCMacro
                MyGC.PostError SPLIT_ALREADY_EXISTS, True, strCodeListTablename
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Cannot place a Can at existing Can"
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Cannot place a Can at existing Can"
            End If
    
            If ObjectOnTDL(iPartCommon) Then
                RemoveOutputs MyGCMacro
                MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Cannot place a Can on DesignedMember on ToDoList"
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Cannot place a Can on DesignedMember on ToDoList"
            End If
                
            Call MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_MemberToSplit).Add(iPartCommon, "1")
        End If

        distCanLen = Sqr((posL2Hull.x - posL3Hull.x) * (posL2Hull.x - posL3Hull.x) + _
                         (posL2Hull.y - posL3Hull.y) * (posL2Hull.y - posL3Hull.y) + _
                         (posL2Hull.z - posL3Hull.z) * (posL2Hull.z - posL3Hull.z))
        If distCanLen > 0.001 Then
            distCanLen = 0.5 * distCanLen
        Else
            distCanLen = 0.5
        End If

        ' create the one split and the extra bounding plane.
        If iAxisPort.portIndex = SPSMemberAxisStart Then
            crStatus = CreateSplitConnection(MyGC, distCanLen, "1")
        Else
            crStatus = CreateSplitConnection(MyGC, -distCanLen, "1")
        End If

    End If
    
    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).count = 0 Then      ' presumably, we have not executed split yet.
        Exit Sub
    End If

    ' integrity checks
    ' 1: is the output collection of Cans count == 1
    ' 2: is MyGC same as my output member's rule.  Error is very unlikely since it traverses same relation.
    ' 3: check that we have one split connections and one output planes
    ' 4: check that a split neighbor is same as my output Can
    ' 5: check that other split neighbor is my neighbor

    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).count <> 1 Then
        MyGC.PostError UNEXPECTED_ERROR, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "Output collection for Can count <> 1"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "Output collection for Can count <> 1"
    End If

    Dim oMemberPartCommon As ISPSMemberPartCommon
    Set oMemberPartCommon = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")
    If Not oMemberPartCommon.Rule Is MyGC Then
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "Inconsistency: Can is not output of this CanRule"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "Inconsistency: Can is not output of this CanRule"
    End If

    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_SplitConnections).count <> 1 Then
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "Output collection for Splits count <> 1"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "Output collection for Splits count <> 1"
    End If
    
    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes).count <> 1 Then
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "Output collection for Planes count <> 1"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "Output collection for Planes count <> 1"
    End If

    Dim oObj1 As Object, oObj2 As Object, oObjNbor As Object

    GetSplitNeighbors MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_SplitConnections).Item("1"), oObj1, oObj2

    ' if any split outputs are nothing, no check to allow location to be re-set by the evaluate.
    If Not (oObj1 Is Nothing Or oObj2 Is Nothing) Then
        If oObj1 Is oMemberPartCommon Then
            Set oObjNbor = oObj2
        ElseIf oObj2 Is oMemberPartCommon Then
            Set oObjNbor = oObj1
        Else
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename, oMemberPartCommon
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Neither split neighbor is the can output"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Neither split neighbor is the can output"
        End If
    
       ' only check neighbor identity if neighbor collection is okay
        If MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors).count = 1 Then
            If Not oObjNbor Is MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors).Item("1") Then
                 MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename, oMemberPartCommon
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Neighbor of Can is not same as split neighbor"
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Neighbor of Can is not same as split neighbor"
            End If
        End If
    End If
    
    Call StubEndEvaluate(MyGC)

    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Const METHOD = "IJGCMigrate_Migrate"
    On Error GoTo ErrorHandler
    
   
    ' migrate will get called when a nbor is split, or my original split
    
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim startX As Double, startY As Double, startZ As Double, endX As Double, endY As Double, endZ As Double
    Dim posMe As IJPoint

    Dim bDeleted As Boolean
    Dim countMembers As Long
    Dim elesMembers As IJElements
    Dim elesReplacing As IJElements
    Dim oMiddleMember As Object, oNbor1 As Object, oNbor2 As Object
    Dim oSO As IJSmartOccurrence
    Dim oMembPart As ISPSMemberPartCommon
    Dim iAxisPort As ISPSAxisEndPort
    Dim iAxisPortNew As ISPSAxisEndPort

        
    Set MyGCMacro = MyGC
    Set elesMembers = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_MemberToSplit)
    countMembers = elesMembers.count

    Set posMe = MyGC
    posMe.GetPoint startX, startY, startZ

    Set iAxisPort = MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_EndPort).Item("1")
    
    If countMembers = 1 Then         ' first time primary split.  replacing the initial part

        ' get the replacing members
        ' find the one that I am PointOn to = oMiddleMember
        ' find the ones at the start and end of oMiddleMember
        '
        ' set the def of oMiddleMember to be the BUCan
        ' clear all inputs
        ' set neighbors so that "1" is the start, and "2" is the end.
        ' set output to be the oMiddleMember

        pMigrateHelper.ObjectsReplacing elesMembers.Item(1), elesReplacing, bDeleted
        
        FindClosestMember startX, startY, startZ, elesReplacing, Nothing, oMiddleMember
        
        Set oMembPart = oMiddleMember
        
        FindClosestMember startX, startY, startZ, elesReplacing, oMiddleMember, oNbor1

        'replace the  port form the original part with the port from the replacing part
        Set iAxisPortNew = oMembPart.AxisPort(iAxisPort.portIndex)
         
        MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_EndPort).Clear
        
        MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_EndPort).Add iAxisPortNew, "1"
        
        elesMembers.Clear
        Set elesMembers = Nothing
        Set elesMembers = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors)          ' the real neighbors
        
        ' we are writing to the BUCan on an interface which is input to BUCan custom semantic
        'the code below prevents (also an assert) that interface being treated as an ouput interface in the CAD
        Set oSO = oNbor1
        oSO.Properties = oSO.Properties Or 32

        elesMembers.Add oNbor1, "1"
        
        ' TR 250592  create input relation with neighbor so that paste asks for a MemberPart.
        If TypeOf iAxisPortNew.ILC Is ISPSSplitMemberConnection Then
            MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_Target).Clear
            MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_Target).Add oNbor1, "1"
        End If
        ' we are writing to the BUCan on an interface which is input to BUCan custom semantic
        'the code below prevents (also an assert) that interface being treated as an ouput interface in the CAD
        
        Set oSO = oMiddleMember
        oSO.Properties = oSO.Properties Or 32

        SetNameRule oMiddleMember
        MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Add oMiddleMember, "1"
        SetBUCanDefinition oMiddleMember
        
    Else ' a neighbor is being replaced.

        Set elesMembers = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors)          ' the real neighbors
        
        pMigrateHelper.ObjectsReplacing elesMembers.Item("1"), elesReplacing, bDeleted
        
        If Not elesReplacing Is Nothing Then
            If elesReplacing.count > 0 Then
                Set oMiddleMember = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")
                FindClosestMember startX, startY, startZ, elesReplacing, oMiddleMember, oNbor1
                elesMembers.Remove ("1")

                ' we are writing to the BUCan on an interface which is input to BUCan custom semantic
                'the code below prevents (also an assert) that interface being treated as an ouput interface in the CAD
                
                Set oSO = oNbor1
                oSO.Properties = oSO.Properties Or 32

                elesMembers.Add oNbor1, "1"
            
                ' TR 250592  create input relation with neighbor so that paste asks for a MemberPart.
                If TypeOf iAxisPort.ILC Is ISPSSplitMemberConnection Then
                    MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_Target).Clear
                    MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_Target).Add oNbor1, "1"
                End If
            
            End If
            Set elesReplacing = Nothing
        End If
       
    End If

    Exit Sub

ErrorHandler:
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
    On Error GoTo ErrorHandler
    
    IJUserAttributeMgmt_OnAttributeChange = OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrToChange, varNewAttrValue)
    Exit Function

ErrorHandler:
    IJUserAttributeMgmt_OnAttributeChange = METHOD & ": " & Err.Description
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    IJUserAttributeMgmt_OnPreCommit = vbNullString
End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
    On Error GoTo ErrorHandler
    
    IJUserAttributeMgmt_OnPreLoad = vbNullString
    
    Dim oAttrs As Collection
    Dim oAttrDescr As IJAttributeDescriptor
    Dim diameterRule As Long
    
    Dim oRule  As IJAttributeDescriptor
    Dim oCanOD As IJAttributeDescriptor
    Dim oCanID As IJAttributeDescriptor
    Dim oCanUD As IJAttributeDescriptor
    
    Dim oL2Method As IJAttributeDescriptor
    Dim oL2Factor As IJAttributeDescriptor
    Dim oL2Length As IJAttributeDescriptor
    
    Dim oConeMethod As IJAttributeDescriptor
    Dim oConeSlope  As IJAttributeDescriptor
    Dim oConeLength As IJAttributeDescriptor
    Dim oConeAngle  As IJAttributeDescriptor
    
    Dim oCanRule As ISPSCanRule
    Set oCanRule = pIJDAttrs
    
    If ifWellDefinedSetReadOnly(oCanRule.DefinitionName, CollAllDisplayedValues) Then
        Exit Function
    End If
    
    For Each oAttrDescr In CollAllDisplayedValues
        Select Case oAttrDescr.attrName
            Case attrDiameterRule:      Set oRule = oAttrDescr
            Case attrOuterDiameter:     Set oCanOD = oAttrDescr
            Case attrInnerDiameter:     Set oCanID = oAttrDescr
            
            Case attrL2CompMethod:      Set oL2Method = oAttrDescr
            Case attrL2Factor:          Set oL2Factor = oAttrDescr
            Case attrL2Length:          Set oL2Length = oAttrDescr
            
            Case attrConeLengthMethod:  Set oConeMethod = oAttrDescr
            Case attrConeSlope:         Set oConeSlope = oAttrDescr
            Case attrConeAngle:         Set oConeAngle = oAttrDescr
            Case attrConeLength:        Set oConeLength = oAttrDescr
        
        End Select
    Next oAttrDescr
    
    oRule.AttrState = AttributeDescriptor_ReadOnly
    RuleSetReadOnly oRule.AttrValue, oCanID, oCanOD
    
    LMethodSetReadOnly oL2Method.AttrValue, oL2Length, oL2Factor
    
    ConeMethodSetReadOnly oConeMethod.AttrValue, oConeLength, oConeSlope, oConeAngle

    
    Exit Function
ErrorHandler:
    IJUserAttributeMgmt_OnPreLoad = Err.Description
End Function

Private Function ISPSCanRuleHelper_GetTubeDiameter(ByVal CanRule As SPSMembers.ISPSCanRule, tubeDiameter As Double) As SPSMembers.SPSCanRuleStatus
    ISPSCanRuleHelper_GetTubeDiameter = GetTubeDiameter(CanRule, False, tubeDiameter)
End Function


Private Function ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions(ByVal CanRule As ISPSCanRule) As SPSCanRuleStatus

    ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions = UpdateOutputCrossSectionDimensions(False, CanRule)

    Exit Function

End Function

' interface on myself that relates to diameter only.
Private Function ISPSCanRuleHelper_GetCanRuleCrossSectionInterface(ByVal CanRule As ISPSCanRule, _
                ByRef strInputXSInterfaceName As String) As SPSCanRuleStatus
    
    strInputXSInterfaceName = crNameCanRuleXSectionInterface        ' my GC input interface
    ' TODO: change Can interfaces to separate changes that influence diameter from length parameters
    ISPSCanRuleHelper_GetCanRuleCrossSectionInterface = SPSCanRule_Ok
    Exit Function

End Function

' what interface to listen to neighbors' size
Private Function ISPSCanRuleHelper_GetNeighborUpdateInterface(ByVal CanRule As ISPSCanRule, _
                ByRef strNeighborInterfaceName As String) As SPSCanRuleStatus

    strNeighborInterfaceName = crName_BuiltUpTubeInterface
    ISPSCanRuleHelper_GetNeighborUpdateInterface = SPSCanRule_Ok
    Exit Function
 
End Function

' interface that the ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions will update on the BUCan
Private Function ISPSCanRuleHelper_GetBUCanCrossSectionInterface(ByVal CanRule As ISPSCanRule, _
                ByRef strCrossSectionInterface As String) As SPSCanRuleStatus

    strCrossSectionInterface = crName_BuiltUpTubeInterface
    ISPSCanRuleHelper_GetBUCanCrossSectionInterface = SPSCanRule_Ok
    Exit Function
    
End Function


