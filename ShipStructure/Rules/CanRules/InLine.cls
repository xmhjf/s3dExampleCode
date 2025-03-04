VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InLine"
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
'   April 22 2009       RP   Added disconnection code for secondary membs to reattach to memb sys
'   June 05, 2009       MH   165811 detect overlapping can conditions
'   June 26, 2009       GG   163724 NameRules are required for DM's driven by CanRules
'   June 26, 2009       RP   166645 Modified input/output interface definition  for secondaries, plane and can member
'                            also added code to call UpdatePosition during evaluate.
'   July 17, 2009       GG   166187 Attribute management code needs to validate user keyed in values and update related values
'   July 23, 2009       MH   163732 update parameters during compute
'*****************************************************************************************************************

Option Explicit

Private Const MODULE = "InLineCan"


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
    Dim strRlnName As String
    Dim oIJDObject As IJDObject
    
    'This code is called for disconnect of InputRel or OutputRel
    'CanRule is Origin for InputRel and OutputRel
    '
    'Use cases:
    '   delete canRule
    '   delete primary member-system
    '   delete secondary member-system
    '   delete canRule and secondary FC
    '   delete canRule and secondary memberSystem
    '   set secondary member-system FC to unsupported
    '   install neighbor can
    '   delete neighbor can

    Set oRelationShip = pDispatchOfRelationship
    strRlnName = oRelationShip.Name
    ''MsgBox strRlnName & ", bIsOriginDeleted=" & bIsOriginDeleted & ", bIsDestDeleted=" & bIsDestinationDeleted

    If bIsOriginDeleted = True Then    'CanRule is being deleted

    'Delete the planes that are input to the split connections.  That deletes the split connectons,
    'which will delete the BUCan in the split semantic.

        If Left(strRlnName, Len(StructCanRuleCollectionNames.StructCanRule_Planes)) = StructCanRuleCollectionNames.StructCanRule_Planes Then
            Set oIJDObject = oRelationShip.Destination
            oIJDObject.Remove
    
        ElseIf Left(strRlnName, Len(StructCanRuleCollectionNames.StructCanRule_Primary)) = StructCanRuleCollectionNames.StructCanRule_Primary Then
    
            'Reset secondary member FC's to be in relation to this CanRule's primary memberSystem.
            ResetMemberFCs oRelationShip.Origin, bIsDestinationDeleted
    
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
    ' inputs related to the Column.   Order of listing the inputs is important !!
    '
    ' input 1               : the member system of the member system being split
    ' compute triggering    : SuppingNotify3 is pre-part notification
    Call MyDefinition.AddInput(StructCanRuleCollectionNames.StructCanRule_Primary, "Select Primary Member System", "ISPSMemberSystemLinear AND [SM3DCanRules.CanFilter,IsTube]", 1, 1, "ISPSMemberSystemPhysicalAxis")

    ' input 2               : the member system of the secondary member systems contributing to my length
    ' compute triggering    : SuppingNotify2 is post-part notification
    Call MyDefinition.AddInput(StructCanRuleCollectionNames.StructCanRule_Secondary, "Select Secondary Member Systems", "ISPSMemberSystemLinear", 0, 100, "ISPSMemberSystemEndEndNotify ISPSMemberSystemLinear")

    ' retrieved inputs      : the initial member part along the primary that is split
    ' compute triggering    : NotifyInput triggers me to change split locations, and/or tube-diameter
    Call MyDefinition.AddControlledInput(StructCanRuleCollectionNames.StructCanRule_MemberToSplit, "ISPSCanRule")
    
    ' retrieved inputs      : the member parts along the Column
    ' compute triggering    : NotifyInput triggers me to change split locations, and/or tube-diameter
    Call MyDefinition.AddControlledInput(StructCanRuleCollectionNames.StructCanRule_Neighbors, "IStructCrossSectionDesignProperties")

    '
    ' output which is the BuiltUp Can
    '
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
    Call MyDefinition.AddSelfOutput("IJUASMCanRuleInLine")     ' declare this as self-output to keep params in synch

    ' parameters
    Call MyDefinition.AddParameter(attrDiameterRule, "DiameterRule", GCCodeList, UNIT_SCALAR, Nothing, Nothing, Nothing, DiameterRule_User)
    
    Call MyDefinition.AddParameter("CanOD", "CanOD", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1.5249)
    Call MyDefinition.AddParameter("CanID", "CanID", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 1.4732)
    Call MyDefinition.AddParameter("CanThickness", "CanThickness", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.0254)
    Call MyDefinition.AddParameter("CanMaterial", "CanMaterial", GCChar, 0, 0, 0, 0, "Steel - Carbon")
    Call MyDefinition.AddParameter("CanGrade", "CanGrade", GCChar, 0, 0, 0, 0, "A")
    Call MyDefinition.AddParameter("MinExtensionDistance", "MinExtensionDistance", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.3048)
    Call MyDefinition.AddParameter("L2Method", "L2Method", GCCodeList, UNIT_SCALAR, SCALAR_NULL, Nothing, Nothing, TubeExtension_CLFactor)
    Call MyDefinition.AddParameter("L2Factor", "L2Factor", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.25)
    Call MyDefinition.AddParameter("L2Length", "L2Length", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.381)
    Call MyDefinition.AddParameter("L3Method", "L3Method", GCCodeList, UNIT_SCALAR, SCALAR_NULL, Nothing, Nothing, TubeExtension_CLFactor)
    Call MyDefinition.AddParameter("L3Factor", "L3Factor", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.25)
    Call MyDefinition.AddParameter("L3Length", "L3Length", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.381)
    Call MyDefinition.AddParameter("Cone1Method", "Cone1Method", GCLong, UNIT_SCALAR, SCALAR_NULL, Nothing, Nothing, ConeMethod_Length)
    Call MyDefinition.AddParameter("Cone1Slope", "Cone1Slope", GCDouble, UNIT_SCALAR, SCALAR_NULL, 0, 0, 4#)
    Call MyDefinition.AddParameter("Cone1Angle", "Cone1Angle", GCDouble, UNIT_ANGLE, ANGLE_RADIAN, 0, 0, 0.1745)
    Call MyDefinition.AddParameter("Cone1Length", "Cone1Length", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.61)
    
    Call MyDefinition.AddParameter("Cone2Method", "Cone2Method", GCLong, UNIT_SCALAR, SCALAR_NULL, Nothing, Nothing, ConeMethod_Length)
    Call MyDefinition.AddParameter("Cone2Slope", "Cone2Slope", GCDouble, UNIT_SCALAR, SCALAR_NULL, 0, 0, 4#)
    Call MyDefinition.AddParameter("Cone2Angle", "Cone2Angle", GCDouble, UNIT_ANGLE, ANGLE_RADIAN, 0, 0, 0.1745)
    Call MyDefinition.AddParameter("Cone2Length", "Cone2Length", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.61)

    Call MyDefinition.AddParameter("Cone1Thickness", "Cone1Thickness", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.02)
    Call MyDefinition.AddParameter("Cone1Material", "Cone1Material", GCChar, 0, 0, 0, 0, "Steel - Carbon")
    Call MyDefinition.AddParameter("Cone1Grade", "Cone1Grade", GCChar, 0, 0, 0, 0, "A")

    Call MyDefinition.AddParameter("Cone2Thickness", "Cone2Thickness", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.02)
    Call MyDefinition.AddParameter("Cone2Material", "Cone2Material", GCChar, 0, 0, 0, 0, "Steel - Carbon")
    Call MyDefinition.AddParameter("Cone2Grade", "Cone2Grade", GCChar, 0, 0, 0, 0, "A")

    Call MyDefinition.AddParameter("ChamferSlope", "ChamferSlope", GCDouble, UNIT_SCALAR, SCALAR_NULL, 0, 0, 4#)
    Call MyDefinition.AddParameter("RoundoffDistance", "RoundoffDistance", GCDouble, UNIT_DISTANCE, DISTANCE_METER, 0, 0, 0.03)
    
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
    Dim distCanLen As Double        ' initial can length based on secondary member sizes only
    Dim iIJPoint As IJPoint
    Dim iCanRule As ISPSCanRule
    Dim pMemberSystem As iSPSMemberSystem
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim oBUCan As ISPSDesignedMember
    Dim crStatus As SPSCanRuleStatus
    Dim posL2Hull As IJDPosition, posL2Centerline As IJDPosition
    Dim posL3Hull As IJDPosition, posL3Centerline As IJDPosition
    
    Set iCanRule = MyGC
    Set MyGCMacro = MyGC
    Set iIJPoint = iCanRule
    
    ' check inputs
    If MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_Primary).count = 0 Then
        MyGC.PostError MISSING_MANDATORY_INPUT, True, strCodeListTablename
        Err.Raise E_FAIL
    End If
    
    'update the position of the canrule to be consistent with that of the primary member
    crStatus = iCanRule.UpdatePosition
    If crStatus <> SPSCanRule_Ok Then
        MyGC.PostError GetCodeListErrorNumber(crStatus), True, strCodeListTablename
        Err.Raise E_FAIL
    End If

    Set pMemberSystem = MyGC.Inputs(StructCanRuleCollectionNames.StructCanRule_Primary).Item("1")

    If MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors).count = 0 Then        ' have not migrated yet

        ' call computeMinMaxPoints to get average position of located secondary members
        iCanRule.Services.ComputeMinMaxPoints 0, posL2Hull, posL2Centerline, posL3Centerline, posL3Hull, crStatus
        
        ' if I have no output yet, then find the part to split and set it as input.
        If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).count = 0 And MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_MemberToSplit).count = 0 Then
            Dim elesParts As IJElements, elesDummy As IJElements
            Dim iPartCommon As ISPSMemberPartCommon

            Dim iMF As New SPSMemberFactory
            Dim iMemberFeatureServices As ISPSMemberFeatureServices
            
            Set iMemberFeatureServices = iMF.CreateMemberFeatureServices

            iMemberFeatureServices.FindPartsAlongMemberSystem pMemberSystem, posL2Hull, posL3Hull, elesParts, elesDummy
            
            If elesParts.count <> 1 Then
                RemoveOutputs MyGCMacro
                MyGC.PostError SPLIT_ALREADY_EXISTS, True, strCodeListTablename
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Cannot place a Can across existing splits."
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Cannot place a Can across existing splits."
            End If

            Set iPartCommon = elesParts.Item(1)
            
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

            posL2Centerline.x = 0.5 * (posL2Hull.x + posL3Hull.x)
            posL2Centerline.y = 0.5 * (posL2Hull.y + posL3Hull.y)
            posL2Centerline.z = 0.5 * (posL2Hull.z + posL3Hull.z)
            
            crStatus = iCanRule.SetPosition(posL2Centerline.x, posL2Centerline.y, posL2Centerline.z)
            
            Call MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_MemberToSplit).Add(iPartCommon, "1")
            
        End If
    
        distCanLen = Sqr((posL2Hull.x - posL3Hull.x) * (posL2Hull.x - posL3Hull.x) + _
                         (posL2Hull.y - posL3Hull.y) * (posL2Hull.y - posL3Hull.y) + _
                         (posL2Hull.z - posL3Hull.z) * (posL2Hull.z - posL3Hull.z))
        If distCanLen > 0.001 Then
            distCanLen = 0.5 * distCanLen
            Else
            distCanLen = 0.2
        End If
    
        ' create or update the two splits and exit early.            ' create the two splits and exit early.
        crStatus = CreateSplitConnection(MyGC, -distCanLen, "1")
        crStatus = CreateSplitConnection(MyGC, distCanLen, "2")
        
    End If

    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).count = 0 Then      ' presumably, we have not executed split yet.
        Exit Sub
    End If

    ' integrity checks
    ' 1: is the output collection of Cans count == 1
    ' 2: is MyGC same as my output member's rule.  Error is very unlikely since it traverses same relation.
    ' 3: check that we have two split connections and two output planes
    ' 4: check that common member of splits are same as my output Can
    ' 5: check that split neighbors are same as my neighbors

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

    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_SplitConnections).count <> 2 Then
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "Output collection for Splits count <> 2"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "Output collection for Splits count <> 2"
    End If
    
    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes).count <> 2 Then
        MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename
        WriteToErrorLog E_FAIL, MODULE, METHOD, "Output collection for Planes count <> 2"
        Err.Raise E_FAIL, MODULE & ":" & METHOD, "Output collection for Planes count <> 2"
    End If

    Dim oObj1 As Object, oObj2 As Object, oObj3 As Object, oObj4 As Object, oObjCommon As Object
    Dim oObjNbor1 As Object, oObjNbor2 As Object

    GetSplitNeighbors MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_SplitConnections).Item("1"), oObj1, oObj2
    
    GetSplitNeighbors MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_SplitConnections).Item("2"), oObj3, oObj4

    ' if any split outputs are nothing, no check to allow location to be re-set by the evaluate.
    If Not (oObj1 Is Nothing Or oObj2 Is Nothing Or oObj3 Is Nothing Or oObj4 Is Nothing) Then
        If oObj1 Is oObj3 Then
            Set oObjCommon = oObj1
            Set oObjNbor1 = oObj2
            Set oObjNbor2 = oObj4
        ElseIf oObj1 Is oObj4 Then
            Set oObjCommon = oObj1
            Set oObjNbor1 = oObj2
            Set oObjNbor2 = oObj3
        ElseIf oObj2 Is oObj3 Then
            Set oObjCommon = oObj2
            Set oObjNbor1 = oObj1
            Set oObjNbor2 = oObj4
        ElseIf oObj2 Is oObj4 Then
            Set oObjCommon = oObj2
            Set oObjNbor1 = oObj1
            Set oObjNbor2 = oObj3
        Else
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename, oMemberPartCommon
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Splits do not share a common output member"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Splits do not share a common output member"
        End If
    
        If Not oObjCommon Is oMemberPartCommon Then
            MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename, oMemberPartCommon
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Splits common output member is not can output"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Splits common output member is not can output"
        End If

        ' only check neighbor identity if neighbor collection is okay
        If MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors).count = 2 Then
            If Not oObjNbor1 Is MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors).Item("1") Then
                MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename, oObjNbor1
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Neighbor1 is not same as split neighbor"
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Neighbor1 is not same as split neighbor"
            End If
            If Not oObjNbor2 Is MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors).Item("2") Then
                MyGC.PostError INVALID_PARAMETERS, True, strCodeListTablename, oObjNbor2
                WriteToErrorLog E_FAIL, MODULE, METHOD, "Neighbor2 is not same as split neighbor"
                Err.Raise E_FAIL, MODULE & ":" & METHOD, "Neighbor2 is not same as split neighbor"
            End If
        End If
    End If
    
    Call InLineEvaluate(MyGC)

    Exit Sub
ErrorHandler:
    
    Err.Raise E_FAIL
End Sub
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Const METHOD = "IJGCMigrate_Migrate"
    On Error GoTo ErrorHandler
    
    ' migrate will get called when a nbor is split, or my original split
    
    Dim MyCanRule As ISPSCanRule
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim oCurve As IJCurve
    Dim startX As Double, startY As Double, startZ As Double, endX As Double, endY As Double, endZ As Double
    Dim posMe As IJPoint

    Dim bDeleted As Boolean
    Dim countMembers As Long
    Dim elesMembers As IJElements
    Dim elesReplacing As IJElements
    Dim oMiddleMember As Object, oNbor1 As Object, oNbor2 As Object
    Dim oMemberPartCommon As ISPSMemberPartCommon
    Dim oSO As IJSmartOccurrence
    
    Set MyCanRule = MyGC
    Set MyGCMacro = MyGC
    Set elesMembers = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_MemberToSplit)
    countMembers = elesMembers.count

    Set posMe = MyGC
    posMe.GetPoint startX, startY, startZ

    If countMembers = 1 Then         ' first time primary split.  replacing the initial part

        ' get the replacing members
        ' find the one that I am PointOn to = oMiddleMember
        ' find the ones at the start and end of oMiddleMember
        '
        ' set the def of oMiddleMember to be the BUCan
        ' clear all inputs
        ' set neighbors so that "1" is the start, and "2" is the end.
        ' set output to be the oMiddleMember

        If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).count > 0 Then      ' should never happen.  already have an output member with a MemberToSplit.
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Inconsistency: MemberToSplit cannot co-exist with output member"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Inconsistency: MemberToSplit cannot co-exist with output member"
        End If

        pMigrateHelper.ObjectsReplacing elesMembers.Item("1"), elesReplacing, bDeleted
        
        FindClosestMember startX, startY, startZ, elesReplacing, Nothing, oMiddleMember

        Set oSO = oMiddleMember
        oSO.Properties = oSO.Properties Or 32       ' tell SO to use interfaces declared as input as input
        
        Set oCurve = oMiddleMember
        oCurve.EndPoints startX, startY, startZ, endX, endY, endZ

        FindClosestMember startX, startY, startZ, elesReplacing, oMiddleMember, oNbor1
        FindClosestMember endX, endY, endZ, elesReplacing, oMiddleMember, oNbor2

        If oNbor1 Is oNbor2 Then        'should not happen.  trying to place a can near end of member system ?
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Unexpected result: oNbor1 is oNbor2"
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Unexpected result: oNbor1 is oNbor2"
        End If
        
        elesMembers.Clear
        Set elesMembers = Nothing
        Set elesMembers = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors)          'get the neighbors collection
        
        ' we are writing to the BUCan on an interface which is input to BUCan custom semantic
        'the code below prevents (also an assert) that interface being treated as an ouput interface in the CAD
        
        Set oSO = oNbor1
        oSO.Properties = oSO.Properties Or 32
        Set oSO = oNbor2
        oSO.Properties = oSO.Properties Or 32
    
        elesMembers.Add oNbor1, "1"
        elesMembers.Add oNbor2, "2"

        SetNameRule oMiddleMember
        MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Add oMiddleMember, "1"
        SetBUCanDefinition oMiddleMember
        
    Else

        ' check whether the output can is being replaced
        Dim bCanReplaced As Boolean

        Set oMiddleMember = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")
        
        If oMiddleMember Is Nothing Then
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Output Can missing."
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Output Can missing."
        End If

        bCanReplaced = False
        pMigrateHelper.ObjectsReplacing oMiddleMember, elesReplacing, bDeleted
        If Not elesReplacing Is Nothing Then
            bCanReplaced = True
            If Not bDeleted And elesReplacing.count = 1 Then
                If elesReplacing.Item(1) Is oMiddleMember Then
                    bCanReplaced = False
                End If
            End If
        End If
        If bCanReplaced Then
            WriteToErrorLog E_FAIL, MODULE, METHOD, "Splitting a Can is not supported."
            Err.Raise E_FAIL, MODULE & ":" & METHOD, "Splitting a Can is not supported."
        End If
    
        Set elesMembers = Nothing
        ' the current neighbors, if any.  might be less than two for InLineCan if a split is beyond end of MemberSystem
        Set elesMembers = MyGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors)
                        
        Set oCurve = oMiddleMember
        oCurve.EndPoints startX, startY, startZ, endX, endY, endZ
       
        If elesMembers.count > 0 Then
        pMigrateHelper.ObjectsReplacing elesMembers.Item("1"), elesReplacing, bDeleted
        If Not elesReplacing Is Nothing Then
            If elesReplacing.count > 0 Then
                FindClosestMember startX, startY, startZ, elesReplacing, oMiddleMember, oNbor1
                elesMembers.Remove ("1")
                Set oSO = oNbor1
                oSO.Properties = oSO.Properties Or 32

                elesMembers.Add oNbor1, "1"
            End If
            Set elesReplacing = Nothing
        End If
        End If
        
        If elesMembers.count > 1 Then
            pMigrateHelper.ObjectsReplacing elesMembers.Item("2"), elesReplacing, bDeleted
            If Not elesReplacing Is Nothing Then
                If elesReplacing.count > 0 Then
                    FindClosestMember endX, endY, endZ, elesReplacing, oMiddleMember, oNbor2
                    elesMembers.Remove ("2")
    
                    Set oSO = oNbor2
                    oSO.Properties = oSO.Properties Or 32
                    
                    elesMembers.Add oNbor2, "2"
                End If
                Set elesReplacing = Nothing
            End If
        End If
      
    End If

    Exit Sub

'   ToDoListNotify(INSULATIONERROR_TDLCODELISTNAME, INSULATIONERROR_NOSELECTION, iInsulation

ErrorHandler:
    WriteToErrorLog E_FAIL, MODULE, METHOD, "Unexpected error"
    Err.Raise E_FAIL, MODULE & ":" & METHOD, "Unexpected error"
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

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    On Error GoTo ErrorHandler
    
    IJUserAttributeMgmt_OnPreLoad = vbNullString
    
    Dim oAttrDescr As IJAttributeDescriptor
    Dim diameterRule As Long
    
    Dim oRule  As IJAttributeDescriptor
    Dim oCanOD As IJAttributeDescriptor
    Dim oCanID As IJAttributeDescriptor
    Dim oCanUD As IJAttributeDescriptor
    
    Dim oL2Method As IJAttributeDescriptor
    Dim oL2Factor As IJAttributeDescriptor
    Dim oL2Length As IJAttributeDescriptor
    
    Dim oL3Method As IJAttributeDescriptor
    Dim oL3Factor As IJAttributeDescriptor
    Dim oL3Length As IJAttributeDescriptor
    
    Dim oCone1Method As IJAttributeDescriptor
    Dim oCone1Slope  As IJAttributeDescriptor
    Dim oCone1Length As IJAttributeDescriptor
    Dim oCone1Angle  As IJAttributeDescriptor
    
    Dim oCone2Method As IJAttributeDescriptor
    Dim oCone2Slope  As IJAttributeDescriptor
    Dim oCone2Length As IJAttributeDescriptor
    Dim oCone2Angle  As IJAttributeDescriptor
    
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
            
            Case attrL3CompMethod:      Set oL3Method = oAttrDescr
            Case attrL3Factor:          Set oL3Factor = oAttrDescr
            Case attrL3Length:          Set oL3Length = oAttrDescr
            
            Case attrCone1LengthMethod: Set oCone1Method = oAttrDescr
            Case attrCone1Slope:        Set oCone1Slope = oAttrDescr
            Case attrCone1Angle:        Set oCone1Angle = oAttrDescr
            Case attrCone1Length:       Set oCone1Length = oAttrDescr
            
            Case attrCone2LengthMethod: Set oCone2Method = oAttrDescr
            Case attrCone2Slope:        Set oCone2Slope = oAttrDescr
            Case attrCone2Angle:        Set oCone2Angle = oAttrDescr
            Case attrCone2Length:       Set oCone2Length = oAttrDescr
        
        End Select
    Next oAttrDescr
    
    oRule.AttrState = AttributeDescriptor_ReadOnly
    RuleSetReadOnly oRule.AttrValue, oCanID, oCanOD
    
    LMethodSetReadOnly oL2Method.AttrValue, oL2Length, oL2Factor
    LMethodSetReadOnly oL3Method.AttrValue, oL3Length, oL3Factor
    
    ConeMethodSetReadOnly oCone1Method.AttrValue, oCone1Length, oCone1Slope, oCone1Angle
    ConeMethodSetReadOnly oCone2Method.AttrValue, oCone2Length, oCone2Slope, oCone2Angle
    
    Exit Function
ErrorHandler:
    IJUserAttributeMgmt_OnPreLoad = Err.Description
End Function

Private Function ISPSCanRuleHelper_GetTubeDiameter(ByVal CanRule As SPSMembers.ISPSCanRule, tubeDiameter As Double) As SPSMembers.SPSCanRuleStatus
    ISPSCanRuleHelper_GetTubeDiameter = GetTubeDiameter(CanRule, True, tubeDiameter)
End Function

Private Function ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions(ByVal CanRule As ISPSCanRule) As SPSCanRuleStatus

    ISPSCanRuleHelper_UpdateOutputCrossSectionDimensions = UpdateOutputCrossSectionDimensions(True, CanRule)

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



