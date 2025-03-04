VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GussetPlateConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : GussetPlateAsmConnDef.cls
'
'Author : RP
'
'Description :
'    Macro for creating gusset plate  connection
'
'History:
'
' 08/14/03   RP             Substituted interface name in place of guid
'                           in IJDPropertyDescriptions->AddProperty. The guid of
'                           a virtual interface may change during bulkload
' 09/22/03   CE             Hide any persistent construction planes
' 02/02/04   JS             Added code to verify the assembly connection
'                           and check if an AC already exists between
'                           the supported or supporting members TR#52040
' 02/05/04   JS             Changed all generic reference collection relations to
'                           specific relations. This includes the relations
'                           to the split axis end ports, split axis along ports
'                           and the cutback planes (if a plane exists).
'04/12/04    RP             Check for proper orientation is moved to common.bas
'                           Also moved all checks to the first CMConditional()
'                           Also removed code from the unused plane member
' 08/27/04  MH              CMMigrate
' 06/13/06  RP              Changes due to impact from curved members. Validate method checks for linear members.
'                           Eror is returned when imput members are not linear type  (DI#84001)
' 27/Sep/06 AS              TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
'***************************************************************************************************************

Option Explicit

Private Const MODULE = "GussetPlateConnDef"

Const m_ItemProgId As String = "SPSACMacros.GussetPlateConnDef"
Const m_ItemName As String = "SPSACMacros.GussetPlateConnDef"
Private Const strSourceFile As String = "GussetPlateAsmConnDef.cls"
Private m_oLocalizer As IJLocalizer
Implements ISPSFACInputHelper
Implements IJDUserSymbolServices ' Declaration of the User Symbol Services interface
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    pIH.ReportError
End Sub


Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler

  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs -
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{E43FD681-1B37-4CC1-BD94-F399F43F946F}"     ' CStructAssemblyConnection
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  Set pAD = Nothing
  
  ' tr 74803
  Dim pCADefinition As IJCADefinition
  Set pCADefinition = pDefinition
  Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
  Set pCADefinition = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  ' Listens to IJUASPSPlateTypeCategory
  pAPDs.AddProperty "PlateTypeCategory", 1, "IJUASPSPlateTypeCategory"

  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
   
  '_________________________________________________________________________________________________________________________________________
  'Cutback Plane
  Set pMemberDescription = pMemberDescriptions.AddMember("GussetPlateAsmConnCutbackPlane", 1, "CMConstructGussetPlateAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGussetPlateAsmConnCutbackPlane"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGussetPlateAsmConnCutbackPlane"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateGussetPlateAsmConnCutbackPlane"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJPlane.
  pPropertyDescriptions.AddProperty "ComputeGussetPlateAsmConnCutbackPlane", 1, "{4317C6B3-D265-11D1-9558-0060973D4824}", "CMComputeGussetPlateAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB
   
  '_______________________________________________________________________________________________________________________________
  ' Cutback
  Set pMemberDescription = pMemberDescriptions.AddMember("GussetPlateAsmConnCutback", 2, "CMConstructGussetPlateAsmConnCutback", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGussetPlateAsmConnCutback"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGussetPlateAsmConnCutback"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateGussetPlateAsmConnCutback"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJSmartOccurrence. This is just to make sure that the parent computes before the child.
  pPropertyDescriptions.AddProperty "ComputeGussetPlateAsmConnCutback", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeGussetPlateAsmConnCutback", imsCOOKIE_ID_USS_LIB
  '________________________________________________________________________________________________________________________________
  ' GussetPlate
  Set pMemberDescription = pMemberDescriptions.AddMember("GussetPlateAsmConnPlate", 3, "CMConstructGussetPlateAsmConnPlate", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGussetPlateAsmConnPlate"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGussetPlateAsmConnPlate"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateGussetPlateAsmConnPlate"
  Set pPropertyDescriptions = pMemberDescription
  '_____________________________________________________________________________________________________________________________
  
  Set pMemberDescriptions = Nothing
  Set pMemberDescription = Nothing
  Set pPropertyDescriptions = Nothing

  Exit Sub
ErrorHandler:
  HandleError MODULE, MT
End Sub


'
' Aggregator custom methods
'
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructAggregator"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMComputeGussetPlateAsmConnProperties(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMComputeGussetPlateAsmConnProperties"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    
    MigrateTheAggregator pAggregatorDescription, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'
'Cutback Plane custom methods
'
Public Sub CMConditionalGussetPlateAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalGussetPlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    'unused member so always return false
    bIsNeeded = False
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMConstructGussetPlateAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructGussetPlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    'unused member so this method doesn't do anything
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMSetInputGussetPlateAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputGussetPlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    'unused member so this method doesn't do anything
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMComputeGussetPlateAsmConnCutbackPlane(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeGussetPlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    'unused member so this method doesn't do anything
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMMigrateGussetPlateAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateGussetPlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    
    'MsgBox METHOD
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'
'Cutback custom methods
'
Public Sub CMConditionalGussetPlateAsmConnCutback(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalGussetPlateAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPart As ISPSMemberPartPrismatic
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingPort As ISPSSplitAxisPort
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim strError As String

    bIsNeeded = False 'initialize to false
    
    ' Retrieve the inputs of the custom assembly occurrence
    Set oStructConn = pMemberDescription.CAO
    
    oStructConn.enumPorts colPorts
    
    Set oSuppedPort = GetEndPort(colPorts)
    
    If colPorts.count <> 2 Then
        Dim iJDObject As iJDObject
        Set iJDObject = oStructConn
        iJDObject.Remove

        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_NOOF_INPUTS_GSASM, oStructConn, Nothing
        Err.Raise E_FAIL
    
    End If
    
    If oSuppedPort Is colPorts.Item(1) Then
        Set oSuppingPort = colPorts.Item(2)
    Else
        Set oSuppingPort = colPorts.Item(1)
    End If
    
    If (oSuppedPort Is Nothing) Or (oSuppingPort Is Nothing) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_INPUT_GUSSETPLATE, oStructConn, Nothing
        Err.Raise E_FAIL
        
    End If
    
    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed and the assembly connection added to the ToDo list
    If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
        Err.Raise E_FAIL
        
    End If
    
    
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    If Not (ValidSectionType(oSuppedPart) And ValidSectionType(oSuppingPart)) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_UNKNOWNSEC_ENCOUNT, oStructConn, Nothing
        Err.Raise E_FAIL
        
    End If
    
    'check if Supping and Supported have their weak or strong axes  parallel
    bIsNeeded = CheckForGussetPlateAsmConn(oSuppedPart, oSuppingPart)
    
    If bIsNeeded = False Then ' Members are not in the same plane
        ' Return error so that the AC is put in the ToDo list
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_MEMBS_NOTIN_SAMEPLANE, oStructConn, Nothing
        Err.Raise E_FAIL
        
    End If
    Exit Sub
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If

End Sub


Public Sub CMConstructGussetPlateAsmConnCutback(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructGussetPlateAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim oFeatureFactory As New StructFeatureFactory
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim iUserType As IJDUserType
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oDesignParent As IJDesignParent
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPort As ISPSSplitAxisPort
    Dim oCAChildObjs As IJDMemberObjects
    
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Set oCAChildObjs = pMemberDescription.CAO
    oStructConn.enumPorts colPorts
    
    Set oSuppedPort = GetEndPort(colPorts)
    
    If oSuppedPort Is colPorts.Item(1) Then
        Set oSuppingPort = colPorts.Item(2)
    Else
        Set oSuppingPort = colPorts.Item(1)
    End If
    
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    'create name
    GenerateNameForFeature oSOChild

    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTPlanarCutbackFeatureOcc"
    oSOChild.RootSelection = "PlanarCutbackFeature"


    Set pObj = oSOChild ' Return the newly created object
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oSuppingPort, CONST_ISPSSplitAxisAlongPort, "RefColl" ' Supporting port
    ' Connect the Smart Occurrence to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMSetInputGussetPlateAsmConnCutback(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputGussetPlateAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
  
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    Set pIJAttribsChild = pMemberDesc.object
    If Not pIJAttribsChild Is Nothing Then
        If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback")) Then
            'This probably when the parent and child are just constructed and the child's def attributes are
            'not copied to occ
            Set oAttrbs = oSmartOccCAO.ItemObject
            ' Copy defaults from parent
            CopyValuesToSOFromItem pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback"), oAttrbs.CollectionOfAttributes("IJUASPSPlanarCutback")
        End If
    End If
    Set oSmartOccCAO = Nothing
    Set oSmartOccChild = Nothing
    Set pIJAttribsCAO = Nothing
    Set pIJAttribsChild = Nothing

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMComputeGussetPlateAsmConnCutback(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeGussetPlateAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
        
    'Nothing to compute here
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMMigrateGussetPlateAsmConnCutback(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateGussetPlateAsmConnCutback"
    On Error GoTo ErrorHandler
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Dim pReferencesCollection As IJDReferencesCollection
    Dim bIsInputMigrated As Boolean
    Dim oPoint As IJPoint
    Dim ii As Integer, eleCount As Integer
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    
    Dim oOld As Object
    Dim oReplacing() As Object

    'MsgBox METHOD

    Set pSmartOccurrence = pMemberDesc.object
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)
    
    GetPositionFromRefColl pReferencesCollection, oPoint

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    ReDim oReplacing(1 To eleCount)

    For ii = 1 To eleCount

        Set oOld = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        
        Call pMigrateHelper.ObjectsReplacing(oOld, pObjectCollectionReplacing, bIsDeleted)
    
        If Not pObjectCollectionReplacing Is Nothing Then
            bIsInputMigrated = True
            SelectReplacingObject pObjectCollectionReplacing, oPoint, oReplacing(ii)
        Else
            Set oReplacing(ii) = oOld
        End If
        
        Set oOld = Nothing
        Set pObjectCollectionReplacing = Nothing
        
    Next ii
     
    If bIsInputMigrated Then
        Call pReferencesCollection.IJDEditJDArgument.RemoveAll
        pReferencesCollection.IJDEditJDArgument.SetEntity 1, oReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), CONST_ISPSSplitAxisAlongPort, "RefColl" ' Supporting port
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub



'
' Plate Smart Occurrence custom methods
'
Public Sub CMConditionalGussetPlateAsmConnPlate(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalGussetPlateAsmConnPlate"
    On Error GoTo ErrorHandler
  
    bIsNeeded = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMConstructGussetPlateAsmConnPlate(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructGussetPlateAsmConnPlate"
    On Error GoTo ErrorHandler
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl   As IJDReferencesCollection
    Dim oSmartOccCAO      As IJSmartOccurrence
    Dim oSymbolOfCAO                As IJDSymbol
    Dim oDesignParent As IJDesignParent
    Dim PlateThickness As Double
    Dim oAttrbs As IJDAttributes
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPort As ISPSSplitAxisPort
    Dim oCAChildObjs As IJDMemberObjects
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    
    ' Retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Set oCAChildObjs = pMemberDescription.CAO
    oStructConn.enumPorts colPorts
    
    Set oSuppedPort = GetEndPort(colPorts)
    
    If oSuppedPort Is colPorts.Item(1) Then
        Set oSuppingPort = colPorts.Item(2)
    Else
        Set oSuppingPort = colPorts.Item(1)
    End If
    
    Set oDesignParent = pMemberDescription.CAO
    
    Set oFactory = New StructCustomPlatePartFactory
    ' Create the custom Plate that is the Base Plate.
    Set oPlate = oFactory.CreateCustomPlatePart(pResourceManager)

    Set oAttrbs = oSmartOccCAO.ItemObject
    On Error Resume Next
        Dim materialProxy As CollectionProxy
        'the interface below is added in V6. It may not exist in older catalog
        'New code and old catalog is unusual. The resume next above takes care of this though
        Set materialProxy = oAttrbs.CollectionOfAttributes("IJUASPSPlateMaterial")
    On Error GoTo ErrorHandler
    
    If Not materialProxy Is Nothing Then
        Dim strMaterial As String, strGrade As String
        strMaterial = materialProxy.Item("SPSMaterial").Value
        strGrade = materialProxy.Item("SPSGrade").Value
        SetPlateMaterial oPlate, strMaterial, strGrade
    Else
        'use hard coded material, as done prior to V6
        SetPlateMaterial oPlate, "Steel - Carbon", "A"
    End If
    
    ' Get default from the parent
    PlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value

    ' Set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, PlateThickness
    
    ' Initialize the SmartOccurrence
    Set oSmartOcc = oPlate
    oSmartOcc.RootSelectorClass = "GussetPlatePartType2"

    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)
    oRefColl.IJDEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefColl.IJDEditJDArgument.SetEntity 2, oSuppingPort, CONST_ISPSSplitAxisAlongPort, "RefColl"
    oRefColl.IJDEditJDArgument.SetEntity 3, oSuppedPort.Part, CONST_ISPSMemberPartGeometry, "RefColl"   ' Cutback member

    ' Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    ' Create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the Smart Occurrence to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefColl
    Set pObj = oPlate ' Return the created plate object to the custom assembly
  
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMSetInputGussetPlateAsmConnPlate(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputGussetPlateAsmConnPlate"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim oCustomPlate As StructCustomPlatePart
  
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsCAO = oSmartOccCAO.ItemObject
    Set pIJAttribsChild = pMemberDesc.object
    If Not pIJAttribsChild Is Nothing Then
        If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim")) Then
            CopyValuesToSOFromItem pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim"), pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlatePartDim")
        End If
        If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartCutDim")) Then
            CopyValuesToSOFromItem pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartCutDim"), pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlatePartCutDim")
        End If
        If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSGussetPlatePartType2")) Then
            CopyValuesToSOFromItem pIJAttribsChild.CollectionOfAttributes("IJUASPSGussetPlatePartType2"), pIJAttribsCAO.CollectionOfAttributes("IJUASPSGussetPlatePartType2")
        End If
    
    End If
    
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    On Error Resume Next ' Suppress the error that VB throws when QI fails
    Set oCustomPlate = pMemberDesc.object
    On Error GoTo ErrorHandler
    
    If Not oCustomPlate Is Nothing Then
        ' Checks if the occurrence attributes are Empty
        If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory")) Then
            ' Copies defintion values to occurrence values
            Dim oItemAttrbs As IJDAttributes
            Set oItemAttrbs = oSmartOccCAO.ItemObject
            CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory"), oItemAttrbs.CollectionOfAttributes("IJUASPSPlateTypeCategory")
        End If
        
        ' Copy plate category and type from parent to child
        oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Category").Value
        oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Type").Value
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMMigrateGussetPlateAsmConnPlate(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateGussetPlateAsmConnPlate"
    On Error GoTo ErrorHandler
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Dim pReferencesCollection As IJDReferencesCollection
    Dim bIsInputMigrated As Boolean
    Dim oPoint As IJPoint
    Dim ii As Integer, eleCount As Integer
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    
    Dim oOld As Object
    Dim oReplacing() As Object

    'MsgBox METHOD

    Set pSmartOccurrence = pMemberDesc.object
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)
    
    GetPositionFromRefColl pReferencesCollection, oPoint

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    ReDim oReplacing(1 To eleCount)

    For ii = 1 To eleCount

        Set oOld = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        
        Call pMigrateHelper.ObjectsReplacing(oOld, pObjectCollectionReplacing, bIsDeleted)
    
        If Not pObjectCollectionReplacing Is Nothing Then
            bIsInputMigrated = True
            SelectReplacingObject pObjectCollectionReplacing, oPoint, oReplacing(ii)
        Else
            Set oReplacing(ii) = oOld
        End If
        
        Set oOld = Nothing
        Set pObjectCollectionReplacing = Nothing
        
    Next ii
     
    If bIsInputMigrated Then
        Call pReferencesCollection.IJDEditJDArgument.RemoveAll
        pReferencesCollection.IJDEditJDArgument.SetEntity 1, oReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), CONST_ISPSSplitAxisAlongPort, "RefColl"
        pReferencesCollection.IJDEditJDArgument.SetEntity 3, oReplacing(3), CONST_ISPSMemberPartGeometry, "RefColl"   ' Cutback member
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'
' Generic Custom assembly routines
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' Method in charge of the creation of the symbol definition object
  Const MT = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  
  Dim pDefinition As IJDSymbolDefinition
  Dim pFact As IJCAFactory
  Set pFact = New CAFactory
  Set pDefinition = pFact.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_ItemProgId
  pDefinition.CodeBase = CodeBase
    
  ' Initialize the definition
  IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
  
  ' Persistence behavior
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Returned symbol definition
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function
  
ErrorHandler:
  HandleError MODULE, MT
End Function


Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function


' Obsolete method.
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub


' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function


'
' Interference detection routines
'
Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub


'
' Attribute management routines
'
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
End Function


Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
End Function


Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
End Function


Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
End Property


Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal inputObjs As SP3DStructInterfaces.IJElements, relationObjs As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const MT = "ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler

    Dim oInputObj1 As Object, oInputObj2 As Object
    Dim count As Integer
    Dim portCol As IJElements

    ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_Ok
    Set portCol = New JObjectCollection
    'filter out ports to portCol
    For count = 1 To inputObjs.count
        If TypeOf inputObjs.Item(count) Is IJPort Then
            portCol.Add inputObjs.Item(count)
        End If
    Next count
    
    '2 ports are expected
    If portCol.count = 2 Then
        Set oInputObj1 = portCol.Item(1)
        Set oInputObj2 = portCol.Item(2)
        If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            If TypeOf oInputObj2 Is ISPSSplitAxisAlongPort Then
                'Make sure all members are linear
                Dim oPort1 As ISPSSplitAxisPort
                Dim oPort2 As ISPSSplitAxisPort
            
                Set oPort1 = oInputObj1
                Set oPort2 = oInputObj2
                
                If (TypeOf oPort1.Part Is ISPSMemberPartLinear) And (TypeOf oPort2.Part Is ISPSMemberPartLinear) Then
                    If relationObjs Is Nothing Then
                        Set relationObjs = New JObjectCollection
                    End If
                    relationObjs.Clear
                    relationObjs.Add oInputObj1
                    relationObjs.Add oInputObj2
                Else
                    ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
                End If
            Else
                ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
            End If
        Else
            ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
        End If
    Else
        ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
    End If
    
    Exit Property

ErrorHandler:
    HandleError MODULE, MT
End Property

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
