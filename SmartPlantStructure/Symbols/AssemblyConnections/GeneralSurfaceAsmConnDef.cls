VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeneralSurfaceAsmConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : GeneralSurfaceAsmConnDef.cls
'
'Author : R. Peringoth
'
'Description :
'    Macro for creating trimmed connection
'
'History:
'
' 01/30/04  RP   created
' 02/05/04  JS   Added code to check if duplicate
'                            ACs are being attached to a member port
'                            and added the specific relations to the
'                            reference collection
' 02/16/04  RP   Instead of IJSmartOccurrence, IJUASPSPlanarCutback
'                            is output for the cutback member. This way the parent
'                            gets notified when user edits cutback properties through
'                            the feature objects GUI and it can reset the edits.
' 08/27/04  MH   CMMigrate
' 05/11/06  AS   Added Pad and replaced the Planar cutback with SurfaceTrim to handle dynamicsurfacefind (CR#79947)
' 07/26/06  AS   TR#102438 Gen Surface AC does not work correctly against IS Plates. Changed inputs to Trim and Pad from surface to port
' 27/Sep/06 AS   TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
' 25/Oct/06 AS   TR#107418 Replaced IJUAGenSurfACPad I/F with IJDAttributes in InitializeSymbol definition and Added check in the
'                Member conditionals to prevent error during synchronization of 6.1 Model
' 01-Feb-07 AS   TR#109129 GenSurface AC should validate the geometric condition of supg and supd
' 07-Jun-07 MH   TR 121116 added CheckForUndefinedValueAndRaiseError
'********************************************************************

Option Explicit

Private Const MODULE = "GeneralSurfaceAsmConnDef"

Const m_ItemProgId As String = "SPSACMacros.GeneralSurfaceAsmConnDef"
Const m_ItemName As String = "SPSACMacros.GeneralSurfaceAsmConnDef"
Private Const strSourceFile As String = "GeneralSurfaceAsmConnDef.cls"

Private m_oLocalizer As IJLocalizer

Implements ISPSFACInputHelper
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJUserAttributeMgmtParent
Implements IJStructCustomFoulCheck

Private Enum GenSurfACMembers
  PlanarCutback = 1
  PadPlate = 2
  SurfaceTrim = 3
End Enum

'*************************************************************************
'Function
'DefinitionInputs
'
'Abstract
'Sets up the inputs for the base plate FrameConnection
'
'Arguments
'pIH As IJDInputsHelper - Input
'
'Exceptions
'
'***************************************************************************

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  Dim oInput As IJDInput
  Dim oInputs As IJDInputs
  
  Set oInputs = pIH.Definition 'optional Plane
  Set oInput = New DInput
  oInput.Name = "Plane"
  oInput.index = 1
  oInput.Properties = igDESCRIPTION_OPTIONAL
  oInputs.Add oInput
  
  Set oInput = Nothing
  Set oInputs = Nothing
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub




'*************************************************************************
'Function
'IJDUserSymbolServices_InitializeSymbolDefinition
'
'Abstract
'Bulk load called method that establishes the Base Plate assembly connection
'definition
'
'Arguments
'pDefinition - Symbol definition created for the base plate assembly connection
'
'Exceptions
'
'***************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler
  
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs for the base plate connection (the helper class is initialized to the
  ' symbol definition and is an aid to help set the inputs)
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{E43FD681-1B37-4CC1-BD94-F399F43F946F}"     'CStructAssemblyConnection
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
  
  ' add the interfaces here. so that semantic here is executed when those interfaces are modified
  ' Listens to IJUASPSGeneralSurfaceAsmConn, GenSurfAsmConnPad, PlateSizeParams- recompute the AC when anyone of these change
  pAPDs.AddProperty "GeneralSurfaceAsmConn", 1, "IJDAttributes"
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
  
  '_________________________________________________________________________________________________________________________________________
  ' Cutback Smart Occurrence
  Set pMemberDescription = pMemberDescriptions.AddMember("GeneralSurfaceAsmConnCutback", 1, "CMConstructGeneralSurfaceAsmConnCutback", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGeneralSurfaceAsmConnCutback"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGeneralSurfaceAsmConnCutback"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateGeneralSurfaceAsmConnCutback"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJUASPSPlanarCutback (Method to prevent direct edit of cutback properties from child object's GUI)
  pPropertyDescriptions.AddProperty "ComputeGeneralSurfaceAsmConnCutback", 1, "IJUASPSPlanarCutback", "CMComputeGeneralSurfaceAsmConnCutback", imsCOOKIE_ID_USS_LIB
  
  '_________________________________________________________________________________________________________________________________________
  ' Pad Smart Occurrence' Part of CR #79947
  Set pMemberDescription = pMemberDescriptions.AddMember("GeneralSurfaceAsmConnPad", 2, "CMConstructGeneralSurfaceAsmConnPad", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGeneralSurfaceAsmConnPad"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGeneralSurfaceAsmConnPad"
    
  '_________________________________________________________________________________________________________________________________________
  ' Surface Trim Smart Occurrence Part of CR #79947. Use SUrface Trim as the trimming feature
  Set pMemberDescription = pMemberDescriptions.AddMember("GeneralSurfaceAsmConnTrim", 3, "CMConstructGeneralSurfaceAsmConnTrim", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGeneralSurfaceAsmConnTrim"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGeneralSurfaceAsmConnTrim"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJUASPSSurfaceTrim
  pPropertyDescriptions.AddProperty "ComputeGeneralSurfaceAsmConnCutback", 1, "IJUASPSSurfaceTrim", "CMComputeGeneralSurfaceAsmConnCutback", imsCOOKIE_ID_USS_LIB
    
  Set pMemberDescriptions = Nothing
  Set pMemberDescription = Nothing
  Set pPropertyDescriptions = Nothing

'MsgBox MT
  Exit Sub
ErrorHandler:
  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMFinalConstructAggregator
'
'Abstract
'Permits one-time initialization of the user-defined aggregator class.
'
'Arguments
'IJDAggregatorDescription interface of the CustomAssemblyDefinition wrapper
'
'Exceptions
'
'***************************************************************************
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructAggregator"
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

'*************************************************************************
'Function
'CMConditionalGeneralSurfaceAsmConnCutback
'
'Abstract
'Callback method to determine whether the current member is needed as an output.
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'bIsNeeded - boolean which is always returned as true to indicate that a
'            cutback is always needed.
'
'Exceptions
'
'***************************************************************************
Public Sub CMConditionalGeneralSurfaceAsmConnCutback(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalGeneralSurfaceAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim oSuppedPort As ISPSSplitAxisEndPort
    Dim oMembPort As ISPSSplitAxisPort
    Dim oSuppingPort As IJPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    
    Dim strError As String
    Dim colPorts As IJElements
    
    'Set to false as it is now replaced by SrfaceTrim
    Dim oAttribsParent As IJDAttributes
    Dim oInterfaceColl As CollectionProxy
    
    Set oStructConn = pMemberDescription.CAO
    Set oAttribsParent = oStructConn
    On Error Resume Next
    Set oInterfaceColl = oAttribsParent.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    Err.Clear
    On Error GoTo ErrorHandler
    
    'If the new Pad I/F exists on the parent attributes dont need the cutback
    If Not oInterfaceColl Is Nothing Then
        bIsNeeded = False
    End If
    Set oInterfaceColl = Nothing
    
    If bIsNeeded Then
        ' Retrieve the inputs of the custom assembly occurrence
        GetACSupdAndSupgPort oStructConn, oSuppedPort
        ' Verify we do not have an assembly connection already attached to the end ports
        
        '   because if we do then this asssembly connection needs to have its relations
        '   to the ports severed and the assembly connection added to the ToDo list
        
        If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
            
            SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
            Err.Raise E_FAIL
        End If
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


'*************************************************************************
'Function
'CMConstructGeneralSurfaceAsmConnCutback
'
'Abstract
'Callback method for constructing the Base Plate Plane for the custom assembly
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'pResourceManager - resource manager for the plant model database
'pObj - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMConstructGeneralSurfaceAsmConnCutback(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructGeneralSurfaceAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim oFeatureFactory As New StructFeatureFactory
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oRefColl   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oSmartOccFac As IJSmartOccurrenceFactory
    Dim oDesignParent As IJDesignParent
    Dim oSuppingPlane As IJPlane
    Dim oPlate As IJStructCustomPlatePart
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppedPort As IJPort

    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    oStructConn.enumPorts colPorts
    If colPorts.count = 1 Then
        Set oSuppedPort = colPorts.Item(1)
        Set oRefColl = GetRefCollFromSmartOccurrence(oSOParent)
        On Error Resume Next
        ' Check if there is a plane in the ref coll
        ' if there is on we use that for positioning the plate
        Set oSuppingPlane = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        On Error GoTo ErrorHandler
    ElseIf colPorts.count = 2 Then ' a plane is also input
        Set oSuppedPort = GetEndPort(colPorts)
        If oSuppedPort Is colPorts.Item(1) Then
            Set oSuppingPlane = colPorts.Item(2)
        Else
            Set oSuppingPlane = colPorts.Item(1)
        End If
    End If
    
    Set oSmartOccFac = New SmartOccurrenceFactory
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    ' Generate name
    GenerateNameForFeature oSOChild
    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTPlanarCutbackFeatureOcc"
    oSOChild.RootSelection = "PlanarCutbackFeature"
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oSuppingPlane, IJPlane, "SPSSuppPlaneToRC_DEST"
    ' Connect the Smart Occurrence to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    Set pObj = oSOChild 'Return the newly created object
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMSetInputGeneralSurfaceAsmConnCutback
'
'Abstract
'Callback method used to set inputs on the parent SmartOccurrence prior to each CM evaluate
'
'Arguments
'pMemberDesc - Metadata describing the components of the custom
'              assembly which include cutback. This metdata information was defined
'              at bulkload time via the
'              IJDUserSymbolServices_InitializeSymbolDefinition function.
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputGeneralSurfaceAsmConnCutback(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputGeneralSurfaceAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim pIJAttribsParent    As IJDAttributes
    Dim pIJAttribsChild    As IJDAttributes
    Dim dblClearance As Double
    
    Set pIJAttribsParent = pMemberDesc.CAO
    Set pIJAttribsChild = pMemberDesc.object
    dblClearance = pIJAttribsParent.CollectionOfAttributes("IJUASPSGeneralSurfaceAsmConn").Item("Clearance").Value
    
    pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("Clearance").Value = dblClearance
    pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("SquaredEnd").Value = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMComputeGeneralSurfaceAsmConnCutback
'
'Abstract
'Callback Method to notify the base plate connection that the cutback
'needs to be recomputed
'
'Arguments
'pPropertyDescriptions - collection of the property descriptions for the custom
'                        assembly defined at bulk load time
'                        (see IJDUserSymbolServices_InitializeSymbolDefinition).
'pObject - Base plate custom assembly's Smart Occurrence instance
'
'Exceptions
' E_FAIL -raised when any of the input ports are missing which
'           forces the GeneralSurface to the ToDo list
'
'***************************************************************************
Public Sub CMComputeGeneralSurfaceAsmConnCutback(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeGeneralSurfaceAsmConnCutback"
    On Error GoTo ErrorHandler
    
    'Nothing to compute
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMMigrateGeneralSurfaceAsmConnCutback(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateGeneralSurfaceAsmConnCutback"
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
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), IJPlane, "SPSSuppPlaneToRC_DEST"
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'*************************************************************************
'Function
'CMSetInputGeneralSurfaceAsmConnPad
'
'Abstract
'Callback method used to set inputs on the parent SmartOccurrence prior to each CM evaluate
'
'Arguments
'pMemberDesc - Metadata describing the components of the custom
'              assembly which include cutback. This metdata information was defined
'              at bulkload time via the
'              IJDUserSymbolServices_InitializeSymbolDefinition function.
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputGeneralSurfaceAsmConnPad(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputGeneralSurfaceAsmConnPad"
    'MsgBox MT
    On Error GoTo ErrorHandler
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oIJAttribsParent    As IJDAttributes
    Dim oIJAttribsChild     As IJDAttributes
    Dim lSizingRule         As Integer
    Dim dClearance          As Double
    Dim bWithPad            As Boolean
    Dim oInterfaceColl As CollectionProxy
    Dim oStructConn As IJAppConnection
    Dim oStructPlate As IJStructPlate
    Dim oInterfaceCollOcc As CollectionProxy
        
    Set oStructConn = pMemberDesc.CAO
    Set oIJAttribsParent = pMemberDesc.CAO
    
    Set oStructPlate = pMemberDesc.object
    Set oSmartOccChild = pMemberDesc.object
    Set oIJAttribsChild = oSmartOccChild
    
    Set oInterfaceColl = oIJAttribsParent.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    lSizingRule = oInterfaceColl.Item("SizingRule").Value
    dClearance = oInterfaceColl.Item("Offset").Value
    bWithPad = oInterfaceColl.Item("WithPad").Value
        
    Set oInterfaceColl = oIJAttribsChild.CollectionOfAttributes("IJUAPlateSizeParams")
    If lSizingRule > 0 Then
        oInterfaceColl.Item("SizingRule").Value = lSizingRule
    End If
    oInterfaceColl.Item("Offset").Value = dClearance
    
    'MsgBox "leaving" + MT
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub     'CMSetInputGeneralSurfaceAsmConnPad


'*************************************************************************
'Function
'CMConditionalGeneralSurfaceAsmConnPad
'
'Abstract
'Callback method to determine whether the Pad is needed as an output. Currently
' only checks for the WithPad attribute
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'bIsNeeded - boolean which is always returned as true to indicate that a
'            cutback is always needed.
'
'Exceptions
'
'***************************************************************************
Public Sub CMConditionalGeneralSurfaceAsmConnPad(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalGeneralSurfaceAsmConnPad"
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim oSuppedPort As ISPSSplitAxisEndPort
    Dim oMembPort As ISPSSplitAxisPort
    Dim oSuppingPort As IJPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    
    Dim strError As String
    Dim colPorts As IJElements
    Dim oIJAttribs    As IJDAttributes
    
    Dim oInterfaceAttrColl As CollectionProxy
        
    Dim bWithPad As Boolean
    Dim iIdx As Integer
    
    ' Retrieve the inputs of the custom assembly occurrence
    
    Set oStructConn = pMemberDescription.CAO
    Set oIJAttribs = pMemberDescription.CAO

    ' addedtthis code to initialize the attribuites on Assy Conn
    Dim oIJAttribsDef As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    'Dim oInterfaceCollDef As CollectionProxy
    Dim oInterfaceCollOcc As CollectionProxy
        
    Set oSmartOccCAO = oStructConn
    
    Set oIJAttribsDef = oSmartOccCAO.ItemObject
    
    On Error Resume Next
    'Set oInterfaceCollDef = oIJAttribsDef.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    Set oInterfaceCollOcc = oIJAttribs.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    Err.Clear
    On Error GoTo ErrorHandler
    If oInterfaceCollOcc Is Nothing Then
       bIsNeeded = False
       Exit Sub
    End If
    'CopyValuesToSOFromItem oInterfaceCollOcc, oInterfaceCollDef
    
    bWithPad = False
    'check the withPad attribute value
    bWithPad = oInterfaceCollOcc.Item("WithPad").Value
       
    If bWithPad Then
        GetACSupdAndSupgPort oStructConn, oSuppedPort

        ' Verify we do not have an assembly connection already attached to the end ports
        '   because if we do then this asssembly connection needs to have its relations
        '   to the ports severed and the assembly connection added to the ToDo list
        If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
            SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
            Err.Raise E_FAIL
            
        End If
        
        'Check if the section is valid. Otherwise wont be able to get Pad size
        Set oMembPort = oSuppedPort
        Set oSuppedPart = oMembPort.Part
        If Not ValidSectionType(oSuppedPart) Then
            SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_UNKNOWNSEC_ENCOUNT, oStructConn, Nothing
            Err.Raise E_FAIL
            
        End If
    End If
    
    bIsNeeded = bWithPad
  '  MsgBox MT + "conditional for pad returned " + CStr(bIsNeeded)
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

'*************************************************************************
'Function
'CMConstructGeneralSurfaceAsmConnPad
'
'Abstract
'Callback method for constructing the Pad for the custom assembly
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'pResourceManager - resource manager for the plant model database
'pObj - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMConstructGeneralSurfaceAsmConnPad(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructGeneralSurfaceAsmConnPad"
'    MsgBox MT
    On Error GoTo ErrorHandler
    'SO and Ref Coll
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl   As IJDReferencesCollection
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oAttrbs As IJDAttributes
    Dim oEditJDArgument  As IJDEditJDArgument
    Dim oSmartOccCAO      As IJSmartOccurrence
    Dim oDesignParent As IJDesignParent
    
    'inputs
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPort As IJPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oStructConn As IJAppConnection
    Dim oSupgPort As Object
    Dim oSupgSurface As Object
    Dim colPorts As IJElements
    
    'Plate and its properties
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim lPadType As Long
    Dim materialProxy As CollectionProxy
    Dim dPlateThickness As Double
    Dim oPlatePort As IJPort
    
    'For Surface Trim
    Dim oStructSymbolTools As IJStructSymbolTools
    Dim oAssyMembers As IJDMemberObjects
    Dim oTrimSO As IJSmartOccurrence
    Dim oTrimRC As IJDReferencesCollection
    
    Dim strError As String
    
    ' Retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    
    Set oFactory = New StructCustomPlatePartFactory
    ' Create the custom Plate that is the Pad
    Set oPlate = oFactory.CreateCustomPlatePart(pResourceManager)
    Set oFactory = Nothing
    
    Set oAttrbs = oSmartOccCAO.ItemObject
        
    Set materialProxy = oAttrbs.CollectionOfAttributes("IJUASPSPlateMaterial")
    If Not materialProxy Is Nothing Then
        Dim strMaterial As String, strGrade As String
        strMaterial = materialProxy.Item("SPSMaterial").Value
        strGrade = materialProxy.Item("SPSGrade").Value
        SetPlateMaterial oPlate, strMaterial, strGrade
    Else
        SetPlateMaterial oPlate, "Steel - Carbon", "A"
    End If
    
    ' Get default from the parent
    dPlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value
    ' Set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, dPlateThickness
    
    Set oSmartOcc = oPlate
    ' Initialize the SmartOccurrence
    'Set the root selector to the correct Smart class depending on the Pad Type attribute
    lPadType = oAttrbs.CollectionOfAttributes("IJUAGenSurfAsmConnPad").Item("PadType").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, lPadType, "StructACPadType", 32

   'TODO: Change the logic to setthe root selector to the correct shape pad SC after the
   ' catalog content is avl.
    If lPadType = 2 Then
       oSmartOcc.RootSelectorClass = "PadPlateRectangular"
    ElseIf lPadType = 3 Then
       oSmartOcc.RootSelectorClass = "PadPlateRectangular" ' "CircularPadPlatePart"
    ElseIf lPadType = 4 Then
       oSmartOcc.RootSelectorClass = "PadPlateRectangular" ' "TriangularPadPlatePart"
    Else
       oSmartOcc.RootSelectorClass = "PadPlates"
    End If
    
    GetACSupdAndSupgPort oStructConn, oSuppedPort, oSupgPort, oSupgSurface
    If oSupgPort Is Nothing Then
        'In case of equipment foundations and other guys whoi do not offer ports, just use the suppingsurface
        Set oSupgPort = oSupgSurface
    End If
    ' Create the reference collection
    Set oRefCollChild = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)
    Set oEditJDArgument = oRefCollChild.IJDEditJDArgument
    oEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
    oEditJDArgument.SetEntity 2, oSupgPort, ConstIJSurface, "SurfaceToRC_DEST"
    
    ' Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    ' Create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefCollChild
    
    'call the evaluate to generate the plat's faces
    'The feature needs the face as input for trimming the supported member
    oSmartOcc.EvaluateSmartOccurrenceCES
    
    Set pObj = oPlate ' Return the created plate object to the custom assembly
    
    'MsgBox "leaving" + MT
    Exit Sub

ErrorHandler:
    Err.Raise ReportError(Err, strSourceFile, MT, strError).Number
    
    End Sub 'CMConstructGeneralSurfaceAsmConnPad


'*************************************************************************
'Function
'CMConditionalGeneralSurfaceAsmConnTrim
'
'Abstract
'Callback method to determine whether the Surface Trim is needed as an output. Currently
' only checks for the WithPad attribute
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'bIsNeeded - boolean which is always returned as true to indicate that a
'            cutback is always needed.
'
'Exceptions
'
'***************************************************************************
Public Sub CMConditionalGeneralSurfaceAsmConnTrim(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalGeneralSurfaceAsmConnTrim"
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim oSuppedPort As ISPSSplitAxisEndPort
    Dim oSupgPort As IJPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSOParent As IJSmartOccurrence
    Dim oRefColl   As IJDReferencesCollection
    Dim colPorts As IJElements
    Dim oSupgSurface As Object
    
    Dim strError As String
    
    'No trim -default
    bIsNeeded = False
        
    ' Retrieve the inputs of the custom assembly occurrence
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = oSOParent
    
    'Set to false as it is now replaced by SrfaceTrim
    Dim oAttribsParent As IJDAttributes
    Dim oInterfaceColl As CollectionProxy
    
    Set oAttribsParent = oStructConn
    On Error Resume Next
    Set oInterfaceColl = oAttribsParent.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    Err.Clear
    On Error GoTo ErrorHandler
    
    'If the new Pad I/F does not exist on the parent attributes dont need the surfacetrim
    If oInterfaceColl Is Nothing Then
        bIsNeeded = False
        Exit Sub
    End If
    Set oInterfaceColl = Nothing
    
    GetACSupdAndSupgPort oStructConn, oSuppedPort, oSupgPort, oSupgSurface
    
    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed and the assembly connection added to the ToDo list
    If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
        Err.Raise E_FAIL

    End If
    
    'Now check if the input surface is valid for Surface Trim- type and geometric condition
    bIsNeeded = IsSurfaceTypeAcceptableForSurfaceTrim(oSuppedPort, oSupgSurface)
    'MsgBox MT
    
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

'*************************************************************************
'Function
'CMSetInputGeneralSurfaceAsmConnTrim
'
'Abstract
'Callback method used to set inputs on the parent SmartOccurrence prior to each CM evaluate
'
'Arguments
'pMemberDesc - Metadata describing the components of the custom
'              assembly which include cutback. This metdata information was defined
'              at bulkload time via the
'              IJDUserSymbolServices_InitializeSymbolDefinition function.
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputGeneralSurfaceAsmConnTrim(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputGeneralSurfaceAsmConnTrim"
    On Error GoTo ErrorHandler
    Dim oAttribsParent    As IJDAttributes
    Dim oAttribsChild    As IJDAttributes
    Dim dOffset As Double
    Dim dOccValue As Double
    Dim colSurfaceTrim As CollectionProxy
    
    Dim oStructConn As IJAppConnection
    Dim oStructSymbolTools As IJStructSymbolTools
    Dim oPadPlate As IJStructCustomPlatePart
    Dim oPlatePort As IJPort
    Dim oSupdPort As Object
    Dim oSupgPort As Object
    Dim oSupgSurface As Object
    Dim oAssyMembers As IJDMemberObjects
    Dim oTrimSO   As IJSmartOccurrence
    Dim oTrimRC As IJDReferencesCollection
    
    Set oAttribsParent = pMemberDesc.CAO
    Set oAttribsChild = pMemberDesc.object
    Set oTrimSO = oAttribsChild
    Set oTrimRC = GetRefCollFromSmartOccurrence(oTrimSO)
    
    'get the trim object if it exist
    Set oAssyMembers = oAttribsParent
    
    'Also change the supg surface of the trim to the the Pad offset surface
    'get the offset port from the plate object if it exist
    Set oPadPlate = oAssyMembers.Item(GenSurfACMembers.PadPlate)  'second member of the CA is the plate
    'get the plate's port that is on the supported member side
    If Not oPadPlate Is Nothing Then
        Set oStructSymbolTools = New StructSymbolTools
        oStructSymbolTools.GetLateBindingPort oPadPlate, JS_TOPOLOGY_PROXY_LFACE, 320153976, 0, 138, -1, oPlatePort
        Set oStructSymbolTools = Nothing
        oTrimRC.IJDEditJDArgument.SetEntity 2, oPlatePort, ConstIJSurface, "SurfaceToRC_DEST"
    Else
        Set oStructConn = oAttribsParent
        GetACSupdAndSupgPort oStructConn, oSupdPort, oSupgPort, oSupgSurface    'Always connect to port. Otherwise error
        If oSupgPort Is Nothing Then
        'In case of equipment foundations and other guys whoi do not offer ports, just use the suppingsurface
            Set oSupgPort = oSupgSurface
        End If
        oTrimRC.IJDEditJDArgument.SetEntity 2, oSupgPort, ConstIJSurface, "SurfaceToRC_DEST"
        'oTrimRC.IJDEditJDArgument.SetEntity 2, oSupgSurface, ConstIJSurface, "SurfaceToRC_DEST"
    End If
        
    'MsgBox "leaving" + MT + "clearance =" + Str(dOffset)
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConstructGeneralSurfaceAsmConnTrim
'
'Abstract
'Callback method for constructing the SurfaceTrim for the custom assembly
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'pResourceManager - resource manager for the plant model database
'pObj - assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMConstructGeneralSurfaceAsmConnTrim(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructGeneralSurfaceAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim oFeatureFactory As New StructFeatureFactory
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oRefColl   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oSmartOccFac As IJSmartOccurrenceFactory
    Dim oDesignParent As IJDesignParent
    Dim oSupgPort As Object
    Dim oSuppingSurface As Object
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppedPort As IJPort
    Dim oStructSymbolTools As IJStructSymbolTools
    
    Dim oAssMembs As IJDMemberObjects
    
    Dim oPlate As IJStructCustomPlatePart
    Dim oPlatePort As IJPort

    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = oSOParent
    
    'construct the cmn struct feature object
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oFeatureFactory = Nothing
    
    'add a design child of the AC
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    ' Generate name
    GenerateNameForFeature oSOChild
    
    'Set user type. don't know why is it needed?
    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTSurfaceTrimOcc"
    
    'root selector is surface trim
    oSOChild.RootSelection = "SurfaceTrim"
    
    'now create the ref collection for inputs
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)
    Set oSymFactory = Nothing
    
    GetACSupdAndSupgPort oStructConn, oSuppedPort, oSupgPort, oSuppingSurface
    If oSupgPort Is Nothing Then
        'In case of equipment foundations and other guys whoi do not offer ports, just use the suppingsurface
        Set oSupgPort = oSuppingSurface
    End If
    
    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oSupgPort, ConstIJSurface, "SurfaceToRC_DEST"
    'oRefCollChild.IJDEditJDArgument.SetEntity 2, oSuppingSurface, ConstIJSurface, "SurfaceToRC_DEST"
    ' Connect the Smart Occurrence to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    Set pObj = oSOChild 'Return the newly created object
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub 'CMConstructGeneralSurfaceAsmConnTrim

'*************************************************************************
'Function
'IJDUserSymbolServices
'   InstanciateDefinition
'
'Abstract
'Instantiates a persistent symbol definition object and initializes it for the first time,
'returning a pointer (ppSymbolDefDisp) to the IDispatch interface of the initialized symbol definition.
'
'Arguments:
'codeBase specifies the URL (or UNC) of the .cab file that can provides the dll associated to the symbol definition object (ActiveX® control packaging).
'definitionParameters  Definition parameters.
'pResourceMgr  resource manager to which the symbol definition will be connected.
'
'Return:
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions:
'
'***************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
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


'*************************************************************************
'Function
'IJDUserSymbolServices
'   GetDefinitionName
'
'Abstract
'Used during the execution of IJDDefinitionCollection::GetDefinitionByProgId to get the definition name
'based upon the definitionParameters passed in. It returns the definition name (pDefName) if it already
'exists within the collection. The name of a definition is the identifier of the definition object
'in the definition collection and assures its uniqueness in the given resource manager.
'
'Arguments
'definitionParameters
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function


'*************************************************************************
'Function
'
'IJDUserSymbolServices
'   InvokeRepresentation
'
'Abstract
'Obsolete method
'
'***************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub


'*************************************************************************
'Function
'IJDUserSymbolServices_EditOccurence
'
'Abstract
'Obsolete method --instead you can record your custom command within the
'definition (see IJDCommandDescription interface)
'
'***************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function


'*************************************************************************
'Function
'IJStructCustomFoulCheck
'   GetConnectedParts
'
'Abstract
'Adds objects that are related with the SmartOccurrence ( input objects ) to the list of connected objects
'
'Arguments
'pPartObject is the SmartOccurrence object
'pIJMonUnks is the list of monikers
'
'Return
'
'Exceptions
'
'***************************************************************************
Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
End Sub


'*************************************************************************
'Function
'IJStructCustomFoulCheck
'   GetFoulInterfaceType
'
'Abstract
'Returns type of interference participant that this object is.
'
'Arguments
'pFoulInterfaceType is the InterferenceChecking type
'
'Return
'
'Exceptions
'
'***************************************************************************
Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnAttributeChange
'
'Abstract
'Gets called for each attribute change on the property page
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of attributes as currently displayed ( prior to Commit )
'pAttrToChange is which attribute is being edited
'varNewAttrValue is the value given by the user.
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
End Function


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnPreCommit
'
'Abstract
'Gets called before the attribute changes are committed to allow a check of validity.
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of attributes as currently displayed ( prior to Commit )
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
End Function


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnPreLoad
'
'Abstract
'Gets called prior to display of attributes on the property page to set readOnly status
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of IJAttributeDescriptor's
'
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Dim ii As Long
    Dim bWithPad As Boolean
    Dim oCollProxy As CollectionProxy
    
    Set oCollProxy = pIJDAttrs.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    
    bWithPad = oCollProxy.Item("WithPad").Value
    If Not bWithPad = True Then
        'If not with pad then gray out the attributes regarding pad
        Dim pAttrColl As Collection
        Set pAttrColl = CollAllDisplayedValues
        Dim pAttrDescr As IJAttributeDescriptor
        For ii = 1 To pAttrColl.count
            Set pAttrDescr = pAttrColl.Item(ii)
            If pAttrDescr.InterfaceName = "IJUAGenSurfAsmConnPad" Then
                If pAttrDescr.AttrName = "SizingRule" Or pAttrDescr.AttrName = "PadType" Or pAttrDescr.AttrName = "Offset" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            End If
        Next
    End If
    
End Function


'*************************************************************************
'Function
'ISPSFACInputHelper
'   ValidateObjects
'
'Abstract
'Validates the given objects and returns the objects to be used for assemblyConnection inputs
'
'Arguments
'inputObjs is a collection of objects
'relationObjs is the collection to be used to build relationships
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
End Function


Private Function IJUserAttributeMgmtParent_OnAttributeChange(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
'
End Function

Private Function IJUserAttributeMgmtParent_OnPreCommit(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
'
End Function

Private Function IJUserAttributeMgmtParent_OnPreLoad(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    'get the parent
    Dim oParent As IJDMemberObjects
    Dim oAttribMgmntChild As IJUserAttributeMgmtParent
    Dim oSmartItem As IJSmartItem
    Dim ii As Long
    Dim oChild As Object
    
    Set oParent = pParent
    
    If Not oParent Is Nothing Then
       
       'Handle Planar cutback attributes
        Set oChild = oParent.ItemByDispid(1)
        If oChild Is pIJDAttrs Then

            Dim pAttrColl As Collection
            Set pAttrColl = CollAllDisplayedValues
            Dim pAttrDescr As IJAttributeDescriptor
            For ii = 1 To pAttrColl.count
                Set pAttrDescr = pAttrColl.Item(ii)
                If pAttrDescr.InterfaceName = "IJUASPSPlanarCutback" Then
                    'parent controlling  child's property
                    'make property read only in child's GUI
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            Next

        End If
        
    End If

End Function

'*************************************************************************
'Function
'ISPSFACInputHelper
'   UserAttributeMgmt
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
Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
End Property


Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal inputObjs As SP3DStructInterfaces.IJElements, relationObjs As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const MT = "ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler

    Dim oInputObj1 As Object
    Dim oInputObj2 As Object
    
    ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_Ok
    If inputObjs.count = 2 Then
        Set oInputObj1 = inputObjs.Item(1)
        Set oInputObj2 = inputObjs.Item(2)
        If Not (oInputObj1 Is Nothing Or oInputObj2 Is Nothing) Then
            If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
                'Now check if the input surface is valid for Surface Trim- type and geometric condition
                If IsSurfaceTypeAcceptableForSurfaceTrim(oInputObj1, oInputObj2) Then
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
            ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
        End If
    Else
        ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
    End If
    Exit Property
ErrorHandler:
    HandleError MODULE, MT
End Property

'*************************************************************************
'Function
'GetACSupdAndSupgPort
'
'Abstract
'Utility emthod to get the SUpd port and Supg Surface from the AC.
'
'Arguments
'oStructConn - Assy Conn
'oSupdPort - SplitAxisPort - Supd Port
'oSupgPort - SplitAxisPort - Supd Port
'oSupgSurface - Supporting surface for the AC
'***************************************************************************

Private Sub GetACSupdAndSupgPort(oStructConn As IJAppConnection, oSupdPort As ISPSSplitAxisPort, Optional oSupgPort As Object, Optional oSupgSurface As Object)
    Const MT = "GetACSupdAndSupgPort"
    
    Dim strError    As String
    Dim colPorts    As IJElements
    Dim oSmartOcc   As IJSmartOccurrence
    Dim oRefColl    As IJDReferencesCollection
    
    If Not oStructConn Is Nothing Then
    
        oStructConn.enumPorts colPorts
    
        If colPorts.count = 1 Then
            'Check if a valid member part is selcted
            Set oSupdPort = colPorts.Item(1)
            If oSupdPort Is Nothing Then
                SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_INPUTS_GENERALSURFACE, oStructConn, Nothing
                Err.Raise E_FAIL
            
            End If
            Set oSmartOcc = oStructConn
            Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
            If Not IsMissing(oSupgSurface) Then
                Set oSupgSurface = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
            End If
    
        ElseIf colPorts.count = 2 Then ' a surface is also input
            Set oSupdPort = GetEndPort(colPorts)
            If oSupdPort Is Nothing Then
                SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_INPUTS_GENERALSURFACE, oStructConn, Nothing
                Err.Raise E_FAIL
                
            End If
            
            If Not IsMissing(oSupgPort) Then
                If oSupdPort Is colPorts.Item(1) Then
                    Set oSupgPort = colPorts.Item(2)
                Else
                    Set oSupgPort = colPorts.Item(1)
                End If
                Set oSupgSurface = oSupgPort.Geometry
            End If
        Else
                SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_NOOF_INPUTS_GENERALSURFACE, oStructConn, Nothing
                Err.Raise E_FAIL
            
        End If
    
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


Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
