VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndPlateAsmConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : EndPlateAsmConnDef.cls
'
'Author : RP
'
'Description :
'    Macro for creating end plate  connection
'
'History:
'
' 08/14/03   RP   Substituted interface name in place of guid
'                           in IJDPropertyDescriptions->AddProperty. The guid of
'                           a virtual interface may change during bulkload
' 02/05/04  JS     Added code to validate there are not two
'                           ACs competing for the same member axis end port
'                           (i.e., we do not have duplicate ACs attached)
' 02/05/04  JS      Changed all generic reference collection relations to
'                           specific relations. This includes the relations
'                           to the split axis end ports, split axis along ports
'                           and the cutback planes (if a plane exists).
' 08/27/04  MH     CMMigrate
' 27/Sep/06 AS      TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
'
'********************************************************************

Option Explicit

Private Const MODULE = "EndPlateAsmConnDef"

Const m_ItemProgId As String = "SPSACMacros.EndPlateAsmConnDef"
Const m_ItemName As String = "SPSACMacros.EndPlateAsmConnDef"
Private Const strSourceFile As String = "EndPlateAsmConnDef.cls"
Private m_oLocalizer As IJLocalizer
Implements ISPSFACInputHelper
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
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
  ' Listens to IJUASPSEndPlateAsmConn
  pAPDs.AddProperty "EndPlateAsmConnProperties", 1, "IJUASPSEndPlateAsmConn", "CMComputeEndPlateAsmConnProperties", imsCOOKIE_ID_USS_LIB
  ' Listens to IJUASPSPlateTypeCategory
  pAPDs.AddProperty "PlateTypeCategory", 2, "IJUASPSPlateTypeCategory"
  
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
  '_________________________________________________________________________________________________________________________________________
  ' Plate Smart Occurrence
  Set pMemberDescription = pMemberDescriptions.AddMember("EndPlateAsmConnPlate", 1, "CMConstructEndPlateAsmConnPlate", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndPlateAsmConnPlate"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputEndPlateAsmConnPlate"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateEndPlateAsmConnPlate"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs  IJUASPSPlatePartDim
  pPropertyDescriptions.AddProperty "EndPlateSize", 1, "IJUASPSPlatePartDim", "CMComputeEndPlateSize", imsCOOKIE_ID_USS_LIB
  ' Outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "EndPlateThickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeEndPlateThickness", imsCOOKIE_ID_USS_LIB
  ' Outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "EndPlatePosition", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeEndPlatePosition", imsCOOKIE_ID_USS_LIB
  '_______________________________________________________________________________________________________________________________________
  ' Cope Smart Occurrence
  Set pMemberDescription = pMemberDescriptions.AddMember("EndPlateAsmConnCope", 2, "CMConstructEndPlateAsmConnCope", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndPlateAsmConnCope"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputEndPlateAsmConnCope"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateEndPlateAsmConnCope"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJSmartOccurrence. This is just to make sure that the parent computes before the child.
  pPropertyDescriptions.AddProperty "ComputeEndPlateAsmConnCope", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeEndPlateAsmConnCope", imsCOOKIE_ID_USS_LIB
  '___________________________________________________________________________________________________________________________________
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


Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    
    MigrateTheAggregator pAggregatorDescription, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'
' Plate Smart Occurrence custom methods
'
Public Sub CMComputeEndPlateAsmConnProperties(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMComputeEndPlateAsmConnProperties"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Sub CMMigrateEndPlateAsmConnPlate(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateEndPlateAsmConnPlate"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub


Public Sub CMConditionalEndPlateAsmConnPlate(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalEndPlateAsmConnPlate"
    On Error GoTo ErrorHandler
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPart As ISPSMemberPartPrismatic
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingPort As ISPSSplitAxisPort
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements

    Dim strError As String
    ' Retrieve the inputs of the custom assembly occurrence
    Set oStructConn = pMemberDescription.CAO
    
    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed and the assembly connection added to the ToDo list
    If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
    oStructConn.enumPorts colPorts
    
    If colPorts.count < 2 Then
        Dim iJDObject As iJDObject
        Set iJDObject = oStructConn
        iJDObject.Remove
        Err.Raise E_FAIL
    End If

    ' One and only one port for this assembly connection
    Set oSuppedPort = colPorts.Item(1)
    Set oSuppingPort = colPorts.Item(2)
    
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    If Not (ValidSectionType(oSuppedPart) And ValidSectionType(oSuppingPart)) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_UNKNOWNSEC_ENCOUNT, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
    bIsNeeded = False ' Set to false until plate is implemented for this connection
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


Public Sub CMConstructEndPlateAsmConnPlate(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructEndPlateAsmConnPlate"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMSetInputEndPlateAsmConnPlate(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputEndPlateAsmConnPlate"
  On Error GoTo ErrorHandler
  Dim pIJAttribsCAO As IJDAttributes
  Dim oSmartOccCAO As IJSmartOccurrence
  Dim oCustomPlate As StructCustomPlatePart
  
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


Public Sub CMComputeEndPlateSize(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeEndPlateSize"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMComputeEndPlateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeEndPlateThickness"
    On Error GoTo ErrorHandler
    ' This is a notification function and is called when the plate thickness is changed.
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMComputeEndPlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeEndPlatePosition"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

'
'Cope Smart Occurrence custom methods
'
Public Sub CMConditionalEndPlateAsmConnCope(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalEndPlateAsmConnCope"
    On Error GoTo ErrorHandler
    bIsNeeded = True
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMConstructEndPlateAsmConnCope(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructEndPlateAsmConnCope"
    On Error GoTo ErrorHandler
    Dim oFeatureFactory As New StructFeatureFactory
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oDesignParent As IJDesignParent
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppedPort As IJPort ' ISPSSplitAxisEndPort
    Dim oSuppingPort As IJPort ' ISPSSplitAxisAlongPort
    
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    oStructConn.enumPorts colPorts

    Set oSuppedPort = colPorts.Item(1)
    Set oSuppingPort = colPorts.Item(2)
    
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    '  Create name
    GenerateNameForFeature oSOChild

    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTCopeFeatureOcc"
    oSOChild.RootSelection = "CopeFeature"
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oSuppingPort, CONST_ISPSSplitAxisAlongPort, "RefColl"

    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild

    Set pObj = oSOChild ' return the newly created object
    Set oStructConn = Nothing
    Set oSymFactory = Nothing
    Set oSOParent = Nothing
    Set oSOChild = Nothing
    Set iUserType = Nothing
  
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMSetInputEndPlateAsmConnCope(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputEndPlateAsmConnCope"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
  
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsChild = pMemberDesc.object
    
    If Not pIJAttribsChild Is Nothing Then
        If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSCope")) Then
            ' This is probably executed when the parent and child are just constructed and the child's occ attributes are
            ' empty. In this case defaults are copied from the parent to the child
            Set oAttrbs = oSmartOccCAO.ItemObject
            ' Copy defaults from parent
            CopyValuesToSOFromItem pIJAttribsChild.CollectionOfAttributes("IJUASPSCope"), oAttrbs.CollectionOfAttributes("IJUASPSCope")
        End If
    End If
    
    Set oAttrbs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIJAttribsCAO = Nothing
    Set pIJAttribsChild = Nothing

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Sub CMComputeEndPlateAsmConnCope(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeEndPlateAsmConnCope"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMMigrateEndPlateAsmConnCope(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateEndPlateAsmConnCope"
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
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'
' The following methods are generic for all Custom assemblies
'
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' Method in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
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
    
  'returned symbol definition
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


Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
End Sub


Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub


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

Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal inputObjs As SP3DStructInterfaces.IJElements, relationObjs As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFACInputHelperStatus
Const MT = "ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler
    Dim oFrmConn As ISPSFrameConnection
    Dim oInputObj1 As Object, oInputObj2 As Object
    Dim portIdx As SPSMemberAxisPortIndex
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
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
    
    '  make sure there are only two ports
    If portCol.count = 2 Then
        Set oInputObj1 = portCol.Item(1)
        Set oInputObj2 = portCol.Item(2)
        If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            Set oSuppedPort = oInputObj1
            If TypeOf oInputObj2 Is ISPSSplitAxisAlongPort Then
                Set oSuppingPort = oInputObj2
                If relationObjs Is Nothing Then
                    Set relationObjs = New JObjectCollection
                End If
                relationObjs.Clear
                relationObjs.Add oSuppedPort
                relationObjs.Add oSuppingPort
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
