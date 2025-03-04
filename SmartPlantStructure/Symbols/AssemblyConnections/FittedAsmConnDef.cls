VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FittedAsmConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : FittedAsmConnDef.cls
'
'Author : RP
'
'Description :
'    Macro for creating fitted end assembly connection
'
'History:
' 07/22/03   RP   Wrote it
' 02/02/04  JS      Added code to verify the assembly connection
'                           and check if an AC already exists between
'                           the supported or supporting members TR#52040
' 02/05/04  JS      Changed all generic reference collection relations to
'                           specific relations. This includes the relations
'                           to the split axis end ports, split axis along ports
'                           and the cutback planes (if a plane exists).
' 08/27/04  MH     CMMigrate
' 27/Sep/06 AS   TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
'********************************************************************

Option Explicit

Private Const MODULE = "FittedAsmConnDef"

Const m_ItemProgId As String = "SPSACMacros.FittedAsmConnDef"
Const m_ItemName As String = "SPSACMacros.FittedAsmConnDef"
Private Const strSourceFile As String = "FittedAsmConnDef.cls"
Private m_oLocalizer As IJLocalizer
Implements ISPSFACInputHelper
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck

'*************************************************************************
'Function
'DefinitionInputs
'
'Abstract
'Sets any required inputs for symbol evaluation
'
'Arguments
'IJDInputsHelper defined in CommonApp
'
'Return
'
'Exceptions
'
'***************************************************************************

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  Exit Sub
  
ErrorHandler:
  pIH.ReportError
End Sub



'*************************************************************************
'Function
'IJDUserSymbolServices_InitializeSymbolDefinition
'
'Abstract
'Cleans the previous definition up and initializes the new one (Input, Output, Representation,
'RepresenationEvaluation, ...) constructing the symbol definition by (re)defining the inputs, outputs,
'representations, and representation.
'Note:  The previous setting of the definition must be reset in this method before constructing the new definition.
'
'Arguments:
'pDefinition  Symbol definition passed by reference that will be initialized in this method.
'
'Return
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions
'
'***************************************************************************

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
    'No property to be notified as this connection doesn't have any occurrence property
    Set pAPDs = Nothing
    
    ' Define the members
    Dim pMemberDescriptions As IJDMemberDescriptions
    Dim pMemberDescription As IJDMemberDescription
    Dim pPropertyDescriptions As IJDPropertyDescriptions
    
    Set pMemberDescriptions = pDefinition
    ' Remove all the previous member descriptions
    pMemberDescriptions.RemoveAll
     
    'Cope
    Set pMemberDescription = pMemberDescriptions.AddMember("FittedAsmConnCope", 1, "CMConstructFittedAsmConnCope", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
    
    pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalFittedAsmConnCope"
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputFittedAsmConnCope"
    pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateFittedAsmConnCope"
    Set pPropertyDescriptions = pMemberDescription
    'outputs IJSmartOccurrence. This is just to make sure that the parent computes before the child.Is this necessary? Not sure now.
    pPropertyDescriptions.AddProperty "ComputeFittedAsmConnCope", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeFittedAsmConnCope", imsCOOKIE_ID_USS_LIB
    Set pMemberDescriptions = Nothing
    Set pMemberDescription = Nothing
    Set pPropertyDescriptions = Nothing
    
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
'Return
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
'CMConditionalFittedAsmConnCope
'
'Abstract
'Determines whether the current member is needed as an output.
'
'Arguments
'IJDMemberDescription interface of the member
'Boolean set to True if the Member is needed.
'
'Return
'
'Exceptions
'
'***************************************************************************

Public Sub CMConditionalFittedAsmConnCope(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalFittedAsmConnCope"
  On Error GoTo ErrorHandler

    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPart As ISPSMemberPartPrismatic
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingPort As ISPSSplitAxisPort
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements

    Dim strError As String
    'retrieve the inputs of the custom assembly occurrence
    Set oStructConn = pMemberDescription.CAO
    oStructConn.enumPorts colPorts
    
    If colPorts.count <> 2 Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_NOOF_INPUTS_FC, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    Set oSuppedPort = GetEndPort(colPorts)
    
    If oSuppedPort Is colPorts.Item(1) Then
        Set oSuppingPort = colPorts.Item(2)
    Else
        Set oSuppingPort = colPorts.Item(1)
    End If
    If (oSuppedPort Is Nothing) Or (oSuppingPort Is Nothing) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_INPUT_FITTED, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    If Not (ValidSectionType(oSuppedPart) And ValidSectionType(oSuppingPart)) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_UNKNOWNSEC_ENCOUNT, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed and the assembly connection added to the ToDo list
    If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
        Err.Raise E_FAIL
    End If

    bIsNeeded = True
  
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
'CMConstructFittedAsmConnCope
'
'Abstract
'Creates the output member object
'
'Arguments
'IJDMemberDescription interface of the member
'pResourceManager used to construct the member
'pObject is the constructed object
'
'Return
'
'Exceptions
'
'***************************************************************************

Public Sub CMConstructFittedAsmConnCope(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructFittedAsmConnCope"
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
    Dim oSuppedPort As ISPSSplitAxisEndPort
    Dim oSuppingPort As ISPSSplitAxisAlongPort

    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
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
ErrorHandler:      HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMSetInputFittedAsmConnCope
'
'Abstract
'Used to set inputs on the child SmartOccurrence prior to CM evaluate
'
'Arguments
'IJDMemberDescription interface of the wrapped output member
'
'Return
'
'
'Exceptions
'
'***************************************************************************

Public Sub CMSetInputFittedAsmConnCope(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputFittedAsmConnCope"
  On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim iSOCompute As IJSOCompute
    Dim bCopyAttribNeeded As Boolean
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsChild = pMemberDesc.object
    Set iSOCompute = pMemberDesc.CAO
    bCopyAttribNeeded = False
    
    If iSOCompute.PreviousSmartItem <> vbNullString Then
        If oSmartOccCAO.Item <> iSOCompute.PreviousSmartItem Then
            bCopyAttribNeeded = True
        End If
    End If

    If Not pIJAttribsChild Is Nothing Then
        If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSCope")) Then
        'This is probably executed when the parent and child are just constructed and the child's occ attributes are
        'empty. In this case defaults are copied from the parent to the child
            bCopyAttribNeeded = True
        End If
    End If
        
    If bCopyAttribNeeded Then
        Set oAttrbs = oSmartOccCAO.ItemObject
        'copy defaults from parent
        CopyValuesToSOFromItem pIJAttribsChild.CollectionOfAttributes("IJUASPSCope"), oAttrbs.CollectionOfAttributes("IJUASPSCope")
    End If

    Set oAttrbs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIJAttribsCAO = Nothing
    Set pIJAttribsChild = Nothing

  Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMComputeFittedAsmConnCope
'
'Abstract
'Evaluates the member property
'
'Arguments
'IJDPropertyDescription interface describing the property to be evaluated
'pObject is the object whose property is being computed
'
'Return
'
'Exceptions
'
'***************************************************************************

Public Sub CMComputeFittedAsmConnCope(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeFittedAsmConnCope"
    On Error GoTo ErrorHandler
        
    'Nothing to compute
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMMigrateFittedAsmConnCope(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateFittedAsmConnCope"
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

        MigrateMemberObject pMemberDesc.object, pMigrateHelper

    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

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
     ' This method is in charge of the creation of the symbol definition object
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
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
 ' Obsolete method.
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
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
    Set ISPSFACInputHelper_UserAttributeMgmt = Me
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
