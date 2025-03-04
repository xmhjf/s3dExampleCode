VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlanarCutbackDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : PlanarCutbackDef.cls
'
'Author : RP
'
'Description :
'    Module for creating cutback for the member part
'
' 02/16/04  RP    Added a parameter, resource manager, to the
'                           cope and cutback creation methods. When resouce
'                           manager is nothing, the plane or wire body is not
'                           persisted.
'                           Moved the code that hides the plane/wirebody to the
'                           CMConstruct()
' 08/27/04  MH   CMMigrate
' 09/27/06  AS   TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
'
' 03/14/08  SS   DI#134831  Changed the code from CreateObject() to SP3DCreateObject()
'                           as the symbol is no longer registered.
'
'**********************************************************************************************************************Option Explicit

Option Explicit

Private Const MODULE = "PlanarCutbackDef"
Private Const strSourceFile As String = "PlanarCutbackDef.cls"
' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFeatureMacros.PlanarCutbackDef"
Const m_ItemName As String = "SPSFeatureMacros.PlanarCutbackDef"
Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1
Dim bOnPreLoad As Boolean
Private m_oLocalizer As IJLocalizer
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  Dim oInput As IJDInput
  Dim oInputs As IJDInputs
  
  pIH.SetInput "Supported" 'member part
  pIH.SetInput "Supporting1" 'plane,plate or member part
  
  Set oInputs = pIH.Definition
  Set oInput = New DInput
  oInput.Name = "Supporting2" 'optional plane, plate or member part
  oInput.index = 3
  oInput.Properties = igDESCRIPTION_OPTIONAL
  oInputs.Add oInput
  
  Set oInput = Nothing
  Set oInputs = Nothing
  
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub


Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler
  Dim strIIDPlanarCutback As String
  Dim oMetaData As IJDAttributeMetaData
  Dim oObject As IJDObject
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
  
  'Create a custom component
  'pDefinition.SupportUpdate = igSYMBOL_SUPPORT_UPDATE
  'pDefinition.SupportUpdate = igSYMBOL_CACHE_OPTION_NOT_SHARED

  ' Define the inputs -
  ' They are identical to the class inputs (i.e. penetrated and Penetrating)
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
'  pIH.InitAs m_FamilyProgid
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{CA64DDA7-EB18-4346-89C1-21C45FA8B836}"  'CStructFeature2
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  Set oObject = pDefinition
  Set oMetaData = oObject.ResourceManager
  strIIDPlanarCutback = oMetaData.IID("IJUASPSPlanarCutback")
  pAPDs.AddProperty "IJUASPSPlanarCutback", 1, strIIDPlanarCutback
  Set oObject = Nothing
  Set oMetaData = Nothing
'  pAPDs.AddProperty "IJUASPSPlanarCutback", 1, "{28718714-CFB9-4ced-BF89-DB40D2547EB4}"
'  pAPDs.AddProperty "IJUASPSPlanarCutback", 1, "{10191AD6-664B-4714-AE13-3B7B8308A656}"
  Set pAPDs = Nothing

'CRE - create a CutBackPlane drive here, should create general case for this
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
   
  Set pMemberDescription = pMemberDescriptions.AddMember("PlanarCutback", 1, "CMConstructCutback", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputCutback"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseCutback"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalCutback"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateCutback"
  
  Set pPropertyDescriptions = pMemberDescription
 ' outputs IJPlane = {4317C6B3-D265-11D1-9558-0060973D4824}
  pPropertyDescriptions.AddProperty "ComputeCutback", 1, "{4317C6B3-D265-11D1-9558-0060973D4824}", "CMComputeCutback", imsCOOKIE_ID_USS_LIB

  Exit Sub
  
ErrorHandler:  HandleError MODULE, MT
  
End Sub


Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler
    
  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalCutback(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalCutback"
  On Error GoTo ErrorHandler
    
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oRefColl   As IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSmartOcc As IJSmartOccurrence
    Dim strError As String

    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    If IsReadOnlyObject(oSuppedPart) Then
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_READONLY_OBJECT, oSmartOcc, Nothing
        Err.Raise SPS_MACRO_WARNING
        
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
Public Sub CMConstructCutback(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const MT = "CMConstructCutback"
    On Error GoTo ErrorHandler
    Dim oControlFlags As IJControlFlags
    
    'MsgBox MT
    
    ComputePlanarCutback pMemberDescription.CAO, pObj
    
    'hide the plane
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
      
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
Public Sub CMMigrateCutback(ByVal pMemberDescription As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateCutback"
    On Error GoTo ErrorHandler
    
    'MsgBox METHOD
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Set pSmartOccurrence = pMemberDescription.CAO

    Dim pReferencesCollection As IJDReferencesCollection
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)

    Dim pSplitAxisPort As ISPSSplitAxisPort
    Set pSplitAxisPort = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(1)

    Dim pMemberPartPrismatic As ISPSMemberPartPrismatic
    Set pMemberPartPrismatic = pSplitAxisPort.Part

    Call pMemberPartPrismatic.AddCutbackSurface(pSplitAxisPort.portIndex, pMemberDescription.object)

    Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMSetInputCutback(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputCutback"
     On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMComputeCutback(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMComputeCutback"
    On Error GoTo ErrorHandler
    
    ComputePlanarCutback pPropertyDescriptions.CAO, pObject
    
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub

Public Sub CMReleaseCutback(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseCutback"
On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler: HandleError MODULE, MT

End Sub


' The following methods are generic for all the Custom assembly
'
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
  Const MT = "CMFinalConstructCutout"
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
ErrorHandler:      HandleError MODULE, MT
End Function
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

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FEATUREMACROS_ERROR, "ERROR")
    
    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long
    
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
'            IJUserAttributeMgmt_OnAttributeChange = "ERROR::Bad Value"
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
        End If
    End If
    
    IJUserAttributeMgmt_OnAttributeChange = ""
   
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_FEATUREMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String
    
    'get the parent
    Dim oParent As IJSmartOccurrence
    Dim oAttribMgmntParent As IJUserAttributeMgmtParent
    Dim oSmartItem As IJSmartItem
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pCollectionHelper As IMSRelation.DCollectionHelper
    Dim pRelationshipHelper As DRelationshipHelper
    Dim pRevision As IJRevision
    
    Set pRelationHelper = pIJDAttrs
    Set pCollectionHelper = pRelationHelper.CollectionRelations("{BCBFB3C0-98C2-11D1-93DE-08003670A902}", "toAssembly1")
    If Not pCollectionHelper Is Nothing Then
        If pCollectionHelper.Count = 1 Then
            Dim mum As Variant
            mum = 1
            Set oParent = pCollectionHelper.Item(mum)
        End If
    End If
    
    If Not oParent Is Nothing Then
    
        Set oSmartItem = oParent.ItemObject
        On Error Resume Next
        Set oAttribMgmntParent = SP3DCreateObject(oSmartItem.Definition)
        On Error GoTo ErrorHandler
        If Not oAttribMgmntParent Is Nothing Then
            ErrStr = oAttribMgmntParent.OnPreLoad(oParent, pIJDAttrs, CollAllDisplayedValues)
        End If
    End If
        
    Set pAttrColl = CollAllDisplayedValues
    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
    Next
    
    bOnPreLoad = False

    IJUserAttributeMgmt_OnPreLoad = ""
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_ERROR, "ERROR")

    Dim dInputs As IJDInputs
    Dim CurrentInput As IJDInput
    Dim oAttribute As IJDAttribute
    Dim PC As DParameterContent
    Dim bvalid As Boolean
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = pIJDAttrs
    Dim oSymbolDef As IJDSymbolDefinition
    Dim ErrMessage As String
    Set oSymbolDef = oSymbolOcc.IJDSymbolDefinition(2)
    Set dInputs = oSymbolDef.IJDInputs
    Set PC = New DParameterContent
    
    Set oAttribute = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item(sAttributeName)
    If oAttribute.Value <> "" Then
        If oAttribute.AttributeInfo.Type = igString Then    ' check for string type here
        Else
            PC.UomValue = oAttribute.Value
            Set CurrentInput = Nothing
            bvalid = True
            On Error Resume Next
            Set CurrentInput = dInputs.GetInputByName(oAttribute.AttributeInfo.Name)
            If Not CurrentInput Is Nothing Then
                CurrentInput.IJDInputDuringGame.Definition = oSymbolDef
                CurrentInput.IJDInputStdCustomMethod.InvokeCMCheck PC, bvalid, ErrMessage
                CurrentInput.IJDInputDuringGame.Definition = Nothing
                Set oSymbolOcc = Nothing
                Set oSymbolDef = Nothing
                If bvalid = False Then
'                    UserAttributeMgmt_Validate = "Symbol CMCheck Failed"
                    UserAttributeMgmt_Validate = ErrMessage
                    Exit Function
                Else
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
' get the list of interfaces implemented by the schema from IJDAttributes
' make sure that you are not looking into a system interface
' from the input interfaceName and propertyName, get the property type from catalog info
' select case on the property types, and in there, mention the valid attribute values for each propertyName
    Dim InterfaceID As Variant
'    Dim oAttrObj As IJDAttribute
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
'    Dim oAttrCol As IJDAttributesCol
    Dim oAttrCol As IJDInfosCol
    Dim IsInterfaceFound As Boolean
    Dim AttrCount As Long
    Dim AttrType As Long
    
    Set oAttributeMetaData = pIJDAttrs
    IsInterfaceFound = False
    For Each InterfaceID In pIJDAttrs
        Set oInterfaceInfo = Nothing
        Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(InterfaceID)
        If (oInterfaceInfo.IsHardCoded = False) Then
            If (oInterfaceInfo.Name = sInterfaceName) Then
                IsInterfaceFound = True
                Exit For
            End If
        End If
    Next
    
'    Set oAttributeMetaData = Nothing
    Set oInterfaceInfo = Nothing
    
    If IsInterfaceFound = False Then
        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
    If sAttributeName = "Clearance" Then
        Dim pStructConn As IJAppConnection
        Dim depth As Double
        Dim SecWidth As Double
        Dim colPorts As IJElements
        Dim oCrossAttrs As IJDAttributes
        Dim pPort As IJPort
        Dim pDesignChild As IJDesignChild
        Dim pDesignParent As IJDesignParent
        Dim pMemberPart As ISPSMemberPartPrismatic
    
        Set pDesignChild = pIJDAttrs
        If Not pDesignChild Is Nothing Then
            Set pDesignParent = pDesignChild.GetParent
            If Not pDesignParent Is Nothing Then
                Set pStructConn = pDesignParent
                If Not pStructConn Is Nothing Then
                    pStructConn.enumPorts colPorts
                    ' get the cross section depth from the member part
                    If colPorts.Count > 0 Then
                        Set pPort = colPorts.Item(1)
                        If Not pPort Is Nothing Then
                            Set pMemberPart = pPort.Connectable
                            If Not pMemberPart Is Nothing Then
                                Set oCrossAttrs = pMemberPart.CrossSection.Definition
                                If Not oCrossAttrs Is Nothing Then
                                    depth = oCrossAttrs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If varAttributeValue > (0.5 * depth) Then
            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_CLEAREANCE_MEMBLENGTH, "Clearance > half the member length")
            Exit Function
        End If
    End If
    
'    Set oAttrCol = pIJDAttrs.CollectionOfAttributes(InterfaceID)
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.Name = sAttributeName Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (varAttributeValue < 0#) Then
'                            UserAttributeMgmt_Validate = sAttributeName
                           UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_VALUE_MUSTBE_POSITIVE, "Value must be equal to or greater than Zero")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
            End Select
        End If
    Next
    
    UserAttributeMgmt_Validate = ""
    Set oAttributeMetaData = Nothing

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function


Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
