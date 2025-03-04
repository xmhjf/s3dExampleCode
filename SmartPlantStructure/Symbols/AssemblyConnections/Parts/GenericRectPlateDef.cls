VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericRectPlateDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2003 - 2004 Intergraph Corporation. All rights reserved.
'
'
'
'   Custom assembly definition for the plate part
'   --------------------------------------------
'
'
'History
'    RP         03/26/03      Creation
'    SS         03/14/08      DI#134831  Changed the code from CreateObject() to SP3DCreateObject()
'                             as the symbol is no longer registered.
'---------------------------------------------------------------------------------------

Option Explicit

Private Const MODULE = "GenericRectPlateDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSPartMacros.GenericRectPlateDef"
Const m_ItemName As String = "SPSPartMacros.GenericRectPlateDef"
' Declaration of the User Symbol Services interface
Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1
Dim bOnPreLoad As Boolean

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
'  pIH.InitAs m_FamilyProgid
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{A46498E6-9116-42B1-8A18-031415C07428}"     'CStructCustomPlatePart
  pAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAggregator"
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  'listens to IJUASPSPlatePartDim
  pAPDs.AddProperty "PlateProperties", 1, "IJUASPSPlatePartDim"
  'listens to IJUASPSPlateThickness
  pAPDs.AddProperty "PlateThickness", 2, "IJUASPSPlateThickness"
  
  Set pAPDs = Nothing
   
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
  Set pMemberDescriptions = Nothing
  'No assembly members for this parent

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'Aggregator custom methods.......................................................
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler
  'Get the smart plate object
  Dim oCustomPlate As IJStructCustomPlatePart
  Set oCustomPlate = pAggregatorDescription.CAO
  
  'Set the generation pattern of the smart plate geometry
    SetCustomPlatePartGenerationAE oCustomPlate
    
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMSetInputAggregator"
On Error GoTo ErrorHandler
    Dim pIJSymbolOfCA0 As IMSSymbolEntities.IJDSymbol
 
    Dim i As Long, lcnt As Long
    Dim oRCSymbol As Object
    Set pIJSymbolOfCA0 = pAggregatorDescription.CAO
    Set oRCSymbol = pIJSymbolOfCA0.IJDReferencesArg.GetReferencesCollection

    Dim pIRCAsm As IJDReferencesCollection
    Dim oRefColl1 As IMSSymbolEntities.IJDReferencesCollection
    Set pIRCAsm = GetRefCollFromSmartOccurrence(pAggregatorDescription.CAO)

    If pIRCAsm Is Nothing Then
    Exit Sub
    End If

    If Not oRCSymbol Is Nothing Then
        pIJSymbolOfCA0.IJDReferencesArg.SetReferencesCollection pIRCAsm
    Exit Sub 'already have one get out unless we want to change it's contents
    Else 'create a new or initial one

    ' Create the new symbol input reference collection and copy the inputs reference

        pIJSymbolOfCA0.IJDReferencesArg.SetReferencesCollection pIRCAsm

    End If

  
Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
End Sub



'
' The following methods are generic for all the Custom assembly
'
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
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
ErrorHandler:  HandleError MODULE, MT
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

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
    Call GetRelatedParts(pPartObject, pIJMonUnks)
End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = StandardGraphicEntity
End Sub


Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & App.EXEName
    
    IJUserAttributeMgmt_OnAttributeChange = oLocalizer.GetString(IDS_PARTMACROS_ERROR, "ERROR")

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
    Set oLocalizer = Nothing

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & App.EXEName
    
    IJUserAttributeMgmt_OnPreLoad = oLocalizer.GetString(IDS_PARTMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String
    
    Set pAttrColl = CollAllDisplayedValues
    
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
    
    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
    Next
    
    bOnPreLoad = False
    Set oLocalizer = Nothing

    IJUserAttributeMgmt_OnPreLoad = ""
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler
    Dim oLocalizer As IJLocalizer
    Set oLocalizer = New IMSLocalizer.Localizer
    oLocalizer.Initialize App.Path & "\" & App.EXEName
    
    UserAttributeMgmt_Validate = oLocalizer.GetString(IDS_PARTMACROS_ERROR, "ERROR")

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
        UserAttributeMgmt_Validate = oLocalizer.GetString(IDS_PARTMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
'    Set oAttrCol = pIJDAttrs.CollectionOfAttributes(InterfaceID)
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.Name = sAttributeName Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (varAttributeValue <= 0#) Then
'                            UserAttributeMgmt_Validate = sAttributeName
                            UserAttributeMgmt_Validate = oLocalizer.GetString(IDS_PARTMACROS_VALUE_NEGATIVE, "Negative Attribute Value")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
            End Select
        End If
    Next
    
    UserAttributeMgmt_Validate = ""
    Set oAttributeMetaData = Nothing
    Set oLocalizer = Nothing

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function
