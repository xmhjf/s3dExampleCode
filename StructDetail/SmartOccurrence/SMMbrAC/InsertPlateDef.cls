VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InsertPlateDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : InsertPlateDef.cls
'
'  Description : Control creation of the Member Items for the InsertPlateDef Class
'
'  Author      : Alligators
'
'  History     :
'    30/JUL/2012 - Created
'    26/AUG/2013 - skcheeka - TR-237880 : Added proper checks to handle cases where the
'                                         penetrating and penetrated ports in a collection gets reversed.
'
'*********************************************************************************************

Const m_sClassName As String = "InsertPlateDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

'*Require by S:\StructuralModeling\Client\AsmConnCommands
Implements ISPSFACInputHelper
Implements IJUserAttributeMgmt

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ItemInputs
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    'CStructAssemblyConnection
    pAD.AggregatorClsid = "{E43FD681-1B37-4CC1-BD94-F399F43F946F}"
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CM_FinalConstructAggregator"
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateAggregator"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    
    Dim sDispId As String
    
    Dim lDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    '******************************************************************************
    ' if any more item members are needed to be added please add only at the bottom
       
    ' should be responsible for creation of Slot
    lDispId = 1
    Set oMemDesc = pMDs.AddMember("SlotCutOut", lDispId, "CMConstructSlot", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsSlotNeeded"
    Set oMemDesc = Nothing
    
    lDispId = 2
    Set oMemDesc = pMDs.AddMember("Chamfer1", lDispId, "CMConstructLeftChamfer", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsChamfer"
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsLeftChamferNeeded"
    Set oMemDesc = Nothing

    lDispId = 3
    Set oMemDesc = pMDs.AddMember("Chamfer2", lDispId, "CMConstructRightChamfer", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsChamfer"
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsRightChamferNeeded"
    Set oMemDesc = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Sub

' ** Start CM **
'*********************************************************************************************
' If needed Add Custom Methods HERE
'*********************************************************************************************

Public Sub CM_IsSlotNeeded(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    bIsNeeded = True
    
    Dim lDispId As Long
    Dim oAppConnection As IJAppConnection
    
    lDispId = 0
    lDispId = pMD.dispid
    
    On Error Resume Next
    Set oAppConnection = pMD.CAO
    On Error GoTo ErrorHandler
    
    ' When the penetrated or penetrating objects ports are modified, the correspondig slot needs to be updated...
    If (bIsNeeded) And (Not oAppConnection Is Nothing) Then
        
        Dim oMemObj As IJDMemberObjects
        Set oMemObj = oAppConnection
        
        On Error Resume Next
        Dim oItemMember As Object
        Set oItemMember = oMemObj.ItemByDispid(lDispId)
        Set oMemObj = Nothing
        On Error GoTo ErrorHandler
        
        If Not oItemMember Is Nothing Then
            If TypeOf oItemMember Is IJStructFeature Then
                Dim oFeatureSmartOcc As IJSmartOccurrence
                Set oFeatureSmartOcc = oItemMember
                oFeatureSmartOcc.Update
                Set oFeatureSmartOcc = Nothing
            End If
            Set oItemMember = Nothing
        End If
    End If
   
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CM_IsSlotNeeded").Number
End Sub


Public Sub CMConstructSlot(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    strError = "Retrieving CAO and constructing slot."
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oPenetratingPort As IJPort
    Dim oPenetratedPort As IJPort
    GetPenetratedAndPenetratingPorts oAppConnection, oPenetratedPort, oPenetratingPort
        
    Dim oSlot As New StructDetailObjects.Slot
    Dim strStartClass As String
    Dim pSystemParent As IJSystemChild
    
    strStartClass = "InsertPlateSlot"
    Set pSystemParent = pMemberDescription.CAO
    Call oSlot.Create(pResourceManager, oPenetratingPort.Connectable, oPenetratedPort.Connectable, strStartClass, pSystemParent)
                               
    strError = "Setting Slot to private variable"
    Set pObject = oSlot.object
    
    oSlot.AddSlotToCut
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructSlot", strError).Number
End Sub

Public Sub CM_IsLeftChamferNeeded(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    bIsNeeded = False
    
    Dim sACItemName As String
    
    ' Get AC Smart Item Name
    Parent_SmartItemName pMD.CAO, sACItemName
    
    'Create Chamfer if Smart Item is Chamfered AC
    If sACItemName = "Chamfered_InsertPlate_AC" Then
        bIsNeeded = True
    End If
   
    Dim lDispId As Long
    Dim oAppConnection As IJAppConnection
    
    lDispId = 0
    lDispId = pMD.dispid
    
    On Error Resume Next
    Set oAppConnection = pMD.CAO
    On Error GoTo ErrorHandler
    
    ' When the penetrated or penetrating objects ports are modified, or any change on AC Property page attributes
    ' the correspondig chamfer needs to be updated...
    If (bIsNeeded) And (Not oAppConnection Is Nothing) Then
        
        Dim oMemObj As IJDMemberObjects
        Set oMemObj = oAppConnection
        
        On Error Resume Next
        Dim oItemMember As Object
        Set oItemMember = oMemObj.ItemByDispid(lDispId)
        Set oMemObj = Nothing
        On Error GoTo ErrorHandler
        
        If Not oItemMember Is Nothing Then
            If TypeOf oItemMember Is IJChamfer Then
                Dim oChamferSmartOcc As IJSmartOccurrence
                Set oChamferSmartOcc = oItemMember
                oChamferSmartOcc.Update
                Set oChamferSmartOcc = Nothing
            End If
            Set oItemMember = Nothing
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CM_IsLeftChamferNeeded").Number
End Sub

Public Sub CM_IsRightChamferNeeded(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    bIsNeeded = False
    
    Dim sACItemName As String
    
    ' Get AC Smart Item Name
    Parent_SmartItemName pMD.CAO, sACItemName
    
    'Create Chamfer if Smart Item is Chamfered AC
    If sACItemName = "Chamfered_InsertPlate_AC" Then
        bIsNeeded = True
    End If
   
    Dim lDispId As Long
    Dim oAppConnection As IJAppConnection
    
    lDispId = 0
    lDispId = pMD.dispid
    
    On Error Resume Next
    Set oAppConnection = pMD.CAO
    On Error GoTo ErrorHandler
    
    ' When the penetrated or penetrating objects ports are modified, or any change on AC Property page attributes
    ' the correspondig chamfer needs to be updated...
    If (bIsNeeded) And (Not oAppConnection Is Nothing) Then
        
        Dim oMemObj As IJDMemberObjects
        Set oMemObj = oAppConnection
        
        On Error Resume Next
        Dim oItemMember As Object
        Set oItemMember = oMemObj.ItemByDispid(lDispId)
        Set oMemObj = Nothing
        On Error GoTo ErrorHandler
        
        If Not oItemMember Is Nothing Then
            If TypeOf oItemMember Is IJChamfer Then
                Dim oChamferSmartOcc As IJSmartOccurrence
                Set oChamferSmartOcc = oItemMember
                oChamferSmartOcc.Update
                Set oChamferSmartOcc = Nothing
            End If
            Set oItemMember = Nothing
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CM_IsRightChamferNeeded").Number
End Sub

Public Sub CMConstructLeftChamfer(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler

    Dim strError As String
    strError = "Retrieving CAO and constructing slot."
   
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oPenetratingPort As IJPort
    Dim oPenetratedPort As IJPort
    GetPenetratedAndPenetratingPorts oAppConnection, oPenetratedPort, oPenetratingPort

    Dim pSystemParent As IJSystemChild
    Set pSystemParent = pMemberDescription.CAO

    Dim oChamfer As New StructDetailObjects.Chamfer

    Dim oSlotmappingRule As IJSlotMappingRule
    Set oSlotmappingRule = CreateSlotMappingRuleSymbolInstance()

    Dim oBasePort As IJPort
    Dim oMappedPorts As JCmnShp_CollectionAlias
    Set oMappedPorts = New Collection

    oSlotmappingRule.GetEmulatedPorts oPenetratingPort.Connectable, oPenetratedPort.Connectable, oBasePort, oMappedPorts
 
    Dim oPortBeforeCut As IJPort
    Dim oMemberPort As IJPort
    Dim oMemberObj As IJDMemberObjects
    Dim oSlotOcc As Object

    Set oMemberObj = pMemberDescription.CAO
    Set oSlotOcc = oMemberObj.ItemByDispid(1)

    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    Dim oFeatureUtils As GSCADSDCreateModifyUtilities.IJSDFeatureAttributes
    
    'Try to get Slot Cut-out Port(Late Port)whihc acta as an input for Chamfer creation
    Set oStructProfilePart = oPenetratedPort.Connectable
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    oStructEndCutUtil.GetLatePortForFeatureSegment oSlotOcc, 257, oMemberPort
    
    Set oPortBeforeCut = RelatedPortBeforeCut(oMappedPorts.Item(CStr(JXSEC_WEB_LEFT)))
    
    oChamfer.Create pResourceManager, oPortBeforeCut, oMemberPort, "RootChamfer", pSystemParent
    Set pObject = oChamfer.object
    
    Dim oChamferObj As IJChamfer
    Set oChamferObj = oChamfer.object
    oChamferObj.ExtensionType = ChamferEntireEdge
    Set oChamferObj = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructLeftChamfer", strError).Number
End Sub

Public Sub CMConstructRightChamfer(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    strError = "Retrieving CAO and constructing slot."
   
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = pMemberDescription.CAO
    
    Dim oPenetratingPort As IJPort
    Dim oPenetratedPort As IJPort
    GetPenetratedAndPenetratingPorts oAppConnection, oPenetratedPort, oPenetratingPort

    Dim pSystemParent As IJSystemChild
    Set pSystemParent = pMemberDescription.CAO

    Dim oChamfer As New StructDetailObjects.Chamfer

    Dim oSlotmappingRule As IJSlotMappingRule
    Set oSlotmappingRule = CreateSlotMappingRuleSymbolInstance()

    Dim oBasePort As IJPort
    Dim oMappedPorts As JCmnShp_CollectionAlias
    Set oMappedPorts = New Collection

    oSlotmappingRule.GetEmulatedPorts oPenetratingPort.Connectable, oPenetratedPort.Connectable, oBasePort, oMappedPorts
    
    Dim oPortBeforeCut As IJPort
    Dim oMemberPort As IJPort
    Dim oMemberObj As IJDMemberObjects
    Dim oSlotOcc As Object

    Set oMemberObj = pMemberDescription.CAO
    Set oSlotOcc = oMemberObj.ItemByDispid(1)

    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    Dim oFeatureUtils As GSCADSDCreateModifyUtilities.IJSDFeatureAttributes
    
    'Try to get Slot Cut-out Port(Late Port)whihc acta as an input for Chamfer creation
    Set oStructProfilePart = oPenetratedPort.Connectable
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    oStructEndCutUtil.GetLatePortForFeatureSegment oSlotOcc, 258, oMemberPort

    Set oPortBeforeCut = RelatedPortBeforeCut(oMappedPorts.Item(CStr(JXSEC_WEB_RIGHT)))

    oChamfer.Create pResourceManager, oPortBeforeCut, oMemberPort, "RootChamfer", pSystemParent
    Set pObject = oChamfer.object
    
    Dim oChamferObj As IJChamfer
    Set oChamferObj = oChamfer.object
    oChamferObj.ExtensionType = ChamferEntireEdge
    Set oChamferObj = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructRightChamfer", strError).Number
End Sub

Public Sub CMCopyQuestionsChamfer(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Putting questions Answer to Chamfer Rules"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    Dim oChamferedPort As IJStructPort
    
    ' Get the PlatePort on which chamfer has to be created
    GetChamferedPort pMemberDescription, oChamferedPort
        
    'Set Chamfered Type Answer based on chamfered port
    If oChamferedPort.ContextID And CTX_BASE Then
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "ChamferRules.RootChamferSel", "ChamferType", "Obj1Base"
    ElseIf oChamferedPort.ContextID And CTX_OFFSET Then
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "ChamferRules.RootChamferSel", "ChamferType", "Obj1Offset"
    End If

    Set pCopyAnswerHelper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsToMemebers", strError).Number
End Sub
'*********************************************************************************************
' Method      : CM_FinalConstructAggregator
' Description :
'
'*********************************************************************************************
Public Sub CM_FinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::CM_FinalConstructAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Final Construct of Assembly Connection"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : CM_MigrateAggregator
' Description :
'
'*********************************************************************************************
Public Sub CM_MigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = m_DefinitionProgid & "::CM_MigrateAggregator"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Mirgrating Assembly Connection Inputs"
    
    MigrateAssemblyConnection pAggregatorDescription, pMigrateHelper

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

' ** End CM **

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation

    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Dim pAPDs As IJDPropertyDescriptions
    Set pAD = pDefinition
    Set pAPDs = pDefinition
    pAPDs.RemoveAll ' Removes all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Removes all the previous Member descriptions
    ItemMembers pMDs
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_DefinitionProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean

End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Implements ISPSFACInputHelper
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'*********************************************************************************************
' Method      : ISPSFACInputHelper_UserAttributeMgmt
' Description :
'
'*********************************************************************************************
Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
    Const METHOD = m_DefinitionProgid & "::ISPSFACInputHelper_UserAttributeMgmt"
    
    Set ISPSFACInputHelper_UserAttributeMgmt = Me
    
End Property

'*********************************************************************************************
' Method      : ISPSFACInputHelper_ValidateObjects
' Description :
'
'*********************************************************************************************
Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal oInputObjs As SP3DStructInterfaces.IJElements, oRelationObjs As SP3DStructInterfaces.IJElements) _
                            As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const METHOD = m_DefinitionProgid & "::ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler
    
    ISPSFACInputHelper_ValidateObjects = InputHelper_ValidateObjectsForSplit(oInputObjs, oRelationObjs)
    
    Exit Property
ErrorHandler:
    HandleError MODULE, METHOD
    
End Property

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Implements IJUserAttributeMgmt
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'*********************************************************************************************
' Method      : IJUserAttributeMgmt_OnAttributeChange
' Description :
'
'*********************************************************************************************
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, _
                        ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnAttributeChange"

End Function

'*********************************************************************************************
' Method      : IJUserAttributeMgmt_OnPreCommit
' Description :
'
'*********************************************************************************************
Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreCommit"

End Function

'*********************************************************************************************
' Method      : IJUserAttributeMgmt_OnPreLoad
' Description :
'
'*********************************************************************************************
Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = m_DefinitionProgid & "::IJUserAttributeMgmt_OnPreLoad"

End Function

'*********************************************************************************************
' Method      : InputHelper_ValidateObjectsForSplit
' Description :
'
'*********************************************************************************************
Private Function InputHelper_ValidateObjectsForSplit(ByVal oInputObjs As SP3DStructInterfaces.IJElements, _
                            oRelationObjs As SP3DStructInterfaces.IJElements) _
                            As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const METHOD = "::InputHelper_ValidateObjectsForSplit"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim lCount As Long
    Dim oPortCol As IJElements
    
    Dim ePortIdx As SPSMemberAxisPortIndex
    
    Dim oInputObj1 As Object
    Dim oInputObj2 As Object
    Dim oFrmConn As ISPSFrameConnection
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppingPort As ISPSSplitAxisPort
    
    
    InputHelper_ValidateObjectsForSplit = SPSFACInputHelper_Ok
    Set oPortCol = New JObjectCollection
    'filter out ports to portCol
    For lCount = 1 To oInputObjs.Count
        If TypeOf oInputObjs.Item(lCount) Is IJPort Then
            oPortCol.Add oInputObjs.Item(lCount)
        End If
    Next lCount
    
    '  make sure there are only two ports
    If oPortCol.Count = 2 Then
        Set oInputObj1 = oPortCol.Item(1)
        Set oInputObj2 = oPortCol.Item(2)
        
        If Not TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            Dim oTemp As Object
            Set oTemp = oInputObj1
            oInputObj1 = oInputObj2
            oInputObj2 = oTemp
        End If
        If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            Set oSuppedPort = oInputObj1
            If TypeOf oInputObj2 Is ISPSSplitAxisAlongPort Or TypeOf oInputObj2 Is ISPSSplitAxisEndPort Then
                Set oSuppingPort = oInputObj2
                
                If oRelationObjs Is Nothing Then
                    Set oRelationObjs = New JObjectCollection
                End If
                
                oRelationObjs.Clear
                oRelationObjs.Add oSuppedPort
                oRelationObjs.Add oSuppingPort
            Else
                InputHelper_ValidateObjectsForSplit = SPSFACInputHelper_InvalidTypeOfObject
            End If
        Else
            InputHelper_ValidateObjectsForSplit = SPSFACInputHelper_InvalidTypeOfObject
        End If
    Else
        InputHelper_ValidateObjectsForSplit = SPSFACInputHelper_BadNumberOfObjects
    End If
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Function

