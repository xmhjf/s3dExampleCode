VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrFreeEndCutDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "FreeEndCutRules.MbrFreeEndCutDef"
Const m_ItemName As String = CUSTOMERID + "FreeEndCutRules.MbrFreeEndCutDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "FreeEndCutRules\MbrFreeEndCutDef.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput FE_INPUT_BOUNDED
    pIH.SetInput FE_INPUT_BOUNDING
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    Dim strError As String
    Dim pMemDesc As IJDMemberDescription
'Dim zz As SMFreeEndCutRules.FreeEndCutDefCM

       
    '--------------------------
    'Web Cut for Part 1
    strError = "Constructing Web Cut."
    Set pMemDesc = pMDs.AddMember("FreeEndWebCut", 1, _
                                  "CMConstruct_FreeEndWebCut", _
                                  CMLIBRARY_FREEENDCUTDEFCM)
                                  
    'Copy Q/A in the Free End Cut selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs CMLIBRARY_FREEENDCUTDEFCM, "CMCopyWCQuestionsToMembers"
    pMemDesc.SetCMConditional CMLIBRARY_FREEENDCUTDEFCM, _
                              "CMConditional_FreeEndWebCut"

    '--------------------------
    'Top flange cut for part
    strError = "Constructing top Flange Cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("FreeEndTopFlangeCut", 2, _
                                  "CMConstruct_FreeEndTopFlangeCut", _
                                  CMLIBRARY_FREEENDCUTDEFCM)

    'Copy Q/A in the Free End Cut selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs CMLIBRARY_FREEENDCUTDEFCM, "CMCopyFCQuestionsToMembers"
    pMemDesc.SetCMSetInputs CMLIBRARY_FREEENDCUTDEFCM, "CMSetTopFlange"
    pMemDesc.SetCMConditional CMLIBRARY_FREEENDCUTDEFCM, _
                              "CMConditional_FreeEndTopFlangeCut"
    
    '--------------------------
    'Bottom flange cut for part
    strError = "Constructing bottom Flange Cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("FreeEndBottomFlangeCut", 3, _
                                  "CMConstruct_FreeEndBottomFlangeCut", _
                                  CMLIBRARY_FREEENDCUTDEFCM)

    'Copy Q/A in the Free End Cut selector that must be sent to the endcut.
    pMemDesc.SetCMSetInputs CMLIBRARY_FREEENDCUTDEFCM, "CMCopyFCQuestionsToMembers"
    pMemDesc.SetCMSetInputs CMLIBRARY_FREEENDCUTDEFCM, "CMSetBottomFlange"
    pMemDesc.SetCMConditional CMLIBRARY_FREEENDCUTDEFCM, _
                              "CMConditional_FreeEndBottomFlangeCut"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructFEC(ByVal pAggregatorDescription As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    ' TODO - Add any Final Construction requirements here
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMFinalConstructFEC").Number
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)

  
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
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
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
    'pAD.AggregatorClsid = "{29D085BB-89FD-4964-AA9B-C34EFE68DB02}"
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
    
    
    ' We want free end cuts to be copied when the member end cut is copied.
    ' It cannot be done thro AssemblyMembersR relationship because it belongs to commonApp.
    Dim pCADefinition As IJCADefinition
    Set pCADefinition = pDefinition
    pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_PROPAGATE
    Set pCADefinition = Nothing
    
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous member descriptions
    ItemMembers pMDs
  
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_ItemProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    MigrateTheAggregator pAggregatorDescription, pMigrateHelper
    Exit Sub

ErrorHandler:
'    HandleError MODULE, METHOD
End Sub



