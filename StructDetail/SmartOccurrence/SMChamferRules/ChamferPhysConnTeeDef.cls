VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChamferPhysConnTeeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "ChamferRules.ChamferPhysConnTeeDef"
Const m_ItemName As String = CUSTOMERID + "ChamferRules.ChamferPhysConnTeeDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "ChamferRules\ChamferPhysConnTeeDef.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add input description below
    
    ' Part to be chamfered
    ' Opposite Part - determines if a chamfer is needed and if so what size
    
    pIH.SetInput CHAMFER_PART
    pIH.SetInput OPPOSITE_PART
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    ' TODO - Add aggregator description below
    
'*********** CMFinalConstructChamfer needs to be built.
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTeeChamfer"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    ' TODO - Add member description below
    Dim pMemDesc As IJDMemberDescription

    Set pMemDesc = pMDs.AddMember("ChamferPhysConn1", 1, "CMConstructTeePC", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsTeeChamfer"

    Set pMemDesc = Nothing
   
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

Public Sub CMCopyQuestionsTeeChamfer(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Calling Copy answer to copy assembly questions"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.CopyAnswer "Connections,Ingr.SP3D.Content.Structure.ChamferTeeWeldSelectorRule", "ChamferThickness", _
                                 CUSTOMERID + "ChamferRules.RootTeeChamferSel", "ChamferThickness"
                                
    Set pCopyAnswerHelper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsTeeChamfer", strError).Number
End Sub

Public Sub CMFinalConstructTeeChamfer(ByVal pAggregatorDescription As IJDAggregatorDescription)
  Dim strError As String
  On Error GoTo ErrorHandler
  
  Dim clsChamferCM As New ChamferDefCM
  clsChamferCM.Chamfer_FinalConstruct pAggregatorDescription
  Set clsChamferCM = Nothing
    
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMFinalConstructTeeChamfer", strError).Number
End Sub
Public Sub CMConstructTeePC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
  Dim strError As String
  On Error GoTo ErrorHandler
  
  Dim clsChamferCM As New ChamferDefCM
  Dim oPhysConn As Object
  Set oPhysConn = clsChamferCM.CAConstruct_PhysConn(pMemberDescription, _
                                                    pResourceManager, _
                                                    "ChamferTeeWeld")
  Set pObject = oPhysConn
  Set clsChamferCM = Nothing
  Set oPhysConn = Nothing
  
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructTeePC", strError).Number
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
    pIH.Definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
     
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

Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

'********************************************************************************************
' Method      : CM_IsPCNeeded
' Description :
'********************************************************************************************
Public Sub CM_IsPCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    
  Const METHOD = "::CM_IsPCNeeded"
  On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
