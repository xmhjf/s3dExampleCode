VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CutFlange1_Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "EdgeFeatureRules.CutFlange1_Def"
Const m_ItemName As String = CUSTOMERID + "EdgeFeatureRules.CutFlange1_Def"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\" + CUSTOMERID + "EdgeFeatureRules\CutFlange1_Def.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_EDGE
    pIH.SetInput INPUT_POINT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler

    pAD.SetCMFinalConstruct CUSTOMERID + "EdgeFeatureRules.CutFlange1_Def", "CMFinalConstructEF"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim oMD As IJDMemberDescription
    
    ' Items Members are:
    '   InsertPlate (the InsertPlate will create the Physical Connections)
    Set oMD = pMDs.AddMember("InsertPlate", 1, "CMConstructPlate", imsCOOKIE_ID_USS_LIB)
    oMD.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMInsertPlate"
    Set oMD = Nothing
     
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

Public Sub CMInsertPlate(ByRef pMD As IJDMemberDescription, _
                         ByRef bIsNeeded As Boolean)
 On Error GoTo ErrorHandler
 
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_InsertPlate) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    
    Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMInsertPlate").Number
End Sub

Public Sub CMConstructPlate(ByVal pMemberDescription As IJDMemberDescription, _
                            ByVal pResourceManager As IUnknown, _
                            ByRef pObject As Object)
    On Error GoTo ErrorHandler
    CreateInsertPlate pMemberDescription, pResourceManager, _
                      False, "InsertPlate_CutEdge", pObject
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPlate").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructEF(ByVal pAggregatorDescription As IJDAggregatorDescription)
    On Error GoTo ErrorHandler

    Dim oEdgeFeatureWrapper As New StructDetailObjects.EdgeFeature
  
    Set oEdgeFeatureWrapper.object = pAggregatorDescription.CAO
    oEdgeFeatureWrapper.AddFeatureToCut
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMFinalConstructEF").Number
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
    On Error GoTo ErrorHandler

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
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous member descriptions
    ItemMembers pMDs
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJDUserSymbolServices_InitializeSymbolDefinition").Number
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

