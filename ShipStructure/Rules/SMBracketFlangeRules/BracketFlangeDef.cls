VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketFlangeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Definition for the Bracket Flange.
'

Const m_ItemProgid   As String = "SMBcktFlangeRules.BracketFlangeDef"
Const m_ItemName     As String = "SMBcktFlangeRules.BracketFlangeDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SMBracketFlangeRules\BracketFlangeDef.cls"
Private ErrorMsg As String
Private METHOD   As String

Implements IJDUserSymbolServices
' *******************************************************************************************
Public Sub ItemInputs( _
    pIH As IJDInputsHelper _
)
    On Error GoTo ErrorHandler
    METHOD = "ItemInputs"
    
    'what are the inputs?
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, ErrorMsg).Number
End Sub
' *******************************************************************************************
Public Sub ItemAggregator( _
    pAD As IJDAggregatorDescription _
)
    On Error GoTo ErrorHandler
    METHOD = "ItemAggregator"
    
    ' TODO - Add aggregator description below
  
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, _
                            "CMFinalConstructFlange"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, ErrorMsg).Number
End Sub
' *******************************************************************************************
Public Sub ItemMembers( _
    pMDs As IJDMemberDescriptions _
)
    On Error GoTo ErrorHandler
    METHOD = "ItemMembers"
    
    ' TODO - Add member description below
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, ErrorMsg).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructFlange( _
    ByVal pAggregatorDescription As IJDAggregatorDescription _
)
  Dim clsBracketCM As New BracketFlangeDefCM
  clsBracketCM.BracketFlange_FinalConstruct pAggregatorDescription

End Sub
' ** End CM **
' *******************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' *******************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function
' *******************************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition( _
    pDefinition As IJDSymbolDefinition _
)
 
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
    pMDs.RemoveAll ' Remove all the previous member  descriptions
    ItemMembers pMDs
  
End Sub
' *******************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition( _
    ByVal CodeBase As String, _
    ByVal defParams As Variant, _
    ByVal ActiveConnection As Object _
) As Object

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
' *******************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation( _
    ByVal sblOcc As Object, _
    ByVal repName As String, _
    ByVal outputcoll As Object, _
    ByRef arrayOfInputs() _
)
End Sub
' *******************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence( _
    ByRef pSymbolOccurence As Object, _
    ByVal transactionMgr As Object _
) As Boolean
End Function
' *******************************************************************************************
'Public Sub ReportError( _
'    Optional ByVal sFunctionName As String, _
'    Optional ByVal sErrorName As String _
')
'    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
'                & " - " & "::" & sFunctionName & " - " & sErrorName
'End Sub
' *******************************************************************************************




