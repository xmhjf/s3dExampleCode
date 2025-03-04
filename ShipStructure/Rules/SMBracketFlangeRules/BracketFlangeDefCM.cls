VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketFlangeDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SMBracketFlangeRules\BracketFlangeDefCM"
Private ErrorMsg As String
Private METHOD   As String
' *******************************************************************************************
' METHOD:  BracketFlange_FinalConstruct
'
' DESCRIPTION: Final construction of the Bracket Flange.
' *******************************************************************************************
Public Sub BracketFlange_FinalConstruct( _
    ByVal pAggregatorDescription As IJDAggregatorDescription _
)
    On Error GoTo ErrorHandler
    METHOD = "BracketFlange_FinalConstruct"
    
    'Add any specific code here
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, ErrorMsg).Number

End Sub
' *******************************************************************************************
' METHOD:  CMCreateBracketFlange
'
' DESCRIPTION:  Conditional that determines if there is a bracket flange
'
' *******************************************************************************************
Public Sub CMCreateBracketFlange( _
    ByRef pMD As IJDMemberDescription, _
    ByRef bIsNeeded As Boolean _
)
   
    bIsNeeded = True
    
End Sub
' *******************************************************************************************


