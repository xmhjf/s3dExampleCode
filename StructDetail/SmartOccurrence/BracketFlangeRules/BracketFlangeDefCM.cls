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

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketFlangeRules\BracketDefCM"
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
    On Error GoTo ErrorHandler
    METHOD = "CMCreateBracketFlange"
    
    bIsNeeded = True
    
    Dim oBracketByPlane As Object
    Set oBracketByPlane = pMD.CAO

    'define interface from which to get the question
    Const strIASelectionRuleRoot As String = "IABracketRules_RootBracketByPlaneSel"
    
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String

    'use helper to get the answer from the question
    Call pHelper.GetCustomAttributeParamValue( _
                                        oBracketByPlane, _
                                        strIASelectionRuleRoot, _
                                        "PlaceBracketFlange", _
                                        retValue, _
                                        retValueType, _
                                        sTableName, _
                                        lCodeList, _
                                        sShortDesc, _
                                        sLongDesc)
    
    Dim sPlaceBracketFlange As String
    
    'check if answer is from code list; if so, get short desc, if not, use retValue
    If Len(Trim(sTableName)) > 0 Then
        sPlaceBracketFlange = sShortDesc
    Else
        sPlaceBracketFlange = retValue
    End If
        
    If sPlaceBracketFlange = "Yes" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub
' *******************************************************************************************


