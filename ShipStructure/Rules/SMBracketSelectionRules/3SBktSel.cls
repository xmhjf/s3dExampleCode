VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt3SSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
Option Explicit

Const m_SelectorProgid As String = "BktSelRules.Bkt3SSel"
Const m_SelectorName As String = "BktSelRules.Bkt3SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BracketRules\3SBktSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' inputs:
    '   Plane that bracket is on
    '   Bracket Plate System
    '   Group of supports

    pIH.SetInput INPUT_BRACKETPLANE
    
    pIH.SetInput INPUT_SUPPORTS
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler

    pQH.SetQuestion QUESTION_BracketContourType, _
                    DEFAULT_BracketContourType_LINEAR, _
                    CODELIST_BracketContourType

  
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
     Dim sBracketName As String
    Dim sBracketContourType As String
    Dim sBracketSupport1Type As ShpStrBktSupportConnectionType
    Dim sBracketSupport2Type As ShpStrBktSupportConnectionType
    Dim sBracketSupport3Type As ShpStrBktSupportConnectionType
  
    '1. Get inputs for deciding bracket selection...
    sBracketSupport1Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 1)
    sBracketSupport2Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 2)
    sBracketSupport3Type = CM_SetBracketContourTypeSup(pSL.SmartOccurrence, 3)
  
    sBracketContourType = pSL.SelectorAnswer("BktSelRules.Bkt3SSel", QUESTION_BracketContourType)
    
    Select Case sBracketSupport3Type
    
    Case Trimmed, Lapped
    
         If sBracketSupport1Type = Trimmed And sBracketSupport2Type = Trimmed And sBracketSupport3Type = Trimmed Then
         
             Select Case sBracketContourType
                Case "Linear"
                    pSL.Add "BPS_3SL"
                    pSL.Add "BPS_3SL_OF"
                    pSL.Add "BPS_3SL_IF"
                    pSL.Add "BPS_3SL_FT"
                    pSL.Add "BPS_3SL_OT"
                    pSL.Add "BPS_3SLC_FB"
                Case "Curved"
                   
                    pSL.Add "BPS_3SC"
                    pSL.Add "BPS_3SC_OT"
                    pSL.Add "BPS_3SC_Sym"
                End Select
         ElseIf sBracketSupport1Type = Trimmed And sBracketSupport2Type = Trimmed And sBracketSupport3Type = Lapped Then
            Select Case sBracketContourType
                Case "Linear"
                    pSL.Add "BPS_3SLLS3"
                   
                Case "Curved"
                    
                    pSL.Add "BPS_3SCLS3"
                   
                    pSL.Add "BPS_3SCLS3_Sym"
                End Select
         ElseIf sBracketSupport1Type = Trimmed And sBracketSupport2Type = Lapped And sBracketSupport3Type = Lapped Then
                Select Case sBracketContourType
                  Case "Linear"
                      ' Additional Method could be added to check the 3rd support type
                      ' and conditions to filter selection set...
                      pSL.Add "BPS_3SLLS2S3"
                     
                  Case "Curved"
                      pSL.Add "BPS_3SCLS2S3"
                     pSL.Add "BPS_3SCLS2S3_Sym"
                       
                End Select
         ElseIf sBracketSupport1Type = Trimmed And sBracketSupport2Type = Lapped And sBracketSupport3Type = Trimmed Then
                Select Case sBracketContourType
                  Case "Linear"
                      ' Additional Method could be added to check the 3rd support type
                      ' and conditions to filter selection set...
                     
                      pSL.Add "BPS_3SLLS2"
                      pSL.Add "BPS_3SLLS2_OF"
                      pSL.Add "BPS_3SLLS2_IF"
                      pSL.Add "BPS_3SLLS2_FT"
                      pSL.Add "BPS_3SLLS2_OT"
                      pSL.Add "BPS_3SLCLS2_FB"
                      
                      
                  Case "Curved"
                    
                      pSL.Add "BPS_3SCLS2"
                      pSL.Add "BPS_3SCLS2_OT"
                      pSL.Add "BPS_3SCLS2_Sym"
                       
                End Select
         ElseIf sBracketSupport1Type = Lapped And sBracketSupport2Type = Trimmed And sBracketSupport3Type = Trimmed Then
                 Select Case sBracketContourType
                  Case "Linear"
                      ' Additional Method could be added to check the 3rd support type
                      ' and conditions to filter selection set...
                     
                      pSL.Add "BPS_3SLLS1"
                      pSL.Add "BPS_3SLLS1_OF"
                      pSL.Add "BPS_3SLLS1_IF"
                      pSL.Add "BPS_3SLLS1_FT"
                      pSL.Add "BPS_3SLLS1_OT"
                      pSL.Add "BPS_3SLCLS1_FB"
                      
                     
                  Case "Curved"
                      
                      pSL.Add "BPS_3SCLS1"
                      pSL.Add "BPS_3SCLS1_OT"
                      pSL.Add "BPS_3SCLS1_Sym"
                       
                End Select
     
          
          ElseIf sBracketSupport1Type = Lapped And sBracketSupport2Type = Lapped And sBracketSupport3Type = Lapped Then
                 Select Case sBracketContourType
                    Case "Linear"
                        ' Additional Method could be added to check the 3rd support type
                        ' and conditions to filter selection set...
                        pSL.Add "BPS_3SLLS1S2S3"
                      
                    Case "Curved"
                        pSL.Add "BPS_3SCLS1S2S3"
                        pSL.Add "BPS_3SCLS1S2S3_Sym"
                   End Select
            
      
        End If
    Case Reference
        If sBracketSupport1Type = Trimmed And sBracketSupport2Type = Trimmed Then
             Select Case sBracketContourType
                Case "Linear"
                      pSL.Add "BPS_3SL_SB"
                        pSL.Add "BPS_3SL_SB_Far"
                Case "Curved"
                     pSL.Add "BPS_3SC_SB"
                    pSL.Add "BPS_3SLC_SB_Far"
                End Select
         
         ElseIf sBracketSupport1Type = Trimmed And sBracketSupport2Type = Lapped Then
               Select Case sBracketContourType
                Case "Linear"
                      pSL.Add "BPS_3SLLS2_SB"
                      pSL.Add "BPS_3SLLS2_SB_Far"
                Case "Curved"
                     pSL.Add "BPS_3SCLS2_SB"
                     pSL.Add "BPS_3SLCLS2_SB_Far"
                End Select
         ElseIf sBracketSupport1Type = Lapped And sBracketSupport2Type = Trimmed Then
                Select Case sBracketContourType
                Case "Linear"
                      pSL.Add "BPS_3SLLS1_SB"
                      pSL.Add "BPS_3SLLS1_SB_Far"
                Case "Curved"
                     pSL.Add "BPS_3SCLS1_SB"
                     pSL.Add "BPS_3SLCLS1_SB_Far"
                End Select
         
         
        End If
        
    End Select
    
    
   
    
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub


' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************

' ** End CM **



' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************



