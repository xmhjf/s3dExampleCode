VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FaceRatholeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2014-16, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : FaceRatholeParm.cls
'
'  Description :
'       It computes the item paramaters in the context of the smart occurrence.
'
'       - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'       - Outputs defined by name the collection of parameters
'
'  Author      : Alligators
'
'  History     :
'    30/MAY/2014 - Creation
'
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature
'   07 Jan 2016 -GHM      - TR-279987- Updated GetCornerFatureEdgeLengths() method with proper corner feature input object
'*********************************************************************************************

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.FaceRatholeParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.FaceRatholeParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\FaceRatholeParm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
    
    pOH.SetOutput "Ulength"
    pOH.SetOutput "Vlength"
    pOH.SetOutput "Radius"
    pOH.SetOutput "OffsetU"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number

End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
 
    Dim dLength1 As Double
    Dim dLength2 As Double
    
    Dim dShorter As Double
        
    'Get Edge Lengths
    GetCornerFatureEdgeLengths pPRL.SmartOccurrence, Nothing, pPRL.InputObject(INPUT_PORT2EDGE), pPRL.InputObject(INPUT_PORT3EDGE), _
                                dLength1, dLength2, pPRL.InputObject(INPUT_PORT1FACE)
    
    If dLength1 < dLength2 Then
        dShorter = dLength1
    Else
        dShorter = dLength2
    End If
    
    'Set Parameter Values
             pPRL.Add "Ulength", dShorter
             pPRL.Add "Vlength", dShorter + 0.01
             pPRL.Add "Radius", dShorter / 2
             pPRL.Add "OffsetU", dShorter
    
'    Select Case dShorter
    
'        Case Is <= 0.02
'            pPRL.Add "Ulength", dShorter / 2
'            pPRL.Add "Vlength", dShorter / 2
'            pPRL.Add "Radius", dShorter / 8
'        Case 0.02 To 0.2
'            pPRL.Add "Ulength", 0.015
'            pPRL.Add "Vlength", 0.015
'            pPRL.Add "Radius", 0.005
'
'        Case 0.2 To 0.28
'            pPRL.Add "Ulength", 0.025
'            pPRL.Add "Vlength", 0.025
'            pPRL.Add "Radius", 0.008
'        Case 0.28 To 0.4
'            pPRL.Add "Ulength", 0.035
'            pPRL.Add "Vlength", 0.035
'            pPRL.Add "Radius", 0.01
'        Case Is > 0.4
'            pPRL.Add "Ulength", 0.05
'            pPRL.Add "Vlength", 0.05
'            pPRL.Add "Radius", 0.015
   ' End Select
    
    Exit Sub
    
    
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
 
End Sub
  
' ** Start CM **
'*********************************************************************************************
' If needed Add Custom Method HERE
'*********************************************************************************************
' ** End CM **

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
  
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
    On Error Resume Next
      
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pPR.IJDInputs.RemoveAllInput
    pPR.IJDRepresentations.RemoveAllRepresentation
      
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractParameterRule pPR
    
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pPR
    pIH.InitAs m_FamilyProgid
    ParameterRuleInputs pIH
    
    Dim pOH As IJDOutputsHelper
    Set pOH = New OutputHelper
    pOH.Representation = pPR.IJDRepresentations.Item(1)
    pOH.InitAs m_FamilyProgid
    ParameterRuleOutputs pOH
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
    Dim pPRL As IJDParameterLogic
    Set pPRL = New ParameterLogic
    pPRL.Representation = pRep
    ParameterRuleLogic pPRL
    
End Sub

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

