' DI-CP-235025  StrDet: Error in Partclass View creation during bulkloading on Oracle DB 
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffEndByMemberSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMAssyConRul
'  File        : StiffEndByMemberSel.cls
'
'  Description :
'               Selector:
'               ---------
'               It selects the list of available Items or SmartClasses in the context of
'               the smart occurrence.
'
'               - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'               - Questions are defined to parameterize the selection
'
'
'  Author      : Alligators
'
'  History     :
'    08/NOV/2011 - Created
'
'**********************************************************************************************************************

Option Explicit

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.StiffEndByMemberSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.StiffEndByMemberSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffEndByMemberSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs", strError).Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler

    Dim strError As String
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim iPort2TopoContext As IMSStructConnection.eUSER_CTX_FLAGS  'Bounding object port
    Dim iPort2CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim iBoundingStiffenerPortContext As enmPortBasicContext

    strError = "Retrieving assembly connection from Smart Occurrence."
    Set oAssyConn = New StructDetailObjects.AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence

    'Already know that bounded stiffener is connected at an end, so to determine which definition
    'rule to fire we simply need to determine port type of bounding stiffener based on
    'cross section type and topological entity.
    strError = "Retrieving assembly connection bounding port info."
    Call oAssyConn.Port2Topology(iPort2TopoContext, iPort2CrossSectEntity)  'Bounding stiffener is port 2
    iBoundingStiffenerPortContext = PortBasicContext(iPort2TopoContext)
    

    Select Case iBoundingStiffenerPortContext
    
        Case PORT_BASIC_CONTEXT_LATERAL
        
            pSL.Add "StiffEndByMbrFace"
            
        Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
        
            pSL.Add "StiffenerEndByMemberEnd"
        
    End Select
    
    
    Set oAssyConn = Nothing

    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
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
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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



