VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffEndByBeamSel"
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

Const m_SelectorProgID As String = "AssyConnRules.StiffEndByBeamSel"
Const m_SelectorName As String = "AssyConnRules.StiffEndByBeamSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\StiffEndByBeamSel.cls"

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
    
    Dim strError As String
    
    'Define questions
    strError = "Setting questions."
    SetQuestionEndCutType pQH
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", strError).Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Dim iPort2TopoContext As IMSStructConnection.eUSER_CTX_FLAGS  'Bounding object port
    Dim iPort2CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim oBeam As StructDetailObjects.BeamPart
    Dim iBeamPortContext As enmPortBasicContext
    Dim iBeamPortType As enmProfilePortType

    strError = "Retrieving assembly connection and beam part from Smart Occurrence."
    Set oAssyConn = New StructDetailObjects.AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence
    Set oBeam = New StructDetailObjects.BeamPart
    Set oBeam.object = oAssyConn.ConnectedObject2 'Bounding part
    
    'Already know that stiffener is connected at an end, so to determine which definition rule to
    'fire we simply need to determine port context and cross section type of bounding beam.
    strError = "Retrieving assembly connection bounding port info."
    Call oAssyConn.Port2Topology(iPort2TopoContext, iPort2CrossSectEntity)  'Beam is port 2
    iBeamPortContext = PortBasicContext(iPort2TopoContext)
    iBeamPortType = ProfilePartPortType(oBeam.SectionType, iPort2CrossSectEntity)
  
    'get the assembly connection ports, used to determine name of the bounding port
    Dim oPortBounding As IJPort
    Set oPortBounding = pSL.InputObject(INPUT_BOUNDING_OR_PENETRATING_OBJECT)
    
    Dim oPortBounded As IJPort
    Set oPortBounded = pSL.InputObject(INPUT_BOUNDED_OR_PENETRATED_OBJECT)
    
    'Get the name of the port, used to determine side for the symbol
    Dim dPortValue As Long
    Dim oPartInfo As New StructDetailObjects.Helper
    dPortValue = oPartInfo.GetBoundingProfileFace(oPortBounded, oPortBounding)
    Select Case iBeamPortContext
      Case PORT_BASIC_CONTEXT_LATERAL
            Select Case oBeam.SectionType
                Case "P", "R", "HalfR", "RB", "RT", "SqTu", "SB"
                    pSL.Add "StiffenerEndToBeamFlange"
                Case "B", "BUT", "BUL2", "BUTL2", "BUTL3", "EA", "UA", "FB", "C", "C_SS", "CSType", "C_STYPE", "T_XType", "TSType", "I", "I_STYPE", "ISType"
                    If dPortValue = JXSEC_TOP Or dPortValue = JXSEC_BOTTOM Then
                        pSL.Add "StiffenerEndToBeamFlange"
                    Else
                        pSL.Add "StiffenerEndToBeamWeb"
                    End If
            End Select
       Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
            strError = "Invalid bounding object port type:" & vbNewLine & _
                            "     Cannot bound a stiffener end by a beam end."
            GoTo ErrorHandler  'This configuration is not supported by GSCAD
        Case Else
        strError = "Unknown value: iBeamPortContext=" & iBeamPortContext
        GoTo ErrorHandler
    End Select
    
    Set oAssyConn = Nothing
    Set oBeam = Nothing
    
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
    pIH.Definition = pSelector
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

