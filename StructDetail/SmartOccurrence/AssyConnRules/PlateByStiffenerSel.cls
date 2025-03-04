VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateByStiffenerSel"
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

Const m_SelectorProgID As String = "AssyConnRules.PlateByStiffenerSel"
Const m_SelectorName As String = "AssyConnRules.PlateByStiffenerSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\PlateByStiffenerSel.cls"

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
    Dim iPort1TopoContext As IMSStructConnection.eUSER_CTX_FLAGS  'Bounded object port
    Dim iPort1CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim iPort2TopoContext As IMSStructConnection.eUSER_CTX_FLAGS  'Bounding object port
    Dim iPort2CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim oStiffener As StructDetailObjects.ProfilePart
    Dim iStiffenerPortType As enmProfilePortType
    Dim iStiffenerPortContext As enmPortBasicContext
    Dim iPlatePortContext As enmPortBasicContext

    strError = "Retrieving assembly connection and stiffener part from Smart Occurrence."
    Set oAssyConn = New StructDetailObjects.AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence
    Set oStiffener = New StructDetailObjects.ProfilePart
    Set oStiffener.object = oAssyConn.ConnectedObject2 'Bounding part

    'Determine which definition rule to fire based on port contexts of bounding
    'stiffener and bounded plate as well as port type of bounding stiffener.
    strError = "Retrieving assembly connection bounded and bounding port info."
    Call oAssyConn.Port1Topology(iPort1TopoContext, iPort1CrossSectEntity)  'Plate is port 1
    iPlatePortContext = PortBasicContext(iPort1TopoContext)
    Call oAssyConn.Port2Topology(iPort2TopoContext, iPort2CrossSectEntity)  'Stiffener is port 2
    iStiffenerPortContext = PortBasicContext(iPort2TopoContext)
    iStiffenerPortType = ProfilePartPortType(oStiffener.SectionType, iPort2CrossSectEntity)
    
    Select Case iPlatePortContext
        Case PORT_BASIC_CONTEXT_LATERAL
            Select Case iStiffenerPortContext
                Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
                    pSL.Add "PlateEdgeToStiffenerEdge"
                Case PORT_BASIC_CONTEXT_LATERAL
                    Select Case iStiffenerPortType
                        Case PROFILE_PORTTYPE_FACE
                            pSL.Add "PlateEdgeToStiffenerFace"
                        Case PROFILE_PORTTYPE_EDGE
                            pSL.Add "PlateEdgeToStiffenerEdge"
                    End Select
            End Select
        Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
            ' This next entry is for a plate face lapped onto a profile face
            ' The present definition for PlateFace to PlateFace is sufficient for
            ' this because all we need is a lap connection between the faces.
            pSL.Add "PlateFaceToPlateFace"
    End Select
    
    Set oAssyConn = Nothing
    Set oStiffener = Nothing

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

