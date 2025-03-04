VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoundedConnSel"
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

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.BoundedConnSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.BoundedConnSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\BoundedConnSel.cls"

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
    Dim iPort1TopoContext As IMSStructConnection.eUSER_CTX_FLAGS
    Dim iPort1CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim iStiffenerPortContext As enmPortBasicContext
    Dim iPlatePortContext As enmPortBasicContext
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    Dim oAppConn As IJAppConnection


    strError = "Retrieving assembly connection from Smart Occurrence."
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence
    Set oAppConn = oAssyConn.object

    'Determine next selector based on types of connected objects and if bounded object is
    'a stiffener, its basic port context (base, offset, or lateral).
    If TypeOf oAppConn Is IJAppConnection Then
        Set oConnAttrbs = New SDConnectionUtils
        
        If GreaterThanOrEqualTo(oConnAttrbs.get_AuxiliaryPorts(oAppConn).Count, 1) Then
            pSL.Add "GenericAssyConn"
            Set oConnAttrbs = Nothing
            Exit Sub
        Else
            strError = "Querying assembly connection for bounded object type."
            Select Case oAssyConn.ConnectedObject1Type  'Bounded object
                Case SDOBJECT_PLATE
                    strError = "Querying assembly connection for bounding object type and choosing next Smart Class."
                    Select Case oAssyConn.ConnectedObject2Type  'Bounding object
                        Case SDOBJECT_PLATE
                            pSL.Add "PlateByPlate"
                        Case SDOBJECT_STIFFENER
                            pSL.Add "PlateByStiffener"
                        Case SDOBJECT_BEAM
                            pSL.Add "PlateByBeam"
                        Case SDOBJECT_MEMBER
                            Call oAssyConn.Port1Topology(iPort1TopoContext, iPort1CrossSectEntity)
                            iPlatePortContext = PortBasicContext(iPort1TopoContext)
                            
                            Select Case iPlatePortContext
                                Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
                                    pSL.Add "MemberByPlate"
                                Case PORT_BASIC_CONTEXT_LATERAL
                                    pSL.Add "PlateByMember"
                            End Select
                    End Select
                    
                Case SDOBJECT_STIFFENER
                    strError = "Retrieving assembly connection bounded port info."
                    Call oAssyConn.Port1Topology(iPort1TopoContext, iPort1CrossSectEntity)
                    iStiffenerPortContext = PortBasicContext(iPort1TopoContext)
                    
                    Select Case iStiffenerPortContext
                        Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
                            Select Case oAssyConn.ConnectedObject2Type  'Bounding object
                                Case SDOBJECT_PLATE
                                    pSL.Add "StiffenerEndByPlate"
                                Case SDOBJECT_STIFFENER
                                    pSL.Add "StiffenerEndByStiffener"
                                Case SDOBJECT_BEAM
                                    pSL.Add "StiffenerEndByBeam"
                                                            
                                'HACK: Using the existing 'StiffenerEndByStiffener' selector rule for generating Assembly Connections.
                                'This would be replaced with the proper content.
                                'Reference : DI-203173    Generate Assembly Connections for Profile Bound/Split By Member functionality
                                Case SDOBJECT_MEMBER
                                    pSL.Add "StiffenerEndByMember"
                            End Select
                        Case PORT_BASIC_CONTEXT_LATERAL
                            Select Case oAssyConn.ConnectedObject2Type  'Bounding object
                                Case SDOBJECT_PLATE
                                    pSL.Add "StiffenerByPlate"
                                Case SDOBJECT_STIFFENER, SDOBJECT_BEAM 'Not supported in GSCAD
                                    strError = "Invalid bounding object type:" & vbNewLine & _
                                                    "     Cannot bound a stiffener along its length to a beam or another stiffener."
                                    GoTo ErrorHandler
                            End Select
                    End Select
                    
                Case SDOBJECT_BEAM
                    Select Case oAssyConn.ConnectedObject2Type  'Bounding object
                        Case SDOBJECT_PLATE
                            pSL.Add "BeamByPlate"
                        Case SDOBJECT_STIFFENER
                            pSL.Add "BeamByStiffener"
                        Case SDOBJECT_BEAM
                            pSL.Add "BeamByBeam"
                    End Select
                
                Case SDOBJECT_MEMBER
                    Select Case oAssyConn.ConnectedObject2Type  'Bounding object
                        Case SDOBJECT_PLATE
                            pSL.Add "MemberByPlate"
                    End Select
                    
                Case Else
                    strError = "Invalid assembly connection bounded object type found."
                    GoTo ErrorHandler
            End Select
        End If
    Else
        strError = "Not a valid Assenbly connection"
        GoTo ErrorHandler
    End If
        
    Set oAssyConn = Nothing
    Set oAppConn = Nothing
    
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

