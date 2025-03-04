VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberByPlateSel"
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

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.MemberByPlateSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.MemberByPlateSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\MemberByPlateSel.cls"

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
    Dim iPort2TopoContext As IMSStructConnection.eUSER_CTX_FLAGS  'Bounding object port
    Dim iPort1CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim iPort2CrossSectEntity As IMSProfileEntity.JXSEC_CODE
    Dim iMemberPortType As enmProfilePortType
    Dim iPlatePortContext As enmPortBasicContext

    strError = "Retrieving assembly connection and stiffener part from Smart Occurrence."
    Set oAssyConn = New StructDetailObjects.AssemblyConn
    Set oAssyConn.object = pSL.SmartOccurrence

    'Determine which definition rule to fire based on port contexts of bounding
    'plate and bounded member as well as port type of bounded member.
    strError = "Retrieving assembly connection bounded and bounding port info."
    
    Select Case oAssyConn.ConnectedObject1Type
        Case SDOBJECT_PLATE
            If oAssyConn.ConnectedObject2Type = SDOBJECT_MEMBER Then
                Call oAssyConn.Port1Topology(iPort1TopoContext, iPort1CrossSectEntity)  'Plate is port 1
                iPlatePortContext = PortBasicContext(iPort1TopoContext)
            End If
        Case SDOBJECT_MEMBER
            If oAssyConn.ConnectedObject2Type = SDOBJECT_PLATE Then
                Call oAssyConn.Port2Topology(iPort2TopoContext, iPort2CrossSectEntity)  'Plate is port 2
                iPlatePortContext = PortBasicContext(iPort2TopoContext)
            End If
    End Select
    
    Select Case iPlatePortContext
        Case PORT_BASIC_CONTEXT_BASE, PORT_BASIC_CONTEXT_OFFSET
        
            iMemberPortType = GetMemberPartPortType(oAssyConn)
            
            Select Case iMemberPortType
                Case PROFILE_PORTTYPE_FACE
                     pSL.Add "MemberFaceToPlateFace"
                Case PROFILE_PORTTYPE_EDGE
                     pSL.Add "MemberEdgeToPlateFace"
                Case Else
                    strError = "Invalid bounding object type:" & vbNewLine & _
                                "     Cannot support lapped weld for a plate edge by a member face ."
                    GoTo ErrorHandler
            End Select
        Case PORT_BASIC_CONTEXT_LATERAL
            strError = "Invalid bounding object type:" & vbNewLine & _
                            "     Cannot support lapped weld for a manual connection with plate edge by a member face ."
'            pSL.Add "PlateEdgeToMemberFace"
            GoTo ErrorHandler
            
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

'********************************************************************************************************
'METHOD: GetMemberPartPortType
'
'DESCRIPTION: This helper method gets all the face ports of the member part overlapping with
'             plate port. Always considers the first port in the collection.
'             Returns the port type of this particular port in terms of enmProfilePortType.
'********************************************************************************************************
Private Function GetMemberPartPortType(ByVal oAssyConn As AssemblyConn) As enmProfilePortType

    On Error GoTo ErrorHandler
    Const METHOD_NAME = "GetMemberPartPortType"
    
    Dim eMemberPartPortType As enmProfilePortType
    eMemberPartPortType = PROFILE_PORTTYPE_FACE
    
    Dim oMemberPart As Object
    Dim oPlatePartPort As IJPort

    Select Case oAssyConn.ConnectedObject1Type
        Case SDOBJECT_PLATE
            Set oPlatePartPort = oAssyConn.Port1
            Set oMemberPart = oAssyConn.ConnectedObject2
        Case SDOBJECT_MEMBER
            Set oMemberPart = oAssyConn.ConnectedObject1
            Set oPlatePartPort = oAssyConn.Port2
    End Select
    
    Dim oSDMemberPart As StructDetailObjects.MemberPart
    Set oSDMemberPart = New StructDetailObjects.MemberPart
    
    Dim strMemberSectionType As String
    
    If TypeOf oMemberPart Is ISPSMemberPartPrismatic Then
        Set oSDMemberPart.object = oMemberPart
        
        strMemberSectionType = oSDMemberPart.SectionType
    End If
    
    Dim oFacePortColl As Collection
    Set oFacePortColl = GetFacePortsOfMemberOverlappingWithPlate( _
                            oPlatePartPort, _
                            oMemberPart)
    
    Dim ePortCrossSectionCode As IMSProfileEntity.JXSEC_CODE
    ePortCrossSectionCode = JXSEC_UNKNOWN
    
    If Not oFacePortColl Is Nothing Then
        If oFacePortColl.Count > 0 Then
           Call GetMemberPortCrossSectionCode(oFacePortColl.Item(1), _
                                ePortCrossSectionCode)

           Call MemberPartPortType(strMemberSectionType, _
                                 ePortCrossSectionCode, _
                                 eMemberPartPortType)
        End If
    End If

PROC_EXIT:
    Set oSDMemberPart = Nothing
    Set oFacePortColl = Nothing
    Set oMemberPart = Nothing
    Set oPlatePartPort = Nothing
    
    GetMemberPartPortType = eMemberPartPortType
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD_NAME).Number
    GoTo PROC_EXIT
End Function

