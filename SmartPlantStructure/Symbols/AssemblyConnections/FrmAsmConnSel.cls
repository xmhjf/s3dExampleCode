VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrmAsmConnSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2003-2015 Intergraph Corporation. All rights reserved.
'
'File : FrmAsmConnSel.cls
'
'Author : R. Peringoth
'
'Description :
'    Macro for selecting a child connection type based on  number and type of inputs
'
'History:
'
'   03-Nov-14  MDT/GH   CR-CP-250198  Updated Selection to select MbrStdAC when both Inputs are Axis along
'   07-Aug-15  svsmylav CR-272424: Added code to the selector logic to handle Generic assembly connection.
'   09-Sep-15  knukala CR-226692: Rule-Based selection of standard member ACs .
'   30-Oct-15  svsmylav TR-280722: code added to identify member split by plate case,
'                       if so, added 'MemberByMember' child selector.
'   11-Dec-15  mchandak TR-283209: Executing the selection rule on Dummy Smart Occurrence via new method AnyAvailableSelectionsOnDummySO().
'   8-Jan-15   hgajula  DM-CP-287215    Should be able to revert back to Plant Rules for Assemebly connections in MHE
'********************************************************************

Option Explicit

Const m_SelectorProgid As String = "SPSACMacros.FrmAsmConnSel"
Const m_SelectorName As String = "SPSACMacros.FrmAsmConnSel"
Const m_FamilyProgid As String = "SPSACMacros.Family"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    
    Exit Sub
ErrorHandler:
    pIH.ReportError
End Sub


Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler

    
    Exit Sub
ErrorHandler:
    pQH.ReportError
End Sub


Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim count As Integer
    Dim oPlane As IJPlane
    Dim oMemberPort As ISPSSplitAxisPort
    Dim oSupportedPort As ISPSSplitAxisPort
    Dim oBoundedPart As ISPSMemberPartPrismatic
    Set oStructConn = pSL.SmartOccurrence
    oStructConn.enumPorts colPorts
    count = colPorts.count
    Dim oClassMoniker As IUnknown
    Dim oSmartClass As IJSmartClass
    
   'Get model type
    Dim eModelType As ProductType
    Dim oVersionHelper As VersionHelper
    Set oVersionHelper = New VersionHelper
    oVersionHelper.GetSiteTypeFromPOMIfExists eModelType
    
    'Check if it is a generic assembly connection
    Dim bGenericAC As Boolean
    Dim bInsertPlateAC As Boolean
    
    bGenericAC = False
    bInsertPlateAC = False
    If count > 1 Then
        Dim oPort1 As IJPort
        Dim oPort2 As IJPort

        Set oPort1 = colPorts.Item(1)
        Set oPort2 = colPorts.Item(2)
        
        If oPort1.Connectable Is oPort2.Connectable Then
            bGenericAC = True
        Else
            bInsertPlateAC = IsMbrSplitByPlateCase(oPort1, oPort2)
        End If
        If bGenericAC Then
                If PartClassExistsInDB("MbrStdAC", oClassMoniker) Then
                    If TypeOf oClassMoniker Is IJSmartClass Then
                        Set oSmartClass = oClassMoniker
                        If AnyAvailableSelectionsOnDummySO(pSL.SmartOccurrence, oSmartClass) Then
                            Set oSupportedPort = GetEndPort(colPorts)
                            If Not oSupportedPort Is Nothing Then
                                Set oBoundedPart = oSupportedPort.Part
                                If bIsTrafficItem(oBoundedPart) Then
                                    pSL.Add "MbrStdAC"
                                    pSL.Add "MemberByMember"
                                Else
                                    pSL.Add "MemberByMember"
                                    pSL.Add "MbrStdAC"
                                End If
                            End If
                        Else
                            pSL.Add "MemberByMember"
                        End If
                    Else
                        If eModelType = ProductType_PLANT Then
                            'if model type is plant, then need to check if member endcuts data
                            'is bulkloaded to the catalog database
                            If PartClassExistsInDB("MemberByMember", Nothing) And _
                                    PartClassExistsInDB("GenAssyConnSel", Nothing) Then
                                pSL.Add "MemberByMember"
                            Else
                                'Cannot proceed further when generic AC command invoked, so raise error
                                GoTo ErrorHandler
                            End If
                        ElseIf eModelType = ProductType_MARINE Or eModelType = ProductType_MHE Then
                            'Marine or MHE models to use member endcuts data existing in the catalog database
                            pSL.Add "MemberByMember"
                        End If
                    End If
                Else
                    If eModelType = ProductType_PLANT Then
                        'if model type is plant, then need to check if member endcuts data
                        'is bulkloaded to the catalog database
                        If PartClassExistsInDB("MemberByMember", Nothing) And _
                                PartClassExistsInDB("GenAssyConnSel", Nothing) Then
                            pSL.Add "MemberByMember"
                        Else
                            'Cannot proceed further when generic AC command invoked, so raise error
                            GoTo ErrorHandler
                        End If
                    ElseIf eModelType = ProductType_MARINE Or eModelType = ProductType_MHE Then
                        'Marine or MHE models to use member endcuts data existing in the catalog database
                        pSL.Add "MemberByMember"
                    End If
                End If

        ElseIf bInsertPlateAC Then
            pSL.Add "MemberByMember"
        End If
    End If
    
    If Not (bGenericAC Or bInsertPlateAC) Then
        'For Axis-assembly connection below code will be executed
        If count = 1 Then
            pSL.Add "MemberEnd"
        ElseIf count = 2 Then
            Set oSupportedPort = GetEndPort(colPorts)
            ' to support SplitNone Assembly connection
            If oSupportedPort Is Nothing Then
                If PartClassExistsInDB("MbrStdAC", Nothing) And (TypeOf colPorts.Item(1) Is ISPSSplitAxisAlongPort And TypeOf colPorts.Item(2) Is ISPSSplitAxisAlongPort) Then
                    pSL.Add "MbrStdAC"
                End If
            Else
                Set oBoundedPart = oSupportedPort.Part
                If oSupportedPort Is colPorts.Item(1) Then
                    If TypeOf colPorts.Item(2) Is IJPlane Then
                        Set oPlane = colPorts.Item(2) ' Succeeds if Ist input is a plane
                    ElseIf TypeOf colPorts.Item(2) Is ISPSSplitAxisPort Then
                        Set oMemberPort = colPorts.Item(2) ' Succeeds if Ist input is a member part
                    End If
                Else
                    If TypeOf colPorts.Item(1) Is IJPlane Then
                        Set oPlane = colPorts.Item(1) ' Succeeds if 2nd input is a plane
                    ElseIf TypeOf colPorts.Item(1) Is ISPSSplitAxisPort Then
                        Set oMemberPort = colPorts.Item(1) ' Succeeds if 2nd input is a member part
                    End If
                End If
                
                If Not oPlane Is Nothing Then
                    If PartClassExistsInDB("MbrStdAC", oClassMoniker) Then
                        If TypeOf oClassMoniker Is IJSmartClass Then
                            Set oSmartClass = oClassMoniker
                            If AnyAvailableSelectionsOnDummySO(pSL.SmartOccurrence, oSmartClass) Then
                                If bIsTrafficItem(oBoundedPart) Then
                                    pSL.Add "MbrStdAC"
                                    pSL.Add "MemberEnd"
                                Else
                                    pSL.Add "MemberEnd"
                                    pSL.Add "MbrStdAC"
                                End If
                            Else
                                pSL.Add "MemberEnd"
                            End If
                        Else
                            pSL.Add "MemberEnd"
                        End If
                    Else
                       pSL.Add "MemberEnd"
                    End If
                    
                ElseIf Not oMemberPort Is Nothing Then
                    If PartClassExistsInDB("MbrStdAC", oClassMoniker) Then
                        If TypeOf oClassMoniker Is IJSmartClass Then
                            Set oSmartClass = oClassMoniker
                            If AnyAvailableSelectionsOnDummySO(pSL.SmartOccurrence, oSmartClass) Then
                                If eModelType = ProductType_PLANT Or eModelType = ProductType_MHE Then
                                    If PartClassExistsInDB("SPS_MemberByMember", Nothing) Then
                                        pSL.Add "SPS_MemberByMember"
                                    End If
                                    pSL.Add "MemberByMember"
                                    pSL.Add "MbrStdAC"
                                Else
                                    If bIsTrafficItem(oBoundedPart) Then
                                        pSL.Add "MbrStdAC"
                                        pSL.Add "MemberByMember"
                                        pSL.Add "SPS_MemberByMember"
                                    Else
                                        pSL.Add "MemberByMember"
                                        pSL.Add "SPS_MemberByMember"
                                        pSL.Add "MbrStdAC"
                                    End If
                                End If
                            Else
                                 If eModelType = ProductType_PLANT Or eModelType = ProductType_MHE Then
                                   If PartClassExistsInDB("SPS_MemberByMember", Nothing) Then
                                        pSL.Add "SPS_MemberByMember"
                                   End If
                                   pSL.Add "MemberByMember"
                                Else
                                    pSL.Add "MemberByMember"
                                    pSL.Add "SPS_MemberByMember"
                                End If
                            End If
                        Else
                             If eModelType = ProductType_PLANT Or eModelType = ProductType_MHE Then
                                If PartClassExistsInDB("SPS_MemberByMember", Nothing) Then
                                    pSL.Add "SPS_MemberByMember"
                                End If
                                pSL.Add "MemberByMember"
                            Else
                                pSL.Add "MemberByMember"
                                pSL.Add "SPS_MemberByMember"
                            End If
                        End If
                    Else
                     If eModelType = ProductType_PLANT Or eModelType = ProductType_MHE Then
                            If PartClassExistsInDB("SPS_MemberByMember", Nothing) Then
                                pSL.Add "SPS_MemberByMember"
                            End If
                            pSL.Add "MemberByMember"
                        Else
                            pSL.Add "MemberByMember"
                            pSL.Add "SPS_MemberByMember"
                        End If
                    End If
                End If
            End If
        ElseIf count = 3 Then
            pSL.Add "MemberBy2Members"
        End If
    End If
 
    Set oStructConn = Nothing
    Set colPorts = Nothing
    Exit Sub
    
ErrorHandler:
    pSL.ReportError
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
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
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
'*************************************************************************
'Function
'AnyAvailableSelectionsOnDummySO
'
'Abstract
'Returns a flag(True/False) if there is any valid selection by executing a selection rule(of smart class obj) on dummy Smart Occurrence.
'
'  Running a selection on existing smart occurence might internally change the user type of itself, thereby
'  we create a dummy Assembly Connection,
'  assign all the needed properties to the Asembly connection from passed persisted Smart Occureence
'  Create a dummy Smart occurence(thorugh dummy Assembly connection)- nothing but we are exactly creating a copy of persisted Smart Occurence passed
'  and execute the selection rule on this dummy smart occurence by passing the Smart Class object(on which the selection needs to be executed)
'  After the execution WE NEED TO DELETE the dummy Assyembly connection(created above)
'
'Arguments
'persistedSmartOccurenrce is input to the function via which dummy AC and dummy SO are created
'catalogSmartClassObj is input on which selection rule nees to be executed
'
'Return
'True, if any valid selections are available False otherwise
'
'***************************************************************************
Private Function AnyAvailableSelectionsOnDummySO(persistedSmartOccurenrce As IJSmartOccurrence, catalogSmartClassObj As IJSmartClass) As Boolean

    Const METHOD = "AnyAvailableSelectionsOnDummySO"
    On Error GoTo ErrorHandler
    
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oDummyStructConn As IJAppConnection
    Dim oDummySmartOcc As IJSmartOccurrence
    Dim oDummyAssConn As IJStructAssemblyConnection
    Dim oAssConnFactory As New StructConnectionsFactory
    Dim oSelHlper As IJSmartOccurrenceHelper
    Dim oListofSelColl As IJDOutputCollection
    
    Dim oModelObj As iJDObject
    Dim oSymbolFactory  As IJDSymbolEntitiesFactory
    Dim oDummyRefColl As IJDReferencesCollection
    
    AnyAvailableSelectionsOnDummySO = False
    
    Set oStructConn = persistedSmartOccurenrce
    oStructConn.enumPorts colPorts
    Set oModelObj = oStructConn
    Set oStructConn = Nothing
    Set oSymbolFactory = New DSymbolEntitiesFactory
    
    'Create Dummy Ref Coll and Dummy Assembly connection
    Set oDummyRefColl = oSymbolFactory.CreateEntity(referencesCollection, oModelObj.ResourceManager)
    Set oDummyAssConn = oAssConnFactory.CreateStructAssemblyConnection(oModelObj.ResourceManager)

    Set oModelObj = Nothing
    
    Set oDummyStructConn = oDummyAssConn
    oDummyStructConn.addPort colPorts.Item(1)
    oDummyStructConn.addPort colPorts.Item(2)

    Set oDummySmartOcc = oDummyStructConn
    Set oSelHlper = New CSmartOccurrenceCES
    
    'connect Fref collection to the dummy smart occurrence
    ConnectSmartOccurrence oDummySmartOcc, oDummyRefColl
    
    Dim oReferencesCollection As IJDReferencesCollection
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(persistedSmartOccurenrce)
    
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oDummyEditJDArgument As IJDEditJDArgument
    
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    Set oDummyEditJDArgument = oDummyRefColl.IJDEditJDArgument
    
    'Fill the same reference Args in the dummy ref coll from the ref coll of persisted smart occurence
    'in other words we are creating a copy of persisted ref coll so that it can be attached to the dummy Smart Occurrence
    'dummy smart occ itself should be the copy of persisted smart occurence
    Dim iIndex As Long
    For iIndex = 1 To oEditJDArgument.GetCount
        oDummyEditJDArgument.SetEntity iIndex, oEditJDArgument.GetEntityByIndex(iIndex)
    Next
    
    Set oSelHlper = New CSmartOccurrenceCES
    Set oListofSelColl = oSelHlper.ExecuteSelectionRule(catalogSmartClassObj, catalogSmartClassObj.SelectionRuleDef, oDummySmartOcc)

    If oListofSelColl.count > 0 Then
        AnyAvailableSelectionsOnDummySO = True
    End If

    Set oSelHlper = Nothing
    Set oListofSelColl = Nothing
    
    'Remove or deleted whatever we have created above
    Dim oRelationHelper As IMSRelation.DRelationHelper
    Dim oRevisionMgr As IJRevision
    
    Set oRelationHelper = oDummySmartOcc
    Set oRevisionMgr = New JRevision

    oDummyRefColl.IJDEditJDArgument.RemoveAll
    
    Set oDummyRefColl = Nothing
    Set oRelationHelper = Nothing
    
    oDummyStructConn.removePort colPorts.Item(2)
    oDummyStructConn.removePort colPorts.Item(1)

    oRevisionMgr.Delete oDummyAssConn
    
    Set oRevisionMgr = Nothing
    Set oDummyStructConn = Nothing
    Set oDummyAssConn = Nothing
    Set oDummySmartOcc = Nothing
    Set oAssConnFactory = Nothing
    Set oSymbolFactory = Nothing
    
    Exit Function
    
ErrorHandler:
    HandleError m_SelectorProgid, METHOD
    
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************






