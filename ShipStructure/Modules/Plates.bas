Attribute VB_Name = "Plates"
Sub PlateSystem_AddNewBoundaries(oPlateSystem As Object, pGeometricConstructionMacro As IJGeometricConstructionMacro)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "PlateSystem_AddNewBoundaries"
    Call DebugIn("PlateSystem_AddNewBoundaries")
    
    Dim pElementsOfRegularBoundaries As IJElements: Set pElementsOfRegularBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
    Dim pElementsOfSubtitutedBoundaries As IJElements: Set pElementsOfSubtitutedBoundaries = SubstituteBoundariesBySystems(pGeometricConstructionMacro, pElementsOfRegularBoundaries)
    Dim pStructApplyOperation As IJDStructApplyOperation: Set pStructApplyOperation = oPlateSystem
    Dim pElementsOfOperators As IJElements: Set pElementsOfOperators = New JObjectCollection
    Call pStructApplyOperation.GetOperation("PlateEntity.PlateBound_AE.1", pElementsOfOperators)
    
    Dim bAreNewBoundaries As Boolean: Let bAreNewBoundaries = False
    Dim i As Integer
    For i = 1 To pElementsOfSubtitutedBoundaries.Count
        If Not pElementsOfOperators.Contains(pElementsOfSubtitutedBoundaries(i)) Then
            If Not GeometricConstructionMacro_IsOutputUnused(pGeometricConstructionMacro, pElementsOfRegularBoundaries(i)) Then
                Dim sKey As String: sKey = pElementsOfSubtitutedBoundaries.GetKey(pElementsOfSubtitutedBoundaries(i))
                Dim oRegularBoundary As Object: Set oRegularBoundary = pElementsOfRegularBoundaries(sKey)
                If pElementsOfOperators.Contains(oRegularBoundary) Then
                    Dim lIndex As Long: lIndex = pElementsOfOperators.GetIndex(oRegularBoundary)
                    pElementsOfOperators.Remove (lIndex)
                End If
                Call pElementsOfOperators.Add(pElementsOfSubtitutedBoundaries(i))
                Let bAreNewBoundaries = True
            End If
        End If
    Next
    
    If bAreNewBoundaries Then
        Call pStructApplyOperation.SetOperation("PlateEntity.PlateBound_AE.1", pElementsOfOperators)
    End If
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
'
' private functions
'

Public Sub PlateSystem_BoundByFrameConnection(oRootPlateSystem As Object, oFrameConnection As Object)

End Sub

Public Function GetMemberFromPort(pPort As IJPort) As Object
    ' prepare result
    Dim oMember As Object
    
    ' retrieve connectable from port
    Dim oConnectable As Object: Set oConnectable = pPort.Connectable
                    
    If TypeOf oConnectable Is IJPlateSystem Then
        ' retrieve built-up
        Set oMember = PlateSystem_GetBuiltUp(oConnectable)
    Else
        ' retrieve rolled member
        Set oMember = oConnectable
    End If
    
    ' return reault
    Set GetMemberFromPort = oMember
End Function

Public Function PlateSystem_GetElementsOfBoundaries(oPlateSystem As Object) As IJElements
    ' prepare result
    Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = New JObjectCollection
    
    ' get info about the bound operation
    Dim pStructApplyOperation As IJDStructApplyOperation: Set pStructApplyOperation = oPlateSystem
    Call pStructApplyOperation.GetOperation("PlateEntity.PlateBound_AE.1", pElementsOfBoundaries)
    
    ' return result
    Set PlateSystem_GetElementsOfBoundaries = pElementsOfBoundaries
End Function
Public Sub PlateSystem_SetElementsOfBoundaries(oPlateSystem As Object, pElementsOfBoundaries)
    Dim pStructApplyOperation As IJDStructApplyOperation: Set pStructApplyOperation = oPlateSystem
    Call pStructApplyOperation.SetOperation("PlateEntity.PlateBound_AE.1", pElementsOfBoundaries)
End Sub
Public Sub PlateSystem_BoundByEdgeOfAPS(oPlateSystemToBound As Object, oPlateSystemOfBoundary As Object, pPositionOfNode As IJDPosition, oBoundary As Object, oBuiltUp As Object)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "PlateSystem_BoundByEdgeOfAPS"
    Call DebugIn("PlateSystem_BoundByEdgeOfAPS")
   
    ' retrieve the boundaries of the plate system
    Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = PlateSystem_GetElementsOfBoundaries(oPlateSystemToBound)
            
    ' loop on boundaries to find a frame connection at the node location
    Dim iBoundary As Integer
    For iBoundary = 1 To pElementsOfBoundaries.Count
        ' just worry about frame and split connections
        If TypeOf pElementsOfBoundaries(iBoundary) Is ISPSFrameConnection _
        Or TypeOf pElementsOfBoundaries(iBoundary) Is ISPSSplitMemberConnection Then
            ' connection encountered, get its point
            Dim pPointOfBoundary As IJPoint
            If TypeOf pElementsOfBoundaries(iBoundary) Is ISPSFrameConnection Then
                Dim pFrameConnection As ISPSFrameConnection: Set pFrameConnection = pElementsOfBoundaries(iBoundary)
                Set pPointOfBoundary = pFrameConnection.Joint.Point
            Else
                Dim pSplitMemberConnection As ISPSSplitMemberConnection: Set pSplitMemberConnection = pElementsOfBoundaries(iBoundary)
                Set pPointOfBoundary = pSplitMemberConnection
            End If
            
            ' verify that the frame connection belongs to the built-up
            If Position_FromPoint(pPointOfBoundary).DistPt(pPositionOfNode) < EPSILON Then
                Let sComment = "retrieve the edge port of the APS corresponding to the bounding surface"
                Dim pPortOfEdge As IJPort: Set pPortOfEdge = PlateSystem_GetEdgePortFromBoundingSurface(oPlateSystemOfBoundary, oBoundary)
                        
                Let sComment = "replace the frame connection by the edge port"
                On Error Resume Next
                Call pElementsOfBoundaries.Remove(iBoundary)
                Call pElementsOfBoundaries.Add(pPortOfEdge)
                Call PlateSystem_SetElementsOfBoundaries(oPlateSystemToBound, pElementsOfBoundaries)
                On Error GoTo 0
                Exit For
            End If
        End If
    Next
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub

Public Sub PlateSystem_AddAuxPortToFCOfMember(oPort As IJPort, oPlateSystemOfBoundary As Object, pPositionOfNode As IJDPosition, oMember As Object, bAdd As Boolean)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "PlateSystem_BoundByEdgeOfAPS"
    Call DebugIn("PlateSystem_AddEdgeToFCOfMember")

    ' retrieve the boundaries of the plate system
     Dim oConnection As Object:
    If bGetExactSolution Then
        Set oConnection = BuiltUp_RetrieveConnectionAtNode(oMember, pPositionOfNode)
    Else
        Set oConnection = BuiltUp_RetrieveConnectionAtNode2(oMember, pPositionOfNode)
    End If

    Dim oMemberConn As ISPSMemberConnection
    Dim oconnectablePlateSystem As IJPlateSystem
               
    If TypeOf oConnection Is ISPSMemberConnection Then
        Set oMemberConn = oConnection
        If bAdd Then
            oMemberConn.AddAuxiliaryPort oPort
        Else
            oMemberConn.RemoveAuxiliaryPort oPort ' 224054  remove aux port if the member is no more bounded.
        End If
        On Error Resume Next
        '227617 Add secondary member is not updaing the plate geometry
        'Added this kludge to trigger Plate Part semantic so that ACs are generated
        'when the plate was already detailed.
        If bAdd Then
            Set oconnectablePlateSystem = oPort.Connectable
            If Not oconnectablePlateSystem Is Nothing Then
                UpdateGeneratedPartsOfSystem oconnectablePlateSystem
            End If
        End If
        On Error GoTo ErrorHandler
    End If

     Set oMemberConn = Nothing
     Set oConnection = Nothing
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Public Sub PlateSystem_UnBoundFromEdgeOfAPS(oPlateSystemToBound As Object, oPlateSystemOfBoundary As Object, _
                                                                    pPositionOfNode As IJDPosition, oBuiltUp As Object, pGeometricConstruction As IJGeometricConstruction)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "PlateSystem_UnBoundFromEdgeOfAPS"
    Call DebugIn("PlateSystem_UnBoundFromEdgeOfAPS")
    
    ' retrieve its boundaries
    Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = PlateSystem_GetElementsOfBoundaries(oPlateSystemToBound)
               
    ' loop on boundaries to find an edge port of the APS
    Dim iBoundary As Integer
    For iBoundary = 1 To pElementsOfBoundaries.Count
        ' just worry about ports
        If TypeOf pElementsOfBoundaries(iBoundary) Is IJPort Then
            ' port encountered
            Dim pPort As IJPort: Set pPort = pElementsOfBoundaries(iBoundary)
                
            ' verify that the port is an edge port (or a port being deleted) coming from the bounding plate system
            If (pPort.Type = PortEdge Or pPort.Type = 0) _
            And pPort.Connectable Is oPlateSystemOfBoundary Then
                Dim bIsFrameConnectionAvailable As Boolean: Let bIsFrameConnectionAvailable = False
                
                ' retrieve the connection at node
                Dim oConnection As Object: Set oConnection = BuiltUp_RetrieveConnectionAtNode(oBuiltUp, pPositionOfNode)
                If Not oConnection Is Nothing Then
                    ' verify that the frame connection is not operator of the incoming plate
                    If Not pElementsOfBoundaries.Contains(oConnection) Then Let bIsFrameConnectionAvailable = True
                End If
                
                ' remove the PlateBound_EdgeOper relationship explicitely (bug in SM3d)
                Call Entity_DisconnectRelatedEntity(pPort, "IJNoUpdatePortTrigger", "PlateBound_EdgeOper_DEST")
                                
                ' remove the edge port from the boundary list
                 Call pElementsOfBoundaries.Remove(iBoundary)
                 
                If bIsFrameConnectionAvailable Then
                    ' add the frame connection to the boundary list
                    Call pElementsOfBoundaries.Add(oConnection)
                Else
                    ' connect the PlateSystem to the GCMacro to remember it is missing a FrameConnection
                    If pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Count = 0 Then
                        Call pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Add(oPlateSystemToBound)
                    End If
                End If
                        
                ' update boundary list
                Call PlateSystem_SetElementsOfBoundaries(oPlateSystemToBound, pElementsOfBoundaries)
                        
                Exit For
            End If
        End If
    Next
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
'''Public Function PlateSystem_GetSupport(pPlateSystem As IJPlateSystem) As Object
'''    Dim pStructCustomGeometry As IJDStructCustomGeometry
'''    Set pStructCustomGeometry = pPlateSystem
'''
'''    Dim sProgid As String
'''    Dim pElementsOfParents As IJElements: Set pElementsOfParents = New JObjectCollection
'''    Call pStructCustomGeometry.GetCustomGeometry(sProgid, pElementsOfParents)
'''    Debug.Print "Progid= " + sProgid
'''
'''    Set PlateSystem_GetSupport = pElementsOfParents(1)
'''End Function
Public Function PlateSystem_GetEdgePortFromBoundingSurface(oPlateSystem As Object, oBoundingSurface As Object) As IJPort
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "PlateSystem_GetEdgePortFromBoundingSurface"
    Call DebugIn("PlateSystem_GetEdgePortFromBoundingSurface")
    
    ' prepare result
    Dim pPort As IJPort
    
    ' instantiate helper
    Dim pPortHelper As IPortHelper
    Set pPortHelper = New PortHelper
   
    ' retrieve operator for bound
    Dim pElementsOfOperators As IJElements: Set pElementsOfOperators = New JObjectCollection
    If True Then
        Dim pStructApplyOperation As IJDStructApplyOperation: Set pStructApplyOperation = oPlateSystem
        Call pStructApplyOperation.GetOperation("PlateEntity.PlateBound_AE.1", pElementsOfOperators)
    End If
    
    ' retrieve BoundAE
    Dim oBoundAE As Object: Set oBoundAE = PlateSystem_GetBoundAE(oPlateSystem)
    
    Dim pAssocRelation As IJDAssocRelation
    Set pAssocRelation = oBoundAE
    
    Dim pRelationshipColl As IJDRelationshipCol
    Set pRelationshipColl = pAssocRelation.CollectionRelations("IJPlateBound", "PlateBound_OPER2_ORIG")
    
    ' compute OperationId
    Dim lOperationId As Long: Let lOperationId = pPortHelper.GetObjectIdAsLong(oBoundAE)

    ' retrieve operator id corresponding to the bounding surface
    Dim lOperatorId As Long
    If True Then
        Dim iRelationShip As Integer
        For iRelationShip = 1 To pRelationshipColl.Count
            Dim pRelationship As IJDRelationship: Set pRelationship = pRelationshipColl.Item(iRelationShip)
            If pRelationship.Target Is oBoundingSurface Then
                Let lOperatorId = Val(pRelationship.Name)
                Exit For
            End If
        Next
    End If
    
    ' try first to retrieve an existing port
    If True Then
        Dim pConnectable As IJConnectable: Set pConnectable = oPlateSystem
        Dim pElementsOfPorts As IJElements: Call pConnectable.enumConnectedPorts(pElementsOfPorts, PortEdge)
        Dim iPort As Integer
        For iPort = 1 To pElementsOfPorts.Count
            Dim pStructPort As IJStructPort: Set pStructPort = pElementsOfPorts(iPort)
            If pStructPort.ProxyType = JS_TOPOLOGY_PROXY_LEDGE _
            And pStructPort.OperationID = lOperationId _
            And pStructPort.OperatorID = lOperatorId _
            And pStructPort.SectionID = -1 Then
                Set pPort = pStructPort
                Exit For
            End If
        Next
    End If
        
    If pPort Is Nothing Then
        'create the port
        
        ' encode moniker of proxy
        Dim pMonikerOfProxy As IMoniker
        If True Then
            Dim lContext As Long: Let lContext = 0
            Dim lXid As Long: Let lXid = -1
            Call pPortHelper.EncodeTopologyProxyMoniker(JS_TOPOLOGY_PROXY_LEDGE, lContext, lOperationId, lOperatorId, lXid, pMonikerOfProxy)
        End If
        
        ' create composite moniker of port
        Dim pMonikerOfPort As IMoniker: Set pMonikerOfPort = pPortHelper.CreateCompositeMonikerEx1(pMonikerOfProxy, "SL_")
        
        ' retrieve or create the port
        Set pPort = pPortHelper.GetPort(oPlateSystem, pMonikerOfPort)
    End If
    
    ' return result
    Set PlateSystem_GetEdgePortFromBoundingSurface = pPort
    Call DebugOut
    Exit Function
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Function
Public Function PlateSystem_GetFacePortFromBoundingSurface(oMemberSystem As Object, oBoundingSurface As Object) As IJPort
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "PlateSystem_GetEdgePortFromBoundingSurface"
    Call DebugIn("PlateSystem_GetEdgePortFromBoundingSurface")
    
    ' prepare result
    Dim pPort As IJPort
    
    ' instantiate helper
    Dim pPortHelper As IPortHelper
    Set pPortHelper = New PortHelper
     
    Dim pElementsOfOperators As IJElements: Set pElementsOfOperators = New JObjectCollection
  
    
    ' try first to retrieve an existing port
    If True Then
        Dim pConnectable As IJConnectable: Set pConnectable = oBoundingSurface
        Dim pElementsOfPorts As IJElements: Call pConnectable.enumConnectedPorts(pElementsOfPorts, PortFace)
  
        Dim iPort As Integer
        For iPort = 1 To pElementsOfPorts.Count
            Dim pStructPort As IJStructPort: Set pStructPort = pElementsOfPorts(iPort)
            If pStructPort.ContextID = CTX_NPLUS Then
                Set pPort = pStructPort
                Exit For
            End If
        Next
    End If
        
     If pPort Is Nothing Then
        'create the port
        
        ' encode moniker of proxy
        Dim pMonikerOfProxy As IMoniker
        If True Then
            Dim lContext As Long: Let lContext = 8
            Dim lXid As Long: Let lXid = -1
            Call pPortHelper.EncodeTopologyProxyMoniker(JS_TOPOLOGY_PROXY_LFACE, lContext, 0, 0, lXid, pMonikerOfProxy)
        End If
        
        ' create composite moniker of port
        Dim pMonikerOfPort As IMoniker: Set pMonikerOfPort = pPortHelper.CreateCompositeMonikerEx1(pMonikerOfProxy, "SL_")
        
        ' retrieve or create the port
        Set pPort = pPortHelper.GetPort(oBoundingSurface, pMonikerOfPort)
    End If
    ' return result
    Set PlateSystem_GetFacePortFromBoundingSurface = pPort
    Call DebugOut
    Exit Function
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Function
Public Function PlateSystem_GetBuiltUp(oPlateSystem As Object) As Object
    Dim pElementsOfBuiltUps As IJElements: Set pElementsOfBuiltUps = Entity_GetElementsOfRelatedEntities(oPlateSystem, "IJShpStrDesignChild", "ShpStrDesignParent")
    If pElementsOfBuiltUps.Count > 0 Then Set PlateSystem_GetBuiltUp = pElementsOfBuiltUps(1)
End Function
Function PlateSystem_GetBoundAE(oPlateSystem As Object) As Object
    ' prepare the result
    Dim oBoundAE As Object
    
    ' retrieve bounded geometry
    Dim pPlateGeometryHelper As IJPlateGeometryHelper: Set pPlateGeometryHelper = New PlateGeometryHelper
    Dim oPlateGeometry As Object: Set oPlateGeometry = pPlateGeometryHelper.geometryBeforeCutout(oPlateSystem)
    
    ' retrieve bound AE
    Set oBoundAE = Entity_GetElementsOfRelatedEntities(oPlateGeometry, "IJStructGeometry", "StructOperation_RSLT1_DEST")(1)

    ' return result
    Set PlateSystem_GetBoundAE = oBoundAE
End Function
Public Function PlateSystem_GetMacro(pPlateSystem As IJPlateSystem) As IJGeometricConstructionMacro
    If TypeOf pPlateSystem Is IJDStructCustomGeometry Then
        Dim pStructCustomGeometry As IJDStructCustomGeometry
        Set pStructCustomGeometry = pPlateSystem
        
        Dim sProgid As String
        Dim pElementsOfParents As IJElements: Set pElementsOfParents = New JObjectCollection
        On Error Resume Next
        Call pStructCustomGeometry.GetCustomGeometry(sProgid, pElementsOfParents)
        On Error GoTo 0
        
        If Err.Number = 0 And Not pElementsOfParents Is Nothing Then
            If pElementsOfParents.Count > 0 Then
                Dim pElementsOfGrandParents As IJElements
                Set pElementsOfGrandParents = Entity_GetElementsOfRelatedEntities(pElementsOfParents(1), "IJGeometry", "ConstructionForOutput")
                If Not pElementsOfGrandParents Is Nothing Then
                    If pElementsOfGrandParents.Count > 0 Then
                        Dim oGeometricConstructionMacro As Object
                        Set oGeometricConstructionMacro = pElementsOfGrandParents.Item(1)
                          
                        If TypeOf oGeometricConstructionMacro Is IJGeometricConstructionMacro Then
                            Set PlateSystem_GetMacro = oGeometricConstructionMacro
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
Public Function SubstituteBoundariesBySystems(pGeometricConstruction As IJGeometricConstruction, pElementsOfBoundaries As IJElements) As IJElements
    Dim pGeometricConstructionMacro As IJGeometricConstructionMacro
    Set pGeometricConstructionMacro = pGeometricConstruction
    
    ' retrieve names of controlled inputs
    Dim pCollectionOfNamesOfControlledInputs As Collection
    Set pCollectionOfNamesOfControlledInputs = GeometricConstruction_GetNamesOfControlledInputs(pGeometricConstruction)
    
    'prepare a new collection of boundaries to be filled
    Dim pElementsOfNewBoundaries As IJElements: Set pElementsOfNewBoundaries = New JObjectCollection
    
    ' substitute some boundaries by the corresponding PlateSystems
    If True Then
        ' loop on boundaries
        Dim i As Integer
        For i = 1 To pElementsOfBoundaries.Count
            ' retrieve the boundary
            Dim oBoundary As Object: Set oBoundary = pElementsOfBoundaries(i)
            
            ' get the key of the boundary
            Dim sKey As String: Let sKey = pGeometricConstructionMacro.Outputs("Boundary").GetKey(oBoundary)
            
            ' check if the key corresponds to the name of a controlled input
            If Collection_ContainsName(pCollectionOfNamesOfControlledInputs, sKey) Then
                ' retrieve the collection of the corresponding controlled input
                Dim pElementsOfInputs As IJElements: Set pElementsOfInputs = pGeometricConstruction.ControlledInputs(sKey)
                If Not pElementsOfInputs Is Nothing Then
                    If pElementsOfInputs.Count = 1 Then
                        If TypeOf pElementsOfInputs(1) Is IJPort Then
                            'substitute the boundary by the corresponding PlateSystem
                            Dim pPort As IJPort
                            Set pPort = pElementsOfInputs(1)
                            If pPort.Type = PortEdge And TypeOf pPort.Connectable Is IJPlateSystem Then
                                Set oBoundary = pPort
                            ElseIf pPort.Type = PortFace And TypeOf pPort.Connectable Is IJPlateSystem Then
                                Set oBoundary = pPort.Connectable
                            ElseIf pPort.Type = PortEdge And TypeOf pPort.Connectable Is IJStiffener Then
                                Set oBoundary = pPort.Connectable
                            ElseIf pPort.Type = PortFace And TypeOf pPort.Connectable Is ISPSMemberPartPrismatic Then
                                Dim pMemberPartPrismatic As ISPSMemberPartPrismatic
                                Set pMemberPartPrismatic = pPort.Connectable
                                Set oBoundary = pMemberPartPrismatic.MemberSystem
                            ElseIf pPort.Type = PortFace And TypeOf pPort.Connectable Is IJPlatePart Then
                                Dim oPlatePartUtils As IJPlatePartAttributes
                                Set oPlatePartUtils = New PlatePartUtils
                                On Error Resume Next
                                Set oBoundary = oPlatePartUtils.GetPartRootPlateSystem(pPort.Connectable)
                            End If
                        ElseIf TypeOf pElementsOfInputs(1) Is IJPlateSystem Then
                            'MsgBox "PlateSystem as controlled input"
                            Set oBoundary = pElementsOfInputs(1)
                        End If
                    End If
                End If
            End If
            
            ' add the boundary to the new collection
            Call pElementsOfNewBoundaries.Add(oBoundary, sKey)
        Next
    End If
    
    ' return result
    Set SubstituteBoundariesBySystems = pElementsOfNewBoundaries
End Function


Private Sub UpdateGeneratedPartsOfSystem(oSystem As Object)
    
    Dim oStructDetailHelper As New StructDetailHelper
    Dim oPartsEnum As IEnumUnknown
    Dim oConvertUtils As New CONVERTUTILITIESLib.CCollectionConversions
    Dim oPartsCol As Collection
    Dim oPart As Object
    Dim oIJPartGeometryState As IJPartGeometryState
    Dim bIsDetailed As Boolean
    Dim IID_IJTriggerGenPlatePartBounded As String
    Dim strProgID As String
    Dim oAssocHelper As New AssocHelper
    Dim oPartGeometry As Object
    
    IID_IJTriggerGenPlatePartBounded = "{2449F501-4FFB-460b-BFDB-ACAAE49D1DCC}"
    strProgID = "CreatePlatePart.GeneratePlatePart_AE.1"
    oStructDetailHelper.GetPartsDerivedFromSystem oSystem, oPartsEnum, True
    Call oConvertUtils.CreateVBCollectionFromIEnumUnknown(oPartsEnum, oPartsCol)
    For Each oPart In oPartsCol
        Set oIJPartGeometryState = oPart
        If Not oIJPartGeometryState Is Nothing Then
            If oIJPartGeometryState.PartGeometryState = PartGeometryStateType.DetailedPart Then
                Set oPartGeometry = oStructDetailHelper.GetGeomAfterOpByProgID(strProgID, oPart)
                If Not oPartGeometry Is Nothing Then
                    oAssocHelper.UpdateObject oPartGeometry, IID_IJTriggerGenPlatePartBounded
                End If
            End If
        End If
    Next
End Sub
