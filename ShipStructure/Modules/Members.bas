Attribute VB_Name = "Members"
Option Explicit
Const E_FAIL = &H80004005
Const m_sModule = "GCSidePlate.Members"

Private m_MemberSystemReference As Object
Private Property Get MemberSystemReference() As Object
    Set MemberSystemReference = m_MemberSystemReference
End Property
Private Property Set MemberSystemReference(oMemberSystem As Object)
    Set m_MemberSystemReference = oMemberSystem
End Property
Public Sub MemberSystemReference_Set(oMemberOriginal As Object, Optional oMemberReplacing As Object)
    Set MemberSystemReference = Nothing
    
    If Not oMemberOriginal Is Nothing Then
        Dim pMemberPartCommon As ISPSMemberPartCommon
        Set pMemberPartCommon = oMemberOriginal
        If Not pMemberPartCommon Is Nothing Then
            If Not pMemberPartCommon.MemberSystem Is Nothing Then
                ' store oMemberOriginal
                Set MemberSystemReference = pMemberPartCommon.MemberSystem
                
            ElseIf Not oMemberReplacing Is Nothing Then
                Set pMemberPartCommon = Nothing
                Set pMemberPartCommon = oMemberReplacing
                If Not pMemberPartCommon Is Nothing Then
                    If Not pMemberPartCommon.MemberSystem Is Nothing Then
                        ' store oMemberReplacing
                        Set MemberSystemReference = pMemberPartCommon.MemberSystem
                    End If
                End If
            End If
            
        End If
    End If
   
End Sub
Public Function MemberSystemReference_Get(oMemberPart As Object) As Object
    Set MemberSystemReference_Get = Nothing
    
    If Not oMemberPart Is Nothing Then
        Dim pMemberPartCommon As ISPSMemberPartCommon
        Set pMemberPartCommon = oMemberPart
        If Not pMemberPartCommon Is Nothing Then
            If Not pMemberPartCommon.MemberSystem Is Nothing Then
                Set MemberSystemReference_Get = pMemberPartCommon.MemberSystem
            Else
                Set MemberSystemReference_Get = MemberSystemReference
            End If
        End If
    End If
    
End Function

Public Sub Member_UnboundOnRemove(pGeometricConstructionMacro As IJGeometricConstructionMacro, oRemovedMember As Object)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Member_UnboundOnRemove"
    Call DebugIn("Member_UnboundOnRemove")
    
    ' use the GCMacro as a GC
    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pGeometricConstructionMacro
    Dim oAdvancedPlateSystem As Object
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        ' retrieve advanced plate system
        Set oAdvancedPlateSystem = pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1)
    End If
    If oAdvancedPlateSystem Is Nothing Then
        'APS has been removed, do nothing
        Call DebugOut
        Exit Sub
    End If

    ' retrieve the input index of the secondary membert parts
    Dim iInputIndexOfPrimaryMemberParts As Integer: Let iInputIndexOfPrimaryMemberParts = MacroDefinition_GetInputIndex(pGeometricConstruction.definition, "MemberParts")
    Dim iInputIndexOfSecondaryMemberParts As Integer: Let iInputIndexOfSecondaryMemberParts = MacroDefinition_GetInputIndex(pGeometricConstruction.definition, "SecondaryMemberParts")
        
    Let sComment = "retrieve info about the primary members"
    Dim pElementsOfPrimaryMembers As IJElements: Set pElementsOfPrimaryMembers = Elements_GetValidElements(pGeometricConstruction.Inputs("MemberParts"))
    Dim iCountOfPrimaryMembers As Integer: Let iCountOfPrimaryMembers = pElementsOfPrimaryMembers.Count
    
    Dim sKeyOfRemovedMember As String: Let sKeyOfRemovedMember = ""
    Dim k As Integer
    
    If iCountOfPrimaryMembers > 0 Then
        Let sComment = "retrieve info about the boundaries"
        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
        Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
        Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)
        Dim sKeysOfPrimaryMembers() As String: Let sKeysOfPrimaryMembers = Elements_GetKeys(pElementsOfPrimaryMembers)

        For k = 1 To iCountOfPrimaryMembers
            Dim omb As Object
            Set omb = pElementsOfPrimaryMembers(k)
            If oRemovedMember Is omb Then
                sKeyOfRemovedMember = pElementsOfPrimaryMembers.GetKey(omb)
            End If
        Next
        
        
        Let sComment = "retrieve position of node"
        Dim pPositionOfNode As IJDPosition
        
        ' if something goes wrong, we will at least try to retrieve the node
        On Error Resume Next
        
        ' retrieve lines of member axes
        Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = Members_GetLinesOfMemberSystemAxes(pElementsOfPrimaryMembers)

        ' compute the node
        Set pPositionOfNode = GetPositionAtCommonExtremityOfLines(pLinesOfMemberAxes)
        
        On Error GoTo 0
    
        ' retrieve the node, if not computed
        If pPositionOfNode Is Nothing Then Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstruction)
             
        If sKeyOfRemovedMember <> "" Then
            Call DebugValue("sKeyOfRemovedMember ", sKeyOfRemovedMember)
           
            ' un-bound removed primary member
             If Not IsDeleted(oRemovedMember) Then
                ' build prefix for keys
                Dim sPrefixOfBoundaryKey As String: Let sPrefixOfBoundaryKey = "TR" + CStr(iInputIndexOfPrimaryMemberParts)
                Let sComment = "un-bound the removed primary Member"
                 DebugMsg ("un-bound the removed primary Member")
                 
                 Call Member_BoundByOrUnBoundFromEdgeOfAPS(oRemovedMember, pGeometricConstruction, oAdvancedPlateSystem, Nothing, pElementsOfBoundaries(sPrefixOfBoundaryKey + "." + sKeyOfRemovedMember), pPositionOfNode)
                     
                 If Not oAdvancedPlateSystem Is Nothing Then
                    Dim sAxisPortkey As String: Let sAxisPortkey = "AxisPort" + CStr(iInputIndexOfPrimaryMemberParts)
                    Dim sFacePortkey As String: Let sFacePortkey = "FacePort" + CStr(iInputIndexOfPrimaryMemberParts)
                     Call DebugValue("disconnect the AxisPort and the FacePort controlledInput", iInputIndexOfPrimaryMemberParts)
                     ' disconnect the AxisPort and the FacePort. If the Member is just removed, we do not want to maintain controlled inputs to him
                                         ' try to retrieve the FacePort corresponding to the removed member
                        Dim oAxisPort As Object: On Error Resume Next: Set oAxisPort = Nothing: Set oAxisPort = pGeometricConstruction.ControlledInputs(sAxisPortkey)(sKeyOfRemovedMember): On Error GoTo 0
                        If Not oAxisPort Is Nothing Then Call pGeometricConstruction.ControlledInputs(sAxisPortkey).Remove(sKeyOfRemovedMember)
                        Dim oFacePort As Object: On Error Resume Next: Set oFacePort = Nothing: Set oFacePort = pGeometricConstruction.ControlledInputs(sFacePortkey)(sKeyOfRemovedMember): On Error GoTo 0
                        If Not oFacePort Is Nothing Then Call pGeometricConstruction.ControlledInputs(sFacePortkey).Remove(sKeyOfRemovedMember)
                 End If
             End If
         End If

        If sKeyOfRemovedMember <> "" Then
            'Remove member was primary, do not search in secondary
            Call DebugOut
            Exit Sub
        End If
        
        Let sComment = "retrieve info about the secondary members"
        Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = Elements_GetValidElements(pGeometricConstruction.Inputs("SecondaryMemberParts"))
        Dim iCountOfSecondaryMembers As Integer: Let iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count

        If iCountOfSecondaryMembers > 0 Then
            Dim sKeysOfSecondaryMembers() As String: Let sKeysOfSecondaryMembers = Elements_GetKeys(pElementsOfSecondaryMembers)
            Let sComment = "retrieve info about the boundaries"
            Dim pElementsOfPseudoBoundaries As IJElements: Set pElementsOfPseudoBoundaries = pGeometricConstructionMacro.Outputs("PseudoBoundary")
            Dim iCountOfPseudoBoundaries As Integer: Let iCountOfPseudoBoundaries = pElementsOfPseudoBoundaries.Count
            Dim sKeysOfPseudoBoundaries() As String: If iCountOfPseudoBoundaries > 0 Then Let sKeysOfPseudoBoundaries = Elements_GetKeys(pElementsOfPseudoBoundaries)
           
            For k = 1 To iCountOfSecondaryMembers
                Set omb = pElementsOfSecondaryMembers(k)
                If oRemovedMember Is omb Then
                    sKeyOfRemovedMember = pElementsOfSecondaryMembers.GetKey(omb)
                End If
            Next
            
            If sKeyOfRemovedMember <> "" Then
                Call DebugValue("sKeyOfRemovedMember ", sKeyOfRemovedMember)
                
                ' un-bound removed secondary member
                If Not IsDeleted(oRemovedMember) Then
                    Let sComment = "un-bound removed secondary members"
                    DebugMsg ("un-bound removed secondary members")
                    Call Member_BoundByOrUnBoundFromFaceOfAPS(oRemovedMember, False, oAdvancedPlateSystem, pPositionOfNode)
                                
                    If Not oAdvancedPlateSystem Is Nothing Then
                        ' disconnect the AxisPort. If the Member is just removed, we do not want to maintain controlled inputs to him
                        Let sAxisPortkey = "AxisPort" + CStr(iInputIndexOfSecondaryMemberParts)
                        Call DebugValue("disconnect the AxisPort controlledInput", sAxisPortkey)
                        Call pGeometricConstruction.ControlledInputs(sAxisPortkey).Remove(sKeyOfRemovedMember)
                    End If
                End If
            End If
        End If
        End If
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub

Public Sub Members_UnboundOnDelete(pGeometricConstructionMacro As IJGeometricConstructionMacro, oAdvancedPlateSystem As Object)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Members_UnboundOnDelete"
    Call DebugIn("Members_UnboundOnDelete")
    
    ' use the GCMacro as a GC
    Dim pGeometricConstruction As IJGeometricConstruction: Set pGeometricConstruction = pGeometricConstructionMacro
    
    ' retrieve the input index of the secondary membert parts
    Dim iInputIndexOfSecondaryMemberParts As Integer: Let iInputIndexOfSecondaryMemberParts = MacroDefinition_GetInputIndex(pGeometricConstruction.definition, "SecondaryMemberParts")
        
    Let sComment = "retrieve info about the primary members"
    Dim pElementsOfPrimaryMembers As IJElements: Set pElementsOfPrimaryMembers = Elements_GetValidElements(pGeometricConstruction.Inputs("MemberParts"))
    Dim iCountOfPrimaryMembers As Integer: Let iCountOfPrimaryMembers = pElementsOfPrimaryMembers.Count
    Dim sKeysOfPrimaryMembers() As String: If iCountOfPrimaryMembers > 0 Then Let sKeysOfPrimaryMembers = Elements_GetDummyKeys(pElementsOfPrimaryMembers)
    
    Let sComment = "retrieve position of node"
    Dim pPositionOfNode As IJDPosition
    If iCountOfPrimaryMembers > 0 Then
        ' if something goes wrong, we will at least try to retrieve the node
        On Error Resume Next
        
        ' retrieve lines of member axes
        Dim pLinesOfMemberAxes() As IJLine: Let pLinesOfMemberAxes = Members_GetLinesOfMemberSystemAxes(pElementsOfPrimaryMembers)

        ' compute the node
        Set pPositionOfNode = GetPositionAtCommonExtremityOfLines(pLinesOfMemberAxes)
        
        On Error GoTo 0
    
        ' retrieve the node, if not computed
        If pPositionOfNode Is Nothing Then Set pPositionOfNode = GeometricConstructionMacro_RetrievePositionOfNode(pGeometricConstruction)
    End If
    
    If iCountOfPrimaryMembers > 0 Then
        Let sComment = "retrieve info about the boundaries"
        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = pGeometricConstructionMacro.Outputs("Boundary")
        Dim iCountOfBoundaries As Integer: Let iCountOfBoundaries = pElementsOfBoundaries.Count
        Dim sKeysOfBoundaries() As String: If iCountOfBoundaries > 0 Then Let sKeysOfBoundaries = Elements_GetKeys(pElementsOfBoundaries)

        ' un-bound removed primary members
        Call Members_UnboundFromEdgeOfAPS(pElementsOfPrimaryMembers, iCountOfPrimaryMembers, sKeysOfPrimaryMembers, _
                                          pElementsOfBoundaries, iCountOfBoundaries, sKeysOfBoundaries, _
                                          pGeometricConstruction, 1, oAdvancedPlateSystem, pPositionOfNode)
    End If

    Let sComment = "retrieve info about the secondary members"
    Dim pElementsOfSecondaryMembers As IJElements: Set pElementsOfSecondaryMembers = Elements_GetValidElements(pGeometricConstruction.Inputs("SecondaryMemberParts"))
    Dim iCountOfSecondaryMembers As Integer: Let iCountOfSecondaryMembers = pElementsOfSecondaryMembers.Count
    Dim sKeysOfSecondaryMembers() As String: If iCountOfSecondaryMembers > 0 Then Let sKeysOfSecondaryMembers = Elements_GetDummyKeys(pElementsOfSecondaryMembers)

    If iCountOfSecondaryMembers > 0 Then
        Let sComment = "retrieve info about the boundaries"
        Dim pElementsOfPseudoBoundaries As IJElements: Set pElementsOfPseudoBoundaries = pGeometricConstructionMacro.Outputs("PseudoBoundary")
        Dim iCountOfPseudoBoundaries As Integer: Let iCountOfPseudoBoundaries = pElementsOfPseudoBoundaries.Count
        Dim sKeysOfPseudoBoundaries() As String: If iCountOfPseudoBoundaries > 0 Then Let sKeysOfPseudoBoundaries = Elements_GetKeys(pElementsOfPseudoBoundaries)

        Let sComment = "un-bound removed secondary members"
        Call Members_UnboundFromFaceOfAPS(pElementsOfSecondaryMembers, iCountOfSecondaryMembers, sKeysOfSecondaryMembers, _
                                          pElementsOfPseudoBoundaries, iCountOfPseudoBoundaries, sKeysOfPseudoBoundaries, _
                                          pGeometricConstruction, iInputIndexOfSecondaryMemberParts, oAdvancedPlateSystem, pPositionOfNode)

    End If
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Sub Members_BoundByOrUnboundFromEdgeOfAPS(pElementsOfMembers As IJElements, iCountOfMembers As Integer, sKeysOfMembers() As String, _
                                          pElementsOfBoundaries As IJElements, iCountOfBoundaries As Integer, sKeysOfBoundaries() As String, _
                                          pGeometricConstruction As IJGeometricConstruction, iStepIndex As Integer, oAdvancedPlateSystem As Object, pPositionOfNode As IJDPosition)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Members_BoundByOrUnboundFromEdgeOfAPS"
    Call DebugIn("Members_BoundByOrUnboundFromEdgeOfAPS")
    
    ' build prefix for keys
    Dim sPrefixOfBoundaryKey As String: Let sPrefixOfBoundaryKey = "TR" + CStr(iStepIndex)
    
    Let sComment = "retrieve planar support"
    Dim pPlaneOfSupport As IJPlane:
    If True Then
        Dim pGeometricConstructionMacro As IJGeometricConstructionMacro: Set pGeometricConstructionMacro = pGeometricConstruction
        Set pPlaneOfSupport = pGeometricConstructionMacro.Outputs("Support")(1)
    End If
        
    Dim iMember As Integer
    For iMember = 1 To iCountOfMembers
        Let sComment = "bound/un-bound member"
        Call Member_BoundByOrUnBoundFromEdgeOfAPS(pElementsOfMembers(iMember), pGeometricConstruction, oAdvancedPlateSystem, pPlaneOfSupport, pElementsOfBoundaries(sPrefixOfBoundaryKey + "." + sKeysOfMembers(iMember)), pPositionOfNode)
    Next
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Sub Members_BoundByOrUnboundFromFaceOfAPS(pElementsOfMembers As IJElements, oAdvancedPlateSystem As Object, pPositionOfNode As IJDPosition)
    Call DebugIn("Members_BoundByOrUnboundFromFaceOfAPS")
    
    Dim i As Integer
    For i = 1 To pElementsOfMembers.Count
        ' bound/un-bound member
        Call Member_BoundByOrUnBoundFromFaceOfAPS(pElementsOfMembers(i), True, oAdvancedPlateSystem, pPositionOfNode)
    Next
    
    Call DebugOut
End Sub
Sub Member_BoundByOrUnBoundFromEdgeOfAPS(oMember As Object, pGeometricConstructionMacro As IJGeometricConstructionMacro, oPlateSystemAdvanced As Object, pPlaneOfSupport As IJPlane, pPlaneOfBoundary As IJPlane, pPositionOfNode As IJDPosition)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Member_BoundByOrUnBoundFromEdgeOfAPS"
    Call DebugIn("Member_BoundByOrUnBoundFromEdgeOfAPS")
    
    If TypeOf oMember Is ISPSDesignedMember Then
        ' process built-up
        If oPlateSystemAdvanced Is Nothing Then
            Let sComment = "APS has just been deleted, bounding edges have been removed, but bounding FrameConnections have not necessarily been restored"
            Dim oConnection As Object: Set oConnection = BuiltUp_RetrieveConnectionAtNode(oMember, pPositionOfNode)
            If Not oConnection Is Nothing Then Call BuiltUp_BoundPlateSystemsByConnection(oMember, oConnection, pGeometricConstructionMacro)
        Else
            Let sComment = "retrieve web/flange coplanar with APS, if exists"
            Dim oPlateSystemCoplanar As Object: If Not pPlaneOfSupport Is Nothing Then Set oPlateSystemCoplanar = BuiltUp_GetPlateSystemCoplanar(oMember, pPlaneOfSupport)
                    
            Let sComment = "retrieve web/flange already bounded by the APS"
            Dim oPlateSystemBounded As Object:  If Not pPlaneOfBoundary Is Nothing Then Set oPlateSystemBounded = BuiltUp_GetPlateSystemBounded(oMember, oPlateSystemAdvanced, pPlaneOfBoundary)
                    
            Let sComment = "get status of the boundary"
            Dim bIsBoundaryUnused As Boolean: Let bIsBoundaryUnused = False
            If Not pGeometricConstructionMacro Is Nothing Then
                Let bIsBoundaryUnused = GeometricConstructionMacro_IsOutputUnused(pGeometricConstructionMacro, pPlaneOfBoundary)
            End If
            
            Let sComment = "un-bound plate system, which should not be bounded any more"
            If Not oPlateSystemBounded Is Nothing And ((Not oPlateSystemBounded Is oPlateSystemCoplanar) Or bIsBoundaryUnused) Then _
                Call PlateSystem_UnBoundFromEdgeOfAPS(oPlateSystemBounded, oPlateSystemAdvanced, pPositionOfNode, oMember, pGeometricConstructionMacro)
            
            Let sComment = "bound plate system, which needs to be bounded"
            If Not oPlateSystemCoplanar Is Nothing And Not oPlateSystemCoplanar Is oPlateSystemBounded And Not bIsBoundaryUnused Then _
                Call PlateSystem_BoundByEdgeOfAPS(oPlateSystemCoplanar, oPlateSystemAdvanced, pPositionOfNode, pPlaneOfBoundary, oMember)
        End If
    Else
        Dim pPortOfEdge As IJPort: Set pPortOfEdge = PlateSystem_GetEdgePortFromBoundingSurface(oPlateSystemAdvanced, pPlaneOfBoundary)
  
        PlateSystem_AddAuxPortToFCOfMember pPortOfEdge, oPlateSystemAdvanced, pPositionOfNode, oMember, True
        
    End If
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Sub Member_BoundByOrUnBoundFromFaceOfAPS(oMember As Object, bBound As Boolean, oPlateSystemAdvanced As Object, pPositionOfNode As IJDPosition, Optional bGetExactSolution As Boolean = True)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Member_BoundByOrUnBoundFromFaceOfAPS"
    Call DebugIn("Member_BoundByOrUnBoundFromFaceOfAPS")
    Call DebugInput("GetExactSolution", bGetExactSolution)

    If TypeOf oMember Is ISPSDesignedMember Then
        ' process built-up
        
        Let sComment = "retrieve FrameConnection at node"
        Dim oConnection As Object:
        If bGetExactSolution Then
            Set oConnection = BuiltUp_RetrieveConnectionAtNode(oMember, pPositionOfNode)
        Else
            Set oConnection = BuiltUp_RetrieveConnectionAtNode2(oMember, pPositionOfNode)
        End If
        If oConnection Is Nothing Then Exit Sub
        
        ' loop on root plate systems
        Dim i As Integer
        Dim pElementsOfRootPlateSystems As IJElements: Set pElementsOfRootPlateSystems = Entity_GetElementsOfRelatedEntities(oMember, "IJDesignParent", "ShpStrDesignChild")
        For i = 1 To pElementsOfRootPlateSystems.Count
            Dim oRootPlateSystem As Object: Set oRootPlateSystem = pElementsOfRootPlateSystems(i)
            
            Let sComment = "retrieve existing boundaries"
            Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = PlateSystem_GetElementsOfBoundaries(oRootPlateSystem)
            
            If oPlateSystemAdvanced Is Nothing Then
                Let sComment = "APS has just been deleted, re-bound by the Connection"
                If Not pElementsOfBoundaries.Contains(oConnection) Then Call pElementsOfBoundaries.Add(oConnection)
            Else
                If bBound Then
                    Let sComment = "replace the Connection by the APS"
                    If pElementsOfBoundaries.Contains(oConnection) Then Call pElementsOfBoundaries.Remove(pElementsOfBoundaries.GetIndex(oConnection))
                    If Not pElementsOfBoundaries.Contains(oPlateSystemAdvanced) Then Call pElementsOfBoundaries.Add(oPlateSystemAdvanced)
                Else
                    Let sComment = "replace the APS by the Connection"
                    If pElementsOfBoundaries.Contains(oPlateSystemAdvanced) Then Call pElementsOfBoundaries.Remove(pElementsOfBoundaries.GetIndex(oPlateSystemAdvanced))
                    If Not pElementsOfBoundaries.Contains(oConnection) Then Call pElementsOfBoundaries.Add(oConnection)
                End If
            End If
            
            Let sComment = "update boundaries"
            Call PlateSystem_SetElementsOfBoundaries(oRootPlateSystem, pElementsOfBoundaries)
        Next

    Else
        Dim pFacePort As IJPort: Set pFacePort = PlateSystem_GetFacePortFromBoundingSurface(oMember, oPlateSystemAdvanced)
  
        PlateSystem_AddAuxPortToFCOfMember pFacePort, oPlateSystemAdvanced, pPositionOfNode, oMember, bBound
  
        
        ' MsgBox "Rolled members currently not processed"
    End If
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Function Member_IsBoundedBySurface(oMember As Object, oTrimmingSurface As Object) As Boolean
    Dim pElementsOfRootPlateSystems As IJElements: Set pElementsOfRootPlateSystems = Entity_GetElementsOfRelatedEntities(oMember, "IJDesignParent", "ShpStrDesignChild")
    Dim oRootPlateSystem As Object: Set oRootPlateSystem = pElementsOfRootPlateSystems(1)
    Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = PlateSystem_GetElementsOfBoundaries(oRootPlateSystem)
    Member_IsBoundedBySurface = pElementsOfBoundaries.Contains(oTrimmingSurface)
End Function
Sub Member_BoundByOrUnBoundFromSurface(oMember As Object, bBound As Boolean, oPlateSystemAdvanced As Object, oTrimmingSurface As Object, pPositionOfNode As IJDPosition, Optional bGetExactSolution As Boolean = True)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Member_BoundByOrUnBoundFromSurface"
    Call DebugIn("Member_BoundByOrUnBoundFromSurface")
    Call DebugInput("GetExactSolution", bGetExactSolution)
    
    If TypeOf oMember Is ISPSDesignedMember Then
        ' process built-up
        
        Let sComment = "retrieve FrameConnection at node"
        Dim oConnection As Object
        If bGetExactSolution Then
            Set oConnection = BuiltUp_RetrieveConnectionAtNode(oMember, pPositionOfNode)
        Else
            Set oConnection = BuiltUp_RetrieveConnectionAtNode2(oMember, pPositionOfNode)
        End If
        ' If oConnection Is Nothing Then Exit Sub
        
        ' loop on root plate systems
        Dim i As Integer
        Dim pElementsOfRootPlateSystems As IJElements: Set pElementsOfRootPlateSystems = Entity_GetElementsOfRelatedEntities(oMember, "IJDesignParent", "ShpStrDesignChild")
        For i = 1 To pElementsOfRootPlateSystems.Count
            Dim oRootPlateSystem As Object: Set oRootPlateSystem = pElementsOfRootPlateSystems(i)
            
            Let sComment = "retrieve existing boundaries"
            Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = PlateSystem_GetElementsOfBoundaries(oRootPlateSystem)
            
            If oPlateSystemAdvanced Is Nothing Then
                Let sComment = "APS has just been deleted, re-bound by the Connection"
                If Not oConnection Is Nothing Then
                    If Not pElementsOfBoundaries.Contains(oConnection) Then Call pElementsOfBoundaries.Add(oConnection)
                End If
            Else
                If bBound Then
                    Let sComment = "replace the Connection by the Surface"
                    If Not oConnection Is Nothing Then
                        If pElementsOfBoundaries.Contains(oConnection) Then Call pElementsOfBoundaries.Remove(pElementsOfBoundaries.GetIndex(oConnection))
                    End If
                    If Not pElementsOfBoundaries.Contains(oTrimmingSurface) Then Call pElementsOfBoundaries.Add(oTrimmingSurface)
                Else
                    Let sComment = "replace the Surface by the Connection"
                    If pElementsOfBoundaries.Contains(oTrimmingSurface) Then Call pElementsOfBoundaries.Remove(pElementsOfBoundaries.GetIndex(oTrimmingSurface))
                    If Not oConnection Is Nothing Then
                        If Not pElementsOfBoundaries.Contains(oConnection) Then Call pElementsOfBoundaries.Add(oConnection)
                    End If
                End If
            End If
            
            Let sComment = "update boundaries"
            Call PlateSystem_SetElementsOfBoundaries(oRootPlateSystem, pElementsOfBoundaries)
        Next

    Else
        ' MsgBox "Rolled members currently not processed"
    End If
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Sub Members_UnboundFromEdgeOfAPS(pElementsOfMembers As IJElements, iCountOfMembers As Integer, sKeysOfMembers() As String, _
                                 pElementsOfBoundaries As IJElements, iCountOfBoundaries As Integer, sKeysOfBoundaries() As String, _
                                 pGeometricConstruction As IJGeometricConstruction, iStepIndex As Integer, oAdvancedPlateSystem As Object, pPositionOfNode As IJDPosition)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Members_UnboundFromEdgeOfAPS"
    Call DebugIn("Members_UnboundFromEdgeOfAPS")
    
    ' build prefix for keys
    Dim sPrefixOfAxisPortKey As String: Let sPrefixOfAxisPortKey = "AxisPort" + CStr(iStepIndex)
    Dim sPrefixOfFacePortKey As String: Let sPrefixOfFacePortKey = "FacePort" + CStr(iStepIndex)
    Dim sPrefixOfBoundaryKey As String: Let sPrefixOfBoundaryKey = "TR" + CStr(iStepIndex)
    
'''    ' check for added members
'''    Dim i As Integer
'''    For i = 1 To iCountOfMembers
'''        If Not ArrayOfStrings_ContainsString(sKeysOfBoundaries, sPrefixOfBoundaryKey + "." + sKeysOfMembers(i)) Then
'''            ' MsgBox "Added Member #" + sKeysOfMembers(i)
'''        End If
'''    Next i
                
    ' check for removed members
    Dim i As Integer
    For i = 1 To iCountOfBoundaries
       Dim sKeyOfMember As String: Let sKeyOfMember = Mid(sKeysOfBoundaries(i), 5)
       If Mid(sKeysOfBoundaries(i), 1, 4) = sPrefixOfBoundaryKey + "." _
        And Not ArrayOfStrings_ContainsString(sKeysOfMembers, sKeyOfMember) Then
            ' MsgBox "Removed Member #" + sKeyOfMember
            Let sComment = "retrieve the Member"
            Dim oMember As Object
            If True Then
                ' check if the Member is still connected, case where the APS is deleting
                On Error Resume Next: Set oMember = Nothing: Set oMember = pElementsOfMembers(sKeyOfMember): On Error GoTo 0
                If oMember Is Nothing Then
                    ' the Member is removed
                    ' try to retrieve the FacePort corresponding to the removed member
                    Dim oFacePort As Object: On Error Resume Next: Set oFacePort = Nothing: Set oFacePort = pGeometricConstruction.ControlledInputs(sPrefixOfFacePortKey)(sKeyOfMember): On Error GoTo 0
                    
                    If oFacePort Is Nothing Then
                        ' if the FacePort is no more connected, then the member is being deleted and there is nothing to do
                    Else
                        ' retrieve the corresponding rolled or built-up Member
                        Set oMember = GetMemberFromPort(oFacePort)
                    End If
                End If
            End If
            
            If Not oMember Is Nothing Then
                If Not IsDeleted(oMember) Then
                    Let sComment = "un-bound the Member"
                    Call Member_BoundByOrUnBoundFromEdgeOfAPS(oMember, pGeometricConstruction, oAdvancedPlateSystem, Nothing, pElementsOfBoundaries(sKeysOfBoundaries(i)), pPositionOfNode)
                        
                    If Not oAdvancedPlateSystem Is Nothing Then
                        ' disconnect the AxisPort and the FacePort. If the Member is just removed, we do not want to maintain controlled inputs to him
                        Call pGeometricConstruction.ControlledInputs(sPrefixOfAxisPortKey).Remove(sKeyOfMember)
                        Call pGeometricConstruction.ControlledInputs(sPrefixOfFacePortKey).Remove(sKeyOfMember)
                    End If
                End If
            End If
        End If
    Next i
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Sub Members_UnboundFromFaceOfAPS(pElementsOfMembers As IJElements, iCountOfMembers As Integer, sKeysOfMembers() As String, _
                                 pElementsOfBoundaries As IJElements, iCountOfBoundaries As Integer, sKeysOfBoundaries() As String, _
                                 pGeometricConstruction As IJGeometricConstruction, iSmartStepIndex As Integer, oAdvancedPlateSystem As Object, pPositionOfNode As IJDPosition)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Members_UnboundFromFaceOfAPS"
    Call DebugIn("Members_UnboundFromFaceOfAPS")
    
    Dim sPrefixOfAxisPortKey As String: Let sPrefixOfAxisPortKey = "AxisPort" + CStr(iSmartStepIndex)
    
    ' check for removed members
    Dim i As Integer
    For i = 1 To iCountOfBoundaries
        Dim sKeyOfMember As String: Let sKeyOfMember = Mid(sKeysOfBoundaries(i), 5)
        If Not ArrayOfStrings_ContainsString(sKeysOfMembers, sKeyOfMember) Then
            'MsgBox "Removed secondary Member #" + sKeyOfMember
            Let sComment = "retrieve the Member"
            Dim oMember As Object
            If True Then
                ' check if the Member is still connected, case where the APS is deleting
                On Error Resume Next: Set oMember = Nothing: Set oMember = pElementsOfMembers(sKeyOfMember): On Error GoTo 0
                If oMember Is Nothing Then
                    ' the Member is removed
                    ' try to retrieve the AxisPort corresponding to the removed member
                    Dim pPortOfAxis As IJPort: On Error Resume Next: Set pPortOfAxis = Nothing: Set pPortOfAxis = pGeometricConstruction.ControlledInputs(sPrefixOfAxisPortKey)(sKeyOfMember): On Error GoTo 0
                    
                    If pPortOfAxis Is Nothing Then
                        ' if the FacePort is no more connected, then the member is being deleted and there is nothing to do
                    Else
                        ' retrieve the corresponding rolled or built-up Member
                        Set oMember = pPortOfAxis.Connectable
                    End If
                End If
            End If
            
            If Not oMember Is Nothing Then
                If Not IsDeleted(oMember) Then
                    Let sComment = "un-bound the Member"
                    Call Member_BoundByOrUnBoundFromFaceOfAPS(oMember, False, oAdvancedPlateSystem, pPositionOfNode)
                        
                    If Not oAdvancedPlateSystem Is Nothing Then
                        ' disconnect the AxisPort. If the Member is just removed, we do not want to maintain controlled inputs to him
                        Call pGeometricConstruction.ControlledInputs(sPrefixOfAxisPortKey).Remove(sKeyOfMember)
                    End If
                End If
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
Function BuiltUp_RetrieveConnectionAtNode(oBuiltUp As Object, pPositionOfNode As IJDPosition) As Object
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "BuiltUp_RetrieveFrameConnectionAtNode"
    Call DebugIn("BuiltUp_RetrieveConnectionAtNode")
    
    ' initialize result
    Dim oConnection As Object: Set oConnection = Nothing
    
    Let sComment = "retrieve the member system"
    Dim pMemberSystem As ISPSMemberSystem: Set pMemberSystem = MemberSystemReference_Get(oBuiltUp)
    
    Let sComment = "retrieve the frame connections"
'    On Error Resume Next
    If Not pMemberSystem Is Nothing Then
        Dim pFrameConnectionAtStart As ISPSFrameConnection: Set pFrameConnectionAtStart = pMemberSystem.FrameConnectionAtEnd(SPSMemberAxisStart)
        Dim pFrameConnectionAtEnd As ISPSFrameConnection: Set pFrameConnectionAtEnd = pMemberSystem.FrameConnectionAtEnd(SPSMemberAxisEnd)
        Dim pElementsOfSplitConnections As IJElements: Set pElementsOfSplitConnections = pMemberSystem.SplitConnections
    End If
'    On Error GoTo ErrorHandler
    
    Let sComment = "search the connections between the frame connections"
    If Not pFrameConnectionAtStart Is Nothing Then
        If Position_FromPoint(pFrameConnectionAtStart.Joint.Point).DistPt(pPositionOfNode) < EPSILON Then
            Set oConnection = pFrameConnectionAtStart
        Else
            If Not pFrameConnectionAtEnd Is Nothing Then
                If Position_FromPoint(pFrameConnectionAtEnd.Joint.Point).DistPt(pPositionOfNode) < EPSILON Then
                    Set oConnection = pFrameConnectionAtEnd
                End If
            End If
        End If
    End If
    
    Let sComment = "search the connections between the split connections"
    If oConnection Is Nothing Then
        If Not pElementsOfSplitConnections Is Nothing Then
            ' search the connection between the split connections
            Dim i As Integer
            For i = 1 To pElementsOfSplitConnections.Count
                Dim pSplitMemberConnection As ISPSSplitMemberConnection: Set pSplitMemberConnection = pElementsOfSplitConnections(i)
                If Position_FromPoint(pSplitMemberConnection).DistPt(pPositionOfNode) < EPSILON Then
                    Set oConnection = pSplitMemberConnection
                    Exit For
                End If
            Next
        End If
    End If
    
    ' return result
    Set BuiltUp_RetrieveConnectionAtNode = oConnection
    Call DebugOut
    Exit Function
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Function
Function BuiltUp_RetrieveConnectionAtNode2(oBuiltUp As Object, pPositionOfNode As IJDPosition) As Object
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "BuiltUp_RetrieveFrameConnectionAtNode"
    Call DebugIn("BuiltUp_RetrieveConnectionAtNode2")
    
    ' prepare result
    Dim oConnection As Object: Set oConnection = Nothing
    
    Let sComment = "retrieve the member system"
    Dim pMemberSystem As ISPSMemberSystem: Set pMemberSystem = MemberSystemReference_Get(oBuiltUp)
        
    Let sComment = "retrieve the connections"
    Dim pFrameConnectionAtStart As ISPSFrameConnection
    Dim pFrameConnectionAtEnd As ISPSFrameConnection
    Dim pElementsOfSplitConnections As IJElements
    On Error Resume Next
        Set pFrameConnectionAtStart = pMemberSystem.FrameConnectionAtEnd(SPSMemberAxisStart)
        Set pFrameConnectionAtEnd = pMemberSystem.FrameConnectionAtEnd(SPSMemberAxisEnd)
        Set pElementsOfSplitConnections = pMemberSystem.SplitConnections
    On Error GoTo ErrorHandler
    
    Let sComment = "find closest connection"
    Dim dDistanceMini As Double: Let dDistanceMini = 1000000#
    If True Then
        Dim dDistance As Double
        If Not pFrameConnectionAtStart Is Nothing Then
            Let dDistance = Position_FromPoint(pFrameConnectionAtStart.Joint.Point).DistPt(pPositionOfNode)
            Call DebugValue("Distance of FrameConnectionAtStart", dDistance)
            If dDistance < dDistanceMini Then
                Set oConnection = pFrameConnectionAtStart
                Let dDistanceMini = dDistance
            End If
        End If
        
        If Not pFrameConnectionAtEnd Is Nothing Then
            Let dDistance = Position_FromPoint(pFrameConnectionAtEnd.Joint.Point).DistPt(pPositionOfNode)
            Call DebugValue("Distance of FrameConnectionAtEnd", dDistance)
            If dDistance < dDistanceMini Then
                Set oConnection = pFrameConnectionAtEnd
                Let dDistanceMini = dDistance
            End If
        End If
                
        If Not pElementsOfSplitConnections Is Nothing Then
            ' search the connection between the split connections
            Dim i As Integer
            For i = 1 To pElementsOfSplitConnections.Count
                Dim pSplitMemberConnection As ISPSSplitMemberConnection: Set pSplitMemberConnection = pElementsOfSplitConnections(i)
                Let dDistance = Position_FromPoint(pSplitMemberConnection).DistPt(pPositionOfNode)
                Call DebugValue("Distance of SplitMemberConnection", dDistance)
                If dDistance < dDistanceMini Then
                    Set oConnection = pSplitMemberConnection
                    Let dDistanceMini = dDistance
                End If
            Next
        End If
    End If
    
    ' return result
    Set BuiltUp_RetrieveConnectionAtNode2 = oConnection
    Call DebugOutput("Connection", oConnection)
    Call DebugOut
    Exit Function
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Function
Public Function BuiltUp_GetPlateSystemCoplanar(oBuiltUp As Object, pPlaneOfReference As IJPlane) As Object
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "BuiltUp_GetPlateSystemCoplanar"
    Call DebugIn("BuiltUp_GetPlateSystemCoplanar")
    Call DebugInput("BuiltUp", oBuiltUp)
    
    ' prepare result
    Dim oPlateSystemCoplanar As Object
    
    Let sComment = "retrieve the list of plate systems of the built-up"
    Dim pElementsOfRootPlateSystems As IJElements: Set pElementsOfRootPlateSystems = Entity_GetElementsOfRelatedEntities(oBuiltUp, "IJDesignParent", "ShpStrDesignChild")
    
    ' loop on the plate systems
    Dim iPlate As Integer
    For iPlate = 1 To pElementsOfRootPlateSystems.Count
        ' retrieve the current plate
        Dim oPlateSystem As Object: Set oPlateSystem = pElementsOfRootPlateSystems(iPlate)
        
        Let sComment = "retrieve the bound AE"
        Dim oBoundAE As Object: Set oBoundAE = PlateSystem_GetBoundAE(oPlateSystem)
        
        Let sComment = "retrieve the unbounded geometry"
        Dim oUnboundedGeometry As Object: Set oUnboundedGeometry = Entity_GetElementsOfRelatedEntities(oBoundAE, "IJStructOperation", "StructOperation_OPRND_ORIG")(1)
        
        ' retrieve its support
        ' Dim pPlaneOfPlateSystem As IJPlane: Set pPlaneOfPlateSystem = Entity_GetElementsOfRelatedEntities(oPlateSystem, "IJGeometry", "StructToGeometry_DEST")(1)
        
        Let sComment = "only process a web/flange coplanar with the APS or which overlap when no planar"
        If TypeOf oUnboundedGeometry Is IJPlane Then
            If ArePlanesCoplanar(oUnboundedGeometry, pPlaneOfReference) Then
                Set oPlateSystemCoplanar = oPlateSystem
                Exit For
            End If
        Else
            If AreSurfacesOverlapping(oUnboundedGeometry, pPlaneOfReference) Then
                Set oPlateSystemCoplanar = oPlateSystem
                Exit For
            End If
        End If
    Next
    
    ' return result
    Set BuiltUp_GetPlateSystemCoplanar = oPlateSystemCoplanar
    
    Call DebugOutput("PlateSystemCoplanar", BuiltUp_GetPlateSystemCoplanar)
    Call DebugOut
    Exit Function
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Function
Public Function BuiltUp_GetPlateSystem(oBuiltUp As Object) As Object
    Call DebugIn("BuiltUp_GetPlateSystem")
    
    ' retrieve the list of plate systems of the built-up
    Dim pElementsOfRootPlateSystems As IJElements: Set pElementsOfRootPlateSystems = Entity_GetElementsOfRelatedEntities(oBuiltUp, "IJDesignParent", "ShpStrDesignChild")
    
    ' return result
    Set BuiltUp_GetPlateSystem = pElementsOfRootPlateSystems(1)
    Call DebugOutput("PlateSystem", BuiltUp_GetPlateSystem)
    Call DebugOut
End Function
Public Function BuiltUp_GetPlateSystemBounded(oBuiltUp As Object, oPlateSystemOfBoundary As Object, pPlaneOfBoundary As IJPlane) As Object
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "BuiltUp_GetPlateSystemBounded"
    Call DebugIn("BuiltUp_GetPlateSystemBounded")
    Call DebugInput("BuiltUp", oBuiltUp)
    Call DebugInput("PlateSystemOfBoundary", oPlateSystemOfBoundary)
    
    ' prepare result
    Dim oPlateSystemBounded As Object
    
    ' retrieve the list of plate systems of the built-up
    Dim pElementsOfRootPlateSystems As IJElements: Set pElementsOfRootPlateSystems = Entity_GetElementsOfRelatedEntities(oBuiltUp, "IJDesignParent", "ShpStrDesignChild")
    
    ' retrieve the edge port of the APS corresponding to the bounding surface
'''    Dim pPortOfEdge As IJPort: Set pPortOfEdge = PlateSystem_GetEdgePortFromBoundingSurface(oPlateSystemOfBoundary, pPlaneOfBoundary)

    ' loop on the plate systems
    Dim iPlate As Integer
    For iPlate = 1 To pElementsOfRootPlateSystems.Count
        Let sComment = "retrieve the current plate"
        Dim oPlateSystem As Object: Set oPlateSystem = pElementsOfRootPlateSystems(iPlate)
        
        Let sComment = "retrieve its boundaries"
        Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = PlateSystem_GetElementsOfBoundaries(oPlateSystem)
        
        ' loop on boundaries
        Dim iBoundary As Integer
        For iBoundary = 1 To pElementsOfBoundaries.Count
'''            If pElementsOfBoundaries(iBoundary) Is pPortOfEdge Then
'''                Set oPlateSystemBounded = oPlateSystem
'''                Exit For
'''            End If
            If TypeOf pElementsOfBoundaries(iBoundary) Is IJPort Then
                Dim pPort As IJPort: Set pPort = pElementsOfBoundaries(iBoundary)
                If pPort.Connectable Is oPlateSystemOfBoundary Then
                    Set oPlateSystemBounded = oPlateSystem
                    Exit For
                End If
            End If
        Next
        If Not oPlateSystemBounded Is Nothing Then Exit For
    Next
    
    ' return result
    Set BuiltUp_GetPlateSystemBounded = oPlateSystemBounded
    
    Call DebugOutput("PlateSystemBounded", BuiltUp_GetPlateSystemBounded)
    Call DebugOut
    Exit Function
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Function
Private Sub BuiltUp_BoundPlateSystemsByConnection(oBuiltUp As Object, oConnection As Object, pGeometricConstruction As IJGeometricConstruction)
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "BuiltUp_BoundPlateSystemsByFrameConnection"
    Call DebugIn("BuiltUp_BoundPlateSystemsByConnection")
    
    ' retrieve the list of plate systems of the built-up
    Dim pElementsOfRootPlateSystems As IJElements: Set pElementsOfRootPlateSystems = Entity_GetElementsOfRelatedEntities(oBuiltUp, "IJDesignParent", "ShpStrDesignChild")
    
    ' loop on the plate systems
    Dim iPlate As Integer
    For iPlate = 1 To pElementsOfRootPlateSystems.Count
        Let sComment = "retrieve the current plate"
        Dim oPlateSystem As Object: Set oPlateSystem = pElementsOfRootPlateSystems(iPlate)
        
        Dim bIsUnboundedPlateSystem As Boolean: Let bIsUnboundedPlateSystem = True
        If Not pGeometricConstruction Is Nothing Then
            bIsUnboundedPlateSystem = pGeometricConstruction.ControlledInputs("UnBoundedPlateSystems").Contains(oPlateSystem)
        End If
        
        If bIsUnboundedPlateSystem Then
            Let sComment = "retrieve its boundaries"
            Dim pElementsOfBoundaries As IJElements: Set pElementsOfBoundaries = PlateSystem_GetElementsOfBoundaries(oPlateSystem)
            If Not pElementsOfBoundaries.Contains(oConnection) Then
                Call pElementsOfBoundaries.Add(oConnection)
                Call PlateSystem_SetElementsOfBoundaries(oPlateSystem, pElementsOfBoundaries)
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
Private Sub MemberPart_TrimAtNode(oMemberPart As Object, oSurface As Object, pPositionOfNode As IJDPosition)
    Call DebugIn("MemberPart_TrimAtNode")
    
    Dim pMemberPartPrismatic As ISPSMemberPartPrismatic: Set pMemberPartPrismatic = oMemberPart
    
    Dim i As Integer
    For i = SPSMemberAxisStart To SPSMemberAxisEnd
'        If Not pMemberPartPrismatic.Cutbacks(i).Contains(oSurface) Then
'            Dim pPositionOfFrameConnection As IJDPosition:
'            Set pPositionOfFrameConnection = Position_FromPoint(pMemberPartPrismatic.PointAtEnd(i))
'
'            If pPositionOfFrameConnection.DistPt(pPositionOfNode) < EPSILON Then
'                Call pMemberPartPrismatic.AddCutbackSurface(i, oSurface)
'                Let bIsFrameConnectionFound = True
'                Exit For
'            End If
'        End If
        If pMemberPartPrismatic.Cutbacks(i).Count = 0 Then
            Dim pPositionOfFrameConnection As IJDPosition:
            Set pPositionOfFrameConnection = Position_FromPoint(pMemberPartPrismatic.PointAtEnd(i))
               
            If pPositionOfFrameConnection.DistPt(pPositionOfNode) < EPSILON Then
                Call pMemberPartPrismatic.AddCutbackSurface(i, oSurface)
                Exit For
            End If
        End If

    Next
    
    Call DebugOut
End Sub
Private Sub MemberPart_UnTrimAtNode(oMemberPart As Object, oSurface As Object, pPositionOfNode As IJDPosition)
End Sub
Public Function Member_GetLine(oMemberPart As Object, Optional pPOM As IJDPOM = Nothing) As IJLine
    Set Member_GetLine = Nothing
    
    Dim pMemberSystem As ISPSMemberSystem
    Set pMemberSystem = MemberSystemReference_Get(oMemberPart)
    If Not pMemberSystem Is Nothing Then
        Dim pLine As IJLine: Set pLine = pMemberSystem.LogicalAxis.CurveGeometry
        If Not pLine Is Nothing Then
            Set Member_GetLine = Line_FromPositions(pPOM, Position_FromLine(pLine, 0), Position_FromLine(pLine, 1))
        Else
            'MsgBox "Member_GetLine: pLine Is Nothing"
        End If
    Else
        'MsgBox "Member_GetLine: pMemberSystem Is Nothing"
    End If
    
End Function
Public Function MemberPart_GetLine(oMemberPart As Object, Optional pPOM As IJDPOM = Nothing) As IJLine
    Dim pMemberPartCommon As ISPSMemberPartCommon: Set pMemberPartCommon = oMemberPart
    Dim pPointAtStart As IJPoint: Set pPointAtStart = pMemberPartCommon.PointAtEnd(SPSMemberAxisStart)
    Dim pPointAtEnd As IJPoint: Set pPointAtEnd = pMemberPartCommon.PointAtEnd(SPSMemberAxisEnd)
    Set MemberPart_GetLine = Line_FromPositions(pPOM, Position_FromPoint(pPointAtStart), Position_FromPoint(pPointAtEnd))
End Function
Public Function MemberParts_GetLinesOfMemberAxesAtCommonNode(pElementsOfMemberParts As IJElements, pPositionOfCommonNode As IJDPosition) As IJLine()
    Call DebugIn("MemberParts_GetLinesOfMemberAxesAtCommonNode")
    
    ' initialise results
    Dim pLinesOfMemberAxes() As IJLine
    ReDim pLinesOfMemberAxes(1 To pElementsOfMemberParts.Count)
    
    Dim i As Integer
    For i = 1 To pElementsOfMemberParts.Count
        ' retrieve common interface
'''        Dim pMemberPartCommon As ISPSMemberPartCommon: Set pMemberPartCommon = pElementsOfMemberParts(i)

        Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(i))
        Dim i2 As Integer
        For i2 = 0 To 3
            'MsgBox ("i= " + CStr(i) + ", i2= " + CStr(i2))
            If Not pPositionsAtEndsAndJointsOfMember(i2) Is Nothing Then
                If pPositionOfCommonNode.DistPt(pPositionsAtEndsAndJointsOfMember(i2)) < EPSILON Then
                    If i2 = 0 Or i2 = 1 Then
                        ' retrieve logical axis
                        Set pLinesOfMemberAxes(i) = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(0), pPositionsAtEndsAndJointsOfMember(1))
                        'MsgBox ("logical: i= " + CStr(i) + ", i2= " + CStr(i2))
                    Else
                        ' retrieve physical axis
                        Set pLinesOfMemberAxes(i) = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(2), pPositionsAtEndsAndJointsOfMember(3))
                        'MsgBox ("physical: i= " + CStr(i) + ", i2= " + CStr(i2))
                    End If
                    Exit For
                End If
            End If
        Next
        If pLinesOfMemberAxes(i) Is Nothing Then
            'MsgBox "member axis not found"
            Err.Raise E_FAIL
        End If
    Next
    
    ' return result
    Let MemberParts_GetLinesOfMemberAxesAtCommonNode = pLinesOfMemberAxes
    Call DebugOut
End Function
Public Function MemberParts_GetPositionOfCommonNode(pElementsOfMemberParts As IJElements) As IJDPosition
    ' TR-CP-234669 : try to find a common node by looking at the logical ends first.
    ' If no result try by looking at the physical ends only
    Dim pPositionOfNode As IJDPosition: Set pPositionOfNode = MemberParts_GetPositionOfCommonNodeEx(pElementsOfMemberParts, 0)
    If pPositionOfNode Is Nothing Then Set pPositionOfNode = MemberParts_GetPositionOfCommonNodeEx(pElementsOfMemberParts, 2)
    ' return result
    Set MemberParts_GetPositionOfCommonNode = pPositionOfNode
End Function
Public Function MemberParts_GetPositionOfCommonNodeEx(pElementsOfMemberParts As IJElements, iStart As Integer) As IJDPosition
    Call DebugIn("MemberParts_GetPositionOfCommonNodeEx")
    
    ' prepare the result
    Dim pPositionOfCommonNode As IJDPosition: Set pPositionOfCommonNode = Nothing
    
    ' get the positions at ends and joints of the 2 first MemberParts
    Dim pPositionsAtEndsAndJointsOfMember1() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember1 = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(1))
    Dim pPositionsAtEndsAndJointsOfMember2() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember2 = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(2))
    
    ' find a common position between the positions extracted from the 2 first MemberParts
    Dim bCommonNodeIsFound As Boolean: Let bCommonNodeIsFound = False
    Dim i1 As Integer: Dim i2 As Integer
    For i1 = iStart To 3
        For i2 = iStart To 3
            'MsgBox ("i1= " + CStr(i1) + ", i2= " + CStr(i2))
            If Not pPositionsAtEndsAndJointsOfMember1(i1) Is Nothing Then
                If Not pPositionsAtEndsAndJointsOfMember2(i2) Is Nothing Then
                    If pPositionsAtEndsAndJointsOfMember1(i1).DistPt(pPositionsAtEndsAndJointsOfMember2(i2)) < EPSILON Then
                        Set pPositionOfCommonNode = pPositionsAtEndsAndJointsOfMember1(i1)
                        Let bCommonNodeIsFound = True
                        Exit For
                    End If
                End If
            End If
        Next
        If bCommonNodeIsFound Then Exit For
    Next
    
    If bCommonNodeIsFound And pElementsOfMemberParts.Count > 2 Then
        ' verify if we can extract from each remaining MemberPart a position common with the already found common node
        Let bCommonNodeIsFound = False
        Dim i As Integer
        For i = 3 To pElementsOfMemberParts.Count
            Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(i))
            For i2 = iStart To 3
                'MsgBox ("i2= " + CStr(i2))
                If Not pPositionsAtEndsAndJointsOfMember(i2) Is Nothing Then
                    If pPositionOfCommonNode.DistPt(pPositionsAtEndsAndJointsOfMember(i2)) < EPSILON Then
                        Let bCommonNodeIsFound = True
                        Exit For
                    End If
                End If
            Next
            If Not bCommonNodeIsFound Then Exit For
        Next
    End If
     
    
    If Not bCommonNodeIsFound Then
        ' clear the eventually already found common node
        Set pPositionOfCommonNode = Nothing
    End If
    
    ' return result
    Set MemberParts_GetPositionOfCommonNodeEx = pPositionOfCommonNode
    Call DebugOutput("PositionOfCommonNode", MemberParts_GetPositionOfCommonNodeEx)
    Call DebugOut
End Function
Public Function MemberParts_GetPositionOfCommonNodeOnCurve(pElementsOfMemberParts As IJElements, pCurve As IJCurve) As IJDPosition
    Call DebugIn("MemberParts_GetPositionOfCommonNodeOnColumn")
    Call DebugValue("Start position of curve", Position_FromCurve(pCurve, 0))
    Call DebugValue("End   position of curve", Position_FromCurve(pCurve, 1))
    
    ' prepare the result
    Dim pPositionOfCommonNode As IJDPosition: Set pPositionOfCommonNode = Nothing
    
    ' get the positions at ends and joints of the first MemberPart
    Dim pPositionsAtEndsAndJointsOfMember1() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember1 = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(1))
    
    ' find a common position between the positions extracted from the first MemberPart and the Column
    Dim bCommonNodeIsFound As Boolean: Let bCommonNodeIsFound = False
    Dim i1 As Integer: Dim i2 As Integer
    For i1 = 0 To 3
        If Not pPositionsAtEndsAndJointsOfMember1(i1) Is Nothing Then
            If IsPositionOnCurve(pPositionsAtEndsAndJointsOfMember1(i1), pCurve) Then
                Set pPositionOfCommonNode = pPositionsAtEndsAndJointsOfMember1(i1)
                Let bCommonNodeIsFound = True
                Exit For
            End If
        End If
    Next
    
    If bCommonNodeIsFound And pElementsOfMemberParts.Count > 1 Then
        ' verify if we can extract from each remaining MemberPart a position common with the already found common node
        Let bCommonNodeIsFound = False
        Dim i As Integer
        For i = 2 To pElementsOfMemberParts.Count
            Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(i))
            For i2 = 0 To 3
                'MsgBox ("i2= " + CStr(i2))
                If Not pPositionsAtEndsAndJointsOfMember(i2) Is Nothing Then
                    If pPositionOfCommonNode.DistPt(pPositionsAtEndsAndJointsOfMember(i2)) < EPSILON Then
                        Let bCommonNodeIsFound = True
                        Exit For
                    End If
                End If
            Next
            If Not bCommonNodeIsFound Then Exit For
        Next
    End If
     
    
    If Not bCommonNodeIsFound Then
        ' clear the eventually already found common node
        Set pPositionOfCommonNode = Nothing
    End If
    
    ' return result
    Set MemberParts_GetPositionOfCommonNodeOnCurve = pPositionOfCommonNode
    Call DebugOutput("PositionOfCommonNode", MemberParts_GetPositionOfCommonNodeOnCurve)
    Call DebugOut
End Function
Public Function MemberPart_GetPositionsAtJointsAndEnds(oProfile As Object) As IJDPosition()
    Call DebugIn("MemberPart_GetPositionsAtJointsAndEnds")
    Call DebugInput("MemberPart", oProfile)
    
    ' prepare result
    Dim pPositions(0 To 3) As IJDPosition
    
    ' get physical ends
    If TypeOf oProfile Is IJStiffenerSystem Then
'        Dim pDesignChild As IJDesignChild: Set pDesignChild = oProfile
'        Dim oLeafProfileSystem As Object: Set oLeafProfileSystem = pDesignChild.GetParent()
'        Dim pProfileAttributes As IJProfileAttributes: Set pProfileAttributes = New ProfileUtils
'        Dim oLandingCurve As Object: Call pProfileAttributes.GetLandingCurveFromProfile(oLeafProfileSystem, oLandingCurve)
        Dim pProfileAttributes As IJProfileAttributes: Set pProfileAttributes = New ProfileUtils
        Dim oLandingCurve As Object: Call pProfileAttributes.GetRootLandingCurveFromProfile(oProfile, oLandingCurve)
        Dim pWireBody As IJWireBody: Set pWireBody = oLandingCurve
        Set pPositions(2) = Position_FromWireBody(pWireBody, 0)
        Set pPositions(3) = Position_FromWireBody(pWireBody, 1)
    Else
        If TypeOf oProfile Is ISPSMemberSystem Then
            ' retrieve the 2 extreme parts
'            Dim pMemberPartCommonAtStart As ISPSMemberPartCommon:
'            Dim pMemberPartCommonAtEnd As ISPSMemberPartCommon:
'            If True Then
'                Dim pMemberSystem As ISPSMemberSystem: Set pMemberSystem = oProfile
'                Set pMemberPartCommonAtStart = pMemberSystem.MemberPartAtEnd(SPSMemberAxisStart)
'                Set pMemberPartCommonAtEnd = pMemberSystem.MemberPartAtEnd(SPSMemberAxisEnd)
'            End If
'            Set pPositions(2) = Position_FromLine(pMemberPartCommonAtStart.Axis, 0)
'            Set pPositions(3) = Position_FromLine(pMemberPartCommonAtEnd.Axis, 1)
             Set pPositions(2) = Position_FromLine(oProfile, 0)
             Set pPositions(3) = Position_FromLine(oProfile, 1)
        Else ' ISPSMemberPartPrismatic or ISPSDesignedMember
            ' retrieve common interface
            Dim pMemberPartCommon As ISPSMemberPartCommon: Set pMemberPartCommon = oProfile
                
            ' retrieve physical axis
            Dim pLineOfPhysicalAxis As IJLine: Set pLineOfPhysicalAxis = pMemberPartCommon.Axis
            
            Set pPositions(2) = Position_FromLine(pLineOfPhysicalAxis, 0)
            Set pPositions(3) = Position_FromLine(pLineOfPhysicalAxis, 1)
        End If
    End If
    
    ' get logical ends
    If TypeOf oProfile Is IJStiffenerSystem Then
        Set pPositions(0) = pPositions(2)
        Set pPositions(1) = pPositions(3)
    Else
        Dim pMemberSystem As ISPSMemberSystem
        If TypeOf oProfile Is ISPSMemberSystem Then
            Set pMemberSystem = oProfile
        Else
            Set pMemberSystem = pMemberPartCommon.MemberSystem
        End If
        
        'DM-CP-213059 Member cannot be split if it has an APS plate on it
        'Add a test on pMemberSystem that can be Nothing during the migration process.
        If Not pMemberSystem Is Nothing Then
            Dim pFrameConnection0 As ISPSFrameConnection: Set pFrameConnection0 = pMemberSystem.FrameConnectionAtEnd(SPSMemberAxisStart)
            Dim pJoint0 As ISPSAxisJoint: Set pJoint0 = pFrameConnection0.Joint
            Set pPositions(0) = Position_FromPoint(pJoint0.Point)
                    
            Dim pFrameConnection1 As ISPSFrameConnection: Set pFrameConnection1 = pMemberSystem.FrameConnectionAtEnd(SPSMemberAxisEnd)
            Dim pJoint1 As ISPSAxisJoint: Set pJoint1 = pFrameConnection1.Joint
            Set pPositions(1) = Position_FromPoint(pJoint1.Point)
        End If
    End If
    
    ' return result
    Let MemberPart_GetPositionsAtJointsAndEnds = pPositions
    
    If Not pPositions(0) Is Nothing Then Call DebugOutput("Position#0", Position_ToString(pPositions(0)))
    If Not pPositions(1) Is Nothing Then Call DebugOutput("Position#1", Position_ToString(pPositions(1)))
    If Not pPositions(2) Is Nothing Then Call DebugOutput("Position#2", Position_ToString(pPositions(2)))
    If Not pPositions(3) Is Nothing Then Call DebugOutput("Position#3", Position_ToString(pPositions(3)))
    Call DebugOut
End Function
Public Sub Member_UnBoundFromEdgeOfAPS(oMember As Object, pGeometricConstructionMacro As IJGeometricConstructionMacro, oPlateSystemAdvanced As Object, pPositionOfNode As IJDPosition)
    Call DebugIn("Member_UnBoundFromEdgeOfAPS")
    Call DebugInput("Member", oMember)
    Call DebugInput("PlateSystemAdvanced", oPlateSystemAdvanced)
    
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Member_BoundByOrUnBoundFromEdgeOfAPS"
    
    If TypeOf oMember Is ISPSDesignedMember Then
        ' process built-up
        If oPlateSystemAdvanced Is Nothing Then
            Let sComment = "APS has just been deleted, bounding edges have been removed, but bounding FrameConnections have not necessarily been restored"
            Dim oFrameConnection As Object: Set oFrameConnection = BuiltUp_RetrieveConnectionAtNode(oMember, pPositionOfNode)
            If Not oFrameConnection Is Nothing Then Call BuiltUp_BoundPlateSystemsByConnection(oMember, oFrameConnection, pGeometricConstructionMacro)
        Else
            Let sComment = "retrieve web/flange already bounded by the APS"
            Dim oPlateSystemBounded As Object:  Set oPlateSystemBounded = BuiltUp_GetPlateSystemBounded(oMember, oPlateSystemAdvanced, Nothing)
                    
            Let sComment = "un-bound plate system, which should not be bounded any more"
            Call PlateSystem_UnBoundFromEdgeOfAPS(oPlateSystemBounded, oPlateSystemAdvanced, pPositionOfNode, oMember, pGeometricConstructionMacro)
            
        End If
    Else
        ' MsgBox "Rolled members currently not processed"
    End If
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Public Sub Member_UnBoundFromSurface(oMember As Object, pGeometricConstructionMacro As IJGeometricConstructionMacro, oPlateSystemAdvanced As Object, pPositionOfNode As IJDPosition)
    Call DebugIn("Member_UnBoundFromSurface")
    Call DebugInput("Member", oMember)
    Call DebugInput("PlateSystemAdvanced", oPlateSystemAdvanced)
    
    On Error GoTo ErrorHandler: Dim sComment As String: Dim sFunction As String: Let sFunction = "Member_BoundByOrUnBoundFromEdgeOfAPS"
    
    If TypeOf oMember Is ISPSDesignedMember Then
        ' process built-up
        If oPlateSystemAdvanced Is Nothing Then
            'MsgBox "APS is nothing"
            Let sComment = "APS has just been deleted, bounding edges have been removed, but bounding FrameConnections have not necessarily been restored"
            Dim oFrameConnection As Object: Set oFrameConnection = BuiltUp_RetrieveConnectionAtNode(oMember, pPositionOfNode)
            If Not oFrameConnection Is Nothing Then Call BuiltUp_BoundPlateSystemsByConnection(oMember, oFrameConnection, Nothing)
        Else
            'MsgBox "APS still exists"
            Let sComment = "APS has just been deleted, bounding edges have been removed, but bounding FrameConnections have not necessarily been restored"
            Set oFrameConnection = BuiltUp_RetrieveConnectionAtNode(oMember, pPositionOfNode)
            If Not oFrameConnection Is Nothing Then Call BuiltUp_BoundPlateSystemsByConnection(oMember, oFrameConnection, Nothing)
'''            Let sComment = "retrieve web/flange already bounded by the APS"
'''            Dim oPlateSystemBounded As Object:  Set oPlateSystemBounded = BuiltUp_GetPlateSystemBounded(oMember, oPlateSystemAdvanced, Nothing)
'''
'''            Let sComment = "un-bound plate system, which should not be bounded any more"
'''            Call PlateSystem_UnBoundFromEdgeOfAPS(oPlateSystemBounded, oPlateSystemAdvanced, pPositionOfNode, oMember, pGeometricConstructionMacro)
        End If
    Else
        ' MsgBox "Rolled members currently not processed"
    End If
    
    Call DebugOut
    Exit Sub
ErrorHandler:
    Dim sContext As String: Let sContext = m_sModule + "::" + sFunction
    'Call ShowError(sContext, sComment, Err.Number, Err.Description)
    Err.Raise Err.Number, sContext, sComment
End Sub
Function Members_GetLinesOfMemberSystemAxes(pElementsOfMembers As IJElements) As IJLine()
    ' initialise results
    Dim pLinesOfMemberAxes() As IJLine
    
    ReDim pLinesOfMemberAxes(1 To pElementsOfMembers.Count)
    Dim iCount As Integer: iCount = 1
    Dim i As Integer
    For i = 1 To pElementsOfMembers.Count
        Dim pLine As IJLine
        Set pLine = Member_GetLine(pElementsOfMembers(i))
        If Not pLine Is Nothing Then
            Set pLinesOfMemberAxes(iCount) = pLine
            iCount = iCount + 1
        End If
    Next
    ReDim Preserve pLinesOfMemberAxes(1 To iCount - 1)
    
    ' return result
    Let Members_GetLinesOfMemberSystemAxes = pLinesOfMemberAxes
End Function
Function Members_GetLinesOfMemberPartAxes(pElementsOfMembers As IJElements) As IJLine()
    ' initialise results
    Dim pLinesOfMemberAxes() As IJLine
    
    ReDim pLinesOfMemberAxes(1 To pElementsOfMembers.Count)
    Dim i As Integer
    For i = 1 To pElementsOfMembers.Count
        Set pLinesOfMemberAxes(i) = MemberPart_GetLine(pElementsOfMembers(i))
    Next
    
    ' return result
    Let Members_GetLinesOfMemberPartAxes = pLinesOfMemberAxes
End Function
Public Function MemberPart_GetLineOfMemberAxisAtNode(oMemberPart As Object, pPositionOfNode As IJDPosition) As IJLine
    Call DebugIn("MemberPart_GetLineOfMemberAxisAtNode")
    Call DebugInput("MemberPart", oMemberPart)
    Call DebugInput("PositionOfNode", pPositionOfNode)
    
    ' initialise results
    Dim pLineOfMemberAxis As IJLine

    Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(oMemberPart)
    Dim i2 As Integer
    For i2 = 0 To 3
        If Not pPositionsAtEndsAndJointsOfMember(i2) Is Nothing Then
            If pPositionOfNode.DistPt(pPositionsAtEndsAndJointsOfMember(i2)) < EPSILON Then
                If i2 = 0 Or i2 = 1 Then
                    DebugMsg ("Logical axis retrieved")
                    If pPositionOfNode.DistPt(pPositionsAtEndsAndJointsOfMember(0)) < EPSILON Then
                        Set pLineOfMemberAxis = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(0), pPositionsAtEndsAndJointsOfMember(1))
                    Else
                        Set pLineOfMemberAxis = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(1), pPositionsAtEndsAndJointsOfMember(0))
                    End If
                Else
                    DebugMsg ("Physical axis retrieved")
                    If pPositionOfNode.DistPt(pPositionsAtEndsAndJointsOfMember(2)) < EPSILON Then
                        Set pLineOfMemberAxis = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(2), pPositionsAtEndsAndJointsOfMember(3))
                    Else
                        Set pLineOfMemberAxis = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(3), pPositionsAtEndsAndJointsOfMember(2))
                    End If
                End If
                DebugMsg ("Line found")
                Exit For
            End If
        End If
    Next
    
    If pLineOfMemberAxis Is Nothing Then
        'MsgBox "member axis not found"
        Err.Raise E_FAIL
    End If
    
    ' return result
    Set MemberPart_GetLineOfMemberAxisAtNode = pLineOfMemberAxis
    
    Call DebugOut
End Function
Public Function MemberPart_GetLineOfMemberAxisThroughNode(oMemberPart As Object, pPositionOfNode As IJDPosition) As IJLine
    Call DebugIn("MemberPart_GetLineOfMemberAxisThroughNode")
    Call DebugInput("MemberPart", oMemberPart)
    Call DebugInput("PositionOfNode", pPositionOfNode)
    
    ' initialise result
    Dim pLineOfMemberAxis As IJLine

    Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(oMemberPart)
    Dim i2 As Integer
    For i2 = 0 To 3 Step 2
        If Not pPositionsAtEndsAndJointsOfMember(i2) Is Nothing Then
            Dim pLineForProjectionOn As IJLine
            If True Then
                If i2 = 0 Then
                    DebugMsg ("Logical axis re-constructed")
                    If pPositionOfNode.DistPt(pPositionsAtEndsAndJointsOfMember(0)) < pPositionOfNode.DistPt(pPositionsAtEndsAndJointsOfMember(1)) Then
                        Set pLineForProjectionOn = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(0), pPositionsAtEndsAndJointsOfMember(1))
                    Else
                        Set pLineForProjectionOn = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(1), pPositionsAtEndsAndJointsOfMember(0))
                    End If
                Else
                    DebugMsg ("Physical axis re-constructed")
                    If pPositionOfNode.DistPt(pPositionsAtEndsAndJointsOfMember(2)) < pPositionOfNode.DistPt(pPositionsAtEndsAndJointsOfMember(3)) Then
                        Set pLineForProjectionOn = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(2), pPositionsAtEndsAndJointsOfMember(3))
                    Else
                        Set pLineForProjectionOn = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfMember(3), pPositionsAtEndsAndJointsOfMember(2))
                    End If
                End If
            End If
            
            Dim pPositionOfProjectedPoint As IJDPosition: Set pPositionOfProjectedPoint = Position_ProjectOnLine(pPositionOfNode, pLineForProjectionOn)
            If pPositionOfProjectedPoint.DistPt(pPositionOfNode) < EPSILON Then
                DebugMsg ("Line found")
                Set pLineOfMemberAxis = pLineForProjectionOn
                Exit For
            End If
        End If
    Next
    
    If pLineOfMemberAxis Is Nothing Then
        'MsgBox "member axis not found"
        Err.Raise E_FAIL
    End If
    
    ' return result
    Set MemberPart_GetLineOfMemberAxisThroughNode = pLineOfMemberAxis
    
    Call DebugOut
End Function
Public Function MemberParts_GetElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode(pElementsOfMemberParts As IJElements, pPositionOfNode As IJDPosition) As IJElements
    Call DebugIn("MemberParts_GetElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode")
    Call DebugInput("Count of MemberParts", pElementsOfMemberParts.Count)
    
    ' prepare result
    Dim pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode As IJElements: Set pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode = New JObjectCollection
    
    ' collect information about Members
    Dim pArrayOfPositionsAtEndsAndJointsOfMembers() As IJDPosition: ReDim pArrayOfPositionsAtEndsAndJointsOfMembers(1 To pElementsOfMemberParts.Count, 0 To 3)
    Dim i As Integer
    For i = 1 To pElementsOfMemberParts.Count
        Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(i))
        Dim j As Integer
        For j = 0 To 3
            Set pArrayOfPositionsAtEndsAndJointsOfMembers(i, j) = pPositionsAtEndsAndJointsOfMember(j)
        Next
    Next
    
    ' find the first 2 Members which have a common end or are in a configuration of supporting/supported or are intersecting
    Dim pPositionOfCommonNode As IJDPosition: Set pPositionOfCommonNode = Nothing
    If Not pPositionOfNode Is Nothing Then
        Set pPositionOfCommonNode = pPositionOfNode
    Else
        For i = 1 To pElementsOfMemberParts.Count - 1
            For j = i + 1 To pElementsOfMemberParts.Count
                Call DebugMsg("Processing Member#" + CStr(i) + " with Member#" + CStr(j))
                Dim i1 As Integer: Dim j1 As Integer
                For i1 = 2 To 3
                    For j1 = 2 To 3
                        If Not pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1) Is Nothing Then
                            If Not pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1) Is Nothing Then
                                ' test if 2 ends are common
                                If pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1).DistPt(pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1)) < EPSILON Then
                                    ' verify that the 2 lines are not colinear
                                    Dim i2 As Integer: If i1 = 0 Or i1 = 2 Then i2 = i1 + 1 Else i2 = i1 - 1
                                    Dim j2 As Integer: If j1 = 0 Or j1 = 2 Then j2 = j1 + 1 Else j2 = j1 - 1
                                    If Not AreLinesCoLinear(Line_FromPositions(Nothing, _
                                                                               pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1), _
                                                                               pArrayOfPositionsAtEndsAndJointsOfMembers(i, i2)), _
                                                            Line_FromPositions(Nothing, _
                                                                               pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1), _
                                                                               pArrayOfPositionsAtEndsAndJointsOfMembers(j, j2))) Then
                                                                           
                                        Call DebugMsg("Common end found between Position(" + CStr(i1) + ") and Position(" + CStr(j1) + ")")
                                        Set pPositionOfCommonNode = pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1)
                                        Exit For
                                    End If
                                End If
                             End If
                        End If
                    Next j1
                    If Not pPositionOfCommonNode Is Nothing Then Exit For
                Next i1
                If Not pPositionOfCommonNode Is Nothing Then Exit For
            Next j
            If Not pPositionOfCommonNode Is Nothing Then Exit For
        Next
        
        If pPositionOfCommonNode Is Nothing Then
            For i = 1 To pElementsOfMemberParts.Count - 1
                For j = i + 1 To pElementsOfMemberParts.Count
                    Call DebugMsg("Processing Member#" + CStr(i) + " with Member#" + CStr(j))
                    For i1 = 2 To 3
                        For j1 = 2 To 3
                            If Not pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1) Is Nothing Then
                                If Not pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1) Is Nothing Then
                              ' test is one end is along the other
                                    Dim pCurve1 As IJCurve
                                    If (i1 = 0 Or i1 = 2) Then
                                        Set pCurve1 = Line_FromPositions(Nothing, _
                                                                         pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1), _
                                                                         pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1 + 1))
                                        If pCurve1.IsPointOn(pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1).x, _
                                                             pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1).y, _
                                                             pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1).z) Then
                                            Call DebugMsg("Common end found along line from Position(" + CStr(i1) + ") at Position(" + CStr(j1) + ")")
                                            Set pPositionOfCommonNode = pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1)
                                            Exit For
                                        End If
                                    End If
                                        
                                    ' test is one end is along the other(reverse order)
                                    Dim pCurve2 As IJCurve:
                                    If (j1 = 0 Or j1 = 2) Then
                                        Set pCurve2 = Line_FromPositions(Nothing, _
                                                                         pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1), _
                                                                         pArrayOfPositionsAtEndsAndJointsOfMembers(j, j1 + 1))
                                        If pCurve2.IsPointOn(pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1).x, _
                                                             pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1).y, _
                                                             pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1).z) Then
                                            Call DebugMsg("Common end found along line from Position(" + CStr(j1) + ") at Position(" + CStr(i1) + ")")
                                            Set pPositionOfCommonNode = pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1)
                                            Exit For
                                        End If
                                    End If
                                        
                                    ' test if the 2 first Members intersect
                                    If i = 1 And j = 2 And (i1 = 0 Or i1 = 2) And (j1 = 0 Or j1 = 2) Then
                                        ' only look for intersection of physical axes
                                        If i1 = 2 And j1 = 2 Then
                                            Set pPositionOfCommonNode = Position_AtCurvesIntersection(pCurve1, pCurve2, Nothing, GCNear)
                                            If Not pPositionOfCommonNode Is Nothing Then
                                                Call DebugMsg("Common end found at intersection between line from Position(" + CStr(i1) + ") and line from Position(" + CStr(j1) + ")")
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next j1
                        If Not pPositionOfCommonNode Is Nothing Then Exit For
                    Next i1
                    If Not pPositionOfCommonNode Is Nothing Then Exit For
                Next j
                If Not pPositionOfCommonNode Is Nothing Then Exit For
            Next i
        End If
        Call DebugValue("PositionOfCommonEnd", pPositionOfCommonNode)
    End If
    
    ' verify that each Member has the node as end point or is going through the node
    If Not pPositionOfCommonNode Is Nothing Then
        For i = 1 To pElementsOfMemberParts.Count
            Call DebugMsg("Processing Member#" + CStr(i))
            ' retrieve the key of the Member in the collection
            Dim sKey As String: Let sKey = pElementsOfMemberParts.GetKey(pElementsOfMemberParts(i))
            If sKey = "" Then sKey = CStr(i)

            ' build an extended key
            ' when a Member is traversing the node, it is split into 2 lines. The first one get the key of the Member. The second one get the extended key.
            ' the key is made of a letter and a number,which represents the key of the  Member in its original collection, before the collections get merged into a single collection :
            '   - SupportingMemberParts and BracingMemberParts are merged into the a unique collection
            '   - ColumnMemberParts, OrthogonalColumnMemberParts and SecondaryMemberParts are merged into a unique collection
            ' the letter of the key keeps track of the name of the original collection:
            '   - 'S' : "SupportingMemberPart..."
            '   - 'B' : "BracingMemberPart.."
            '   - 'C' : "ColumnMemberPart..."
            '   - 'O' : "OrthogonalColumnMemberPart..."
            '   - 'S' : "SecondaryColumnMemberPart..."
            
            ' verify if the member is connected to the Member
            Dim sExtendedKey As String: Let sExtendedKey = sKey + "|Extended"
            Dim bIsConnectedToNode As Boolean: Let bIsConnectedToNode = False
            For i1 = 2 To 3
                If Not pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1) Is Nothing Then
                    If pArrayOfPositionsAtEndsAndJointsOfMembers(i, i1).DistPt(pPositionOfCommonNode) < EPSILON Then
                        Call DebugMsg("Common end found at Position(" + CStr(i1) + ")")
                        Select Case i1
                            Case 0, 1: Call pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode.Add(Line_FromPositions(Nothing, _
                                                                                                                          pArrayOfPositionsAtEndsAndJointsOfMembers(i, 0), _
                                                                                                                          pArrayOfPositionsAtEndsAndJointsOfMembers(i, 1)), sKey)
                            Case 2, 3: Call pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode.Add(Line_FromPositions(Nothing, _
                                                                                                                          pArrayOfPositionsAtEndsAndJointsOfMembers(i, 2), _
                                                                                                                          pArrayOfPositionsAtEndsAndJointsOfMembers(i, 3)), sKey)
                        End Select
                        bIsConnectedToNode = True
                        Exit For
                    End If
                End If
            Next
            
            ' verify if the Member goes through the node
            If Not bIsConnectedToNode Then
                Dim bIsNodeAlongLine As Boolean: Let bIsNodeAlongLine = False
                If Not pArrayOfPositionsAtEndsAndJointsOfMembers(i, 0) Is Nothing _
                And Not pArrayOfPositionsAtEndsAndJointsOfMembers(i, 1) Is Nothing Then
                    Dim pCurve As IJCurve:
                    Set pCurve = Line_FromPositions(Nothing, _
                                                    pArrayOfPositionsAtEndsAndJointsOfMembers(i, 0), _
                                                    pArrayOfPositionsAtEndsAndJointsOfMembers(i, 1))
                    If pCurve.IsPointOn(pPositionOfCommonNode.x, _
                                        pPositionOfCommonNode.y, _
                                        pPositionOfCommonNode.z) Then
                        Call DebugMsg("Common end found along line from Position(" + CStr(0) + ")")
                        Call pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode.Add(Line_FromPositions(Nothing, _
                                                                                                           pPositionOfCommonNode, _
                                                                                                           pArrayOfPositionsAtEndsAndJointsOfMembers(i, 0)), sKey)
                        Call pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode.Add(Line_FromPositions(Nothing, _
                                                                                                           pPositionOfCommonNode, _
                                                                                                           pArrayOfPositionsAtEndsAndJointsOfMembers(i, 1)), sExtendedKey)
                        bIsNodeAlongLine = True
                    End If
                End If
                
                If Not bIsNodeAlongLine _
                And Not pArrayOfPositionsAtEndsAndJointsOfMembers(i, 2) Is Nothing _
                And Not pArrayOfPositionsAtEndsAndJointsOfMembers(i, 3) Is Nothing Then
                    Set pCurve = Line_FromPositions(Nothing, _
                                                    pArrayOfPositionsAtEndsAndJointsOfMembers(i, 2), _
                                                    pArrayOfPositionsAtEndsAndJointsOfMembers(i, 3))
                
                    If pCurve.IsPointOn(pPositionOfCommonNode.x, _
                                        pPositionOfCommonNode.y, _
                                        pPositionOfCommonNode.z) Then
                        Call DebugMsg("Common end found along line from Position(" + CStr(2) + ")")
                        Call pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode.Add(Line_FromPositions(Nothing, _
                                                                                                           pPositionOfCommonNode, _
                                                                                                           pArrayOfPositionsAtEndsAndJointsOfMembers(i, 2)), sKey)
                        Call pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode.Add(Line_FromPositions(Nothing, _
                                                                                                           pPositionOfCommonNode, _
                                                                                                           pArrayOfPositionsAtEndsAndJointsOfMembers(i, 3)), sExtendedKey)
                        bIsNodeAlongLine = True
                    End If
                End If
                
                If Not bIsNodeAlongLine Then
                    Set pPositionOfCommonNode = Nothing
                    Exit For
                End If
            End If
        Next
    
        ' record the position of the node at the end
        Call pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode.Add(pPositionOfCommonNode, "0")
    End If
        
    ' return result
    Set MemberParts_GetElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode = pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode
    
    Call DebugOutput("ElementsOfPositionOfCommonEndAndLinesOfAxes", pElementsOfLinesOfAxesOfMembersAndPositionOfCommonNode)
    Call DebugOut
End Function
Public Function MemberParts_GetPositionOfCommonEnd(pElementsOfMemberParts As IJElements) As IJDPosition
    Call DebugIn("MemberParts_GetPositionOfCommonNode")
    Call DebugInput("Count of MemberParts", pElementsOfMemberParts.Count)
    
    ' prepare the result
    Dim pPositionOfCommonEnd As IJDPosition: Set pPositionOfCommonEnd = Nothing
    
    ' get the positions at ends and joints of the 2 first MemberParts
    Dim pPositionsAtEndsAndJointsOfMember1() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember1 = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(1))
    Dim pPositionsAtEndsAndJointsOfMember2() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember2 = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(2))
    
    ' find a common position between the positions extracted from the 2 first MemberParts
    Dim bCommonNodeIsFound As Boolean: Let bCommonNodeIsFound = False
    Dim i1 As Integer: Dim i2 As Integer
    For i1 = 0 To 3
        For i2 = 0 To 3
            If Not pPositionsAtEndsAndJointsOfMember1(i1) Is Nothing Then
                If Not pPositionsAtEndsAndJointsOfMember2(i2) Is Nothing Then
                    If pPositionsAtEndsAndJointsOfMember1(i1).DistPt(pPositionsAtEndsAndJointsOfMember2(i2)) < EPSILON Then
                        Call DebugValue("Position found for i1", i1)
                        Call DebugValue("Position found for i2", i2)
                        Set pPositionOfCommonEnd = pPositionsAtEndsAndJointsOfMember1(i1)
                        Let bCommonNodeIsFound = True
                        Exit For
                    End If
                End If
            End If
        Next
        If bCommonNodeIsFound Then Exit For
    Next
    
    If bCommonNodeIsFound And pElementsOfMemberParts.Count > 2 Then
        ' verify if we can extract from each remaining MemberPart a position common with the already found common node
        Let bCommonNodeIsFound = False
        Dim i As Integer
        For i = 3 To pElementsOfMemberParts.Count
            Dim pPositionsAtEndsAndJointsOfMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfMember = MemberPart_GetPositionsAtJointsAndEnds(pElementsOfMemberParts(i))
            For i2 = 0 To 3
                If Not pPositionsAtEndsAndJointsOfMember(i2) Is Nothing Then
                    If pPositionOfCommonEnd.DistPt(pPositionsAtEndsAndJointsOfMember(i2)) < EPSILON Then
                        Call DebugValue("Processing Member", i)
                        Call DebugValue("Position found for i2", i2)
                        Let bCommonNodeIsFound = True
                        Exit For
                    End If
                End If
            Next
            If Not bCommonNodeIsFound Then Exit For
        Next
    End If
     
    
    If Not bCommonNodeIsFound Then
        ' clear the eventually already found common node
        Set pPositionOfCommonEnd = Nothing
    End If
    
    ' return result
    Set MemberParts_GetPositionOfCommonEnd = pPositionOfCommonEnd
    
    Call DebugOutput("PositionOfCommonNode", MemberParts_GetPositionOfCommonEnd)
    Call DebugOut
End Function
Function MemberParts_GetPositionAlong(oSupportingMemberPart As Object, oSupportedMemberPart As Object) As IJDPosition
    Call DebugIn("MemberParts_GetPositionAlong")
    Call DebugInput("SupportingMemberPart", oSupportingMemberPart)
    Call DebugInput("SupportedMemberPart", oSupportedMemberPart)
    
    ' prepare result
    Dim pPositionAlong As IJDPosition: Set pPositionAlong = Nothing
    
    ' get the positions at ends and joints of the 2 first MemberParts
    Dim pPositionsAtEndsAndJointsOfSupportingMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfSupportingMember = MemberPart_GetPositionsAtJointsAndEnds(oSupportingMemberPart)
    Dim pPositionsAtEndsAndJointsOfSupportedMember() As IJDPosition: Let pPositionsAtEndsAndJointsOfSupportedMember = MemberPart_GetPositionsAtJointsAndEnds(oSupportedMemberPart)

    Dim i1 As Integer
    For i1 = 0 To 3 Step 2
        If Not pPositionsAtEndsAndJointsOfSupportingMember(i1) Is Nothing Then
            Dim pLineOfSupportingMember As IJLine
            Set pLineOfSupportingMember = Line_FromPositions(Nothing, pPositionsAtEndsAndJointsOfSupportingMember(i1), _
                                                                      pPositionsAtEndsAndJointsOfSupportingMember(i1 + 1))
            
            Dim i2 As Integer
            For i2 = 0 To 3
                If Not pPositionsAtEndsAndJointsOfSupportedMember(i2) Is Nothing Then
                    Dim pPositionOnSupportingMember As IJDPosition: Set pPositionOnSupportingMember = Position_ProjectOnLine(pPositionsAtEndsAndJointsOfSupportedMember(i2), pLineOfSupportingMember)
                    If pPositionOnSupportingMember.DistPt(pPositionsAtEndsAndJointsOfSupportedMember(i2)) < EPSILON Then
                        Set pPositionAlong = pPositionOnSupportingMember
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    
    If pPositionAlong Is Nothing Then
        'MsgBox "PositionAlong not found"
        Err.Raise E_FAIL
    End If
    
    ' return result
    Set MemberParts_GetPositionAlong = pPositionAlong
    
    Call DebugOutput("PositionAlong", MemberParts_GetPositionAlong)
    Call DebugOut
End Function
Function GetLinesOfMemberAxes(pElementsOfMembers As IJElements) As IJLine()
    ' initialise results
    Dim pLinesOfMemberAxes() As IJLine
    
    ReDim pLinesOfMemberAxes(1 To pElementsOfMembers.Count)
    Dim i As Integer
    For i = 1 To pElementsOfMembers.Count
        Set pLinesOfMemberAxes(i) = Member_GetLine(pElementsOfMembers(i))
    Next
    
    ' return result
    Let GetLinesOfMemberAxes = pLinesOfMemberAxes
End Function


