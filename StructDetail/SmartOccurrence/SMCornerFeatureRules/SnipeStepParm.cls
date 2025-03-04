VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SnipeStepParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'
Option Explicit

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.SnipeStepParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.SnipeStepParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\SnipeStepParm.cls"

Implements IJDUserSymbolServices

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

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  
pOH.SetOutput "Ulength"
pOH.SetOutput "Vlength"

  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
  
    On Error GoTo ErrorHandler
    
    Dim stepLength As Double
    stepLength = 0#
    
    ' --------------------------------------
    ' Get the input ports and their geometry
    ' --------------------------------------
    ' Get the ports
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    
    Set oEdgePort1 = oPRL.InputObject(INPUT_PORT2EDGE)
    Set oEdgePort2 = oPRL.InputObject(INPUT_PORT3EDGE)
    
    'oCornerFeature.GetLocationOfCornerFeature fails in a specific case where a stiffener is bounded to a member and a plate(auxiliary port).
    'In such case oEdgePort1.Geometry and oEdgePort2.geometry is not evaluated in 1st iteration. a TR is filed to handle the situation, until then
    'a snipe 5*5 mm is created in such cases.
    '*********************************************
    On Error Resume Next
    If oEdgePort1.Geometry Is Nothing Or oEdgePort2.Geometry Is Nothing Then
        oPRL.Add "Ulength", 0.005
        oPRL.Add "Vlength", 0.005
    Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    '*****************************************
    
    Dim oStructPort1 As IJStructPort
    Set oStructPort1 = oEdgePort1
    
    If oStructPort1.SectionID >= e_JXSEC_MultipleBounding_5001 And oStructPort1.SectionID <= e_JXSEC_MultipleBounding_5005 Then
        
        Dim dLength1  As Double
        Dim dLength2 As Double
        
        GetCornerFatureEdgeLengths oPRL.SmartOccurrence, Nothing, oEdgePort1, oEdgePort2, _
                                    dLength1, dLength2, oPRL.InputObject(INPUT_PORT1FACE)
    
        If dLength1 < dLength2 Then
            stepLength = dLength1
        Else
            stepLength = dLength2
        End If
        
        oPRL.Add "Ulength", stepLength - 0.000001
        oPRL.Add "Vlength", stepLength - 0.000001
        Exit Sub
    Else
        
        ' --------------------------------------------------------------------------
        ' Get all the faces on each port and find the closest to the corner position
        ' --------------------------------------------------------------------------
        ' Get the attributes on the port
        Dim oFace1 As IJSurfaceBody
        Dim oFace2 As IJSurfaceBody
        
        Set oFace1 = GetFaceNearCornerPosition(oPRL.SmartOccurrence, oEdgePort1)
        Set oFace2 = GetFaceNearCornerPosition(oPRL.SmartOccurrence, oEdgePort2)
        
        If Not oFace1 Is Nothing And Not oFace2 Is Nothing Then
            ' ---------------------------
            ' Find the smaller of the two
            ' ---------------------------
            ' We presume this is the 'step' that must be sniped
            Dim oSmallerFace As IJSurfaceBody
            Dim oLargerFace As IJSurfaceBody
            
            Dim oDModelBody As IJDModelBody
            Dim dAccuracyAchieved As Double
            Dim dLength As Double
            Dim dArea1 As Double
            Dim dArea2 As Double
            Dim dVolume As Double
            
            Set oDModelBody = oFace1
            oDModelBody.GetDimMetrics 0.01, dAccuracyAchieved, dLength, dArea1, dVolume
            
            Set oDModelBody = oFace2
            oDModelBody.GetDimMetrics 0.01, dAccuracyAchieved, dLength, dArea2, dVolume
            
            If dArea1 < dArea2 Then
                Set oSmallerFace = oFace1
                Set oLargerFace = oFace2
            Else
                Set oSmallerFace = oFace2
                Set oLargerFace = oFace1
            End If
            
            ' -------------------------------------------------------------------------------------------
            ' Get all the vertices on the smaller face and find the one furthest from the corner position
            ' -------------------------------------------------------------------------------------------
            Dim oSDOCorner As New StructDetailObjects.CornerFeature
            Set oSDOCorner.object = oPRL.SmartOccurrence
            
            Dim oCornerPos As IJDPosition
            oSDOCorner.GetLocationOfCornerFeature oCornerPos
            
            Dim oVertices As Collection
            Dim oModelUtil As IJSGOModelBodyUtilities
            Set oModelUtil = New SGOModelBodyUtilities
            
            oModelUtil.GetVertices oSmallerFace, oVertices
            
            Dim dist As Double
            Dim maxDist As Double
            Dim oPosition As IJDPosition
            Dim oMaxVert As IJDPosition
            
            maxDist = -1#
            
            For Each oPosition In oVertices
                dist = oPosition.DistPt(oCornerPos)
                
                If dist > maxDist Then
                    maxDist = dist
                    Set oMaxVert = oPosition
                End If
            Next oPosition
            
            ' --------------------------------------------------------------------------------------------
            ' The corner position and the vertex found may form a diagonal.  Project it to the other port.
            ' --------------------------------------------------------------------------------------------
            oModelUtil.GetClosestPointOnBody oLargerFace, oMaxVert, oPosition, stepLength
        End If
    End If
    
    ' ----------------------------------------------------------------------------------
    ' The step is this distance minus a tolerance, to avoid cutting and adjacent surface
    ' ----------------------------------------------------------------------------------
    If stepLength < 0.00001 Then
        stepLength = 0.015
    End If
    If stepLength > 0.035 Then
        stepLength = 0.015
    End If
    
    oPRL.Add "Ulength", stepLength - 0.000001
    oPRL.Add "Vlength", stepLength - 0.000001
    
    Exit Sub
   
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
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
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
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
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim oPRL As IJDParameterLogic
  Set oPRL = New ParameterLogic
  oPRL.Representation = pRep
  ParameterRuleLogic oPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

Private Function GetFaceNearCornerPosition(oCornerFeature As Object, oPort As IJPort) As IJSurfaceBody

    On Error GoTo ErrorHandler
    
    Set GetFaceNearCornerPosition = Nothing
    
    Dim oSDOCorner As New StructDetailObjects.CornerFeature
    Set oSDOCorner.object = oCornerFeature
    
    ' ------------------------------
    ' Get the attributes on the port
    ' ------------------------------
    Dim oStructPort As IJStructPort
    Set oStructPort = oPort
    
    Dim topoType As JS_TOPOLOGY_PROXY_TYPE
    Dim ctxID As eUSER_CTX_FLAGS
    Dim optID As Long
    Dim oprID As Long
    Dim xID As JXSEC_CODE
    
    oStructPort.GetAttributes topoType, ctxID, optID, oprID, xID
    
    ' ----------------------------------------
    ' Get sub ports on the geometry after trim
    ' ----------------------------------------
    Dim oPortsAfterTrim As IJElements
    Dim oStructGraphConnectable As IJStructGraphConnectable
    Dim oStuctConn As IJStructConnectable
    
    Dim oPortElements As IJElements
    Dim filterType As JS_TOPOLOGY_FILTER_TYPE
    Dim isBase As Boolean
    
    If (ctxID And CTX_BASE) = CTX_BASE Then
        isBase = True
    End If
    
    Dim i As Long
    
    Dim nEdges As Long
    Dim nFaces As Long
    Dim nLoops As Long
    Dim nLumps As Long
    Dim nVerts As Long
    Dim nWires As Long
    Dim nShells As Long
    Dim nCoddges As Long
    Dim bSaneBody As Boolean
    Dim oDModelBody As IJDModelBody
    
    
    Dim oOperatorIDs As New Collection
    Dim oGlobalPortElements As IJElements
    Dim j As Long
    Dim subOprID As Long
    Dim bOprFound As Boolean
    Dim oGlobalPorts As New Collection
    
    
    If TypeOf oPort.Connectable Is ISPSMemberPartCommon Then
    
        Set oStructGraphConnectable = oSDOCorner.GetPartObject
        oStructGraphConnectable.enumPortsInGraphByTopologyFilter oPortElements, _
                                                                 JS_TOPOLOGY_FILTER_LCONNECT_PRT_SUB_LFACES, _
                                                                 GeometryByIIDInGraph, _
                                                                 "ISPSPartPrismaticGenerator"
        ' Get the global ports
        oStructGraphConnectable.enumPortsInGraphByTopologyFilter oGlobalPortElements, _
                                                                 JS_TOPOLOGY_FILTER_SOLID_3DCUTOUT_LFACES, _
                                                                 GeometryByIIDInGraph, _
                                                                 "ISPSPartPrismaticGenerator"

    ElseIf TypeOf oPort.Connectable Is IJPlate Then
        Set oStuctConn = oSDOCorner.GetPartObject
        oStuctConn.enumConnectableTransientPorts oPortElements, vbNullString, False, PortFace, JS_TOPOLOGY_FILTER_ALL_LFACES
    Else
    
        Set oStuctConn = oSDOCorner.GetPartObject
        oStuctConn.enumConnectableTransientPorts oPortElements, "ProfilePartActiveEntities.ProfileTrim_AE.1", False, PortFace, JS_TOPOLOGY_FILTER_ALL_LFACES, False

        ' Get the global ports
        oStuctConn.enumConnectableTransientPorts oGlobalPortElements, "ProfilePartActiveEntities.ProfileTrim_AE.1", False, PortFace, JS_TOPOLOGY_FILTER_ALL_LFACES, True
        
    End If
    
    
    ' Collect the sub port operator ids
    If Not oPortElements Is Nothing Then
        For i = 1 To oPortElements.Count
            Set oStructPort = oPortElements.Item(i)
            oOperatorIDs.Add oStructPort.operatorID
        Next i
    End If
    
    ' -------------------------------------------------------
    ' Include global ports if the cut does not have sub-ports
    ' -------------------------------------------------------
    ' If the global port is a new operator, add it to the list
    If Not oGlobalPortElements Is Nothing Then
        For i = 1 To oGlobalPortElements.Count
            Set oStructPort = oGlobalPortElements.Item(i)
            bOprFound = False
            
            For j = 1 To oOperatorIDs.Count
                subOprID = oOperatorIDs.Item(j)
                If subOprID = oStructPort.operatorID Then
                    bOprFound = True
                End If
            Next j
            
            If Not bOprFound Then
                oOperatorIDs.Add oStructPort.operatorID
                oGlobalPorts.Add oStructPort
            End If
        Next i
        
        For i = 1 To oGlobalPorts.Count
            oPortElements.Add oGlobalPorts.Item(i)
        Next i
        
    End If
    
    ' -----------------------------------------------------------------------------------------------------
    ' Find the face closest to the corner position with attributes matching the input port minus section ID
    ' -----------------------------------------------------------------------------------------------------
    Dim oCornerPos As IJDPosition
    Dim oPosOnFace As IJDPosition
    Dim oPortIter As IJPort
    Dim dist As Double
    Dim minDist As Double

    minDist = 1000#
    
    oSDOCorner.GetLocationOfCornerFeature oCornerPos
    
    Dim oModelUtil As IJSGOModelBodyUtilities
    Set oModelUtil = New SGOModelBodyUtilities
    
    ' Loop through all the ports
    If Not oPortElements Is Nothing Then
        For i = 1 To oPortElements.Count
    
            Set oStructPort = oPortElements.Item(i)
        
            ' Skip ports that are not at the same end as the corner feature
            If (((oStructPort.ContextID And CTX_BASE) = CTX_BASE) And isBase) Or _
               (((oStructPort.ContextID And CTX_OFFSET) = CTX_OFFSET) And Not isBase) Then
            
                If oStructPort.ContextID = ctxID And oStructPort.operationID = optID And oStructPort.operatorID = oprID Then
                    Set oPortIter = oStructPort
                    oModelUtil.GetClosestPointOnBody oPortIter.Geometry, oCornerPos, oPosOnFace, dist
                    
                    If dist < minDist Then
                      ' the current Port is the closest port found
                        ' check if the Port geometry contains multiple Lumps
                        nLumps = 0
                        Set oDModelBody = oPortIter.Geometry
                        oDModelBody.CheckTopology vbNullString, nLumps, nShells, nWires, nFaces, _
                                                  nLoops, nCoddges, nEdges, nVerts, bSaneBody
    
                        If nLumps < 2 Then
                            minDist = dist
                            Set GetFaceNearCornerPosition = oPortIter.Geometry
                        Else
                            ' Current closest Port geometry has multiple Lumps
                            ' Need to find the closest Port geometry "Lump"
                            Dim oLump As Object
                            Dim oSurfaceBody As IJSurfaceBody
                            Dim oEnumLumps As IEnumUnknown
                            Dim oCollLumps As JCmnShp_CollectionAlias
                            Dim ConvertUtils As CCollectionConversions
                            Dim oGeomOpsMisc As IMSModelGeomOps.DGeomOpsMisc
                            
                            If oGeomOpsMisc Is Nothing Then
                                Set oGeomOpsMisc = New IMSModelGeomOps.DGeomOpsMisc
                                Set ConvertUtils = New CCollectionConversions
                            End If
                            
                            oGeomOpsMisc.ExplodeModelBodyByLumps Nothing, oDModelBody, oEnumLumps
                            ConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumLumps, oCollLumps
                                        
                            
                            nLumps = 0
                            For Each oLump In oCollLumps
                                nLumps = nLumps + 1
                                Set oSurfaceBody = oLump
                                oModelUtil.GetClosestPointOnBody oSurfaceBody, oCornerPos, oPosOnFace, dist
                                If dist < minDist Then
                                    minDist = dist
                                    Set GetFaceNearCornerPosition = oSurfaceBody
                                End If
                            Next oLump
                        End If
                    End If
                End If
            End If
        Next i
    End If
    
    Exit Function
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "GetProfileEdgeThickness").Number
End Function
