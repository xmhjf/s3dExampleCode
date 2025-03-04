Attribute VB_Name = "Update_Utils"
Option Explicit

Private Const MODULE = m_sProjectPath + "Update_Utils.bas"
Public oGeomUtil As New GSCADStructGeomUtilities.TopologyLocate
'

'**************************************************************************
'**************************************************************************
'**************************************************************************

Public Sub UpdateDependentCorners(oSlotObject As Object)

Const sMETHOD = "UpdateDependentCorners"
On Error GoTo ErrHandler

    ' ----------------------
    ' Check for valid inputs
    ' ----------------------
    Dim oStructFeature As IJStructFeature
    
    If oSlotObject Is Nothing Then
        Exit Sub
    ElseIf TypeOf oSlotObject Is IJStructFeature Then
        Set oStructFeature = oSlotObject
        If oStructFeature.get_StructFeatureType <> SF_Slot Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
            
    ' -------------------------------------------------------------------------------
    ' Retreive all of the Assembly Connections for penetrating and penetrated objects
    ' -------------------------------------------------------------------------------
    Dim oSDO_Helper As New StructDetailObjects.Helper
    Dim oSDO_SlotPart As New StructDetailObjects.Slot
    Set oSDO_SlotPart.object = oStructFeature
    
    Dim nPenetratingACs As Long
    Dim nPenetratedACs As Long
    Dim aPenetratingACData() As ConnectionData
    Dim aPenetratedACData() As ConnectionData
       
    oSDO_Helper.Object_AppConnections oSDO_SlotPart.Penetrating, _
                                      AppConnectionType_Assembly, _
                                      nPenetratingACs, _
                                      aPenetratingACData
                                      
    oSDO_Helper.Object_AppConnections oSDO_SlotPart.Penetrated, _
                                      AppConnectionType_Assembly, _
                                      nPenetratedACs, _
                                      aPenetratedACData
    
    ' -------------------------------------------------------------------
    ' Collect objects connected to both pentrated and penetrating objects
    ' -------------------------------------------------------------------
    Dim i As Long
    Dim j As Long
    Dim oMutualObjects As Collection
    Set oMutualObjects = New Collection
    Dim oPortsToPenetrating As Collection
    Dim oPortsToPenetrated As Collection
    Dim oPenetratingConnectedObject As Object
    Dim oPenetratedConnectedObject As Object
    Set oPortsToPenetrating = New Collection
    Set oPortsToPenetrated = New Collection
    
    For i = 1 To nPenetratingACs
        Set oPenetratingConnectedObject = aPenetratingACData(i).ToConnectable
        
        For j = 1 To nPenetratedACs
            Set oPenetratedConnectedObject = aPenetratedACData(j).ToConnectable
            If oPenetratingConnectedObject Is oPenetratedConnectedObject Then
                oMutualObjects.Add oPenetratedConnectedObject
                oPortsToPenetrating.Add aPenetratingACData(i).ToConnectedPort
                oPortsToPenetrated.Add aPenetratedACData(j).ToConnectedPort
                Exit For
            End If
        Next j
    Next i
                
    Dim oObject As Object
    Dim oFeatures As Collection
    Dim oFeature As IUnknown
    Dim featureType As StructFeatureTypes
    Dim oFacePort As IJPort
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    
    Dim oFeatureAttr As IJSDFeatureAttributes
    Set oFeatureAttr = New SDFeatureUtils
    
    Dim oSDPartSupport As IJPartSupport
    
    Dim eConnectionBehavior As GSCADAppConnections.ConnectionBehavior
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    
    For i = 1 To oMutualObjects.Count
        Set oObject = oMutualObjects.Item(i)
        
        Set oFeatures = Nothing
        Set oSDPartSupport.Part = oObject
        oSDPartSupport.GetFeatures oFeatures

        For j = 1 To oFeatures.Count
            Set oFeature = oFeatures.Item(j)
            If TypeOf oFeature Is IJStructFeature Then
                Set oStructFeature = oFeature
                featureType = oStructFeature.get_StructFeatureType
                If featureType = SF_CornerFeature Then
                    '
                    ' Regular corner features use ports as inputs, while on slot corner features use slot symbol output as inputs
                    ' Use get_CornerCutInputsEx to get inputs for both types of corner features
                    '
                    oFeatureAttr.get_CornerCutInputsEx oFeature, oFacePort, oEdgePort1, oEdgePort2
                    
                    If TypeOf oEdgePort1 Is IJPort And _
                       TypeOf oEdgePort2 Is IJPort Then
                        '
                        ' IsDependentCorner is only applicable to corner features with ports as inputs
                        '
                        If IsDependentCorner(oEdgePort1, _
                                         oEdgePort2, _
                                         oPortsToPenetrating.Item(i), _
                                         oPortsToPenetrated.Item(i)) Then
                            ForceUpdateSmartItem oFeature
                        End If
                    End If
                End If
            End If
       Next j
    
    Next i
    
    Exit Sub
    
ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Function IsDependentCorner(oPort1 As IJStructPort, _
                                  oPort2 As IJStructPort, _
                                  oPortToPenetrating As IJStructPort, _
                                  oPortToPenetrated As IJStructPort) As Boolean
    
    Dim oPort As IJPort
    Set oPort = oPort1
    Dim oMutualObject As Object
    Set oMutualObject = oPort.Connectable
    
    If TypeOf oMutualObject Is IJStiffener Or TypeOf oMutualObject Is IJProfileER Then
        Dim oSDOProfile As StructDetailObjects.ProfilePart
        Set oSDOProfile = New StructDetailObjects.ProfilePart
        Set oSDOProfile.object = oMutualObject
        
        Dim ctx As eUSER_CTX_FLAGS
        Dim filter As eUSER_CTX_FLAGS
        Dim baseType As Base_Port_Types
        filter = CTX_BASE Or CTX_OFFSET
        
        ctx = oPort1.ContextID
        
        If ctx And CTX_BASE Then
            baseType = BPT_Base
        ElseIf ctx And CTX_OFFSET Then
            baseType = BPT_Offset
        Else
            baseType = BPT_Lateral
        End If
        
        Set oPort1 = oSDOProfile.BasePortBeforeTrim(baseType)
        
        ctx = oPort2.ContextID
        
        If ctx And CTX_BASE Then
            baseType = BPT_Base
        ElseIf ctx And CTX_OFFSET Then
            baseType = BPT_Offset
        Else
            baseType = BPT_Lateral
        End If
        
        Set oPort2 = oSDOProfile.BasePortBeforeTrim(ctx And filter)
    Else
        Set oPort1 = RelatedPortBeforeCut(oPort1, True)
        Set oPort2 = RelatedPortBeforeCut(oPort2, True)
    End If
    
    If (oPort1 Is oPortToPenetrating And oPort2 Is oPortToPenetrated) Or _
       (oPort1 Is oPortToPenetrated And oPort2 Is oPortToPenetrating) Then
           
       IsDependentCorner = True
       Exit Function
    End If
    
    IsDependentCorner = False

End Function

Public Sub ForceUpdateSmartItem(oObject As Object)

On Error GoTo ErrorHandler
Const sMETHOD = "ForceUpdateSmartItem"
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This appears to allow Assoc to always recompute WebCut before FlangeCuts
    ' interface IJStructGeometry : {6034AD40-FA0B-11d1-B2FD-080036024603}
    Dim oStructAssocTools As SP3DStructGenericTools.StructAssocTools

    Set oStructAssocTools = New SP3DStructGenericTools.StructAssocTools
    oStructAssocTools.UpdateObject oObject, _
                                   "{6034AD40-FA0B-11d1-B2FD-080036024603}"
    Set oStructAssocTools = Nothing
    Exit Sub
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Sub


' Copied from StructDetailObjects PortUtilities.bas
' Need to request that this is exposed
' Tried adding .bas to project, but could not get past compile errors
Public Function RelatedPortBeforeCut(oPort As IJPort, _
                            Optional bGlobalPorts As Boolean = False) As IJPort
On Error GoTo ErrorHandler
Const MT = "RelatedPortBeforeCut"

    Const strGeneratedPlatePartAEProgId = "CreatePlatePart.GeneratePlatePart_AE.1"
    Const strStandAlonePlatePartAEProgId = "CreatePlatePart.CreatePlatePart_AE.1"
    Const strGeneratedProfilePartAEProgId = "ProfilePartActiveEntities.ProfilePartGeneration_AE.1"
    Const strStandAloneProfilePartTrimAEProgId = "ProfilePartActiveEntities.ProfileTrim_AE.1"
    Const strStandAloneProfilePartAEProgID = "ProfilePartActiveEntities.ProfilePartCreation.1"
    Const strSketchFeatureCutAEProgID = "SketchFeature.SDCutAE.1"

    Dim oStructPort As IJStructPortEx
    Set oStructPort = oPort
    
    If TypeOf oPort.Connectable Is IJPlatePart Then
        Set RelatedPortBeforeCut = oStructPort.RelatedPort(strStandAlonePlatePartAEProgId, _
                                                           False, bGlobalPorts)
        
        If (oPort Is RelatedPortBeforeCut) Then
            Set RelatedPortBeforeCut = oStructPort.RelatedPort(strGeneratedPlatePartAEProgId, _
                                                               False, bGlobalPorts)
        End If
    Else
        ' Generated Profile part with  Bound , stand alone profile part with or without boundary
        If TypeOf oPort.Connectable Is IJProfilePart Then
            Set RelatedPortBeforeCut = oStructPort.RelatedPort(strGeneratedProfilePartAEProgId, _
                                                            False, bGlobalPorts)
             
            ' Generated profile part without Boundary
            If (oPort Is RelatedPortBeforeCut) Then
                Set RelatedPortBeforeCut = oStructPort.RelatedPort(strGeneratedProfilePartAEProgId, _
                                                            False, bGlobalPorts)
            End If
            
            If (oPort Is RelatedPortBeforeCut) Then
                'StandAlone not sure if this is required, it does not harm though
                Set RelatedPortBeforeCut = oStructPort.RelatedPort(strStandAloneProfilePartAEProgID, _
                                                            False, bGlobalPorts)
            End If
            
            ' This may not be required
            If (oPort Is RelatedPortBeforeCut) Then
                Set RelatedPortBeforeCut = oStructPort.RelatedPort(strSketchFeatureCutAEProgID, _
                                                            True, bGlobalPorts)
            End If
        End If
    End If
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function

Function DoesEdgeOverlapPart(ByRef oSlotSymbol As IMSSymbolEntities.IJDSymbol, edgeID As JXSEC_CODE) As Boolean

    DoesEdgeOverlapPart = False
    
    ' Retreive the Collar Input Slot Symbols Output Representation "Slot"
'    Dim oSlotSymbol As IMSSymbolEntities.IJDSymbol
    Dim oSlotSymbDefRep As IJDRepresentation
    Dim oSlotSymbDefReps As IJDRepresentations
    Dim oSlotSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim RepresentationName As String
    
'    Set oSlotSymbol = pMD.CAO
    Set oSlotSymbolDefinition = oSlotSymbol.IJDSymbolDefinition(1)
    Set oSlotSymbDefReps = oSlotSymbolDefinition.IJDRepresentations
        
    RepresentationName = "Slot"
    Set oSlotSymbDefRep = oSlotSymbDefReps.GetRepresentationByName(RepresentationName)
    Set oGeomUtil = New GSCADStructGeomUtilities.TopologyLocate
    ' Retrieve the number of Outputs for the Slot Symbol Representation "Slot"
    Dim n As Long
    
    Dim oRepOutput As IJDOutput
    Dim oIJDOutputs As IJDOutputs
    Dim nOutputCount As Integer
    
    Set oIJDOutputs = oSlotSymbDefRep
    nOutputCount = oIJDOutputs.OutputCount
    
    Dim oSlot As New StructDetailObjects.Slot
    Set oSlot.object = oSlotSymbol
    
    If Not TypeOf oSlot.Penetrated Is IJPlatePart Then
        DoesEdgeOverlapPart = True
        Exit Function
    End If
    
    Dim oPlatePart As New StructDetailObjects.PlatePart
    Set oPlatePart.object = oSlot.Penetrated
    
    Dim oPlateBase As IJSurfaceBody
    Set oPlateBase = Nothing

    Dim dPlateThickness As Double
    Dim dOverlappingTolerance As Double
    Dim oPlateBasePort As IJPort
    
    dPlateThickness = oPlatePart.PlateThickness

    'if stand alone plate part then
    If oPlatePart.plateType = Standalone Then
        Set oPlateBasePort = oPlatePart.BasePortFromOperation(BPT_Base, "CreatePlatePart.CreatePlatePart_AE.1", False)
    Else
    'if system derived then
        Set oPlateBasePort = oPlatePart.BasePortFromOperation(BPT_Base, "CreatePlatePart.GeneratePlatePart_AE.1", False)
    End If
    Set oPlateBase = oPlateBasePort.Geometry
    
    Dim oPenetratingSurface As IJSurfaceBody
    Dim oProfileObj As New StructDetailObjects.ProfilePart
    Dim oPortWithEdgeID As IJPort
    
    Dim oWire As IJWireBody
        
    ' Loop through each output
    For n = 1 To nOutputCount
        Set oRepOutput = oIJDOutputs.GetOutputAtIndex(n)
                
        Dim sOutputName As String
        sOutputName = oRepOutput.Name
        
        ' Check the output for the desired edge name
        If InStr(sOutputName, CStr(edgeID)) > 0 Then
        
            Dim oSlotOutput As Object
            Set oSlotOutput = oSlotSymbol.BindToOutput(RepresentationName, sOutputName)
        
            If TypeOf oSlotOutput Is IJCurve Then
            
                Dim oModelGeomOps As IMSModelGeomOps.DGeomWireFrameBody
                Set oModelGeomOps = New IMSModelGeomOps.DGeomWireFrameBody
                
                Dim pCurves As IJElements
                Set pCurves = New JObjectCollection
                
                pCurves.Add oSlotOutput
                
                Set oWire = oModelGeomOps.CreateSmartWireBodyFromGTypedCurves(Nothing, pCurves)
                
                'Get the profile surface body passing the lateral port for which the PC/Cf is needed
                If TypeOf oSlot.Penetrating Is IJProfile Then
                
                    Set oProfileObj.object = oSlot.Penetrating
                    Set oPortWithEdgeID = oProfileObj.SubPortBeforeTrim(edgeID)
                    Set oPenetratingSurface = oPortWithEdgeID.Geometry
                    
                ElseIf TypeOf oSlot.Penetrating Is IJPlate Then
                    
                    Dim oSlotMappingRule As IJSlotMappingRule
                    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
                    
                    Dim oBasePort As IJPort
                    Dim oMappedPorts As JCmnShp_CollectionAlias
                    Set oMappedPorts = New Collection
                    
                    oSlotMappingRule.GetEmulatedPorts oSlot.Penetrating, oSlot.Penetrated, oBasePort, oMappedPorts
                    
                    Set oPortWithEdgeID = oMappedPorts.Item(CStr(edgeID))
                    Set oPenetratingSurface = oPortWithEdgeID.Geometry
                    
                Else
                    DoesEdgeOverlapPart = True
                    Exit Function
                End If
        
                Dim oSurfUtilities As GSCADShipGeomOps.SGOSurfaceBodyUtilities
                Set oSurfUtilities = New GSCADShipGeomOps.SGOSurfaceBodyUtilities
                
                Dim oSlotEdgeWireOverlapPenetrated As IJWireBody, oSlotEdgeWireOverlapPenetrating As IJWireBody
                
                dOverlappingTolerance = dPlateThickness * 0.6
                'find the wire(Left/Right slot edge wire) overlap with the solid body(penetrated plate)
                oSurfUtilities.FindWireOverlapWithTolerance oPlateBase, oWire, dOverlappingTolerance, Nothing, Nothing, oSlotEdgeWireOverlapPenetrated
                
                If Not oSlotEdgeWireOverlapPenetrated Is Nothing Then
                    'if a wirebody is found, check if the SlotEdge wire overlaps with Profile
                    oSurfUtilities.FindWireOverlapWithTolerance oPenetratingSurface, oSlotEdgeWireOverlapPenetrated, 0.004, Nothing, Nothing, oSlotEdgeWireOverlapPenetrating
                    If Not oSlotEdgeWireOverlapPenetrating Is Nothing Then
                        DoesEdgeOverlapPart = True
                                                Exit Function
                    End If
                End If
                
            End If
        End If
    Next n
    
    Exit Function

End Function
