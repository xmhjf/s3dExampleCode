Attribute VB_Name = "MemberCommon"
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MemberUtilities.bas
'
'Author : D.A. Trent
'
'Description :
'   Utilites for determining Type Of EndCuts to be Placed on Member to Member Assembly Connections
'   most of these Utilities are copied from (they are copied here for convenience)
'   S:\SmartPlantStructure\Symbols\AssemblyConnections\ConnectionUtilities.bas
'
'History:
'*****************************************************************************
Option Explicit


Private Const MODULE = "StructDetail\Data\Include\MemberCommon"
'
Public Const INPUT_BOUNDING = "Bounding"
Public Const INPUT_BOUNDED = "Bounded"
'
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence
'

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error Reporter
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub SetDefaultQuestionsonEndCuts(oCopyAnswerHelper As CopyAnswerHelper, _
                                        sProgId_EndCut As String)
Const METHOD = "::SetDefaultQuestionsonEndCuts"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Setting Default Answers on " & sProgId_EndCut
    oCopyAnswerHelper.PutAnswer sProgId_EndCut, "EndCutType", "W"
    oCopyAnswerHelper.PutAnswer sProgId_EndCut, "WeldPart", "First"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub UpdateWebCutForFlangeCuts(pMemberDescription As IJDMemberDescription)
Const METHOD = "::UpdateWebCutForFlangeCuts"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This appears to allow Assoc to always recompute the WebCut before FlangeCuts
    Dim iIndex As Long
    Dim oWebCut As Object
    Dim oMemberItems As IJElements
    Dim oMemberObjects As IJDMemberObjects
    Dim oStructFeature As IJStructFeature
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    sMsg = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
    Set oMemberObjects = pMemberDescription.CAO
    Set oMemberItems = oMemberObjects.Item(1)
        For iIndex = 1 To oMemberItems.Count
            If TypeOf oMemberItems.Item(iIndex) Is IJStructFeature Then
                Set oStructFeature = oMemberItems.Item(iIndex)
                If oStructFeature.get_StructFeatureType = SF_WebCut Then
                    If pMemberDescription.index < 4 Then
                        Set oWebCut = oStructFeature
                        Exit For
                    ElseIf iIndex > 1 Then
                        Set oWebCut = oStructFeature
                        Exit For
                    End If
                End If
            End If
        Next iIndex
    
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oWebCut
    oSDO_WebCut.ForceUpdateForFlangeCuts
    
    Set oSDO_WebCut = Nothing
    Set oMemberObjects = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Create_PhysConn(oMemberDescription As IJDMemberDescription, _
                                oResourceManager As IUnknown, _
                                sRootSmartClass As String, _
                                eBoundingPort As eUSER_CTX_FLAGS, _
                                eBoundedSubPort As JXSEC_CODE) As Object
Const METHOD = "::Create_PhysConn"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim iIndex As Long
    Dim lOprId As Long
    Dim lIdealEdgeId As Long
    Dim sArrayOfFaceIds() As String
    
    Dim oBoundedPort As IJPort
    Dim oBoundedPart As Object
    
    Dim oBoundingPort As IJPort
    Dim oBoundingPart As Object
    
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort
    Dim oBounding_SplitAxisPort As ISPSSplitAxisPort
    
    Dim oEndCutObject As Object
    Dim oSystemParent As IJSystemChild
    Dim oListMemberFaces As IJElements
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    Dim oBnding_StructProfilePart As IJStructProfilePart
    
    Dim eStructFeature As StructFeatureTypes
    Dim oStructFeature As IJStructFeature
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    
    Set Create_PhysConn = Nothing

    ' Create StructDetailObjects wrapper class for given End Cut
    sMsg = "Setting StructDeatilObjects Wrapper Class for EndCut"
    Set oEndCutObject = oMemberDescription.CAO
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructFeature = oEndCutObject
        eStructFeature = oStructFeature.get_StructFeatureType
        If eStructFeature = SF_WebCut Then
            Set oSDO_WebCut = New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = oEndCutObject
        
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_WebCut.Bounded
            Set oBoundingPart = oSDO_WebCut.Bounding
            
            Set oBounded_SplitAxisPort = oSDO_WebCut.BoundedPort
            
            ' Check if processing WebCut for Bearing Plate case
            ' Bounding Port is NOT ISPSSplitAxisPort
            Set oBoundingPort = oSDO_WebCut.BoundingPort
            If TypeOf oBoundingPort Is ISPSSplitAxisPort Then
                Set oBounding_SplitAxisPort = oSDO_WebCut.BoundingPort
            End If
        
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBoundingPart = oSDO_FlangeCut.Bounding
        
            Set oBounded_SplitAxisPort = oSDO_FlangeCut.BoundedPort
            
            Set oBoundingPort = oSDO_FlangeCut.BoundingPort
            If TypeOf oBoundingPort Is ISPSSplitAxisPort Then
                Set oBounding_SplitAxisPort = oSDO_FlangeCut.BoundingPort
            End If
        
        Else
            sMsg = "EndCut has Unknown StructFeatureType:" & eStructFeature
            GoTo ErrorHandler
        End If
        
    Else
        sMsg = "EndCut is not IJStructFeature ...Type:" & TypeName(oEndCutObject)
        GoTo ErrorHandler
    End If
    
    Set oStructProfilePart = oBoundedPart
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    
    ' The Bounding Port will be: Lateral, Base, or Offset
    sMsg = "Getting Bounding Port for the Physical Connection"
    If eBoundingPort = CTX_NOP Then
        ' Use the Bounding Port in the WebCut inputs
    
    ElseIf eBoundingPort = CTX_VIRTUAL Then
        ' Need to determine Phhysical Connection Port
        '   PlatePart Base or Offset Port
        '   Profile Part Global Lateral Port
        '   Member Part Global Lateral Port
        Virtual_GetBoundingPort oBoundedPart, oBounded_SplitAxisPort, _
                                oBoundingPart, oBoundingPort
    
    ElseIf TypeOf oBoundingPart Is ISPSDesignedMember Then
        ' For SP DesignMember (BuiltUp)
        ' Get the Bounding Idealized Boundary Port
        ' Physical Connections are created only with the Idealized Boundary
        lIdealEdgeId = oStructEndCutUtil.GetIdealizedBoundaryId( _
                                                oBounded_SplitAxisPort, _
                                                oBounding_SplitAxisPort)
        
        Set oBnding_StructProfilePart = oBoundingPart
        oBnding_StructProfilePart.GetSectionFaces False, _
                                                 oListMemberFaces, _
                                                 sArrayOfFaceIds

        ' Search array for matching Port XID to be returned
        If oListMemberFaces.Count > 0 Then
            For iIndex = 1 To oListMemberFaces.Count
                lOprId = Trim(sArrayOfFaceIds(iIndex - 1))
                If lOprId = lIdealEdgeId Then
                    Set oBoundingPort = oListMemberFaces.Item(iIndex)
                    Exit For
                End If
            Next iIndex
        End If
    
    ElseIf Not oBounding_SplitAxisPort Is Nothing Then
        ' Get the Solid Port from the Bounding Member Part
        Set oBoundingPort = Member_GetSolidPort(oBounding_SplitAxisPort)
    End If
    
    ' The Bounded Port might not exist because the Cut has not been applied yet:
    ' Use Late Port Binding to get the Bounded Port
    sMsg = "Getting Bounded Port for the Physical Connection"
    oStructEndCutUtil.GetLatePortForFeatureSegment oEndCutObject, eBoundedSubPort, _
                                                   oBoundedPort
    Set oStructEndCutUtil = Nothing
    
    ' The given End Cut (Web or Flange) is the Parent of the Physical Connection
    sMsg = "Setting system parent to Member Description Custom Assembly"
    Set oSystemParent = oEndCutObject
       
    ' Create physical connection
    sMsg = "Creating Physical Connection"
    Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn
    oSDO_PhysicalConn.Create oResourceManager, oBoundedPort, oBoundingPort, _
                             sRootSmartClass, oSystemParent, ConnectionStandard
                               
    sMsg = "Setting Physical Connection to private variable"
    Set Create_PhysConn = oSDO_PhysicalConn.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub EndCut_FinalConstruct(oCutObject As Object, _
                                 Optional oBounded_SplitAxisPort As ISPSSplitAxisPort = Nothing, _
                                 Optional oBounding_Object As Object = Nothing, _
                                 Optional eEndCutType As eEndCutTypes)
Const METHOD = "::EndCut_FinalConstruct"

    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sBottomFlange As String
    Dim sSelectionRule As String
    
    Dim eStructFeature As StructFeatureTypes

    Dim oBoundedPart As Object
    Dim oBoundingPart As Object
    Dim oEndCutObject As Object
    Dim oStructFeature As IJStructFeature
    Dim oStructProfilePart As IJStructProfilePart

    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition
    Dim oAggregatorDescription As IJDAggregatorDescription
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    ' Check if EndCut or SmartItem given as input
    If TypeOf oCutObject Is IJDAggregatorDescription Then
        ' Get EndCut to be added to cut
        Set oAggregatorDescription = oCutObject
        Set oEndCutObject = oAggregatorDescription.CAO
    Else
        ' assume given object is EndCut to be added
        Set oEndCutObject = oCutObject
    End If

    ' Verify have valid EndCut object
    If TypeOf oEndCutObject Is IJStructFeature Then
        EndCut_InputData oEndCutObject, oBounded_SplitAxisPort, oBoundedPart, _
                         oBounding_Object, oBoundingPart, eEndCutType
    End If
    
    If oBoundedPart Is Nothing Then
        sMsg = "BoundedPart is not valid (Nothing)"
        GoTo ErrorHandler
    
    ElseIf TypeOf oBoundedPart Is IJStructProfilePart Then
        Set oStructProfilePart = oBoundedPart
        oStructProfilePart.PlaceEndCutFeature oEndCutObject, oBounded_SplitAxisPort, eEndCutType
    
    Else
        sMsg = "BoundedPart is not valid Profile/Member object"
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Function IsSplitAngleEndToEndCase(pMD As IJDMemberDescription) As Boolean
Const METHOD = "::IsSplitAngleEndToEndCase"
    
    Dim sMsg As String
    Dim sEndToEndCase As String
    Dim oAppConnection As IJAppConnection
    Dim oCopyAnswerHelper As CopyAnswerHelper
    Dim oCommonHelper As DefinitionHlprs.CommonHelper

    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    On Error GoTo ErrorHandler
    IsSplitAngleEndToEndCase = False
    sEndToEndCase = ""
    
    ' FlangeCut not required for SplitAngle SplitEndToEndCase
    sMsg = "Retreive current Answer for SplitEndCutType from Selector"
    If Not pMD.CAO Is Nothing Then
        Set oSmartOccurrence = pMD.CAO
        
        If Not oSmartOccurrence.SmartItemObject Is Nothing Then
            Set oSmartItem = oSmartOccurrence.SmartItemObject
            
            If Not oSmartItem.Parent Is Nothing Then
                Set oSmartClass = oSmartItem.Parent
                
                If Not oSmartClass.SelectionRuleDef Is Nothing Then
                    Set oSymbolDefinition = oSmartClass.SelectionRuleDef
                    
                    If Not oSymbolDefinition Is Nothing Then
                        Set oCommonHelper = New DefinitionHlprs.CommonHelper
                        sEndToEndCase = oCommonHelper.GetAnswer(oSmartOccurrence, _
                                                                oSymbolDefinition, _
                                                                "SplitEndCutType")
                    End If
                End If
            End If
        End If
    End If

    If LCase(Trim(sEndToEndCase)) = LCase("AngleWebSquareFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebBevelFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebAngleFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("DistanceWebDistanceFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("OffsetWebOffsetFlange") Then
        IsSplitAngleEndToEndCase = True
        Exit Function
    End If
    
    If LCase(Trim(sEndToEndCase)) = LCase("AngleWebSquareFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebBevelFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebAngleFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("DistanceWebDistanceFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("OffsetWebOffsetFlange_Flip") Then
        IsSplitAngleEndToEndCase = True
        Exit Function
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Member_GetSolidPort(oSplitAxisPortObject As Object) As Object
Const METHOD = "::Member_GetSolidPort"
On Error GoTo ErrorHandler

    Dim sMsg As String
    
    Dim lCtxId As Long
    Dim lOprId As Long
    Dim lOptId As Long
    
    Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
    Dim eFilterType As JS_TOPOLOGY_FILTER_TYPE
    Dim eAxisPortIndex As SPSMemberAxisPortIndex
    
    Dim oPort As IJPort
    Dim oPartObject As Object
    Dim oPortObject As Object
    Dim oPortGeometry As Object
    Dim oPortElements As IJElements
    
    Dim oSplitAxisPort As ISPSSplitAxisPort
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    Dim oStructGraphConnectable As IJStructGraphConnectable
    
    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    
    sMsg = "Verifying oSplitAxisPort argument is ISPSSplitAxisPort"
    Set Member_GetSolidPort = Nothing
    If TypeOf oSplitAxisPortObject Is ISPSSplitAxisPort Then
        Set oSplitAxisPort = oSplitAxisPortObject
    Else
        GoTo ErrorHandler
    End If
    
    ' Verify given SplitAxisPort Type
    eAxisPortIndex = oSplitAxisPort.portIndex
    If eAxisPortIndex = SPSMemberAxisStart Then
        eFilterType = JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE
    ElseIf eAxisPortIndex = SPSMemberAxisEnd Then
        eFilterType = JS_TOPOLOGY_FILTER_SOLID_OFFSET_LFACE
    ElseIf eAxisPortIndex = SPSMemberAxisAlong Then
        eFilterType = JS_TOPOLOGY_FILTER_LCONNECT_PRT_LATERAL_LFACES
    Else
        sMsg = "eAxisPortIndex Is NOT Known Type"
        GoTo ErrorHandler
    End If
        
    ' Verify given SplitAxisPort Part
    sMsg = "Verifying oSplitAxisPort.Part is IJStructGraphConnectable"
    Set oPort = oSplitAxisPort
    Set oPartObject = oPort.Connectable
    If TypeOf oPartObject Is IJStructGraphConnectable Then
        Set oStructGraphConnectable = oPartObject
    
    ElseIf TypeOf oPartObject Is ISPSDesignedMember Then
        sMsg = "Else ... Typeof(oPartObject):" & TypeName(oPartObject)
        GoTo ErrorHandler
        
    Else
        sMsg = "Else ... Typeof(oPartObject):" & TypeName(oPartObject)
        GoTo ErrorHandler
    End If
        
    ' Retreive list of Ports from the SPS Member Part's Solid Geometry
    oStructGraphConnectable.enumPortsInGraphByTopologyFilter oPortElements, _
                                                            eFilterType, _
                                                            CurrentGeometry, _
                                                            vbNull
    ' Verify returned List of Port(s) is valid
    If oPortElements Is Nothing Then
        sMsg = "oPortElements Is Nothing"
    ElseIf oPortElements.Count < 1 Then
        sMsg = "oPortElements.Count < 1"
    Else
        Set oPortObject = oPortElements.Item(1)
    
        ' Verify Port to be returned is IJSurface Body
        Set oMemberFactory = New SPSMembers.SPSMemberFactory
        Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
        
        oMemberConnectionServices.GetStructPortInfo oPortObject, _
                                                    ePortType, lCtxId, lOptId, lOprId
        If TypeOf oPortObject Is IJPort Then
            Set oPort = oPortObject
            Set oPortGeometry = oPort.Geometry
            If TypeOf oPortGeometry Is IJSurfaceBody Then
                Set Member_GetSolidPort = oPortObject
            Else
                sMsg = "Else... TypeOf oPortGeometry Is IJSurfaceBody"
                GoTo ErrorHandler
            End If
        Else
            sMsg = "Else... TypeOf oPortObject Is IJPort"
            GoTo ErrorHandler
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
            
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function EndCut_GetCutDepth(ByVal pPRL As IJDParameterLogic) As Double
Const METHOD = "::EndCut_GetCutDepth"

    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sBottomFlange As String
    Dim sSelectionRule As String
    
    Dim dtWeb As Double
    Dim dDepth As Double
    Dim dWidth As Double
    Dim dtFlange As Double
    Dim dCutDepth As Double
    
    Dim eEndCutType As eEndCutTypes
    Dim eStructFeature As StructFeatureTypes

    Dim oBoundedPart As Object
    Dim oEndCutObject As Object
    Dim oStructFeature As IJStructFeature
    Dim oStructProfilePart As IJStructProfilePart
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort

    Dim oSectionAttrbs As IJDAttributes
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    ' Get EndCut to be added to cut
    EndCut_GetCutDepth = 0.1

    Set oEndCutObject = pPRL.SmartOccurrence
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructFeature = oEndCutObject
        eStructFeature = oStructFeature.get_StructFeatureType
        If eStructFeature = SF_WebCut Then
            eEndCutType = WebCut
            Set oSDO_WebCut = New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_WebCut.Bounded
            Set oBounded_SplitAxisPort = oSDO_WebCut.BoundedPort
        
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBounded_SplitAxisPort = oSDO_FlangeCut.BoundedPort
            
            Set oSmartOccurrence = oEndCutObject
            Set oSmartItem = oSmartOccurrence.SmartItemObject
            Set oSmartClass = oSmartItem.Parent
            sSelectionRule = oSmartClass.SelectionRule
            Set oSymbolDefinition = oSmartClass.SelectionRuleDef
    
            Set oCommonHelper = New DefinitionHlprs.CommonHelper
            sBottomFlange = oCommonHelper.GetAnswer(oSmartOccurrence, oSymbolDefinition, "BottomFlange")
            If Trim(LCase(sBottomFlange)) = LCase("No") Then
                eEndCutType = FlangeCutTop
            Else
                eEndCutType = FlangeCutBottom
            End If
            
        End If

    End If
    
    If oBoundedPart Is Nothing Then
        sMsg = "BoundedPart is not valid (Nothing)"
        GoTo ErrorHandler
    
    ElseIf TypeOf oBoundedPart Is ISPSMemberPartPrismatic Then
        Set oMemberPartPrismatic = oBoundedPart
        Set oSectionAttrbs = oMemberPartPrismatic.CrossSection.Definition
        
        'for now assume that the section type is W
        'get section attributes (some sections don't support the interfaces below)
        On Error Resume Next
        sMsg = "Retreiving CrossSection Attributes"
        dDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
        dWidth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    
        dtWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
        dtFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
        
        ' Determine the largest value and use it for the CutDepth
        On Error GoTo ErrorHandler
        dCutDepth = dDepth
        If dWidth > dCutDepth Then
            dCutDepth = dWidth
        End If
        
        If dtWeb > dCutDepth Then
            dCutDepth = dtWeb
        End If
        
        If dtFlange > dCutDepth Then
            dCutDepth = dtFlange
        End If
        
        If dCutDepth > 0.01 Then
            EndCut_GetCutDepth = dCutDepth * 3.5
        End If
    Else
        sMsg = "BoundedPart is not valid Profile/Member object"
        GoTo ErrorHandler
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub EndCut_InputData(oEndCutObject As Object, _
                            oBounded_SplitAxisPort As ISPSSplitAxisPort, _
                            oBoundedPart As Object, _
                            oBounding_Object As Object, _
                            oBoundingPart As Object, _
                            eEndCutType As eEndCutTypes)
Const METHOD = "::EndCut_InputData"

    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sBottomFlange As String
    Dim sSelectionRule As String
    
    Dim eStructFeature As StructFeatureTypes

    Dim oPort As IJPort
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim oBearingPlate As IJSmartPlate
    Dim oStructFeature As IJStructFeature
    Dim oGraphicInputs As JCmnShp_CollectionAlias

    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    Dim oSPAttributes As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    
    If (oEndCutObject Is Nothing) Then Exit Sub

    ' Get EndCut to be added to cut
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructFeature = oEndCutObject
        eStructFeature = oStructFeature.get_StructFeatureType
        If eStructFeature = SF_WebCut Then
            eEndCutType = WebCut
            Set oSDO_WebCut = New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_WebCut.Bounded
            Set oBoundingPart = oSDO_WebCut.Bounding
            Set oBounded_SplitAxisPort = oSDO_WebCut.BoundedPort
            Set oBounding_Object = oSDO_WebCut.BoundingPort
            
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBoundingPart = oSDO_FlangeCut.Bounding
            Set oBounded_SplitAxisPort = oSDO_FlangeCut.BoundedPort
            Set oBounding_Object = oSDO_FlangeCut.BoundingPort
            
            Set oSmartOccurrence = oEndCutObject
            Set oSmartItem = oSmartOccurrence.SmartItemObject
            Set oSmartClass = oSmartItem.Parent
            sSelectionRule = oSmartClass.SelectionRule
            Set oSymbolDefinition = oSmartClass.SelectionRuleDef
    
            Set oCommonHelper = New DefinitionHlprs.CommonHelper
            sBottomFlange = oCommonHelper.GetAnswer(oSmartOccurrence, oSymbolDefinition, "BottomFlange")
            If Trim(LCase(sBottomFlange)) = LCase("No") Then
                eEndCutType = FlangeCutTop
            Else
                eEndCutType = FlangeCutBottom
            End If
            
        End If

    ElseIf TypeOf oEndCutObject Is IJSmartPlate Then
        Set oBearingPlate = oEndCutObject
        Set oSPAttributes = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        oSPAttributes.GetInputs_BearingPlate oBearingPlate, oGraphicInputs
        Set oBounding_Object = oGraphicInputs.Item(1)
        Set oBounded_SplitAxisPort = oGraphicInputs.Item(2)
        
        Set oBoundedPart = oBounded_SplitAxisPort.Part
    
        If TypeOf oBounding_Object Is ISPSSplitAxisPort Then
            Set oSplitAxisPort = oBounding_Object
            Set oBoundingPart = oSplitAxisPort.Part
            
        ElseIf TypeOf oBounding_Object Is IJPort Then
            Set oPort = oBounding_Object
            Set oBoundingPart = oPort.Connectable
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Function Modify_PhysConn(oMemberDescription As IJDMemberDescription, _
                                eBoundingPort As eUSER_CTX_FLAGS, _
                                eBoundedSubPort As JXSEC_CODE) As Object
Const METHOD = "::Modify_PhysConn"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim lXId As Long
    Dim lCtxId As Long
    Dim lOptId As Long
    Dim lOprId As Long
    Dim lPortType As Long
    
    Dim oBoundedPort As Object
    Dim oBoundingPort As Object
    Dim oBoundedPortNew As Object
    
    Dim oPortMoniker As IUnknown
    Dim oSP3D_StructPort As SP3DStructPorts.IJStructPort
    Dim oStructGeomBasicPort As StructGeomBasicPort
    
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart

    Dim oPortHelper As PORTHELPERLib.PortHelper
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
    
    ' Check/Verify Bounded Port has required Xid
    sMsg = "Retreiving Physical Connection Ports"
    Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn
    Set oSDO_PhysicalConn.object = oMemberDescription.object
    
    Set oBoundedPort = oSDO_PhysicalConn.Port1
    Set oBoundingPort = oSDO_PhysicalConn.Port2
    
    sMsg = "Decoding Physical Connection Bounded Port Moniker data"
    Set oStructGeomBasicPort = oBoundedPort
    Set oSP3D_StructPort = oStructGeomBasicPort
    Set oPortMoniker = oSP3D_StructPort.PortMoniker
    
    Set oPortHelper = New PORTHELPERLib.PortHelper
    oPortHelper.DecodeTopologyProxyMoniker oPortMoniker, _
                                           lPortType, lCtxId, lOptId, lOprId, lXId
    If eBoundedSubPort = lXId Then
        Exit Function
    End If
    
    ' The Bounded Port might not exist because the Cut has not been applied yet:
    ' Use Late Port Binding to get the Bounded Port
    sMsg = "Getting Bounded Port for the Physical Connection"
    Set oStructProfilePart = oSDO_PhysicalConn.ConnectedObject1
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    oStructEndCutUtil.GetLatePortForFeatureSegment oMemberDescription.CAO, _
                                                   eBoundedSubPort, oBoundedPortNew
    Set oStructEndCutUtil = Nothing
    
    ' Update the Physical Connection Port
    Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
    oConnectionDefinition.ReplacePhysicalConnectionPort oMemberDescription.object, _
                                                        oBoundedPort, oBoundedPortNew
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Function EndCut_GetTubeDepth(ByVal pPRL As IJDParameterLogic) As Double
Const METHOD = "::EndCut_GetTubeDepth"

    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sBottomFlange As String
    Dim sSelectionRule As String
    
    Dim dtWeb As Double
    Dim dDepth As Double
    Dim dWidth As Double
    Dim dtFlange As Double
    Dim dCutDepth As Double
    
    Dim eEndCutType As eEndCutTypes
    Dim eStructFeature As StructFeatureTypes

    Dim oBoundedPart As Object
    Dim oEndCutObject As Object
    Dim oStructFeature As IJStructFeature
    Dim oStructProfilePart As IJStructProfilePart
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort

    Dim oSectionAttrbs As IJDAttributes
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    ' Get EndCut to be added to cut
    EndCut_GetTubeDepth = 0.1

    Set oEndCutObject = pPRL.SmartOccurrence
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructFeature = oEndCutObject
        eStructFeature = oStructFeature.get_StructFeatureType
        If eStructFeature = SF_WebCut Then
            eEndCutType = WebCut
            Set oSDO_WebCut = New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_WebCut.Bounded
            Set oBounded_SplitAxisPort = oSDO_WebCut.BoundedPort
        
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBounded_SplitAxisPort = oSDO_FlangeCut.BoundedPort
            
            Set oSmartOccurrence = oEndCutObject
            Set oSmartItem = oSmartOccurrence.SmartItemObject
            Set oSmartClass = oSmartItem.Parent
            sSelectionRule = oSmartClass.SelectionRule
            Set oSymbolDefinition = oSmartClass.SelectionRuleDef
    
            Set oCommonHelper = New DefinitionHlprs.CommonHelper
            sBottomFlange = oCommonHelper.GetAnswer(oSmartOccurrence, oSymbolDefinition, "BottomFlange")
            If Trim(LCase(sBottomFlange)) = LCase("No") Then
                eEndCutType = FlangeCutTop
            Else
                eEndCutType = FlangeCutBottom
            End If
            
        End If

    End If
    
    If oBoundedPart Is Nothing Then
        sMsg = "BoundedPart is not valid (Nothing)"
        GoTo ErrorHandler
    
    ElseIf TypeOf oBoundedPart Is ISPSMemberPartPrismatic Then
        Set oMemberPartPrismatic = oBoundedPart
        Set oSectionAttrbs = oMemberPartPrismatic.CrossSection.Definition
        
        'for now assume that the section type is W
        'get section attributes (some sections don't support the interfaces below)
        On Error Resume Next
        sMsg = "Retreiving CrossSection Attributes"
        dDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
        
        ' Determine the largest value and use it for the CutDepth
        On Error GoTo ErrorHandler
        
        If dDepth > 0.01 Then
            EndCut_GetTubeDepth = dDepth
        End If
    Else
        sMsg = "BoundedPart is not valid Profile/Member object"
        GoTo ErrorHandler
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub BearingPlate_FinalConstruct(ByVal pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "::BearingPlate_FinalConstruct"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' =========================================================================
' =========================================================================
' =========================================================================
Public Sub Virtual_GetBoundingPort(oBoundedPart As Object, _
                                   oBoundedSplitAxisPort As Object, _
                                   oBoundingPart As Object, _
                                   oBoundingPort As Object)
Const METHOD = "::Virtual_GetBoundingPort"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lStatus As Long
    Dim nPoints As Long
    
    Dim dX As Double
    Dim dY As Double
    Dim dZ As Double
    Dim dDot As Double
    
    Dim dDistTop As Double
    Dim dDistBottom As Double
    
    Dim oBasePort As IJPort
    Dim oOffsetPort As IJPort
    
    Dim oNormal As IJDVector
    Dim oBaseNormal As IJDVector
    Dim oOffsetNormal As IJDVector
    
    Dim oEndPoint As IJPoint
    Dim oMidPoint As IJDPosition
    Dim oBasePoint As IJDPosition
    Dim oOffsetPoint As IJDPosition
    Dim oTopPosition As IJDPosition
    Dim oBottomPosition As IJDPosition
    Dim oProjectingPoint As IJDPosition
    
    Dim oSplitAxisPort As ISPSSplitAxisPort
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oBoundedWebFlangePoints As Collection
    
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    
    If TypeOf oBoundingPart Is ISPSMemberPartPrismatic Then
        Set oMemberPartPrismatic = oBoundingPart
        Set oSplitAxisPort = oMemberPartPrismatic.AxisPort(SPSMemberAxisAlong)
        Set oBoundingPort = Member_GetSolidPort(oSplitAxisPort)
        
    ElseIf TypeOf oBoundingPart Is IJProfilePart Then
        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oBoundingPart
        Set oBoundingPort = oSDO_ProfilePart.BasePort(BPT_Lateral)
    
    ElseIf TypeOf oBoundingPart Is IJPlatePart Then
        InitEndCutConnectionData oBoundedSplitAxisPort, oBoundingPort, _
                                 oBoundedData, oBoundingData, lStatus, sMsg
        
        ' Get the x,y,z end location of the Bounded Member Part
        Set oSplitAxisPort = oBoundedSplitAxisPort
        Set oMemberPartPrismatic = oBoundedPart
        Set oEndPoint = oMemberPartPrismatic.PointAtEnd(oSplitAxisPort.portIndex)
        oEndPoint.GetPoint dX, dY, dZ
        
        Set oProjectingPoint = New AutoMath.DPosition
        oProjectingPoint.Set dX, dY, dZ
        
        ' Retrieve the 3D points representing the Bounded Member Flange Sections
        ' Top, Top_Flange_Bottom, Bottom_Flange_top, and Bottom locations
        Set oBoundedWebFlangePoints = GetMemberWebFlangePoints(oBoundedPart, _
                                                        oBoundedData.Matrix, _
                                                        eIdealized_WebLeft)
        nPoints = oBoundedWebFlangePoints.Count
        If nPoints < 1 Then
            Set oBoundingPort = oBasePort
            Exit Sub
        End If
        
        Set oTopPosition = oBoundedWebFlangePoints.Item(1)
        Set oBottomPosition = oBoundedWebFlangePoints.Item(nPoints)
        
        ' Get the Base/Offset Ports for the Bounding Plate Part
        Set oSDO_PlatePart = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart.object = oBoundingPart
        Set oBasePort = oSDO_PlatePart.BasePort(BPT_Base)
        Set oOffsetPort = oSDO_PlatePart.BasePort(BPT_Offset)
    
        ' Project the Bounded end Location onto Plate Part Base geometry
        Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
        oTopologyLocate.GetProjectedPointOnModelBody oBasePort.Geometry, _
                                                     oProjectingPoint, _
                                                     oBasePoint, _
                                                     oBaseNormal
        
        ' Project the Bounded end Location onto Plate Part Offset geometry
        oTopologyLocate.GetProjectedPointOnModelBody oOffsetPort.Geometry, _
                                                     oProjectingPoint, _
                                                     oOffsetPoint, _
                                                     oOffsetNormal
        
        ' Determine the Mid-Point between the Plate Part Base and Offset geometry
        Set oMidPoint = New AutoMath.DPosition
        oMidPoint.Set (oBasePoint.x + oOffsetPoint.x) / 2#, _
                      (oBasePoint.y + oOffsetPoint.y) / 2#, _
                      (oBasePoint.z + oOffsetPoint.z) / 2#
        
        ' Determine if Bounding Plate Port is closest to the Bounded Top or Bottom
        dDistTop = oMidPoint.DistPt(oTopPosition)
        dDistBottom = oMidPoint.DistPt(oBottomPosition)
        
        If dDistTop < dDistBottom Then
            ' Bounding Plate Port is closest to Bounded Top point
            Set oNormal = oBottomPosition.Subtract(oTopPosition)
            dDot = oNormal.Dot(oBaseNormal)

            If dDot < 0# Then
                Set oBoundingPort = oOffsetPort
            Else
                Set oBoundingPort = oBasePort
            End If
        Else
            ' Bounding Plate Port is closest to Bounded Bottom point
            Set oNormal = oTopPosition.Subtract(oBottomPosition)
            dDot = oNormal.Dot(oBaseNormal)
            
            If dDot < 0# Then
                Set oBoundingPort = oOffsetPort
            Else
                Set oBoundingPort = oBasePort
            End If
        
        End If

    End If
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

