Attribute VB_Name = "RulesCommon"
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


Private Const MODULE = "StructDetail\Data\SmartOccurrence\MemberAssyConn\RulesCommon"
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
    Dim oStructfeature As IJStructFeature
    
    Dim oSDO_Webcut As StructDetailObjects.WebCut
    
    sMsg = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
    Set oMemberObjects = pMemberDescription.CAO
    Set oMemberItems = oMemberObjects.Item(1)
        For iIndex = 1 To oMemberItems.Count
            If TypeOf oMemberItems.Item(iIndex) Is IJStructFeature Then
                Set oStructfeature = oMemberItems.Item(iIndex)
                If oStructfeature.get_StructFeatureType = SF_WebCut Then
                    If pMemberDescription.index < 4 Then
                        Set oWebCut = oStructfeature
                        Exit For
                    ElseIf iIndex > 1 Then
                        Set oWebCut = oStructfeature
                        Exit For
                    End If
                End If
            End If
        Next iIndex
    
    Set oSDO_Webcut = New StructDetailObjects.WebCut
    Set oSDO_Webcut.object = oWebCut
    oSDO_Webcut.ForceUpdateForFlangeCuts
    
    Set oSDO_Webcut = Nothing
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

    Dim oBoundedPort As IJPort
    Dim oBoundedPart As Object
    
    Dim oBoundingPort As IJPort
    Dim oBoundingPart As Object
    
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort
    Dim oBounding_SplitAxisPort As ISPSSplitAxisPort
    
    Dim oEndCutObject As Object
    Dim oSystemParent As IJSystemChild
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    
    Dim eStructFeature As StructFeatureTypes
    Dim oStructfeature As IJStructFeature
    
    Dim oSDO_Webcut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    
    Set Create_PhysConn = Nothing

    ' Create StructDetailObjects wrapper class for given End Cut
    sMsg = "Setting StructDeatilObjects Wrapper Class for EndCut"
    Set oEndCutObject = oMemberDescription.CAO
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructfeature = oEndCutObject
        eStructFeature = oStructfeature.get_StructFeatureType
        If eStructFeature = SF_WebCut Then
            Set oSDO_Webcut = New StructDetailObjects.WebCut
            Set oSDO_Webcut.object = oEndCutObject
        
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_Webcut.Bounded
            Set oBoundingPart = oSDO_Webcut.Bounding
            
            Set oBounded_SplitAxisPort = oSDO_Webcut.BoundedPort
            Set oBounding_SplitAxisPort = oSDO_Webcut.BoundingPort
        
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBoundingPart = oSDO_FlangeCut.Bounding
        
            Set oBounded_SplitAxisPort = oSDO_FlangeCut.BoundedPort
            Set oBounding_SplitAxisPort = oSDO_FlangeCut.BoundingPort
        
        Else
            sMsg = "EndCut has Unknown StructFeatureType:" & eStructFeature
            GoTo ErrorHandler
        End If
        
    Else
        sMsg = "EndCut is not IJStructFeature ...Type:" & TypeName(oEndCutObject)
        GoTo ErrorHandler
    End If
    
    ' The Bounding Port will be: Lateral, Base, or Offset
    sMsg = "Getting Bounding Port for the Physical Connection"
    Set oBoundingPort = Member_GetSolidPort(oBounding_SplitAxisPort)
    
    ' The Bounded Port might not exist because the Cut has not been applied yet:
    ' Use Late Port Binding to get the Bounded Port
    sMsg = "Getting Bounded Port for the Physical Connection"
    Set oStructProfilePart = oBoundedPart
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
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
Public Sub EndCut_FinalConstruct(ByVal pAggregatorDescription As IJDAggregatorDescription, _
                                 Optional oBounded_SplitAxisPort As ISPSSplitAxisPort = Nothing, _
                                 Optional oBounding_SplitAxisPort As ISPSSplitAxisPort = Nothing, _
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
    Dim oStructfeature As IJStructFeature
    Dim oStructProfilePart As IJStructProfilePart

    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oSDO_Webcut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    If (pAggregatorDescription Is Nothing) Then Exit Sub

    ' Get EndCut to be added to cut
    Set oEndCutObject = pAggregatorDescription.CAO
    If TypeOf oEndCutObject Is IJStructFeature Then
        EndCut_InputData oEndCutObject, oBounded_SplitAxisPort, oBoundedPart, _
                         oBounding_SplitAxisPort, oBoundingPart, eEndCutType
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
    Set oPartObject = oSplitAxisPort.Part
    If TypeOf oPartObject Is IJStructGraphConnectable Then
        Set oStructGraphConnectable = oPartObject
    Else
        sMsg = "Else ... Typeof(oPartObject):" & TypeName(oPartObject)
        GoTo ErrorHandler
    End If
        
    ' Retreive list of Ports from the SPS Member Part's Solid Geometry
    oStructGraphConnectable.enumPortsInGraphByTopologyFilter oPortElements, _
                                                            eFilterType, _
                                                            StableGeometry, _
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
    Dim oStructfeature As IJStructFeature
    Dim oStructProfilePart As IJStructProfilePart
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort

    Dim oSectionAttrbs As IJDAttributes
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oSDO_Webcut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    ' Get EndCut to be added to cut
    EndCut_GetCutDepth = 0.1

    Set oEndCutObject = pPRL.SmartOccurrence
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructfeature = oEndCutObject
        eStructFeature = oStructfeature.get_StructFeatureType
        If eStructFeature = SF_WebCut Then
            eEndCutType = WebCut
            Set oSDO_Webcut = New StructDetailObjects.WebCut
            Set oSDO_Webcut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_Webcut.Bounded
            Set oBounded_SplitAxisPort = oSDO_Webcut.BoundedPort
        
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
            EndCut_GetCutDepth = dCutDepth * 1.5
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
                            oBounding_SplitAxisPort As ISPSSplitAxisPort, _
                            oBoundingPart As Object, _
                            eEndCutType As eEndCutTypes)
Const METHOD = "::EndCut_InputData"

    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sBottomFlange As String
    Dim sSelectionRule As String
    
    Dim eStructFeature As StructFeatureTypes

    Dim oStructfeature As IJStructFeature

    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oSDO_Webcut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    If (oEndCutObject Is Nothing) Then Exit Sub

    ' Get EndCut to be added to cut
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oStructfeature = oEndCutObject
        eStructFeature = oStructfeature.get_StructFeatureType
        If eStructFeature = SF_WebCut Then
            eEndCutType = WebCut
            Set oSDO_Webcut = New StructDetailObjects.WebCut
            Set oSDO_Webcut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_Webcut.Bounded
            Set oBoundingPart = oSDO_Webcut.Bounding
            Set oBounded_SplitAxisPort = oSDO_Webcut.BoundedPort
            Set oBounding_SplitAxisPort = oSDO_Webcut.BoundingPort
            
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBoundingPart = oSDO_FlangeCut.Bounding
            Set oBounded_SplitAxisPort = oSDO_FlangeCut.BoundedPort
            Set oBounding_SplitAxisPort = oSDO_FlangeCut.BoundingPort
            
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


