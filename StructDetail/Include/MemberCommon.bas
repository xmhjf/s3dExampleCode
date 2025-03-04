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
'25/Aug/2011   - Addedd new methods to create InsetBrace
'    15Sep2011  -   pnalugol
'               Modified/Added methods to support Braces on Flange penetrated cases
'    16Sep2011  -   pnalugol
'    Modified/Added methods to support Flush Braces
'    18/April/2012 - CM - TR212654 Modfied Create_PhysConn() to take the last bounding
'                         port as input for any PC creation.
'    27/Jul/2012 - svsmylav
'      CR-216426: Added a new method 'IsCornerFeatureNeededForShapeAtEdge'.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
' 21/July/2015 - GH  SI-CP-275688    Investigate split migration failure
'                   Updated EndCut_InputData()
'*****************************************************************************
Option Explicit


Private Const MODULE = "StructDetail\Data\Include\MemberCommon"
'
Public Const INPUT_BOUNDING = "Bounding"
Public Const INPUT_BOUNDED = "Bounded"
'
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence
Dim oTopologyLocate As New GSCADStructGeomUtilities.TopologyLocate
    
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
'Public Sub UpdateWebCutForFlangeCuts(pMemberDescription As IJDMemberDescription)
'Const METHOD = "::UpdateWebCutForFlangeCuts"
'    On Error GoTo ErrorHandler
'    Dim sMsg As String
'
'    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
'    ' as is used when placing the WebCut as an input to the FlangeCuts
'    ' This appears to allow Assoc to always recompute the WebCut before FlangeCuts
'    Dim iIndex As Long
'    Dim oWebCut As Object
'    Dim oMemberItems As IJElements
'    Dim oMemberObjects As IJDMemberObjects
'    Dim oStructFeature As IJStructFeature
'
'    Dim oSDO_WebCut As StructDetailObjects.WebCut
'
'    sMsg = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
'    Set oMemberObjects = pMemberDescription.CAO
'    Set oMemberItems = oMemberObjects.Item(1)
'        For iIndex = 1 To oMemberItems.Count
'            If TypeOf oMemberItems.Item(iIndex) Is IJStructFeature Then
'                Set oStructFeature = oMemberItems.Item(iIndex)
'                If oStructFeature.get_StructFeatureType = SF_WebCut Then
'                    If pMemberDescription.Index < 4 Then
'                        Set oWebCut = oStructFeature
'                        Exit For
'                    ElseIf iIndex > 1 Then
'                        Set oWebCut = oStructFeature
'                        Exit For
'                    End If
'                End If
'            End If
'        Next iIndex
'
'    Set oSDO_WebCut = New StructDetailObjects.WebCut
'    Set oSDO_WebCut.object = oWebCut
'    oSDO_WebCut.ForceUpdateForFlangeCuts
'
'    Set oSDO_WebCut = Nothing
'    Set oMemberObjects = Nothing
'
'    Exit Sub
'
'ErrorHandler:
'    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
'End Sub

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

    Set Create_PhysConn = Nothing

    ' ----------------------------------------------------------
    ' Create StructDetailObjects wrapper class for given End Cut
    ' ----------------------------------------------------------
    sMsg = "Setting StructDeatilObjects Wrapper Class for EndCut"
    
    Dim oEndCutObject As Object
    Set oEndCutObject = oMemberDescription.CAO
    
    ' ------------------------------------------
    ' Fail, if the input object is not a feature
    ' ------------------------------------------
    If Not TypeOf oEndCutObject Is IJStructFeature Then
        sMsg = "EndCut is not IJStructFeature ...Type:" & TypeName(oEndCutObject)
        GoTo ErrorHandler
    End If
    
    ' ----------------------------------
    ' Get the bounded and bounding ports
    ' ----------------------------------
    Dim oBoundedPart As Object
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    Dim oBoundingPart As Object
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort
    Dim oBounding_SplitAxisPort As ISPSSplitAxisPort
    
    Dim oStructFeature As IJStructFeature
    Dim eStructFeature As StructFeatureTypes
    
    Set oStructFeature = oEndCutObject
    eStructFeature = oStructFeature.get_StructFeatureType
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    
        If eStructFeature = SF_WebCut Then
            Set oSDO_WebCut.object = oEndCutObject
        
            sMsg = "Getting Bounded/Bounding objects from WebCut"
            Set oBoundedPart = oSDO_WebCut.Bounded
            Set oBoundingPart = oSDO_WebCut.Bounding
            
            Set oBoundedPort = oSDO_WebCut.BoundedPort
            If TypeOf oSDO_WebCut.BoundedPort Is ISPSSplitAxisPort Then
            Set oBounded_SplitAxisPort = oSDO_WebCut.BoundedPort
            End If
            
            ' Check if processing WebCut for Bearing Plate case
            ' Bounding Port is NOT ISPSSplitAxisPort
            Set oBoundingPort = oSDO_WebCut.BoundingPort
            If TypeOf oBoundingPort Is ISPSSplitAxisPort Then
                Set oBounding_SplitAxisPort = oSDO_WebCut.BoundingPort
            End If
        
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut.object = oEndCutObject
        Set oSDO_WebCut.object = oSDO_FlangeCut.WebCut
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
            Set oBoundingPart = oSDO_FlangeCut.Bounding
        
            Set oBoundedPort = oSDO_FlangeCut.BoundedPort
            If TypeOf oSDO_FlangeCut.BoundedPort Is ISPSSplitAxisPort Then
            Set oBounded_SplitAxisPort = oSDO_FlangeCut.BoundedPort
            End If
        ' Check if processing FlangeCut for Bearing Plate case
        ' Bounding Port is NOT ISPSSplitAxisPort
            Set oBoundingPort = oSDO_FlangeCut.BoundingPort
            If TypeOf oBoundingPort Is ISPSSplitAxisPort Then
                Set oBounding_SplitAxisPort = oSDO_FlangeCut.BoundingPort
            End If
        Else
            sMsg = "EndCut has Unknown StructFeatureType:" & eStructFeature
            GoTo ErrorHandler
        End If
        
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oStructProfilePart As IJStructProfilePart
    Set oStructProfilePart = oBoundedPart
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    
    ' The Bounding Port will be: Lateral, Base, or Offset
    sMsg = "Getting Bounding Port for the Physical Connection"
    
    ' --------------------------------------------------------------------------
    ' If the input context is "NOP", use the bounding port for the input feature
    ' --------------------------------------------------------------------------
    If eBoundingPort = CTX_NOP Then
        ' Use the Bounding Port in the WebCut inputs
    
    ' ----------------------------------------------------------------------
    ' If the input context is "VIRTUAL", determine which port should be used
    ' ----------------------------------------------------------------------
    ElseIf eBoundingPort = CTX_VIRTUAL Then
        ' Need to determine Phhysical Connection Port
        '   PlatePart Base or Offset Port
        '   Profile Part Global Lateral Port
        '   Member Part Global Lateral Port
        Virtual_GetBoundingPort oBoundedPart, oBounded_SplitAxisPort, _
                                oBoundingPart, oBoundingPort
    ' --------------------------------
    ' If specifically to the base port
    ' --------------------------------
    ElseIf eBoundingPort = CTX_BASE Then
        Set oBoundingPort = GetBaseOffsetOrLateralPortForObject(oBoundingPart, BPT_Base)
    ' ----------------------------------
    ' If specifically to the offset port
    ' ----------------------------------
    ElseIf eBoundingPort = CTX_OFFSET Then
        Set oBoundingPort = GetBaseOffsetOrLateralPortForObject(oBoundingPart, BPT_Offset)
    ' ----------------------------------
    ' If specifically to the offset port
    ' ----------------------------------
    ElseIf eBoundingPort = CTX_LATERAL Then
        Set oBoundingPort = GetBaseOffsetOrLateralPortForObject(oBoundingPart, BPT_Lateral)
    ' ------------------------------------------------------------------------------
    ' If none of the above, and is a design member, determine the idealized boundary
    ' ------------------------------------------------------------------------------
    ElseIf TypeOf oBoundingPart Is ISPSDesignedMember Then
        ' For SP DesignMember (BuiltUp)
        ' Get the Bounding Idealized Boundary Port
        ' Physical Connections are created only with the Idealized Boundary
        Dim lIdealEdgeId As Long
        lIdealEdgeId = oStructEndCutUtil.GetIdealizedBoundaryId( _
                                                oBounded_SplitAxisPort, _
                                                oBounding_SplitAxisPort)
        
        Dim oBnding_StructProfilePart As IJStructProfilePart
        Set oBnding_StructProfilePart = oBoundingPart
        
        Dim sArrayOfFaceIds() As String
        Dim oListMemberFaces As IJElements
        oBnding_StructProfilePart.GetSectionFaces False, _
                                                 oListMemberFaces, _
                                                 sArrayOfFaceIds

        ' Search array for matching Port XID to be returned
        Dim iIndex As Long
        Dim lOprId As Long
        
        If oListMemberFaces.Count > 0 Then
            For iIndex = 1 To oListMemberFaces.Count
                lOprId = Trim(sArrayOfFaceIds(iIndex - 1))
                If lOprId = lIdealEdgeId Then
                    Set oBoundingPort = oListMemberFaces.Item(iIndex)
                    Exit For
                End If
            Next iIndex
        End If
    ' ------------------------------------------------------------------------
    ' If none of the above and is a standard member, get the global solid port
    ' ------------------------------------------------------------------------
    ElseIf Not oBounding_SplitAxisPort Is Nothing Then
        ' Get the global solid port
        Set oBoundingPort = Member_GetSolidPort(oBounding_SplitAxisPort)
    End If
    
    ' ------------------------------------------------
    ' Get the late-binding port for the bounded object
    ' ------------------------------------------------
    ' The Bounded Port might not exist because the Cut has not been applied yet
    sMsg = "Getting Bounded Port for the Physical Connection"
    
    Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
    If TypeOf oBoundedPart Is ISPSMemberPartCommon Then
    oStructEndCutUtil.GetLatePortForFeatureSegment oEndCutObject, eBoundedSubPort, oBoundedPort
    ElseIf TypeOf oBoundedPart Is IJProfile Then
        Set oBoundedProfilePart = New StructDetailObjects.ProfilePart
        Set oBoundedProfilePart.object = oBoundedPart
        Set oBoundedPort = Nothing
        Set oBoundedPort = oBoundedProfilePart.CutoutSubPort(oEndCutObject, eBoundedSubPort)
    End If
    
    'check if the cut is on brace with the Bounded of the parent AC.
    'in such case the Inset cope needs the base/offset ports of the child brace and Bounded member(of the parent AC) to create the PC.
    
    Dim oParentObj As Object
    Dim oDesignChild As IJDesignChild
    Set oDesignChild = oMemberDescription.CAO
    Set oParentObj = oDesignChild.GetParent
    
    Set oDesignChild = Nothing
    
    Dim bCutBTWBraceAndBounded As Boolean, bCutBTWBraceAndBounding As Boolean
    Dim oStructPort As IJStructPort
    
    If TypeOf oParentObj Is IJStructAssemblyConnection Then
        
        IsCutOnInsetBrace oParentObj, bCutBTWBraceAndBounded, bCutBTWBraceAndBounding
        
        Dim oAppConn As IJAppConnection
        Set oAppConn = oParentObj
        
        If bCutBTWBraceAndBounded Then
                Dim bIsTopBrace As Boolean: bIsTopBrace = False
                Dim sBraceType As String
                
                Set oStructPort = oSDO_WebCut.BoundingPort
                If oStructPort.SectionID = 514 Then
                    bIsTopBrace = True
                End If
                
                Set oDesignChild = oBoundedPart
                Dim oParentAC As Object
                Set oParentAC = oDesignChild.GetParent
                
                If bIsTopBrace Then
                    GetSelectorAnswer oParentAC, "TopBraceType", sBraceType
                Else
                    GetSelectorAnswer oParentAC, "BottomBraceType", sBraceType
                End If
                
                If sBraceType = "InsetMember" Then
                    'we use cope cut. pass the base/offset ports of bounded and bounding
                    GetBaseOrOffsetPortsAtAC oParentAC, oBoundingPort, oBoundedPort
                    oStructEndCutUtil.GetLatePortForFeatureSegment oEndCutObject, eBoundedSubPort, oBoundedPort
                End If
        End If
    End If
    
    ' ------------------------------
    ' Get the Last bounding port
    ' which acts as an input to the
    ' PC creation
    ' ------------------------------
    Dim oHelper As StructDetailObjects.Helper
    Dim oLastBoundingPort As IJPort
    
    Set oHelper = New StructDetailObjects.Helper
    Set oLastBoundingPort = oHelper.GetEquivalentLastPort(oBoundingPort)
    
    If oLastBoundingPort Is Nothing Then
        Set oLastBoundingPort = oBoundingPort
    End If
        
    ' ----------------------------------
    ' Set the input object as the parent
    ' ----------------------------------
    sMsg = "Setting system parent to Member Description Custom Assembly"
    
    Dim oSystemParent As IJSystemChild
    Set oSystemParent = oEndCutObject
       
    ' ------------------------------
    ' Create the physical connection
    ' ------------------------------
    sMsg = "Creating Physical Connection"
    
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn
    
    oSDO_PhysicalConn.Create oResourceManager, oBoundedPort, oLastBoundingPort, _
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
                                 Optional oBounded_SplitAxisPort As Object = Nothing, _
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
Public Function Member_GetSolidPort(oSplitAxisPortObject As Object, Optional bStable As Boolean = False) As Object
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
    Dim oProfilePort As IJStructPort
    Dim eContext_Id As eUSER_CTX_FLAGS
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    Dim oStructGraphConnectable As IJStructGraphConnectable
    
    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    
    sMsg = "Verifying oSplitAxisPort argument is ISPSSplitAxisPort"
    Set Member_GetSolidPort = Nothing
    If TypeOf oSplitAxisPortObject Is ISPSSplitAxisPort Then
        Set oSplitAxisPort = oSplitAxisPortObject
        eAxisPortIndex = oSplitAxisPort.PortIndex
    ElseIf TypeOf oSplitAxisPortObject.Connectable Is IJStructProfilePart Then
        Set oProfilePort = oSplitAxisPortObject
        eContext_Id = oProfilePort.ContextID
    Else
        GoTo ErrorHandler
    End If
    
    ' Verify given SplitAxisPort Type
    If eAxisPortIndex = SPSMemberAxisStart Or eContext_Id = CTX_BASE Then
        eFilterType = JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE
    ElseIf eAxisPortIndex = SPSMemberAxisEnd Or eContext_Id = CTX_OFFSET Then
        eFilterType = JS_TOPOLOGY_FILTER_SOLID_OFFSET_LFACE
    ElseIf eAxisPortIndex = SPSMemberAxisAlong Or eContext_Id = CTX_LATERAL Then
        eFilterType = JS_TOPOLOGY_FILTER_LCONNECT_PRT_LATERAL_LFACES
    Else
        sMsg = "eAxisPortIndex Is NOT Known Type"
        GoTo ErrorHandler
    End If
        
    ' Verify given SplitAxisPort Part
    sMsg = "Verifying oSplitAxisPort.Part is IJStructGraphConnectable"
    Set oPort = oSplitAxisPortObject
    Set oPartObject = oPort.Connectable
    Dim geomOption As StructGeometrySelector
    
    If bStable Then
        geomOption = StableGeometry
    Else
        geomOption = CurrentGeometry
    End If
    
    If TypeOf oPartObject Is IJStructGraphConnectable Then
        Set oStructGraphConnectable = oPartObject
        ' Retreive list of Ports from the SPS Member Part's Solid Geometry
        oStructGraphConnectable.enumPortsInGraphByTopologyFilter oPortElements, _
                                                                eFilterType, _
                                                                geomOption, _
                                                                vbNull
    ElseIf TypeOf oPartObject Is IJStructConnectable Then
        Dim oStructConn As IJStructConnectable
        Dim sOpr As String
        Dim bFlag As Boolean
        Set oStructConn = oPartObject
        ' Retreive list of Ports from the Profiles Part's Solid Geometry
        Dim oEnumUmknown As IEnumUnknown
        oStructConn.enumConnectablePortsByOperationAndTopology oPortElements, sOpr, eFilterType, bFlag

    ElseIf TypeOf oPartObject Is ISPSDesignedMember Then
        sMsg = "Else ... Typeof(oPartObject):" & TypeName(oPartObject)
        GoTo ErrorHandler
        
    Else
        sMsg = "Else ... Typeof(oPartObject):" & TypeName(oPartObject)
        GoTo ErrorHandler
    End If
    
    ' Verify returned List of Port(s) is valid
    If oPortElements Is Nothing Then
        sMsg = "oPortElements Is Nothing"
    ElseIf oPortElements.Count < 1 Then
        sMsg = "oPortElements.Count < 1"
    Else
        Set oPortObject = oPortElements.Item(1)
    
        ' Verify Port to be returned is IJSurface Body
        If TypeOf oPortObject.Connectable Is ISPSMemberPartCommon Then
            Set oMemberFactory = New SPSMembers.SPSMemberFactory
            Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
            oMemberConnectionServices.GetStructPortInfo oPortObject, _
                                                    ePortType, lCtxId, lOptId, lOprId
        End If
        
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
 
        
        ElseIf eStructFeature = SF_FlangeCut Then
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            
            sMsg = "Getting Bounded/Bounding objects from FlangeCut"
            Set oBoundedPart = oSDO_FlangeCut.Bounded
   
            
            Set oSmartOccurrence = oEndCutObject
            Set oSmartItem = oSmartOccurrence.SmartItemObject
            Set oSmartClass = oSmartItem.Parent
            sSelectionRule = oSmartClass.SelectionRule
            Set oSymbolDefinition = oSmartClass.SelectionRuleDef
    
            Set oCommonHelper = New DefinitionHlprs.CommonHelper
            GetSelectorAnswer oSmartOccurrence, "BottomFlange", sBottomFlange
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
        Set oSectionAttrbs = oMemberPartPrismatic.CrossSection.definition
        
        'for now assume that the section type is W
        'get section attributes (some sections don't support the interfaces below)
        On Error Resume Next
        sMsg = "Retreiving CrossSection Attributes"
        dDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
        dWidth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    
        dtWeb = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
        dtFlange = oSectionAttrbs.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tf").Value
        
        Err.Clear
        On Error GoTo ErrorHandler
        
        ' Determine the largest value and use it for the CutDepth
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
    ElseIf TypeOf oBoundedPart Is IJProfile Then
        EndCut_GetCutDepth = 1# 'need to enhance as per need

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
                            oBounded_SplitAxisPort As Object, _
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
    
    Set oBounded_SplitAxisPort = Nothing
    Set oBoundedPart = Nothing
    Set oBounding_Object = Nothing
    Set oBoundingPart = Nothing
    eEndCutType = -1
    
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
            GetSelectorAnswer oSmartOccurrence, "BottomFlange", sBottomFlange
            If Trim(LCase(sBottomFlange)) = LCase("No") Then
                eEndCutType = FlangeCutTop
            Else
                eEndCutType = FlangeCutBottom
            End If
            
        End If

    ElseIf TypeOf oEndCutObject Is IJSmartPlate Then
        ' After bearing is plate is added below Code should be updated for passing proper ports when reversed
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
        Set oSectionAttrbs = oMemberPartPrismatic.CrossSection.definition
        
        'for now assume that the section type is W
        'get section attributes (some sections don't support the interfaces below)
        On Error Resume Next
        sMsg = "Retreiving CrossSection Attributes"
        dDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
        
        Err.Clear
        On Error GoTo ErrorHandler
        
        ' Determine the largest value and use it for the CutDepth
        
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
    
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
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
        Set oEndPoint = oMemberPartPrismatic.PointAtEnd(oSplitAxisPort.PortIndex)
        oEndPoint.GetPoint dx, dy, dz
        
        Set oProjectingPoint = New AutoMath.DPosition
        oProjectingPoint.Set dx, dy, dz
        
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

Public Function GetBaseOffsetOrLateralPortForObject(oObject As Object, portType As Base_Port_Types) As IJPort

    If TypeOf oObject Is IJPlate Then
        Dim oSDOPlate As New StructDetailObjects.PlatePart
        Set oSDOPlate.object = oObject
        
        Set GetBaseOffsetOrLateralPortForObject = oSDOPlate.BasePort(portType)
        
    ElseIf TypeOf oObject Is IJStiffener Then
        Dim oSDOProfile As New StructDetailObjects.ProfilePart
        Set oSDOProfile.object = oObject
        
        Set GetBaseOffsetOrLateralPortForObject = oSDOProfile.BasePort(portType)
        
    ElseIf TypeOf oObject Is ISPSMemberPartPrismatic Then
    
        Dim oSDOMember As New StructDetailObjects.MemberPart
        Set oSDOMember.object = oObject
        
        Set GetBaseOffsetOrLateralPortForObject = oSDOMember.BasePort(portType)
        
    End If
    
End Function

' GetBaseOrOffsetPortsAtAC()
' Given an AssemblyConnection, it determines the Base/Offset ports of the Bounded and Bounding near to the AC and returns them
 Public Sub GetBaseOrOffsetPortsAtAC(oAssemblyConnection As IJStructAssemblyConnection, ByRef pBoundingPort As IJPort, ByRef pBoundedPort As IJPort)
    On Error GoTo ErrorHandler
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    Dim sMsg As String
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = oAssemblyConnection
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Dim oSDO_BoundedMember As New StructDetailObjects.MemberPart
    
    If TypeOf oBoundedData.AxisPort.Connectable Is ISPSMemberPartPrismatic Then
        Set oSDO_BoundedMember.object = oBoundedData.AxisPort.Connectable
    End If
    
    If TypeOf oBoundedData.AxisPort Is ISPSSplitAxisPort Then
        Dim oSplitAxisPort As ISPSSplitAxisPort
        Set oSplitAxisPort = oBoundedData.AxisPort
        If oSplitAxisPort.PortIndex = SPSMemberAxisStart Then
            Set pBoundingPort = oSDO_BoundedMember.BasePort(BPT_Base)
        ElseIf oSplitAxisPort.PortIndex = SPSMemberAxisEnd Then
            Set pBoundingPort = oSDO_BoundedMember.BasePort(BPT_Offset)
        Else
            'UNKNOWN CONDITION!!! Need to handle such cases
            'currently set any default port
            Set pBoundingPort = oSDO_BoundedMember.BasePort(BPT_Base)
        End If
    Else
        'Need to Handle such cases if Bounded Part is of Type Designed Member(anything other than Std Mbr)
    End If
    
    Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "GetBaseOrOffsetPortsAtAC", "Error").Number

 End Sub
'***************************************************************************************************************
'IsCutOnInsetBrace()
'Description: This method returns two booleans: the first one true when the connection is between Insetbrace and Bounded of the parent AC.
'The second boolean returns true when the AC is between Insetbrace and bounding of the Parent AC.
'***************************************************************************************************************
Public Sub IsCutOnInsetBrace(oAssembyConnection As IJStructAssemblyConnection, ByRef bCutBTWBraceAndBounded As Boolean, ByRef bCutBTWBraceAndBounding As Boolean)
 Const METHOD = "MemberCommon::IsCutOnInsetBrace"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    bCutBTWBraceAndBounded = False
    bCutBTWBraceAndBounding = False
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = oAssembyConnection
        
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    Dim lStatus As Long

    InitMemberConnectionData oAssembyConnection, oBoundedData, oBoundingData, lStatus, sMsg

    Dim oBoundedChild As IJDesignChild
    Set oBoundedChild = oBoundedData.MemberPart

    Dim oParentBounded As Object
    Set oParentBounded = oBoundedChild.GetParent

    If TypeOf oParentBounded Is IJStructAssemblyConnection Then
        'Get the bounded/bounding of the AC
        Dim oParentAppConnection As IJAppConnection
        Set oParentAppConnection = oParentBounded

        Dim oParentACBoundedData As MemberConnectionData
        Dim oParentACBoundingData As MemberConnectionData

        InitMemberConnectionData oParentAppConnection, oParentACBoundedData, oParentACBoundingData, lStatus, sMsg
        'see if the current WebCut.Bounding is same as Bounded/Bounding of the Parent AC
        If oBoundingData.MemberPart Is oParentACBoundedData.MemberPart Then
            bCutBTWBraceAndBounded = True
        ElseIf oBoundingData.MemberPart Is oParentACBoundingData.MemberPart Then
            bCutBTWBraceAndBounding = True
        End If
        
    End If

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub GetMemberACTopAndBottomShape(oAssemblyObject As Object, _
                                        Optional sBottomAnswerCol As String, _
                                        Optional sBottomShape As String, _
                                        Optional sTopAnswerCol As String, _
                                        Optional sTopShape As String, _
                                        Optional sFaceTopInsideCornerShape As String, _
                                        Optional sFaceTopInsideCornerCol As String, _
                                        Optional sFaceBtmInsideCornerShape As String, _
                                        Optional sFaceBtmInsideCornerCol As String)

    Const METHOD = "MemberCommon::GetMemberACTopAndBottomShape"
    
    On Error GoTo ErrorHandler
    
    sBottomAnswerCol = vbNullString
    sBottomShape = vbNullString
    sTopAnswerCol = vbNullString
    sTopShape = vbNullString
    sFaceTopInsideCornerShape = vbNullString
    sFaceTopInsideCornerCol = vbNullString
    sFaceBtmInsideCornerShape = vbNullString
    sFaceBtmInsideCornerCol = vbNullString
    
    ' ------------------------
    ' Get the AC item and name
    ' ------------------------
    Dim oAC As Object
    Dim sACItemName As String
    
    If TypeOf oAssemblyObject Is IJAppConnection Then
        Set oAC = oAssemblyObject
        Dim oSO As IJSmartOccurrence
        Set oSO = oAssemblyObject
        sACItemName = oSO.Item
    Else
        AssemblyConnection_SmartItemName oAssemblyObject, sACItemName, oAC
    End If
    
    If oAC Is Nothing Then
        Exit Sub
    End If

    ' ---------------------------
    ' Get the connected edge info
    ' ---------------------------
    Dim cTopOrWL As ConnectedEdgeInfo
    Dim cBtmOrWR As ConnectedEdgeInfo
    Dim cITFOrFL As ConnectedEdgeInfo
    Dim cIBFOrFR As ConnectedEdgeInfo
    Dim bPenetratesWeb As Boolean
    Dim oMeasurements As Collection
    
    Select Case LCase(sACItemName)
    
        Case LCase(gsMbrAxisToEdge), LCase(gsStiffEndToMbrEdge), _
             LCase(gsMbrAxisToEdgeAndOutSide2Edge), LCase(gsStiffEndToMbrEdgeAndOutSide2Edge), _
             LCase(gsMbrAxisToFaceAndEdge), LCase(gsStiffEndToMbrFaceAndEdge), _
             LCase(gsMbrAxisToEdgeAndOutSide1Edge), LCase(gsStiffEndToMbrEdgeAndOutSide1Edge), _
             LCase(gsMbrAxisToOutSideAndOutSide1Edge), LCase(gsStiffEndToMbrOutSideAndOutSide1Edge), _
             LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsStiffEndToMbrFaceAndOutSide1Edge)

            Dim oBoundedData As MemberConnectionData
            Dim oBoundingData As MemberConnectionData
            Dim lStatus As Long
            Dim sMsg As String
            
            InitMemberConnectionData oAC, oBoundedData, oBoundingData, lStatus, sMsg

            Set oMeasurements = New Collection
            
            GetConnectedEdgeInfo oAC, _
                                 oBoundedData.AxisPort, _
                                 oBoundingData.AxisPort, _
                                 cTopOrWL, _
                                 cBtmOrWR, _
                                 cITFOrFL, _
                                 cIBFOrFR, _
                                 oMeasurements, _
                                 bPenetratesWeb
    End Select

    Dim cBtmOuterEdge As ConnectedEdgeInfo
    Dim cBtmInnerEdge As ConnectedEdgeInfo
    
    If bPenetratesWeb Then
        cBtmOuterEdge = cBtmOrWR
        cBtmInnerEdge = cIBFOrFR
    Else
        cBtmOuterEdge = cIBFOrFR
        cBtmInnerEdge = cBtmOrWR
    End If
            
    ' -----------------------------------------
    ' Ask for answers based on the AC item name
    ' -----------------------------------------
    Select Case LCase(sACItemName)
        
        ' ---------------------------------------------------
        ' If the the questions are "TopEdge" and "BottomEdge"
        ' ---------------------------------------------------
        Case LCase(gsMbrAxisToEdgeAndEdge), LCase(gsStiffEndToMbrEdgeAndEdge)
            
            GetSelectorAnswer oAC, "ShapeAtTopEdge", sTopShape
            GetSelectorAnswer oAC, "ShapeAtBottomEdge", sBottomShape
            GetSelectorAnswer oAC, "TopInsideCorner", sFaceTopInsideCornerShape
            GetSelectorAnswer oAC, "BottomInsideCorner", sFaceBtmInsideCornerShape
            
            sBottomAnswerCol = gsShapeAtEdgeCol
            sTopAnswerCol = gsShapeAtEdgeCol
            sFaceTopInsideCornerCol = gsInsideCornerCol
            sFaceBtmInsideCornerCol = gsInsideCornerCol
            
            
        Case LCase(gsMbrAxisToOutSideAndOutSide2Edge), LCase(gsStiffEndToMbrOutSideAndOutSide2Edge)
            
            GetSelectorAnswer oAC, "ShapeAtTopEdge", sTopShape
            GetSelectorAnswer oAC, "ShapeAtBottomEdge", sBottomShape
            GetSelectorAnswer oAC, "TopInsideCorner", sFaceTopInsideCornerShape
            GetSelectorAnswer oAC, "BottomInsideCorner", sFaceBtmInsideCornerShape
            
            sBottomAnswerCol = gsShapeAtEdgeOverlapCol
            sTopAnswerCol = gsShapeAtEdgeOverlapCol
            sFaceTopInsideCornerCol = gsInsideCornerCol
            sFaceBtmInsideCornerCol = gsInsideCornerCol
            
        ' ---------------------------------------------------
        ' If the the questions are "TopFace" and "BottomFace"
        ' ---------------------------------------------------
        Case LCase(gsMbrAxisToCenter), LCase(gsStiffEndToMbrCenter)
            
            GetSelectorAnswer oAC, "TopShapeAtFace", sTopShape
            GetSelectorAnswer oAC, "BottomShapeAtFace", sBottomShape
                
            sBottomAnswerCol = gsShapeAtFaceCol
            sTopAnswerCol = gsShapeAtFaceCol
            
        Case LCase(gsMbrAxisToFaceAndOutSideNoEdge), LCase(gsStiffEndToMbrFaceAndOutSideNoEdge)
            
            GetSelectorAnswer oAC, "TopShapeAtFace", sTopShape
            GetSelectorAnswer oAC, "BottomShapeAtFace", sBottomShape
            sBottomAnswerCol = gsShapeAtFaceCol
            sTopAnswerCol = gsShapeAtFaceCol
            
        ' -----------------------------------------------------------
        ' If the questions is "Edge" only, determine if top or bottom
        ' -----------------------------------------------------------
        Case LCase(gsMbrAxisToEdge), LCase(gsStiffEndToMbrEdge)
                
            Select Case cBtmOuterEdge.IntersectingEdge
                
                Case eBounding_Edge.Top_Flange_Right, eBounding_Edge.Top_Flange_Right_Bottom, eBounding_Edge.Web_Right
        
                    GetSelectorAnswer oAC, "ShapeAtEdge", sBottomShape
                    
                    sBottomAnswerCol = gsShapeAtEdgeOutsideCol
                    sTopAnswerCol = vbNullString
                
                Case Else
                
                    GetSelectorAnswer oAC, "ShapeAtEdge", sTopShape
                    
                    sTopAnswerCol = gsShapeAtEdgeOutsideCol
                    sBottomAnswerCol = vbNullString

            End Select
        
        ' ----------------------------------------------------------------------------------
        ' If the questions are "Edge" and "EdgeOverlap", determine if top or bottom overlaps
        ' ----------------------------------------------------------------------------------
        Case LCase(gsMbrAxisToEdgeAndOutSide2Edge), LCase(gsStiffEndToMbrEdgeAndOutSide2Edge)
            
            Select Case cBtmInnerEdge.IntersectingEdge
            
                Case eBounding_Edge.Top_Flange_Right, eBounding_Edge.Top_Flange_Right_Bottom, _
                     eBounding_Edge.Web_Right, eBounding_Edge.Bottom_Flange_Right_Top, eBounding_Edge.Bottom_Flange_Right
                    
                    GetSelectorAnswer oAC, "ShapeAtEdge", sBottomShape
                    GetSelectorAnswer oAC, "ShapeAtEdgeOverlap", sTopShape
                    GetSelectorAnswer oAC, "TopInsideCorner", sFaceTopInsideCornerShape
                    GetSelectorAnswer oAC, "BottomInsideCorner", sFaceBtmInsideCornerShape
                    
                    sBottomAnswerCol = gsShapeAtEdgeCol
                    sTopAnswerCol = gsShapeAtEdgeOverlapCol
                    sFaceTopInsideCornerCol = gsInsideCornerCol
                    sFaceBtmInsideCornerCol = gsInsideCornerCol
                    
                Case Else
                    
                    GetSelectorAnswer oAC, "ShapeAtEdge", sTopShape
                    GetSelectorAnswer oAC, "ShapeAtEdgeOverlap", sBottomShape
                    GetSelectorAnswer oAC, "TopInsideCorner", sFaceTopInsideCornerShape
                    GetSelectorAnswer oAC, "BottomInsideCorner", sFaceBtmInsideCornerShape
                    
                    sBottomAnswerCol = gsShapeAtEdgeOverlapCol
                    sTopAnswerCol = gsShapeAtEdgeCol
                    sFaceTopInsideCornerCol = gsInsideCornerCol
                    sFaceBtmInsideCornerCol = gsInsideCornerCol
                    
            End Select

        ' ---------------------------------------------------------------------------------------
        ' If the questions are "Edge" and "Face", determine whether the edge is "Top" or "Bottom"
        ' ---------------------------------------------------------------------------------------
        ' For these two cases, the bottom or top overlaps an edge and is not completely outside it
        ' The inside can still intersect the web, so seeing which intersects the web (face) is insufficient
        ' Also, the inside of the bounded flange that intersects the face can overlap the corner, so testing to see
        ' if the outer surface is "outside" is also insufficient
        Case LCase(gsMbrAxisToFaceAndEdge), LCase(gsStiffEndToMbrFaceAndEdge), _
             LCase(gsMbrAxisToEdgeAndOutSide1Edge), LCase(gsStiffEndToMbrEdgeAndOutSide1Edge)
        
            Dim bBottomOverlapsEdge As Boolean
            bBottomOverlapsEdge = False
            
            Dim dEdgeLength As Double
            dEdgeLength = 0#
            
            If KeyExists("DimPt21ToPt23", oMeasurements) Then
                dEdgeLength = oMeasurements.Item("DimPt21ToPt23")
            End If
            
            If (cBtmInnerEdge.IntersectingEdge = eBounding_Edge.Web_Right Or _
                cBtmInnerEdge.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top Or _
                cBtmInnerEdge.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                cBtmInnerEdge.IntersectingEdge = eBounding_Edge.Bottom Or _
                cBtmInnerEdge.IntersectingEdge = eBounding_Edge.Web_Right_Bottom Or _
                cBtmInnerEdge.IntersectingEdge = eBounding_Edge.Below) _
                And _
               (cBtmOuterEdge.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right_Top Or _
                cBtmOuterEdge.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                cBtmOuterEdge.IntersectingEdge = eBounding_Edge.Bottom Or _
                cBtmOuterEdge.IntersectingEdge = eBounding_Edge.Web_Right_Bottom Or _
                cBtmOuterEdge.IntersectingEdge = eBounding_Edge.Below) _
                And _
                dEdgeLength > 0.0001 Then
                    
                bBottomOverlapsEdge = True
                
            End If
                    
            If bBottomOverlapsEdge Then
            
                GetSelectorAnswer oAC, "ShapeAtEdge", sBottomShape
                GetSelectorAnswer oAC, "ShapeAtFace", sTopShape
                GetSelectorAnswer oAC, "InsideCorner", sFaceBtmInsideCornerShape
            
                sBottomAnswerCol = gsShapeAtEdgeCol
                sTopAnswerCol = gsShapeAtFaceCol
                sFaceBtmInsideCornerCol = gsInsideCornerCol

            Else
            
                GetSelectorAnswer oAC, "ShapeAtEdge", sTopShape
                GetSelectorAnswer oAC, "ShapeAtFace", sBottomShape
                GetSelectorAnswer oAC, "InsideCorner", sFaceTopInsideCornerShape
                
                sTopAnswerCol = gsShapeAtEdgeCol
                sBottomAnswerCol = gsShapeAtFaceCol
                sFaceTopInsideCornerCol = gsInsideCornerCol
                                
            End If
        
        ' For this case, one of the bounded outer surfaces must touch the face
        Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsStiffEndToMbrFaceAndOutSide1Edge)
            
            Select Case cBtmOuterEdge.IntersectingEdge
            
                Case eBounding_Edge.Web_Right
                
                    GetSelectorAnswer oAC, "ShapeAtEdge", sTopShape
                    GetSelectorAnswer oAC, "ShapeAtFace", sBottomShape
                    GetSelectorAnswer oAC, "InsideCorner", sFaceTopInsideCornerShape

                    sBottomAnswerCol = gsShapeAtFaceCol
                    sTopAnswerCol = gsShapeAtEdgeOverlapCol
                    sFaceTopInsideCornerCol = gsInsideCornerCol
                    
                Case Else
                    
                    GetSelectorAnswer oAC, "ShapeAtEdge", sBottomShape
                    GetSelectorAnswer oAC, "ShapeAtFace", sTopShape
                    GetSelectorAnswer oAC, "InsideCorner", sFaceBtmInsideCornerShape

                    sTopAnswerCol = gsShapeAtFaceCol
                    sBottomAnswerCol = gsShapeAtEdgeOverlapCol
                    sFaceBtmInsideCornerCol = gsInsideCornerCol
                
            End Select
             
        ' For this case, they are both outside, so look to see where the single bounding edge is located relative to the bounded part
        Case LCase(gsMbrAxisToOutSideAndOutSide1Edge), LCase(gsStiffEndToMbrOutSideAndOutSide1Edge)

            Dim oEdgeMap As JCmnShp_CollectionAlias
            Set oEdgeMap = GetEdgeMap(oAC, oBoundingData.AxisPort, oBoundedData.AxisPort)

            If KeyExists(CStr(JXSEC_BOTTOM_FLANGE_RIGHT), oEdgeMap) Then
            
                GetSelectorAnswer oAC, "ShapeAtEdge", sBottomShape
                GetSelectorAnswer oAC, "ShapeAtFace", sTopShape
                GetSelectorAnswer oAC, "InsideCorner", sFaceBtmInsideCornerShape
                
                sBottomAnswerCol = gsShapeAtEdgeOverlapCol
                sTopAnswerCol = gsShapeAtFaceCol
                sFaceBtmInsideCornerCol = gsInsideCornerCol
                
            Else
            
                GetSelectorAnswer oAC, "ShapeAtEdge", sTopShape
                GetSelectorAnswer oAC, "ShapeAtFace", sBottomShape
                GetSelectorAnswer oAC, "InsideCorner", sFaceTopInsideCornerShape
                
                sTopAnswerCol = gsShapeAtEdgeOverlapCol
                sBottomAnswerCol = gsShapeAtFaceCol
                sFaceTopInsideCornerCol = gsInsideCornerCol
            
            End If
            
    End Select

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

Public Function GetMinOverlapToCorner() As Double

    GetMinOverlapToCorner = 0.005

End Function

Public Function GetMinOverlapToEdge() As Double

    GetMinOverlapToEdge = 0.01

End Function

Public Function GetMinPCLength() As Double

    GetMinPCLength = 0.0001 ' 0.1mm
    
End Function

Public Function ProfileHasDetailedRepresentation(oProfile As Object) As Boolean

    ' This implementation should be replaced with one that checks directly for the representation
    ' being used by the system, when and if that option is made available to the user.
    ' For now, users are chaning the nature of the simple representation, and we need to figure out what they
    ' are using if one set of rules is to handle both cases.
        
    Const METHOD = "MemberCommon::ProfileHasDetailedRepresentation"
 
    Dim sMsg As String
    sMsg = "Getting the distance from member top or bottom to web left"

    On Error GoTo ErrorHandler

    ProfileHasDetailedRepresentation = False
    
' temp
' The fillet ports are not being returned. You can see this with the port info tool.
' For testing SHI, I'm temporarily setting this to a constant "True"

 '   ProfileHasDetailedRepresentation = True
    Exit Function
    
    
    Dim oPort As IJPort

    Set oPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_WEB_RIGHT_BOTTOM_CORNER)
    
    If Not oPort Is Nothing Then
        ProfileHasDetailedRepresentation = True
        Exit Function
    End If

    Set oPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_WEB_RIGHT_TOP_CORNER)
    
    If Not oPort Is Nothing Then
        ProfileHasDetailedRepresentation = True
        Exit Function
    End If

    Set oPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER)
    
    If Not oPort Is Nothing Then
        ProfileHasDetailedRepresentation = True
        Exit Function
    End If

    Set oPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER)
    
    If Not oPort Is Nothing Then
        ProfileHasDetailedRepresentation = True
        Exit Function
    End If
    
    Set oPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_BOTTOM_RIGHT_CORNER)
    
    If Not oPort Is Nothing Then
        ProfileHasDetailedRepresentation = True
        Exit Function
    End If
    
    Set oPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_TOP_RIGHT_CORNER)
    
    If Not oPort Is Nothing Then
        ProfileHasDetailedRepresentation = True
        Exit Function
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

Public Function GetDistanceFromTopOrBottomToWebRight(oProfile As Object, bBottomDistance As Boolean) As Double
'This method takes the input as profile or member and gives the distance between Top/Bottom flange
'to WebRight that is nothing but same as thickness of flange(expect S cross sections).

    Const METHOD = "MemberCommon::GetDistanceFromTopOrBottomToWebRight"
 
    Dim sMsg As String
    sMsg = "Getting the distance from member top or bottom to web left"

    On Error GoTo ErrorHandler

    GetDistanceFromTopOrBottomToWebRight = -1#

    If oProfile Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oProfile Is ISPSMemberPartPrismatic Then ' ToDo: accommodate stiffeners
        Exit Function
    End If
       
    Dim oWRPort As IJPort
    Dim oTopOrBtmPort As IJPort
    Set oWRPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_WEB_RIGHT)
            
    If bBottomDistance Then
       Set oTopOrBtmPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_BOTTOM)
    Else
       Set oTopOrBtmPort = GetLateralSubPortBeforeTrim(oProfile, JXSEC_TOP)
    End If
    
    Dim oWebRightGeom As IJDModelBody
    Dim oTopOrBtmGeom As IJDModelBody
    
    If TypeOf oWRPort.Geometry Is IJDModelBody And TypeOf oTopOrBtmPort.Geometry Is IJDModelBody Then
        Set oWebRightGeom = oWRPort.Geometry
        Set oTopOrBtmGeom = oTopOrBtmPort.Geometry

        Dim oPoint1 As IJDPosition
        Dim oPOint2 As IJDPosition
        
        oWebRightGeom.GetMinimumDistance oTopOrBtmGeom, oPoint1, oPOint2, GetDistanceFromTopOrBottomToWebRight
        
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

'*********************************************************************************************
' Method      : IsCornerFeatureNeededForShapeAtEdge
' Description : This is helper method to set bIsNeeded flag in conditionals.
'  oConnection is an AssemblyConnection SmartOccurance.
'*********************************************************************************************
Public Sub IsCornerFeatureNeededForShapeAtEdge(oConnection As IJAppConnection, bIsBottomEdge As Boolean, IsBottomFlange As Boolean, bIsNeeded As Boolean)

    
    Const METHOD = "MemberCommon::IsCornerFeatureNeededForShapeAtEdge"
    
    On Error GoTo ErrorHandler

    Dim sMsg As String
    
    bIsNeeded = False
    
    If oConnection Is Nothing Then
        Exit Sub
    End If
        
    ' ----------------------------------------------------------
    ' Get the Assembly Connection Ports from the IJAppConnection
    ' ----------------------------------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    Dim selString As String
    
    InitMemberConnectionData oConnection, oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        Exit Sub
    End If
    
    ' ---------------------
    ' Not needed for a tube
    ' ---------------------
    If IsTubularMember(oBoundingData.MemberPart) Then
        Exit Sub
    End If
        
    ' -----------------------------------------------
    ' Get the requested shapes at top and bottom edge
    ' -----------------------------------------------
    Dim sBottomAnswerCol As String
    Dim sBottomShape As String
    Dim sTopAnswerCol As String
    Dim sTopShape As String
    Dim sShapeAtEdge As String
    
    GetMemberACTopAndBottomShape oConnection, sBottomAnswerCol, sBottomShape, sTopAnswerCol, sTopShape
     
    ' ---------------------------------------------------------------------------------------------------------------------
    ' If this is the shape at the bottom bounding edge, we want the bottom bounded shape, except for ToEdge case
    ' If this is the shape at the top bounding edge, we want the top bounded shape,  except for ToEdge case
    ' ---------------------------------------------------------------------------------------------------------------------
    Dim bIsToEdgeOnly As Boolean
    bIsToEdgeOnly = False
    
    If sBottomAnswerCol = vbNullString Or sTopAnswerCol = vbNullString Then
        bIsToEdgeOnly = True
    End If
        
    If bIsBottomEdge Then
        If bIsToEdgeOnly Then
            sShapeAtEdge = sTopShape
        Else
            sShapeAtEdge = sBottomShape
        End If
    Else
        If bIsToEdgeOnly Then
            sShapeAtEdge = sBottomShape
        Else
            sShapeAtEdge = sTopShape
        End If
    End If
    
    ' --------------------------------------------------------
    ' Find out if there is any overlap at the edge in question
    ' --------------------------------------------------------
    ' This method is named "inside" corner feature, but it is used also for outside corner features
    ' Due to schema restrictions, the name could not be changed when the usage was changed
    Dim dInsideOverlap As Double
    Dim dOutsideOverlap As Double
    Dim dInsideClearance As Double
    Dim dOutsideClearance As Double
    Dim bIsEdgeToEdge As Boolean
    
    GetEdgeOverlapAndClearance oConnection, bIsBottomEdge, IsBottomFlange, dInsideOverlap, dOutsideOverlap, dInsideClearance, dOutsideClearance, , bIsEdgeToEdge
            
    ' The overlap will be measured on the outside for ToEdge cases and on the inside for other cases
    Dim dOverlap As Double
    If bIsToEdgeOnly Then
        dOverlap = dOutsideOverlap
    Else
        dOverlap = dInsideOverlap
    End If
    
    ' ---------------------------------------------------------
    ' Get the extension/setback setting, and offset, if defined
    ' ---------------------------------------------------------
    Dim sExtend As String
    Dim dOffset As Double
            
    GetSelectorAnswer oConnection, "ExtendOrSetBack", sExtend
    GetSelectorAnswer oConnection, "Offset", dOffset
                                   
    ' ------------------------------------
    ' Based on the "ShapeAtEdge" answer...
    ' ------------------------------------
    Dim minOverlap As Double
    
    Select Case sShapeAtEdge
        
        ' ------------------------
        ' If to the edge or corner
        ' ------------------------
        Case gsFaceToCorner, gsFaceToEdge, gsInsideToEdge, gsOutsideToEdge, gsFaceToInsideCorner
                 
            ' -------------------------------------------------------------------------------------------------
            ' If not edge-to-edge, there must be an overlap of at least 5mm, 10mm if from the inside or outside
            ' Does not apply if not ToEdge case and bounded is extended past the near corner
            ' -------------------------------------------------------------------------------------------------
            ' The tolerance is theoretical (i.e. no user requirements have been seen to drive this)
            ' It is intended to prevent unreasonably small features from being placed, and to prevent 'slivers' (areas where
            ' one or more cuts leave a long thin piece of material which would be weak and unsuited for welding)
                        
            ' Compute minimum overlap per comments above
            minOverlap = GetMinOverlapToCorner()
            If sShapeAtEdge = gsInsideToEdge Or sShapeAtEdge = gsOutsideToEdge Then
                minOverlap = GetMinOverlapToEdge()
            End If
            
            ' ToEdge case
            If bIsToEdgeOnly Then
                ' gsOutsideToEdge case needs feature even if minimum overlap is not met, as long as a bounding edge meets a bounded edge
                If (Not sShapeAtEdge = gsOutsideToEdge) And dOverlap < minOverlap Then
                   Exit Sub
                End If

                If (Not bIsEdgeToEdge) And dOverlap < minOverlap Then
                    Exit Sub
                End If
                
            ' All other cases
            Else
                If (Not bIsEdgeToEdge) And (dOverlap < minOverlap) And (sExtend = vbNullString Or sExtend = gsOffsetNearCorner) Then
                    ' For now, assume anything else produces some overlap
                    Exit Sub
                End If
            End If
            
            
            bIsNeeded = True
            Exit Sub
            
        ' -----------------------------------------------------------------------------------------------------------------
        ' If to the flange starting from face, inside, or outside: a feature is always needed as long as it is edge-to-edge
        ' -----------------------------------------------------------------------------------------------------------------
        Case gsFaceToFlange, gsInsideToFlange, gsOutsideToFlange
            
            If Not bIsEdgeToEdge Then
                Exit Sub
            End If
            
            bIsNeeded = True
            Exit Sub
        
        ' ---------------------------------------------------------------------------------------------------------------------------------
        ' If to the flange starting from the edge: a feature is always needed as long as there is edge-to-edge and a minimum overlap is met
        ' ---------------------------------------------------------------------------------------------------------------------------------
        Case gsCornerToFlange, gsEdgeToFlange

            If Not bIsEdgeToEdge Then
                Exit Sub
            End If
            
            ' -----------------------------------------------
            ' If from corner, there must be an overlap of 5mm
            ' If from edge, there must be an overlap of 10mm
            ' -----------------------------------------------
            minOverlap = GetMinOverlapToCorner()
            If sShapeAtEdge = LCase(gsEdgeToFlange) Then
                minOverlap = GetMinOverlapToEdge()
            End If
            
            ' ToEdge case
            If bIsToEdgeOnly Then
                If dOverlap < minOverlap Then
                   Exit Sub
                End If
            ' All other cases
            Else
                If dOverlap < minOverlap Then
                    If sExtend = vbNullString Or sExtend = gsOffsetNearCorner Then
                        ' For now, assume anything else produces some overlap
                        Exit Sub
                    End If
                End If
            End If
                        
            bIsNeeded = True
            Exit Sub
        
        ' ----------------------------------------------------------
        ' No corner features for cases that always go to the outside
        ' ----------------------------------------------------------
        Case gsOutsideToOutside
            Exit Sub
        ' ----------------------------------------------------------------------------------------
        ' No corner features for cases that go to the outside, unless material outside is extended
        ' ----------------------------------------------------------------------------------------
        Case gsEdgeToOutside, _
             gsFaceToOutsideCorner, _
             gsFaceToOutside, _
             gsInsideCornerToOutside, _
             gsInsideToOutsideCorner, _
             gsInsideToOutside, _
             gsCornerToOutside
            
            ' In the code below we are assuming that any extension from the near corner is large enough, or any offset from a location
            ' other than the near corner is small enough, to leave enough material for the corner feature to cut, and leave
            ' a realistic amount of material behind for welding.  We can enhance the logic to compare the actual offsets to the
            ' flange length in the sketching plane at a later time.
            
            If sExtend = vbNullString Or sExtend = gsOffsetNearCorner Then
                Exit Sub
            End If
            
            ' ------------------------------------------------------------------
            ' For "ToEdge" configuration, no feature is needed, even if extended
            ' ------------------------------------------------------------------
            If bIsToEdgeOnly And (LCase(sShapeAtEdge) = LCase(gsEdgeToOutside) Or LCase(sShapeAtEdge) = LCase(gsCornerToOutside)) Then
                Exit Sub
            End If
            
            bIsNeeded = True
            Exit Sub
            
    End Select
 
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

