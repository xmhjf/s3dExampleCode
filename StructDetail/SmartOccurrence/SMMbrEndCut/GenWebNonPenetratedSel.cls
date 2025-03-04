VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenWebNonPenetratedSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012-2014, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenTopAndBtmEdgeWebSel.cls
'
'Description :
'   Selector for Selection of Member bounded by Generic Port/Object End Cuts
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port/Object will be: (Bounding Object)
'           Plate Base/Offset/Lateral Face Port
'           Profile Base/Offset/ Lateral SubPort
'           MemberPart Base/Offset/ Lateral SubPort
'           Reference(Grid) Plane (IJPlane)
'           Point (IJPoint)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    19/Oct/2011 - mpulikol
'           CR-CP-203633 Performance: Increase speed of generic member assembly connections
'    28/Oct/2011 - svsmylav: Removed intermediate selector GenWebNonPenFullSel and instead added FlushBoundedDepth item.
'    16/may/2012 - hgunturu TR-212435: Modified CMSetPhysicalConnAnswer() method to Create PC for 'Generic_2B_01' item.
'    09/Oct/2012 - dsmamidi CR-CP-180749: Added new webcut smart item "SnipedTube_Gusset" for support SnipedEndCut for Gusset Plate
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'   27/02/2014 - knukala
'                Changed the CMSetPhysicalConnAnswer()to prevent creating unexpected WebPC
'   14/July/2015 -knukala
'           TR-CP-270890  Cannot Place Generic AC when Handrail Post member is bounded to a plate
'*********************************************************************************************

Const m_sClassName As String = "GenWebNonPenetratedSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED

    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"

    'Question: Create Physical Connection ?
    pQH.SetQuestion "CreatePhysicalConnection", "Yes", "BooleanCol", "CMSetPhysicalConnAnswer", CUSTOMERID & "MbrEndCut.GenWebNonPenetratedSel"

    'Question: Create Top Flange Cut ?
    pQH.SetQuestion "CreateTopFlangeCut", "Yes", "BooleanCol", "CMSetTopFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenWebNonPenetratedSel"
    'Question: Create Bottom Flange Cut ?
    pQH.SetQuestion "CreateBottomFlangeCut", "Yes", "BooleanCol", "CMSetBottomFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenWebNonPenetratedSel"
    pQH.SetQuestion "FullDepth", "No", "BooleanCol", "CMSetFullDepthAnswer", CUSTOMERID & "MbrEndCut.GenWebNonPenetratedSel"

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the appropriate standard/normal cut
'*********************************************************************************************
Public Sub SelectorLogic(oSL As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Unknown Error"

    On Error GoTo ErrorHandler
    
    ' ----------------------
    ' Get the web cut object
    ' ----------------------
    Dim oEndCutObject As Object
    Set oEndCutObject = oSL.SmartOccurrence
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    ' ------------------------------------
    ' Get the bounded and bounding objects
    ' ------------------------------------
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    sMsg = "Set the Bounded input"
    Set oBoundedObject = oSDO_WebCut.BoundedPort
    If (oBoundedObject Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding input"
    Set oBoundingObject = oSDO_WebCut.BoundingPort
    If (oBoundingObject Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If

    ' --------------------------------
    ' Check the validity of the inputs
    ' --------------------------------
    Dim lStatus As Long ' reports validity
    Dim oBoundedData As MemberConnectionData ' not used
    Dim oBoundingData As MemberConnectionData ' not used
    
    InitEndCutConnectionData oBoundedObject, oBoundingObject, oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        sMsg = "EndCut Ports are Not valid"
        GoTo ErrorHandler
    End If

    Dim sFullDepth As String
    GetSelectorAnswer oSL, "FullDepth", sFullDepth

    ' -------------------------------------
    ' If the bounding object is connectable
    ' -------------------------------------
    Dim sSelector As String
    sSelector = ""
    Dim IsTubeMember As Boolean
    Dim eBoundingPort As eUSER_CTX_FLAGS
    Dim oStructPort As IJStructPort
    If TypeOf oBoundingObject Is IJPort Then
        ' --------------------------------------------------------------------------------------------------------------
        ' Otherwise, offer to trim flush to the bounding object (depth of bounded or width of bounding) and other shapes
        ' --------------------------------------------------------------------------------------------------------------
        If sFullDepth = "Yes" Then
            If TypeOf oBoundingObject Is IJStructPort Then
                Set oStructPort = oBoundingObject
                    If oStructPort.ContextID = CTX_LATERAL_LFACE Then
                        IsTubeMember = IsTubularMember(oBoundedObject)
                        If IsTubeMember Then
                            ' this item will select only when plate edge is lateral face and bounded must be TUBE
                            oSL.Add "SnipedTube_Gusset"
                        End If
                    End If
                oSL.Add "FlushBoundedDepth"
            Else
                oSL.Add "FlushBoundedDepth"
            End If
        Else
            oSL.Add "GenWebNonPenActualSel"
        End If
    ' ---------------------------------------------------------------------------
    ' If the input is a plane, the actual option is too large.  Ommit this option
    ' ---------------------------------------------------------------------------
    ' Even if the user asked for "actual", it would bound the full depth
    ElseIf TypeOf oBoundingObject Is IJPlane Then
        oSL.Add "FlushBoundedDepth"
        ' other shapes pending
    ' ---------------------------------------------------------------------------------------------------------
    ' If the input is a point, the actual option is too small.  We must trim to the depth of the bounded object
    ' ---------------------------------------------------------------------------------------------------------
    ' "Actual" is simply not valid, and must be ignored
    ElseIf TypeOf oBoundingObject Is IJPoint Then
        oSL.Add "FlushBoundedDepth"
    ' --------------------------------
    ' Default to Flush (bounded depth)
    ' --------------------------------
    Else
        If sFullDepth = "Yes" Then
            oSL.Add "FlushBoundedDepth"
        Else
            oSL.Add "GenWebNonPenActualSel"
        End If
    End If

    Exit Sub
ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
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
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'*********************************************************************************************
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

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMSelector
' Description :
'
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
    
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

'*********************************************************************************************
' Method      : CMSetTopFlangeCutAnswer
' Description : This procedure is invoked before the SelectorLogic procedure in the Selector Class
'
'*********************************************************************************************
Public Sub CMSetTopFlangeCutAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)

    Const METHOD = m_SelectorProgid & "::CMSetTopFlangeCutAnswer"
    On Error GoTo ErrorHandler

    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    oSL.answer("CreateTopFlangeCut") = ComputeFlangeCutAnswer(oSL, False)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
   
End Sub

'*********************************************************************************************
' Method      : CMSetBottomFlangeCutAnswer
' Description : This procedure is invoked before the SelectorLogic procedure in the Selector Class
'
'*********************************************************************************************
Public Sub CMSetBottomFlangeCutAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD = m_SelectorProgid & "::CMSetBottomFlangeCutAnswer"
    On Error GoTo ErrorHandler

    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    oSL.answer("CreateBottomFlangeCut") = ComputeFlangeCutAnswer(oSL, True)

    Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD, "").Number
   
End Sub

'*********************************************************************************************
' Method      : CMSetPhysicalConnAnswer
' Description : This procedure is invoked before the SelectorLogic procedure in the Selector Class
'
'*********************************************************************************************
Public Sub CMSetPhysicalConnAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD = m_SelectorProgid & "::CMSetPhysicalConnAnswer"
    On Error GoTo ErrorHandler
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition

    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)

    Dim answer As String
    answer = "No"
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName, oACObject
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Or sACItemName = "Generic_2B_01" Then
        pSL.answer("CreatePhysicalConnection") = "Yes"
        Exit Sub
    End If
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSL.SmartOccurrence
    Dim IsTubularMbr As Boolean
    Dim eCTXBounding As eUSER_CTX_FLAGS
    Dim oStructPort As IJStructPort
    

    ' ----------------------------------------------------------------------
    ' If the bounding object is not a connectable, the answer is always "no"
    ' ----------------------------------------------------------------------
    If TypeOf oSDO_WebCut.BoundingPort Is IJPort Then
        ' -------------------------------------------
        ' If the bounding part is a member or profile
        ' -------------------------------------------
        If TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_WebCut.Bounding Is IJProfile Then
            ' -----------------------------------------------------------------------
            ' If the web is outside the bounding object, the answer is always "no"
            ' -----------------------------------------------------------------------
            Dim oTop As ConnectedEdgeInfo
            Dim oBottom As ConnectedEdgeInfo
            Dim oInsideTF As ConnectedEdgeInfo
            Dim oInsideBF As ConnectedEdgeInfo

            GetConnectedEdgeInfo oSDO_WebCut.object, oSDO_WebCut.BoundedPort, oSDO_WebCut.BoundingPort, oTop, oBottom, oInsideTF, oInsideBF

            If (Not oTop.IntersectingEdge = Below And Not oTop.IntersectingEdge = Above) Or (Not oBottom.IntersectingEdge = Below And Not oBottom.IntersectingEdge = Above) Then
                answer = "Yes"
            Else
                ' answer always no
            End If
        ' -------------------------
        ' Otherwise default to "No"
        ' -------------------------
        ' Can we do better than this?  Have to determine if it intersects this part *first*
        ElseIf TypeOf oSDO_WebCut.Bounding Is IJPlate Then
            If TypeOf oSDO_WebCut.BoundingPort Is IJStructPort Then
                Set oStructPort = oSDO_WebCut.BoundingPort
                If oStructPort.ContextID = CTX_LATERAL_LFACE Then
                    IsTubularMbr = IsTubularMember(oSDO_WebCut.BoundedPort)
                    If IsTubularMbr Then
                        'when tubular member hits plate lateraledge then "SnipedTube_Gusset" WebCut item is going to be selected.
                        'For this item we always need PC, hence giving PC defaultly
                        answer = "Yes"
                    Else
                        Dim oWebleftPort As IJPort
                        Dim oWebRightPort As IJPort
                        Dim oGeomOpr As GSCADShipGeomOps.SGOModelBodyUtilities
                        Dim oPointOnPort1 As IJDPosition
                        Dim oPointOnPort2 As IJDPosition
                        Dim oPointOnObj1 As IJDPosition
                        Dim oPointOnObj2 As IJDPosition
                        Dim dDistance1 As Double
                        Dim dDistance2 As Double
                        Dim oBndingObject As Object
                        Dim IsIntersecting As Boolean
                        Dim oObjModelBody As IJModelBody, oObjPort As IJPort
                        
                        Set oBndingObject = oSDO_WebCut.Bounding
                        Set oGeomOpr = New SGOModelBodyUtilities

                        If TypeOf oBndingObject Is IJModelBody Then
                            Set oObjModelBody = oBndingObject
                        ElseIf TypeOf oBndingObject Is IJPort Then
                            Set oObjPort = oBndingObject
                            Set oObjModelBody = oObjPort.Geometry
                        Else
                            Set oObjModelBody = oBndingObject
                        End If

                        Set oWebleftPort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_LEFT)
                        Set oWebRightPort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_RIGHT)

                        oGeomOpr.GetClosestPointsBetweenTwoBodies oWebleftPort.Geometry, oObjModelBody, oPointOnPort1, oPointOnObj1, dDistance1
                        oGeomOpr.GetClosestPointsBetweenTwoBodies oWebRightPort.Geometry, oObjModelBody, oPointOnPort2, oPointOnObj2, dDistance2
                        'If the distance between WebLeft/WebRight and the plate part is equal to zero, then answer is set to Yes
                        
                        If Equal(dDistance1, 0) Or Equal(dDistance2, 0) Then
                        
                            answer = "Yes"
                            'If the plate port and the Inner Flange are in the same plane
                            '(that is distance between them is zero), then PC conditional is False.
                            Dim oBngingModelBody As IJDModelBody
                            Dim Position1 As IJDPosition
                            Dim Position2 As IJDPosition
                            Dim oPlateBasePort As IJPort
                            Dim oPlateOffsetPort As IJPort
                            Dim oInnerFlangePort As IJPort
                            Dim dDistFrmBaseToFlng As Double
                            Dim dDistFrmOffsetToFlng As Double
                            Dim oStructGeomUtils As GSCADStructGeomUtilities.PartInfo
                            Dim oSGOModelBodyUtils As GSCADShipGeomOps.SGOModelBodyUtilities
                            Dim oBaseVector As IJDVector
                            Dim oOffsetVector As IJDVector
                            Dim oMemberVector As IJDVector
                            Dim bAU As Boolean
                            Dim bTFL As Boolean
                            Dim bBFL As Boolean
                            Dim bBFR As Boolean
                            Dim bTFR As Boolean
                            
                            Set oPlateBasePort = GetBaseOffsetOrLateralPortForObject(oBndingObject, BPT_Base)
                            Set oPlateOffsetPort = GetBaseOffsetOrLateralPortForObject(oBndingObject, BPT_Offset)
                            Set oStructGeomUtils = New GSCADStructGeomUtilities.PartInfo
                            Set oSGOModelBodyUtils = New GSCADShipGeomOps.SGOModelBodyUtilities
                            'This check works for both case: Plate hitting member Top flnage or member inner flange
                            CrossSection_Flanges oSDO_WebCut.Bounded, bTFL, bBFL, bTFR, bBFR
                            
                            If bTFL Then
                                If oSGOModelBodyUtils.HasIntersectingGeometry(GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_TOP_FLANGE_LEFT).Geometry, oPlateBasePort.Connectable) Then
                                    Set oInnerFlangePort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_TOP_FLANGE_LEFT_BOTTOM)
                                End If
                            ElseIf bTFR Then
                                If oSGOModelBodyUtils.HasIntersectingGeometry(GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_TOP_FLANGE_RIGHT).Geometry, oPlateBasePort.Connectable) Then
                                    Set oInnerFlangePort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
                                End If
                            End If
                            
                            If bBFL Then
                                If oSGOModelBodyUtils.HasIntersectingGeometry(GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_BOTTOM_FLANGE_LEFT).Geometry, oPlateBasePort.Connectable) Then
                                    Set oInnerFlangePort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_BOTTOM_FLANGE_LEFT_TOP)
                                End If
                            ElseIf bBFR Then
                                If oSGOModelBodyUtils.HasIntersectingGeometry(GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_BOTTOM_FLANGE_RIGHT).Geometry, oPlateBasePort.Connectable) Then
                                    Set oInnerFlangePort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
                                End If
                            End If
                            
                            If Not oInnerFlangePort Is Nothing Then
                                Set oBaseVector = oStructGeomUtils.GetPortNormal(oPlateBasePort, bAU)
                                Set oOffsetVector = oStructGeomUtils.GetPortNormal(oPlateOffsetPort, bAU)
                                Set oMemberVector = oStructGeomUtils.GetPortNormal(oInnerFlangePort, bAU)
                                oBaseVector.Length = 1
                                oMemberVector.Length = 1
                                oOffsetVector.Length = 1
    
                                Set oBngingModelBody = oInnerFlangePort.Geometry
                                oBngingModelBody.GetMinimumDistance oPlateBasePort.Geometry, Position1, Position2, dDistFrmBaseToFlng
                                oBngingModelBody.GetMinimumDistance oPlateOffsetPort.Geometry, Position1, Position2, dDistFrmOffsetToFlng
    
                                If Equal(Round((oMemberVector.Dot(oBaseVector)), 4), 1) Then
                                    If Equal(dDistFrmBaseToFlng, 0) Then
                                        answer = "No"
                                    End If
                                ElseIf Equal(Round((oMemberVector.Dot(oOffsetVector)), 4), 1) Then
                                    If Equal(dDistFrmOffsetToFlng, 0) Then
                                        answer = "No"
                                    End If
                                End If
                            End If
                        Else
                            Dim oGeomFactory As IJGeometryFactory
                            Dim oLine As IJLine
                            Set oGeomFactory = New GeometryFactory
                            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oPointOnPort1.x, oPointOnPort1.y, oPointOnPort1.z, oPointOnPort2.x, oPointOnPort2.y, oPointOnPort2.z)
    
                            Dim oGeomMisc As IJGeometryMisc
                            Dim oWire As Object
    
                            Set oGeomMisc = New DGeomOpsMisc
                            oGeomMisc.CreateModelGeometryFromGType Nothing, oLine, Nothing, oWire
    
                            ' --------------------------------------------------------
                            ' If the wire intersects, consider the web as intersecting
                            ' --------------------------------------------------------
                            IsIntersecting = oGeomOpr.HasIntersectingGeometry(oWire, oBndingObject)
                            
                            If IsIntersecting Then
                                answer = "Yes"
                            End If
                        End If
                    End If
                Else
                    ' If the input port is from a bearing plate, presume there is only one boundary and that it intersects
                    ' A PC is needed
                    If TypeOf oSDO_WebCut.Bounding Is IJSmartPlate Then
                        Dim oSmartPlate As IJSmartPlate
                        Set oSmartPlate = oSDO_WebCut.Bounding
                        If oSmartPlate.SmartPlateType = spType_BEARING Then
                            answer = "Yes"
                        End If
                    End If
                    
                    ' Otherwise, if there is only one web cut, presume it itersects and create the PC
                    ' If there is more than web cut the geometry for this one may have been removed.  Do not automatically
                    ' apply a PC or we may get a ToDo record.
                    If (answer = "No") Then
                        Dim nWebCuts As Long
                        nWebCuts = GetNumberOfWebCutsOnAC(pSL.SmartOccurrence)
                        If nWebCuts = 1 Then
                            answer = "Yes"
                        End If
                    End If
                End If
            End If
        Else
            answer = "No"
        End If
    End If
        
    pSL.answer("CreatePhysicalConnection") = answer

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
   
End Sub

'*********************************************************************************************
' Method      : ComputeFlangeCutAnswer
' Description : This procedure is invoked before the SelectorLogic procedure in the Selector Class
'
'*********************************************************************************************
Public Function ComputeFlangeCutAnswer(oSL As IJDSelectorLogic, isBottom As Boolean) As String
    Const METHOD = m_SelectorProgid & "::ComputeFlangeCutAnswer"
    On Error GoTo ErrorHandler

    ComputeFlangeCutAnswer = "No"
    
    ' ------------------------------------
    ' Determine what bounded flanges exist
    ' ------------------------------------
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Set oSDOWebCut.object = oSL.SmartOccurrence
    
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oSDOWebCut.Bounded, bTFL, bBFL, bTFR, bBFR

    ' ----------------------------------
    ' If the bounded object has a flange
    ' ----------------------------------
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oSL.SmartOccurrence, sACItemName, oACObject
    
    If (isBottom And (bBFL Or bBFR)) Or (Not isBottom And (bTFL Or bTFR)) Then
        Dim nWebCuts As Long
        nWebCuts = GetNumberOfWebCutsOnAC(oSL.SmartOccurrence)
        Dim nFlangeCuts As Long
        nFlangeCuts = GetNumberOfFlangeCutsOnAC(oSL.SmartOccurrence)
        
        ' -----------------------------------------------------
        ' Special-case generic ACs predefine the desired answer
        ' -----------------------------------------------------
        ' Set the values, even if the default would match what is desired by the AC
        ' 1) to guard against changes in the default logic affecting the special AC
        ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
        If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
            ComputeFlangeCutAnswer = ComputeFlangeCutAnswerForWPEF_PF_1(oSL, oACObject, isBottom)
        ' ---------------------------------------
        ' If this is the only web cut, then "Yes"
        ' ---------------------------------------
        ElseIf nWebCuts < 2 Then
            If nFlangeCuts > 0 Then
                ComputeFlangeCutAnswer = "No"
            Else
                ComputeFlangeCutAnswer = "Yes"
            End If
        ' -------------------------------------------------------------------
        ' Otherwise, default to "No" unless it intersects the bounding object
        ' -------------------------------------------------------------------
        Else
            Dim oFirstPenetrated As Object
            Set oFirstPenetrated = GetFirstPenetrated(isBottom, oSL.SmartOccurrence)
            
            If oFirstPenetrated Is oSDOWebCut.Bounding Then
                ComputeFlangeCutAnswer = "Yes"
            Else
                ComputeFlangeCutAnswer = "No"
            End If
        End If
    ' ------------------------------------------------
    ' If there is no flange, the answer is always "No"
    ' ------------------------------------------------
    Else
        ComputeFlangeCutAnswer = "No"
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Function

'*********************************************************************************************
' Method      : ComputeFlangeCutAnswerForWPEF_PF_1
' Description : This procedure is invoked before the SelectorLogic procedure in the Selector Class
'
'*********************************************************************************************
Public Function ComputeFlangeCutAnswerForWPEF_PF_1(oSL As IJDSelectorLogic, oACObject As Object, isBottom As Boolean) As String
    Const METHOD = m_SelectorProgid & "::ComputeFlangeCutAnswerForWPEF_PF_1"
    On Error GoTo ErrorHandler
    
    ComputeFlangeCutAnswerForWPEF_PF_1 = "No"
    
    ' -----------------------------------------------------------------------------
    ' Get the web cut on the bounding member (this is web cut for a bounding plate)
    ' -----------------------------------------------------------------------------
    Dim oMemberCut As IJStructFeature
    Dim oMemberObjects As IJDMemberObjects
    Set oMemberObjects = oACObject
    
    ' -------------------------------------------------------------------------------------
    ' If item 1 of the AC is this cut (for the bounding plate) then the member cut is item2
    ' -------------------------------------------------------------------------------------
    If oMemberObjects.ItemByDispid(1) Is oSL.SmartOccurrence Then
        If TypeOf oMemberObjects.ItemByDispid(2) Is IJStructFeature Then
            Set oMemberCut = oMemberObjects.ItemByDispid(2)
        End If
    ' -------------------------------------------------------------------------------------
    ' If item 2 of the AC is this cut (for the bounding plate) then the member cut is item1
    ' -------------------------------------------------------------------------------------
    ElseIf oMemberObjects.ItemByDispid(1) Is oSL.SmartOccurrence Then
        If TypeOf oMemberObjects.ItemByDispid(1) Is IJStructFeature Then
            Set oMemberCut = oMemberObjects.ItemByDispid(1)
        End If
    End If
    
    ' ------------------------------------------
    ' Find out the alias for the bounding member
    ' ------------------------------------------
    If Not oMemberCut Is Nothing Then
        If oMemberCut.get_StructFeatureType = SF_WebCut Then
            Dim oSDOBoundingMemberCut As New StructDetailObjects.WebCut
            Set oSDOBoundingMemberCut.object = oMemberCut
        
            Dim lSectionAlias As Long
            Dim bPenetratesWeb As Boolean
            Dim oEdgeMap As JCmnShp_CollectionAlias
            
            Set oEdgeMap = New Collection
            Set oEdgeMap = GetEdgeMap(oSL.SmartOccurrence, oSDOBoundingMemberCut.BoundingPort, oSDOBoundingMemberCut.BoundedPort, lSectionAlias, bPenetratesWeb)
            ' -------------------------------------------------------------------------------------------------
            ' If it appears to the bounded object as a web with top flange we don't want a top flange cut
            ' If it appears to the bounded object as a web with bottom flange we don't want a bottom flange cut
            ' -------------------------------------------------------------------------------------------------
            If (lSectionAlias = 1 And Not isBottom) Or (lSectionAlias = 3 And isBottom) Then
                ComputeFlangeCutAnswerForWPEF_PF_1 = "No"
            Else
                ComputeFlangeCutAnswerForWPEF_PF_1 = "Yes"
            End If
        End If
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Function

'*********************************************************************************************
' Method      : CMSetFullDepthAnswer
' Description : This procedure is invoked before the SelectorLogic procedure in the Selector Class
'
'*********************************************************************************************
Public Sub CMSetFullDepthAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    Const METHOD = m_SelectorProgid & "::CMSetFullDepthAnswer"
    On Error GoTo ErrorHandler
    
    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' -----------------------------------------------------
    ' Default to full depth if it there is only one web cut
    ' -----------------------------------------------------
    Dim nWebCuts As Long
    nWebCuts = GetNumberOfWebCutsOnAC(oSL.SmartOccurrence)

    If nWebCuts > 1 Then
        oSL.answer("FullDepth") = "No"
    Else
        oSL.answer("FullDepth") = "Yes"
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
   
End Sub

