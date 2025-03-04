VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenFlangeNonPenetratedSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-12, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenFlangeNonPenetratedSel.cls
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
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    18/Oct/2011 - pnalugol - To address 201305: custom method to set answer to create PC modified.
'    19/Oct/2011 - mpulikol
'           CR-CP-203633 Performance: Increase speed of generic member assembly connections
'    07/Feb/2012 - svsmylav
'           TR-CP-209564: Added 'Sniped_LR' and 'Sniped_R' sniped-welded new FlangeCut smart items.
'    16/may/2012 - hgunturu TR-212435: Modified the methods CMSetInsideCornerFeatureAnswer() not to
'                created CornerFeatures and CMSetPhysicalConnAnswer() to create PC for new smart item
'      05/November - dsmamidi
'               Added new smart item selection for member bounded to plate's lateral edge
'               DI-CP-235025  StrDet: Error in Partclass View creation during bulkloading on Oracle DB
'*********************************************************************************************

Const m_sClassName As String = "GenFlangeNonPenetratedSel"
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
    pQH.SetQuestion "CreatePhysicalConnection", "Yes", "BooleanCol", "CMSetPhysicalConnAnswer", CUSTOMERID & "MbrEndCut.GenFlangeNonPenetratedSel"
    pQH.SetQuestion "CreateInsideCornerFeature", "No", "BooleanCol", "CMSetInsideCornerFeatureAnswer", CUSTOMERID & "MbrEndCut.GenFlangeNonPenetratedSel"

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

    ' -------------------------
    ' Get the flange cut object
    ' -------------------------
    Dim oEndCutObject As Object
    Set oEndCutObject = oSL.SmartOccurrence
    
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oEndCutObject

    ' -----------------------------------------------
    ' Get the number of flanges on the bounded object
    ' -----------------------------------------------
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oSDO_FlangeCut.Bounded, bTFL, bBFL, bTFR, bBFR

    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName oSL.SmartOccurrence, sACItemName, oAppConnection
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Or sACItemName = "Generic_2B_01" Then
        If bTFL Or bBFL Then
            oSL.Add "Straight_LR"
        Else
            oSL.Add "Straight_R"
        End If
        Exit Sub
    End If

    ' ------------------------------------
    ' Get the bounded and bounding objects
    ' ------------------------------------
    Dim oBoundedInput As Object
    Dim oBoundingInput As Object
    
    sMsg = "Set the Bounded input"
    Set oBoundedInput = oSDO_FlangeCut.BoundedPort
    If (oBoundedInput Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding input"
    Set oBoundingInput = oSDO_FlangeCut.BoundingPort
    If (oBoundingInput Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If

    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oSDO_FlangeCut.WebCut

    Dim oWebCutBounding As Object
    Set oWebCutBounding = oSDO_WebCut.BoundingPort

    Dim bDifferentWebCutInput As Boolean
    If oWebCutBounding.Connectable Is oBoundingInput.Connectable Then
        bDifferentWebCutInput = False
    Else
        bDifferentWebCutInput = True
    End If

    ' --------------------------------
    ' Check the validity of the inputs
    ' --------------------------------
    Dim lStatus As Long ' reports validity
    Dim oBoundedData As MemberConnectionData ' not used
    Dim oBoundingData As MemberConnectionData ' not used
    
    InitEndCutConnectionData oBoundedInput, oBoundingInput, oBoundedData, oBoundingData, lStatus, sMsg

    If lStatus <> 0 Then
        sMsg = "EndCut Ports are Not valid"
        GoTo ErrorHandler
    End If

    ' -------------------------------------
    ' If the bounding object is connectable
    ' -------------------------------------
    If TypeOf oBoundingInput Is IJPort Then
        ' ---------------------------------------------------
        ' If bounded by a tubular member, pick a special type
        ' ---------------------------------------------------
        If IsTubularMember(oBoundedInput) Then
'            sSelector = "M_Flange_Generic_Tube01"
'            oSL.Add "M_Flange_Generic_Tube01"
        ' ----------------------------------
        ' If bounded by a non-tubular member
        ' ----------------------------------
        ElseIf TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_FlangeCut.Bounding Is IJProfile Then
            ' ------------------------------------------------------------
            ' Determine if the flange is on or outside the bounding object
            ' ------------------------------------------------------------
            Dim bIsOutside As Boolean
            Dim bIsOn As Boolean
            Dim sIsBottom As String
            sIsBottom = oSL.SelectorAnswer(CUSTOMERID & "MbrEndCut.RootGenFlangeSel", "BottomFlange")
            
            Dim eTopOrWL As ConnectedEdgeInfo
            Dim eBottomOrWR As ConnectedEdgeInfo
            Dim eInsideTFOrTFL As ConnectedEdgeInfo
            Dim eInsideBFOrTFR As ConnectedEdgeInfo
                
            GetConnectedEdgeInfo oSDO_FlangeCut.object, oSDO_FlangeCut.BoundedPort, oSDO_FlangeCut.BoundingPort, eTopOrWL, eBottomOrWR, eInsideTFOrTFL, eInsideBFOrTFR
            
            If sIsBottom = "Yes" Then
                If eInsideBFOrTFR.IntersectingEdge = Below Or eBottomOrWR.IntersectingEdge = Above Or eInsideBFOrTFR.CoplanarEdge = Bottom Then
                    bIsOutside = True
                ElseIf eBottomOrWR.CoplanarEdge = Top Then
                    bIsOn = True
                End If
            Else
                If eTopOrWL.IntersectingEdge = Below Or eInsideTFOrTFL.IntersectingEdge = Above Or eInsideTFOrTFL.CoplanarEdge = Top Then
                    bIsOutside = True
                ElseIf eTopOrWL.CoplanarEdge = Bottom Then
                    bIsOn = True
                End If
            End If

            ' ----------------------------------------------------------------
            ' If outside or there is an outside web cut, then snipe by default
            ' ----------------------------------------------------------------
            ' Note: it was thought that we could try fetching the answer from the flange cut input, even if that happed to be an outside cut,
            ' which does not define the question.  In that case the expected answer would be a NULL string.  But when going from no outside
            ' cut to having one, the oSmartOccurrence.ItemObject for the outside cut is Nothing, and this fails.  So we always get the answer
            ' from the center cut.
            Dim oDesignChild As IJDesignChild
            Dim oDesignParent As Object
            Set oDesignChild = oSDO_FlangeCut.object
            Set oDesignParent = oDesignChild.GetParent
            
            Dim sHasTopWeb As String
            Dim sHasBtmWeb As String
            Dim oStructPort As IJStructPort
            Dim IsParallel As Boolean
            GetSelectorAnswer oDesignParent, "CreateTopWebCut", sHasTopWeb
            GetSelectorAnswer oDesignParent, "CreateBottomWebCut", sHasBtmWeb

            'Determine if two new Sniped smart items exist in Catalog
            Dim bSnipWeldItemsFlag As Boolean
            
            Dim eSCType As SmartClassType
            Dim eSCSubType As SmartClassSubType
            Dim sClassName As String
                                  
            eSCType = SMARTTYPE_FLANGECUT
            eSCSubType = 312
            sClassName = "GenFlgNonPenetratdSel"
            If CheckSmartItemExists(eSCType, eSCSubType, sClassName, "Sniped_LR") And _
              CheckSmartItemExists(eSCType, eSCSubType, sClassName, "Sniped_R") Then
              bSnipWeldItemsFlag = True
            Else
              bSnipWeldItemsFlag = False
            End If

            ' The top/bottom cuts don't define questions, so the answer will come back as an empty string
            If bIsOutside Or _
               (sIsBottom = "Yes" And (sHasBtmWeb = "Yes" Or sHasBtmWeb = vbNullString)) Or _
               (sIsBottom = "No" And (sHasTopWeb = "Yes" Or sHasBtmWeb = vbNullString)) Then
                If bTFL Or bBFL Then
                    If bDifferentWebCutInput Then
                        oSL.Add "Straight_LR"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_LR"
                    Else
                        oSL.Add "SnipedSetback_LR"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_LR"
                        oSL.Add "Straight_LR"
                        oSL.Add "StraightSetback_LR"
                    End If
                Else
                    If bDifferentWebCutInput Then
                        oSL.Add "Straight_R"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_R"
                    Else
                        oSL.Add "SnipedSetback_R"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_R"
                        oSL.Add "Straight_R"
                        oSL.Add "StraightSetback_R"
                    End If
                End If
            ' --------------------------------------------
            ' Otherwise, default is straight and connected
            ' --------------------------------------------
            Else
                If bTFL Or bBFL Then
                    If bDifferentWebCutInput Then
                        oSL.Add "Straight_LR"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_R"
                    Else
                        oSL.Add "Straight_LR"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_LR"
                        oSL.Add "StraightSetback_LR"
                        oSL.Add "SnipedSetback_LR"
                    End If
                Else
                    If bDifferentWebCutInput Then
                        oSL.Add "Straight_R"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_R"
                    Else
                        oSL.Add "Straight_R"
                        oSL.Add "StraightSetback_R"
                        If bSnipWeldItemsFlag Then oSL.Add "Sniped_R"
                        oSL.Add "SnipedSetback_R"
                    End If
                End If
            End If
        ' ---------------
        ' If not a member
        ' ---------------
        Else
            If bTFL Or bBFL Then
                'check whether the bounding is plate
                If TypeOf oSDO_FlangeCut.Bounding Is IJPlate Then
                    If TypeOf oSDO_FlangeCut.BoundingPort Is IJStructPort Then
                        Set oStructPort = oSDO_FlangeCut.BoundingPort
                        If oStructPort.ContextID = CTX_LATERAL_LFACE Then
                            ' To avoid selecting this for unneccesary cases like webnormal perpendicular to plate base normal
                            IsParallel = AreTwoportNormalsParallel(oSDO_FlangeCut.Bounding, oSDO_FlangeCut.Bounded)
                            If IsParallel Then
                                 oSL.Add "FlangeCut_Gusset"
                            End If
                        End If
                    End If
                End If
                oSL.Add "Straight_LR"
                oSL.Add "StraightSetback_LR"
                If bSnipWeldItemsFlag Then oSL.Add "Sniped_LR"
                oSL.Add "SnipedSetback_LR"
            Else
                 'check whether the bounding is plate
                If TypeOf oSDO_FlangeCut.Bounding Is IJPlate Then
                    If TypeOf oSDO_WebCut.BoundingPort Is IJStructPort Then
                    Set oStructPort = oSDO_FlangeCut.BoundingPort
                        If oStructPort.ContextID = CTX_LATERAL_LFACE Then
                            ' To avoid selecting this for unneccesary cases like webnormal perpendicular to plate base normal
                            IsParallel = AreTwoportNormalsParallel(oSDO_FlangeCut.Bounding, oSDO_FlangeCut.Bounded)
                            If IsParallel Then
                                oSL.Add "FlangeCut_Gusset"
                            End If
                        End If
                    End If
                End If
                oSL.Add "Straight_R"
                oSL.Add "StraightSetback_R"
                If bSnipWeldItemsFlag Then oSL.Add "Sniped_R"
                oSL.Add "SnipedSetback_R"
            End If
        End If
        
' *******************
' Normally I don't comment-out code instead of deleting it, but not being sure of the intention for this, and
' whether or not it might be important, it is left here for now.  GKH 29Mar2011
'    ' ---------------------------------------------------------------------------
'    ' If the input is a plane, the actual option is too large.  Ommit this option
'    ' ---------------------------------------------------------------------------
'    ElseIf TypeOf oBoundingInput Is IJPlane Then
'        oSL.Add "FlushBoundedDepth"
'        ' other shapes pending
'    ' ---------------------------------------------------------------------------------------------------------
'    ' If the input is a point, the actual option is too small.  We must trim to the depth of the bounded object
'    ' ---------------------------------------------------------------------------------------------------------
'    ElseIf TypeOf oBoundingInput Is IJPoint Then
'        oSL.Add "FlushBoundedDepth"
'    ' --------------------------------
'    ' Default to Flush (bounded depth)
'    ' --------------------------------
'    Else
'        oSL.Add "FlushBoundedDepth"
    End If
' *******************

' ==================================================
' This was overriding the defaults set above, and deleting it (commenting it out) did not seem to have any negative
' side-effects. Normally I don't comment-out code instead of deleting it, but not being sure of the intention for this, and
' whether or not it might be important, it is left here for now. GKH 29Mar2011

'    ' ------------------------------------------------------------------
'    ' If there is no current selection, default to Flush (bounded depth)
'    ' ------------------------------------------------------------------
'    If Len(Trim(sSelector)) > 0 Then
'        ' Get the Current Selection (if exists)
'        Dim oSmartOccurrence As IJSmartOccurrence
'        Dim oSymbolDefinition As IJDSymbolDefinition
'        Dim oSmartOccHelper As GSCADSmartOccurrence.IJSmartOccurrenceHelper
'
'        Set oSmartOccHelper = New GSCADSmartOccurrence.CSmartOccurrenceCES
'        Set oSmartOccurrence = oEndCutObject
'        Set oSymbolDefinition = oSL.Selector
'
'        Dim sCurrentChoice As String
'        sCurrentChoice = oSmartOccHelper.CurrentSelection(oSymbolDefinition, _
'                                                          oEndCutObject)
'        If Len(Trim(sCurrentChoice)) < 1 Then
'            If bTFL Or bBFL Then
'                sSelector = "Straight_LR"
'            Else
'                sSelector = "Straight_R"
'            End If
'
'            oSmartOccHelper.CurrentSelection(oSymbolDefinition, oEndCutObject) = sSelector
'        End If
'    End If
' ==================================================

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
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
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
' Method      : CMSetPhysicalConnAnswer
' Description :
'
'*********************************************************************************************
Public Sub CMSetPhysicalConnAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                   ByRef ppArgument As Object)

    Dim answer As String
    answer = "No"

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName pSL.SmartOccurrence, sACItemName, oACObject
    
    If sACItemName = "WPEF_PF_1" Or GetMbrAssemblyConnectionType(oACObject) = ACType_Stiff_Generic Or sACItemName = "Generic_2B_01" Then
        pSL.answer("CreatePhysicalConnection") = "Yes"
        Exit Sub
    End If
    
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = pSL.SmartOccurrence

    ' ----------------------------------------------------------------------
    ' If the bounding object is not a connectable, the answer is always "no"
    ' ----------------------------------------------------------------------
    If TypeOf oSDO_FlangeCut.BoundingPort Is IJPort Then
        ' --------------------------------
        ' If the bounding part is a member
        ' --------------------------------
        If TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_FlangeCut.Bounding Is IJProfile Then
            ' -----------------------------------------------------------------------
            ' If the flange is outside the bounding object, the answer is always "no"
            ' -----------------------------------------------------------------------
            Dim oTop As ConnectedEdgeInfo
            Dim oBottom As ConnectedEdgeInfo
            Dim oInsideTF As ConnectedEdgeInfo
            Dim oInsideBF As ConnectedEdgeInfo
            
            GetConnectedEdgeInfo oSDO_FlangeCut.object, oSDO_FlangeCut.BoundedPort, oSDO_FlangeCut.BoundingPort, oTop, oBottom, oInsideTF, oInsideBF
            
            Dim sIsBottom As String
            sIsBottom = pSL.SelectorAnswer(CUSTOMERID & "MbrEndCut.RootGenFlangeSel", "BottomFlange")
            
            If sIsBottom = "Yes" And (oInsideBF.IntersectingEdge = Below Or oInsideBF.IntersectingEdge = None Or oInsideBF.CoplanarEdge = Bottom) Then
                ' answer always no
            ElseIf sIsBottom = "No" And (oInsideTF.IntersectingEdge = Above Or oInsideTF.IntersectingEdge = None Or oInsideTF.CoplanarEdge = Top) Then
                ' answer always no
            
            ' --------------------------------------------------
            ' If there is an outside web cut, the answer is "no"
            ' --------------------------------------------------
            Else
                Dim sHasTopWeb As String
                Dim sHasBtmWeb As String
                
                GetSelectorAnswer oSDO_FlangeCut.WebCut, "CreateTopWebCut", sHasTopWeb
                GetSelectorAnswer oSDO_FlangeCut.WebCut, "CreateBottomWebCut", sHasBtmWeb
                
                If sHasBtmWeb = "" And sHasTopWeb = "" Then 'Flange cut must be driven by either Top/Bottom web cuts
                    'Get the parent web cut which should have answers for creating top/bottom webcuts.
                    Dim oParentObj As Object, sParentItemName As String
                    Parent_SmartItemName oSDO_FlangeCut.WebCut, sParentItemName, oParentObj
                    
                    GetSelectorAnswer oParentObj, "CreateTopWebCut", sHasTopWeb
                    GetSelectorAnswer oParentObj, "CreateBottomWebCut", sHasBtmWeb
                End If
    
                If Not (sIsBottom = "Yes" And sHasBtmWeb = "Yes") And Not (sIsBottom = "No" And sHasTopWeb = "Yes") Then
                    answer = "Yes"
                End If
            End If
        
        ElseIf TypeOf oSDO_FlangeCut.Bounding Is IJPlate Then
            Dim oStructPort As IJStructPort
            If TypeOf oSDO_FlangeCut.BoundingPort Is IJStructPort Then
                Set oStructPort = oSDO_FlangeCut.BoundingPort
                If oStructPort.ContextID = CTX_LATERAL_LFACE Then
                    answer = "Yes"
                Else
                    Dim oSGOModelBodyUtilities As New GSCADShipGeomOps.SGOModelBodyUtilities
                    Dim bHasintersection As Boolean
                    Dim oBoundingPort As IJPort
                    Dim oBoundedPort As IJPort

                    Set oBoundingPort = oSDO_FlangeCut.BoundingPort
                    
                        Dim oSL As IJDSelectorLogic
                        Set oSL = GetSelectorLogicForCustomMethod(pInput)
                    
                    sIsBottom = oSL.SelectorAnswer(CUSTOMERID & "MbrEndCut.RootGenFlangeSel", "BottomFlange")
                    

                    If sIsBottom = gsYes Then
                        Set oBoundedPort = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_BOTTOM_FLANGE_RIGHT)
                        bHasintersection = oSGOModelBodyUtilities.HasIntersectingGeometry(oBoundingPort.Geometry, GetExtendedPort(oBoundedPort))

                        If Not bHasintersection Then
                            Set oBoundedPort = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_BOTTOM_FLANGE_LEFT)
                            bHasintersection = oSGOModelBodyUtilities.HasIntersectingGeometry(oBoundingPort.Geometry, GetExtendedPort(oBoundedPort))

                            If Not bHasintersection Then
                                Set oBoundedPort = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_BOTTOM)
                                bHasintersection = oSGOModelBodyUtilities.HasIntersectingGeometry(oBoundingPort.Geometry, GetExtendedPort(oBoundedPort))
                            End If
                        End If
                    Else
                        Set oBoundedPort = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_TOP_FLANGE_RIGHT)
                        
                        bHasintersection = oSGOModelBodyUtilities.HasIntersectingGeometry(oBoundingPort.Geometry, GetExtendedPort(oBoundedPort)) ' oBoundedPort.Geometry)

                        If Not bHasintersection Then
                            Set oBoundedPort = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_TOP_FLANGE_LEFT)
                            bHasintersection = oSGOModelBodyUtilities.HasIntersectingGeometry(oBoundingPort.Geometry, GetExtendedPort(oBoundedPort))

                            If Not bHasintersection Then
                                Set oBoundedPort = GetLateralSubPortBeforeTrim(oSDO_FlangeCut.Bounded, JXSEC_TOP)
                                bHasintersection = oSGOModelBodyUtilities.HasIntersectingGeometry(oBoundingPort.Geometry, GetExtendedPort(oBoundedPort))
                            End If
                        End If
                    End If

                    If bHasintersection Then
                        answer = "Yes"
                    Else
                        answer = "No"
                    End If
                End If
            End If
        ' -------------------------
        ' Otherwise default to "No"
        ' -------------------------
        ' Can we do better than this?  Have to determine if it intersects this part *first*
        Else
            answer = "No"
        End If
    End If
        
    pSL.answer("CreatePhysicalConnection") = answer

End Sub

Public Sub CMSetInsideCornerFeatureAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                          ByRef ppArgument As Object)

    Dim answer As String
    answer = "No"

    ' ----------------------------------------------------------------------
    ' Create/Initialize the selector logic object from the symbol definition
    ' ----------------------------------------------------------------------
    Dim oSL As IJDSelectorLogic
    Set oSL = GetSelectorLogicForCustomMethod(pInput)
    
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oSL.SmartOccurrence

    ' -----------------------------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' -----------------------------------------------------
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName oSL.SmartOccurrence, sACItemName, oACObject
    
    If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Or sACItemName = "Generic_2B_01" Or sACItemName = "Gen_BearingPlateAC_01" Then
        oSL.answer("CreateInsideCornerFeature") = "No"
        Exit Sub
    End If
    
    ' ----------------------------------------------------------------------
    ' If the bounding object is not a connectable, the answer is always "no"
    ' ----------------------------------------------------------------------
    If TypeOf oSDO_FlangeCut.BoundingPort Is IJPort Then
        ' --------------------------------
        ' If the bounding part is a member
        ' --------------------------------
        If TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_FlangeCut.Bounding Is IJProfile Then
            ' --------------------------------------------------
            ' If there is an outside web cut, the answer is "no"
            ' --------------------------------------------------
            Dim sHasTopWeb As String
            Dim sHasBtmWeb As String
            Dim sIsBottom As String
            
            GetSelectorAnswer oSDO_FlangeCut.WebCut, "CreateTopWebCut", sHasTopWeb
            GetSelectorAnswer oSDO_FlangeCut.WebCut, "CreateBottomWebCut", sHasBtmWeb
            
            If sHasBtmWeb = "" And sHasTopWeb = "" Then 'Flange cut must be driven by either Top/Bottom web cuts
                'Get the parent web cut which should have answers for creating top/bottom webcuts.
                Dim oParentObj As Object, sParentItemName As String
                Parent_SmartItemName oSDO_FlangeCut.WebCut, sParentItemName, oParentObj
                
                GetSelectorAnswer oParentObj, "CreateTopWebCut", sHasTopWeb
                GetSelectorAnswer oParentObj, "CreateBottomWebCut", sHasBtmWeb
            End If
            sIsBottom = oSL.SelectorAnswer(CUSTOMERID & "MbrEndCut.RootGenFlangeSel", "BottomFlange")
            
            If (sIsBottom = "Yes" And sHasBtmWeb = "Yes") Or (sIsBottom = "No" And sHasTopWeb = "Yes") Then
                ' default to "no"
            ' -------------------------------------------------------------------------------
            ' Otherwise, if the flange intersects the face of the member, the answer is "Yes"
            ' -------------------------------------------------------------------------------
            Else
                Dim oTop As ConnectedEdgeInfo
                Dim oBottom As ConnectedEdgeInfo
                Dim oInsideTF As ConnectedEdgeInfo
                Dim oInsideBF As ConnectedEdgeInfo
                
                GetConnectedEdgeInfo oSDO_FlangeCut.object, oSDO_FlangeCut.BoundedPort, oSDO_FlangeCut.BoundingPort, oTop, oBottom, oInsideTF, oInsideBF
                  
                If sIsBottom = "Yes" And _
                   (oInsideBF.IntersectingEdge = Web_Right And oBottom.IntersectingEdge = Web_Right) Then
                    answer = "Yes"
                ElseIf sIsBottom = "No" And _
                   (oInsideTF.IntersectingEdge = Web_Right And oTop.IntersectingEdge = Web_Right) Then
                    answer = "Yes"
                End If
            End If
        ' ------------------------------------------------------------
        ' If the flange meets the face of a plate, the answer is "Yes"
        ' ------------------------------------------------------------
        ElseIf TypeOf oSDO_FlangeCut.Bounding Is IJPlate Then
            Dim oBoundingStructPort As IJStructPort
            Set oBoundingStructPort = oSDO_FlangeCut.BoundingPort
             
            If Not ((oBoundingStructPort.ContextID And CTX_LATERAL) = CTX_LATERAL) Then
                answer = "Yes"
            End If
        End If
        
    End If
        
    oSL.answer("CreateInsideCornerFeature") = answer

End Sub



