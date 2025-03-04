VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BottomFlangeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : BottomFlangeSel.cls
'
'  Description :
'       Selector for Selection of Member bounded by Port/Object End Cuts
'       Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    06/Sep/2011 - pnalugol
'           For supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'*********************************************************************************************

Const m_sClassName As String = "BottomFlangeSel"
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
    sMsg = "Defining/Initializing Selector Questions/Answers"

    ' TODO - Add your question below
    '===============================
    pQH.SetQuestion "SelectByRule", "Yes", "BooleanCol"
    pQH.SetQuestion "BottomFlange", "No", "BooleanCol"

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the Basic Type of Member End cut
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim oEndCutObject As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Unknown Error"

    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the FlangeCut Inputs from the StructFeature object
   
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut.object = oEndCutObject

    Set oBoundedObject = oSDO_FlangeCut.BoundedPort
    If (oBoundedObject Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If

    Set oBoundingObject = oSDO_FlangeCut.BoundingPort
    If (oBoundingObject Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean
    
    bIsBoundedTube = IsTubularMember(oSDO_FlangeCut.Bounded)
    bIsBoundingTube = IsTubularMember(oSDO_FlangeCut.Bounding)
    If bIsBoundingTube And (Not bIsBoundedTube) Then
        pSelectorLogic.Add "FlgBtmSnipeOutSide_Tube_FR"
        Exit Sub
    End If
    
    'Select the appropriate Top Flange Cut
    'Possible Top Flange Cuts Currently Include
    '   - Outside Snipe
    '   - Outside Snipe Offset Near
    '   - Square Flange Relief Outside Snipe
    '   - Inside Hook Flange Relief Outside Snipe
    '   - Outside Snipe Offset Top
    '   - Inside Hook Flange Relief
    '   - Square Flange Relief & Square Along Flange Relief
    '   - Outside Hook Flange Relief
    '   - Outside Hook Flange Relief With Snipe
    '   - Hook Face Relief
    '   - Snipe Face Relief
    
    'Get the Section Alias to Determine which Cuts are applicable
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pSelectorLogic.SmartOccurrence, oSDO_FlangeCut.BoundingPort, oSDO_FlangeCut.BoundedPort, lSectionAlias, bPenetratesWeb)
       
    Dim bBoundingBottomFlange As Boolean
    Select Case lSectionAlias
        Case 0, 8, 9, 10, 11, 12, 17
            'Flange, FlangeTopFlangeLeft, FlangeBuiltUpTopFlangeLeft, FlangeBottomFlangeLeft
            'FlangeBuiltUpBottomFlangeLeft, FlangeTopAndBottomLeftFlanges, TwoFlangesTwoFlanges,
            bBoundingBottomFlange = False
        Case 1, 6
            'FlangeTopFlangeRight, FlangeTopFlange
            bBoundingBottomFlange = False
        Case 2
            'FlangeBuiltUpTopFlangeRight
            bBoundingBottomFlange = False
        Case 3, 7
            'FlangeBottomFlangeRight, FlangeBottomFlange
            bBoundingBottomFlange = True
        Case 4
            'FlangeBuiltUpBottomFlangeRight
            bBoundingBottomFlange = True
        Case 5, 13, 19
            'FlangeTopAndBottomRightFlanges, FlangeTopAndBottomFlanges, TwoFlangesBetweenFlanges
            bBoundingBottomFlange = True
        Case 14
            'FlangeLeftAndRightBottomFlanges
            bBoundingBottomFlange = False
        Case 15
            'FlangeLeftAndRightTopFlanges
            bBoundingBottomFlange = False
        Case 16, 18
            'FlangeLeftAndRightFlanges, TwoFlangesBetweenFlanges
            bBoundingBottomFlange = False
        Case 20
            'Tube/Circular Cross Section
        Case Else
            'Unknown Section Alias
    End Select
    
      
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
    
    Dim oAppConn As IJAppConnection
    Do While Not TypeOf oParentObj Is IJAppConnection
        Parent_SmartItemName oParentObj, sParentItemName, oParentObj
    Loop
    Set oAppConn = oParentObj
    
    Dim bTFL As Boolean, bBFL As Boolean, bTFR As Boolean, bBFR As Boolean

    CrossSection_Flanges oSDO_FlangeCut.Bounded, bTFL, bBFL, bTFR, bBFR

    If GetMbrAssemblyConnectionType(oAppConn) = ACType_Mbr_Generic Or _
       GetMbrAssemblyConnectionType(oAppConn) = ACType_Stiff_Generic Then  ''End cut selection for Generic ACs
        If bBoundingBottomFlange Then
            'Flange Relief Cuts are valid  -  Face Relief Cuts may be valid
            pSelectorLogic.Add "FlgSqrFlangeReliefBtm"
            pSelectorLogic.Add "FlgSqrAlgFlangeReliefBtm"
            pSelectorLogic.Add "FlgInsideHookFlangeReliefBtm"
            pSelectorLogic.Add "FlgEdgeHookFlgRlfOutSideBtm"
            pSelectorLogic.Add "FlgSnipeFaceReliefBtm"
            pSelectorLogic.Add "FlgHookFaceReliefBtm"
            'cuts preferred when the bottom goes outside of the bounding
            pSelectorLogic.Add "FlgSqrFlgRlfOutsideSnpBtm_FR"
            pSelectorLogic.Add "FlgInsideHookFlgReliefSnipeBtm_FR"
            pSelectorLogic.Add "FlgEdgeHookFlgRlfOutSideSnpBtm_FR"
            pSelectorLogic.Add "FlgSqrAlgFlgRlfOutsideSnpBtm_FR"
            pSelectorLogic.Add "FlgBtmSnipeOutSide_FR"
            pSelectorLogic.Add "FlgBtmSnipeOutSideSetBack_FR"
            pSelectorLogic.Add "FlgOutsideSnipeBtmOffset"
            
        Else
            'Flange Relief Cuts are not valid  -  Face Relief Cuts are valid
            pSelectorLogic.Add "FlgHookFaceReliefBtm"
            pSelectorLogic.Add "FlgSnipeFaceReliefBtm"
        End If
    Else

        ' -----------------------------
        ' Get the requested bottomshape
        ' -----------------------------
        Dim sBottomAnswerCol As String
        Dim sBottomShape As String
        
        GetMemberACTopAndBottomShape pSelectorLogic.SmartOccurrence, sBottomAnswerCol, sBottomShape
        
        ' -------------------------------------------------------
        ' Call a separate subroutine based on the collection type
        ' -------------------------------------------------------
        Select Case sBottomAnswerCol
            Case gsShapeAtFaceCol
                SelectBtmFlangeCutForShapeAtFace pSelectorLogic, sBottomShape
            Case gsShapeAtEdgeCol
                SelectBtmFlangeCutForShapeAtEdge pSelectorLogic, sBottomShape
            Case gsShapeAtEdgeOverlapCol
                SelectBtmFlangeCutForShapeAtEdgeOverlap pSelectorLogic, sBottomShape
            Case gsShapeAtEdgeOutsideCol, vbNullString ' null string means the bottom cut is not at a face or edge (it's outside)
                SelectBtmFlangeCutForShapeAtEdge_ToOutside pSelectorLogic, sBottomShape
            
        End Select

    End If
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD

End Sub

' ** Start CM **
'*********************************************************************************************
' Method      : SelectBtmFlangeCutForShapeAtEdge
' Description : Selection of Flange top cut when top of the bounded is bounded to a flange of the bounding
'
'*********************************************************************************************
Private Sub SelectBtmFlangeCutForShapeAtEdge(pSelectorLogic As IJDSelectorLogic, sShapeAtEdge As String)
    
    Const METHOD = m_SelectorProgid & "::SelectBtmFlangeCutForShapeAtEdge"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Select Case sShapeAtEdge
        Case gsFaceToOutside, gsFaceToFlange
            pSelectorLogic.Add "FlgSqrFlangeReliefBtm"
        Case gsInsideToOutside, gsInsideToFlange
            pSelectorLogic.Add "FlgInsideHookFlangeReliefBtm"
        Case gsEdgeToFlange, gsEdgeToOutside
            pSelectorLogic.Add "FlgEdgeHookFlgRlfOutSideBtm"
        Case gsCornerToOutside, gsCornerToFlange
            pSelectorLogic.Add "FlgSqrAlgFlangeReliefBtm"
        Case Else
            pSelectorLogic.Add "FlgHookFaceReliefBtm"
            pSelectorLogic.Add "FlgOutsideSnipeBtmOffset"
    End Select
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectBtmFlangeCutForShapeAtFace
' Description : Selection of Flange Top cut when top of the bounded is bounded to face of the bounding
'
'*********************************************************************************************
Private Sub SelectBtmFlangeCutForShapeAtFace(pSelectorLogic As IJDSelectorLogic, sShapeAtFace As String)
    
    Const METHOD = m_SelectorProgid & "::SelectBtmFlangeCutForShapeAtFace"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Select Case sShapeAtFace
        Case gsCope
            pSelectorLogic.Add "FlgHookFaceReliefBtm"
        Case gsSnipe
            pSelectorLogic.Add "FlgSnipeFaceReliefBtm"
        Case Else
            pSelectorLogic.Add "FlgOutsideSnipeBtmOffset"
    End Select
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectBtmFlangeCutForShapeAtEdge_ToOutside
' Description : Selection of Btm web cut when top of the bounded is bounded to edge of the bounding in B/L cases.
'
'*********************************************************************************************
Private Sub SelectBtmFlangeCutForShapeAtEdge_ToOutside(pSelectorLogic As IJDSelectorLogic, sShapeAtEdge As String)
    Const METHOD = m_SelectorProgid & "::SelectBtmFlangeCutForShapeAtEdge_ToOutside"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim sExtendOrSetBack As String
    Dim sShapeOutSide As String
    
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
    
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBack
    GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide

    'when member is bounded to bottom flange of the bounding
    
    If sShapeAtEdge = vbNullString Then
        If sExtendOrSetBack = gsOffsetNearCorner Then
            If sShapeOutSide = gsStraight Then
                pSelectorLogic.Add "FlgBtmSnipeOutSideSetBack_FR"
            Else
                'For now both cases are handled with the same symbol
                pSelectorLogic.Add "FlgBtmSnipeOutSideSetBack_FR"
            End If
        Else
            If sShapeOutSide = gsStraight Then
                pSelectorLogic.Add "FlgBtmSnipeOutSide_FR"
            Else
                'For now both cases are handled with the same symbol
                pSelectorLogic.Add "FlgBtmSnipeOutSide_FR"
            End If
        End If
    Else
        Select Case sShapeAtEdge
            Case gsEdgeToOutside, gsCornerToOutside, gsCornerToFlange
                pSelectorLogic.Add "FlgEdgeHookFlangeReliefInsideBtm"
                
            Case gsOutsideToOutside, gsOutsideToFlange
                pSelectorLogic.Add "FlgBtmOutSideHookFlangeRelief"
            
            Case Else
                pSelectorLogic.Add "FlgEdgeHookFlangeReliefInsideBtm"
                
        End Select
    End If
            
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectBtmFlangeCutForShapeAtOutside
' Description : Selection of Flange Btm cut when top of the bounded is bounded to outside of the bounding in B/L cases.
'
'*********************************************************************************************
Private Sub SelectBtmFlangeCutForShapeAtEdgeOverlap(oSL As IJDSelectorLogic, sShapeAtEdge As String)
    
    Const METHOD = m_SelectorProgid & "::SelectBtmFlangeCutForShapeAtEdgeOverlap"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim sExtendOrSetBack As String
    Dim sShapeOutSide As String
    
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oSL.SmartOccurrence, sParentItemName, oParentObj
    
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBack
    GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
    
    If sShapeOutSide = vbNullString Then
        GetSelectorAnswer oParentObj, "BottomShapeOutside", sShapeOutSide
    End If
                        
    If sExtendOrSetBack = vbNullString Or sExtendOrSetBack = gsOffsetNearCorner Then
        
        Select Case sShapeAtEdge
        
            ' ------------------------------------------------------------------------------------------------
            ' These cases always have corner features, so the outside cut is based only on extension direction
            ' ------------------------------------------------------------------------------------------------
            Case gsNone, gsFaceToInsideCorner, gsFaceToEdge, gsInsideToEdge
                
                oSL.Add "FlgBtmSnipeOutSideSetBack_FR"
                
            ' ----------------------------------------------------------------------
            ' Pick a flange cut to resolve the shape if a corner feature is not used
            ' Otherwise pick one to only trim the outside material
            ' ----------------------------------------------------------------------
            ' It was hoped that we could use the snipe version even if there is a small amount of material outside, so
            ' that if conditions changed the same cut could be used and any overridden settings would not be lost
            ' But we've been unsuccessful getting the snipe case to resolve properly in all cases
            ' For now, we must switch between then based on the answer
            ' The default is Straight, and presumably the user will not want to use Sniped when there is a small amount
            ' of material
            Case gsFaceToOutsideCorner, gsFaceToOutside
                
                If sShapeOutSide = gsSniped Then
                    oSL.Add "FlgSqrFlgRlfOutsideSnpBtm_FR"
                Else
                    oSL.Add "FlgSqrAlgFlangeReliefBtm"
                End If

            Case gsInsideToOutsideCorner, gsInsideToOutside

                If sShapeOutSide = gsSniped Then
                    oSL.Add "FlgInsideHookFlgReliefSnipeBtm_FR"
                Else
                    oSL.Add "FlgInsideHookFlangeReliefBtm"
                End If
        
            Case gsInsideCornerToOutside

                If sShapeOutSide = gsSniped Then
                    oSL.Add "FlgSqrAlgFlgRlfOutsideSnpBtm_FR"
                Else
                    oSL.Add "FlgSqrAlgFlangeReliefBtm"
                End If
            
            Case gsEdgeToOutside
            
                If sShapeOutSide = gsSniped Then
                    oSL.Add "FlgEdgeHookFlgRlfOutSideSnpBtm_FR"
                Else
                    oSL.Add "FlgEdgeHookFlgRlfOutSideBtm"
                End If
        End Select
    Else
        oSL.Add "FlgBtmSnipeOutSide_FR"
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
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

' ********************************************************************************************
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

' ********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

' ********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

' ********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
