VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TopFlangeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : TopFlangeSel.cls
'
'  Description :
'       Selector for Selection of Member bounded by Port/Object End Cuts
'       Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    22/Aug/2011 - svsmylav
'           B-25298(CR-179926): standard member bounded to tube case
'           Added logic to select FlgTopSnipeOutSide_Tube_WL when no-left side flange exists
'           for bounded member.
'    06/Sep/2011 - pnalugol
'           for supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'
'    11/Apr/2012 - svsmylav
'       DM-213229: Updated logic to handle stiffener as bounded object.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'*********************************************************************************************

Const m_sClassName As String = "TopFlangeSel"
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
    
    Dim bTFL As Boolean, bBFL As Boolean, bTFR As Boolean, bBFR As Boolean

    CrossSection_Flanges oSDO_FlangeCut.Bounded, bTFL, bBFL, bTFR, bBFR
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean

    bIsBoundedTube = IsTubularMember(oSDO_FlangeCut.Bounded)
    bIsBoundingTube = IsTubularMember(oSDO_FlangeCut.Bounding)
    If bIsBoundingTube And (Not bIsBoundedTube) Then
        If bTFL Or bBFL Then
            pSelectorLogic.Add "FlgTopSnipeOutSide_Tube_FL"
        Else
            pSelectorLogic.Add "FlgTopSnipeOutSide_Tube_WL"
        End If
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
    Dim bBoundingTopFlange As Boolean
    Select Case lSectionAlias
        Case 0, 8, 9, 10, 11, 12, 17
            'Flange, FlangeTopFlangeLeft, FlangeBuiltUpTopFlangeLeft, FlangeBottomFlangeLeft
            'FlangeBuiltUpBottomFlangeLeft, FlangeTopAndBottomLeftFlanges, TwoFlangesTwoFlanges,
            bBoundingTopFlange = False
        Case 1, 6
            'FlangeTopFlangeRight, FlangeTopFlange
            bBoundingTopFlange = True
        Case 2
            'FlangeBuiltUpTopFlangeRight
            bBoundingTopFlange = True
        Case 3, 7
            'FlangeBottomFlangeRight, FlangeBottomFlange
            bBoundingTopFlange = False
        Case 4
            'FlangeBuiltUpBottomFlangeRight
            bBoundingTopFlange = False
        Case 5, 13, 19
            'FlangeTopAndBottomRightFlanges, FlangeTopAndBottomFlanges, TwoFlangesBetweenFlanges
            bBoundingTopFlange = True
        Case 14
            'FlangeLeftAndRightBottomFlanges
            bBoundingTopFlange = False
        Case 15
            'FlangeLeftAndRightTopFlanges
            bBoundingTopFlange = False
        Case 16, 18
            'FlangeLeftAndRightFlanges, TwoFlangesBetweenFlanges
            bBoundingTopFlange = False
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

    If GetMbrAssemblyConnectionType(oAppConn) = ACType_Mbr_Generic Or _
       GetMbrAssemblyConnectionType(oAppConn) = ACType_Stiff_Generic Then
       
        If bBoundingTopFlange Then
            'Flange Relief Cuts are valid  -  Face Relief Cuts may be valid
            pSelectorLogic.Add "FlgSqrFlangeReliefTop"
            pSelectorLogic.Add "FlgSqrAlgFlangeReliefTop"
            pSelectorLogic.Add "FlgInsideHookFlangeReliefTop"
            pSelectorLogic.Add "FlgEdgeHookFlgRlfOutSideTop"
            pSelectorLogic.Add "FlgSnipeFaceReliefTop"
            pSelectorLogic.Add "FlgHookFaceReliefTop"
            If bTFL Or bBFL Then
                pSelectorLogic.Add "FlgSqrFlgRlfOutsideSnpTop_FL"   'for cross sections which have FL(FlangeLeft)
                pSelectorLogic.Add "FlgInsideHookFlgReliefSnipeTop_FL" 'for cross sections which have FL(FlangeLeft)
                pSelectorLogic.Add "FlgEdgeHookFlgRlfOutSideSnpTop_FL" 'for cross sections which have FL(FlangeLeft)
                pSelectorLogic.Add "FlgSqrAlgFlgRlfOutsideSnpTop_FL"
                pSelectorLogic.Add "FlgTopSnipeOutSide_FL"
                pSelectorLogic.Add "FlgTopSnipeOutSideSetBack_FL"
                pSelectorLogic.Add "FlgOutsideSnipeTopOffset_FL"
            Else
                pSelectorLogic.Add "FlgSqrFlgRlfOutsideSnpTop_WL"
                pSelectorLogic.Add "FlgSqrAlgFlgRlfOutsideSnpTop_WL"
                pSelectorLogic.Add "FlgInsideHookFlgReliefSnipeTop_WL"
                pSelectorLogic.Add "FlgEdgeHookFlgRlfOutSideSnpTop_WL"
                pSelectorLogic.Add "FlgTopSnipeOutSide_WL"
                pSelectorLogic.Add "FlgTopSnipeOutSideSetBack_WL"
                pSelectorLogic.Add "FlgOutsideSnipeTopOffset_WL"
            End If
        Else
            'Flange Relief Cuts are not valid  -  Face Relief Cuts are valid
            pSelectorLogic.Add "FlgHookFaceReliefTop"
            pSelectorLogic.Add "FlgSnipeFaceReliefTop"
        End If
    Else
        ' -----------------------------
        ' Get the requested bottomshape
        ' -----------------------------
        Dim sTopAnswerCol As String
        Dim sTopShape As String
        
        GetMemberACTopAndBottomShape pSelectorLogic.SmartOccurrence, , , sTopAnswerCol, sTopShape
        
        ' -------------------------------------------------------
        ' Call a separate subroutine based on the collection type
        ' -------------------------------------------------------
        Select Case sTopAnswerCol
            Case gsShapeAtFaceCol
                SelectTopFlangeCutForShapeAtFace pSelectorLogic, sTopShape
            Case gsShapeAtEdgeCol
                SelectTopFlangeCutForShapeAtEdge pSelectorLogic, sTopShape
            Case gsShapeAtEdgeOverlapCol
                SelectTopFlangeCutForShapeAtEdgeOverlap pSelectorLogic, sTopShape
            Case gsShapeAtEdgeOutsideCol, vbNullString ' null string means the Top cut is not at a face or edge (it's outside)
                SelectTopFlangeCutForShapeAtEdge_ToOutside pSelectorLogic, sTopShape
        End Select
    End If
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD

End Sub

' ** Start CM **
'*********************************************************************************************
' Method      : SelectTopFlangeCutForShapeAtEdge
' Description : Selection of Flange top cut when top of the bounded is bounded to a flange of the bounding
'
'*********************************************************************************************

Private Sub SelectTopFlangeCutForShapeAtEdge(pSelectorLogic As IJDSelectorLogic, sShapeAtEdge As String)
    
    Const METHOD = m_SelectorProgid & "::SelectTopFlangeCutForShapeAtEdge"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Select Case sShapeAtEdge
        Case gsFaceToOutside, gsFaceToFlange
            pSelectorLogic.Add "FlgSqrFlangeReliefTop"
        Case gsInsideToOutside, gsInsideToFlange
            pSelectorLogic.Add "FlgInsideHookFlangeReliefTop"
        Case gsEdgeToFlange, gsEdgeToOutside
            pSelectorLogic.Add "FlgEdgeHookFlgRlfOutSideTop"
        Case gsCornerToOutside, gsCornerToFlange
            pSelectorLogic.Add "FlgSqrAlgFlangeReliefTop"
        Case Else
            
            Dim oBoundedPort As IJPort
            Dim oBoundingPort As IJPort
            GetBoundingAndBounded pSelectorLogic.SmartOccurrence, oBoundingPort, oBoundedPort
            
            Dim bTFL As Boolean
            Dim bTFR As Boolean
            Dim bBFL As Boolean
            Dim bBFR As Boolean
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

            pSelectorLogic.Add "FlgHookFaceReliefTop"
            
            If bTFL Or bBFL Then
                pSelectorLogic.Add "FlgOutsideSnipeTopOffset_FL"
            Else
                pSelectorLogic.Add "FlgOutsideSnipeTopOffset_WL"
            End If
            
    End Select
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectTopFlangeCutForShapeAtFace
' Description : Selection of Flange Top cut when top of the bounded is bounded to face of the bounding
'
'*********************************************************************************************
Private Sub SelectTopFlangeCutForShapeAtFace(pSelectorLogic As IJDSelectorLogic, sShapeAtFace As String)
    
    Const METHOD = m_SelectorProgid & "::SelectTopFlangeCutForShapeAtFace"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
        
    Select Case sShapeAtFace
        Case gsCope
            pSelectorLogic.Add "FlgHookFaceReliefTop"
        Case gsSnipe
            pSelectorLogic.Add "FlgSnipeFaceReliefTop"
        Case Else
        
            Dim oBoundedPort As IJPort
            Dim oBoundingPort As IJPort
            GetBoundingAndBounded pSelectorLogic.SmartOccurrence, oBoundingPort, oBoundedPort
            
            Dim bTFL As Boolean
            Dim bTFR As Boolean
            Dim bBFL As Boolean
            Dim bBFR As Boolean
            CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
        
            If bTFL Or bBFL Then
                pSelectorLogic.Add "FlgOutsideSnipeTopOffset_FL"
            Else
                pSelectorLogic.Add "FlgOutsideSnipeTopOffset_WL"
            End If
    End Select
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectTopFlangeCutForShapeAtEdgeOverlap
' Description : Selection of Top web cut when top of the bounded is bounded to outside of the bounding edge.
'
'*********************************************************************************************
Private Sub SelectTopFlangeCutForShapeAtEdgeOverlap(oSL As IJDSelectorLogic, sShapeAtEdge As String)
    
    Const METHOD = m_SelectorProgid & "::SelectTopFlangeCutForShapeAtEdgeOverlap"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim sShapeOutSide As String
    Dim sExtendOrSetBack As String
    Dim dOffset As Double

    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oSL.SmartOccurrence, sParentItemName, oParentObj

    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    GetBoundingAndBounded oSL.SmartOccurrence, oBoundingPort, oBoundedPort
    
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR

    GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBack
    GetSelectorAnswer oParentObj, "Offset", dOffset

    If sShapeOutSide = vbNullString Then
        GetSelectorAnswer oParentObj, "TopShapeOutside", sShapeOutSide
    End If

    If sExtendOrSetBack = vbNullString Or sExtendOrSetBack = gsOffsetNearCorner Then
        
        Select Case sShapeAtEdge

            Case gsNone, gsFaceToInsideCorner, gsFaceToEdge, gsInsideToEdge

                If bTFL Or bBFL Then
                    oSL.Add "FlgTopSnipeOutSideSetBack_FL"
                Else
                    oSL.Add "FlgTopSnipeOutSideSetBack_WL"
               End If

            Case gsFaceToOutsideCorner, gsFaceToOutside

                If bTFL Or bBFL Then
                    oSL.Add "FlgSqrFlgRlfOutsideSnpTop_FL"
                Else
                    oSL.Add "FlgSqrFlgRlfOutsideSnpTop_WL"
                End If

            Case gsInsideCornerToOutside
                If bTFL Or bBFL Then
                    If sShapeOutSide = gsSniped Then
                        oSL.Add "FlgSqrAlgFlgRlfOutsideSnpTop_FL"
                    Else
                        oSL.Add "FlgSqrAlgFlangeReliefTop"
                    End If
                Else
                    If sShapeOutSide = gsSniped Then
                        oSL.Add "FlgSqrAlgFlgRlfOutsideSnpTop_WL"
                    Else
                        oSL.Add "FlgSqrAlgFlgRlfOutsideSnpTop_WL" ' ???
                    End If
                End If

            Case gsInsideToOutsideCorner, gsInsideToOutside
                If bTFL Or bBFL Then
                    If sShapeOutSide = gsSniped Then
                        oSL.Add "FlgInsideHookFlgReliefSnipeTop_FL"
                    Else
                        oSL.Add "FlgInsideHookFlangeReliefTop"
                    End If
                Else
                    If sShapeOutSide = gsSniped Then
                        oSL.Add "FlgInsideHookFlgReliefSnipeTop_WL"
                    Else
                        oSL.Add "FlgInsideHookFlgReliefSnipeTop_WL" ' ???
                    End If
                End If

            Case gsEdgeToOutside, gsCornerToOutside
                If bTFL Or bBFL Then
                    If sShapeOutSide = gsSniped Then
                        oSL.Add "FlgEdgeHookFlgRlfOutSideSnpTop_FL"
                    Else
                        oSL.Add "FlgEdgeHookFlgRlfOutSideTop"
                    End If
                Else
                    If sShapeOutSide = gsSniped Then
                        oSL.Add "FlgEdgeHookFlgRlfOutSideSnpTop_WL"
                    Else
                        oSL.Add "FlgEdgeHookFlgRlfOutSideSnpTop_WL" '???
                    End If
                End If
        End Select
    Else
        If bTFL Or bBFL Then
            oSL.Add "FlgTopSnipeOutSide_FL"
        Else
            oSL.Add "FlgTopSnipeOutSide_WL"
        End If
    End If

    
    Exit Sub
ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
End Sub
        
'*********************************************************************************************
' Method      : SelectTopFlangeCutForShapeAtEdge_ToOutside
' Description : Selection of Top web cut when top of the bounded is bounded to edge of the bounding in B/L cases.
'
'*********************************************************************************************
Private Sub SelectTopFlangeCutForShapeAtEdge_ToOutside(pSelectorLogic As IJDSelectorLogic, sShapeAtEdge As String)
    Const METHOD = m_SelectorProgid & "::SelectTopFlangeCutForShapeAtEdge_ToOutside"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    Dim sExtendOrSetBack As String
    Dim sShapeOutSide As String
    
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
    
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBack
    GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    GetBoundingAndBounded pSelectorLogic.SmartOccurrence, oBoundingPort, oBoundedPort
    
    Dim bTFL As Boolean
    Dim bTFR As Boolean
    Dim bBFL As Boolean
    Dim bBFR As Boolean
    CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
    
    If sShapeAtEdge = vbNullString Then
        If sExtendOrSetBack = gsOffsetNearCorner Then
            If sShapeOutSide = gsStraight Then
                If bTFL Or bBFL Then
                    pSelectorLogic.Add "FlgTopSnipeOutSideSetBack_FL"
                Else
                    pSelectorLogic.Add "FlgTopSnipeOutSideSetBack_WL"
                End If
            Else
                'For now both cases are handled with the same symbol
                    If bTFL Or bBFL Then
                        pSelectorLogic.Add "FlgTopSnipeOutSideSetBack_FL"
                    Else
                        pSelectorLogic.Add "FlgTopSnipeOutSideSetBack_WL"
                    End If
            End If
        Else
            If sShapeOutSide = gsStraight Then
                If bTFL Or bBFL Then
                    pSelectorLogic.Add "FlgTopSnipeOutSide_FL"
                Else
                    pSelectorLogic.Add "FlgTopSnipeOutSide_WL"
                End If
            Else
                'For now both cases are handled with the same symbol
                If bTFL Or bBFL Then
                    pSelectorLogic.Add "FlgTopSnipeOutSide_FL"
                Else
                    pSelectorLogic.Add "FlgTopSnipeOutSide_WL"
                End If
            End If
        End If
    Else
    
        Select Case sShapeAtEdge
            Case gsEdgeToOutside, gsCornerToOutside, gsCornerToFlange
                pSelectorLogic.Add "FlgEdgeHookFlangeReliefInsideTop"
                
            Case gsOutsideToOutside, gsOutsideToFlange
                pSelectorLogic.Add "FlgTopOutSideHookFlangeRelief"
            
            Case Else
                pSelectorLogic.Add "FlgEdgeHookFlangeReliefInsideTop"
                
        End Select
        
    End If
            
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
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
