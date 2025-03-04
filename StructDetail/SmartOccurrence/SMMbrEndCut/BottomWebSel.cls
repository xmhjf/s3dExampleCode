VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BottomWebSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : TopWebSel.cls
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
'    15Sep2011  -   pnalugol
'               Modified/Added methods to support Braces on Flange penetrated cases
'    19/Sep/2011 -  CM - TR-CP-202051 "EdgeToFlange and EdgeToOutSide shapes doesn't work for ShapeAtEdge"
'                                     In this file change was done in SelectBtmWebCutForShapeAtEdge()
'                                     to select a appropriate item instead
'                                     of old(which did not exist in Catalog)
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'*********************************************************************************************

Const m_sClassName As String = "BottomWebSel"
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

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the Basic Type of EndCut here
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Unknown Error"
    
    ' ----------------------------
    ' Get a helper for the web cut
    ' ----------------------------
    Dim oSDO_WebCut As StructDetailObjects.WebCut

    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pSelectorLogic.SmartOccurrence

    ' ----------------------------------------------------------
    ' If bounding is tubular, there is currently only one option
    ' ----------------------------------------------------------
    If IsTubularMember(oSDO_WebCut.Bounding) Then
        pSelectorLogic.Add "BottomSnipeOutside_Tube"
        Exit Sub
    End If

    ' ---------------------------
    ' Get the assembly connection
    ' ---------------------------
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pSelectorLogic.SmartOccurrence, sParentItemName, oParentObj
    
    Dim oAppConn As IJAppConnection
    Do While Not TypeOf oParentObj Is IJAppConnection
        Parent_SmartItemName oParentObj, sParentItemName, oParentObj
    Loop
    Set oAppConn = oParentObj
    
    ' -------------------------------------------------------------------------------------------------------------
    ' If not web penetrated, the only current reason the AC definition rule would create this is for an inset brace
    ' -------------------------------------------------------------------------------------------------------------
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMap As New Collection

    Set oEdgeMap = GetEdgeMap(pSelectorLogic.SmartOccurrence, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
    If Not bPenetratesWeb Then
         pSelectorLogic.Add "HookFaceReliefBottom"
        Exit Sub
    End If
    
    ' ---------------------------
    ' If for a generic connection
    ' ---------------------------
    If GetMbrAssemblyConnectionType(oAppConn) = ACType_Mbr_Generic Or GetMbrAssemblyConnectionType(oAppConn) = ACType_Stiff_Generic Then
               
        ' -----------------------------------------------------------------------------------------------
        ' If the alias has bottom flange, make available all the cuts that are related to a bounding edge
        ' -----------------------------------------------------------------------------------------------
        Dim boundingAlias As eBounding_Alias
        boundingAlias = GetBoundingAliasSimplified(lSectionAlias)
    
        Dim bBoundingBottomFlange As Boolean
    
        Select Case boundingAlias
            Case eBounding_Alias.WebBottomFlangeRight, eBounding_Alias.WebBuiltUpBottomFlangeRight, eBounding_Alias.WebTopAndBottomRightFlanges
                bBoundingBottomFlange = True
            Case Else
                bBoundingBottomFlange = False
        End Select
        
        If bBoundingBottomFlange Then
            pSelectorLogic.Add "SquareFlangeReliefBottom"
            pSelectorLogic.Add "SquareAlongFlangeReliefBottom"
            pSelectorLogic.Add "InsideHookFlangeReliefBottom"
            pSelectorLogic.Add "HookFaceReliefBottom"
            pSelectorLogic.Add "SnipeFaceReliefBottom"
            pSelectorLogic.Add "OutsideSnipeBottomOffset"
            pSelectorLogic.Add "BottomSnipeOutside"
            pSelectorLogic.Add "BottomSnipeOutsideSetback"
            pSelectorLogic.Add "EdgeHookFlangeReliefInsideWebBottom"
            pSelectorLogic.Add "SquareFlangeReliefWebBottomCutSnipe"
            pSelectorLogic.Add "SquareAlongFlangeReliefWebBottomCutSnipe"
            pSelectorLogic.Add "EdgeHookFlangeReliefOutsideWebBottom"
            pSelectorLogic.Add "EdgeHookFlangeReliefOutsideWebBottomSnipe"
        ' --------------------------------------------------------------------------
        ' Otherwise, make available all the cuts that are related to a bounding face
        ' --------------------------------------------------------------------------
        Else
            pSelectorLogic.Add "HookFaceReliefBottom"
            pSelectorLogic.Add "SnipeFaceReliefBottom"
            pSelectorLogic.Add "OutsideSnipeBottomOffset"
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
                SelectBtmWebCutForShapeAtFace pSelectorLogic, sBottomShape
            Case gsShapeAtEdgeCol
                SelectBtmWebCutForShapeAtEdge pSelectorLogic, sBottomShape
            Case gsShapeAtEdgeOverlapCol
                SelectBtmWebCutForShapeAtEdgeOverlap pSelectorLogic, sBottomShape
            Case gsShapeAtEdgeOutsideCol, vbNullString ' null string means the bottom cut is not at a face or edge (it's outside)
                SelectBtmWebCutForShapeAtEdge_ToOutside pSelectorLogic, sBottomShape
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

'*********************************************************************************************
' Method      : SelectBtmWebCutForShapeAtEdge
' Description : Selection of Btm web cut when Btm of the bounded is bounded to a flange of the bounding
'
'*********************************************************************************************
Private Sub SelectBtmWebCutForShapeAtEdge(oSL As IJDSelectorLogic, sShapeAtEdge As String)
    
    Const METHOD = m_SelectorProgid & "::SelectBtmWebCutForShapeAtEdge"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Select Case sShapeAtEdge
        
        ' FaceToCorner and FaceToEdge are treated like FaceToOutside when there is no flange and insufficient overlap
        ' FaceToFlange is treated like FaceToOutside when there is no flange
        Case gsFaceToCorner, gsFaceToEdge, gsFaceToFlange, gsFaceToOutside
        
            oSL.Add "SquareFlangeReliefBottom"
    
        ' gsInsideToEdge is treated like InsideToOutside when there is no flange and insufficient overlap
        ' gsInsideToFlange is treated like InsideToOutside when there is no flange
        Case gsInsideToEdge, gsInsideToFlange, gsInsideToOutside
        
            oSL.Add "InsideHookFlangeReliefBottom"
        
        ' CornerToFlange is treated like CornerToOutside when there is no flange
        Case gsCornerToFlange, gsCornerToOutside
        
            oSL.Add "InsideHookFlangeReliefBottom"
            oSL.Add "SquareAlongFlangeReliefBottom"
            
        ' EdgeToFlange is treated like EdgeToOutside when there is no flange
        Case gsEdgeToFlange, gsEdgeToOutside
        
            oSL.Add "EdgeHookFlangeReliefOutsideWebBottom"

    End Select
    
    Exit Sub
    
ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
End Sub
        
'*********************************************************************************************
' Method      : SelectBtmWebCutForShapeAtFace
' Description : Selection of Btm web cut when Btm of the bounded is bounded to face of the bounding
'
'*********************************************************************************************
Private Sub SelectBtmWebCutForShapeAtFace(oSL As IJDSelectorLogic, sShapeAtFace As String)
    
    Const METHOD = m_SelectorProgid & "::SelectBtmWebCutForShapeAtFace"
    
    Dim sMsg As String
    
    On Error GoTo ErrorHandler
    
    Select Case sShapeAtFace
        Case gsCope
            oSL.Add "HookFaceReliefBottom"
        Case gsSnipe
            oSL.Add "SnipeFaceReliefBottom"
            oSL.Add "OutsideSnipeBottomOffset"
        Case Else
            oSL.Add "OutsideSnipeBottomOffset"
    End Select
            
    Exit Sub

ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectBtmWebCutForShapeAtEdge_ToOutside
' Description : Selection of Btm web cut when Btm of the bounded is bounded to edge of the bounding in B/L cases.
'
'*********************************************************************************************
Private Sub SelectBtmWebCutForShapeAtEdge_ToOutside(oSL As IJDSelectorLogic, sShapeAtEdge As String)
    
    Const METHOD = m_SelectorProgid & "::SelectBtmWebCutForShapeAtEdge_ToOutside"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Select Case sShapeAtEdge
        
        ' EdgeToFlange/CornerToFlange treated like EdgeToOutside/CornerToOutside when there is no flange
        ' If there is a flange, we don't expect that
        Case gsEdgeToFlange, gsEdgeToOutside, gsCornerToFlange
            oSL.Add "EdgeHookFlangeReliefInsideWebBottom"

        ' OutsideToEdge is treated like OutsideToOutside when there is no flange and insufficient overlap
        ' OutsideToFlange is treated like OutsideToOutside when there is no flange
        Case gsOutsideToEdge, gsOutsideToFlange, gsOutsideToOutside, gsCornerToOutside
            oSL.Add "BottomOutsideHookFlangeRelief"
        
        ' Null string means that the cut is on the outside
        ' Choose an extended or setback cut
        Case vbNullString
        
            Dim sParentName As String
            Dim oParentObj As Object
            Parent_SmartItemName oSL.SmartOccurrence, sParentName, oParentObj
    
            Dim sExtendOrSetBack As String
            GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBack


            If sExtendOrSetBack = gsOffsetNearCorner Then
                oSL.Add "BottomSnipeOutsideSetback"
            Else
                oSL.Add "BottomSnipeOutside"
            End If
            
    End Select
            
    Exit Sub
    
ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectBtmWebCutForShapeAtEdgeOverlap
' Description : Selection of Btm web cut when Btm of the bounded is bounded to outside of the bounding edge.
'
'*********************************************************************************************
Private Sub SelectBtmWebCutForShapeAtEdgeOverlap(oSL As IJDSelectorLogic, sShapeAtEdge As String)
    
    Const METHOD = m_SelectorProgid & "::SelectBtmWebCutForShapeAtEdgeOverlap"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim sExtendOrSetBack As String
    Dim dOffset As Double
    Dim sShapeOutSide As String
    
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oSL.SmartOccurrence, sParentItemName, oParentObj
    
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBack
    GetSelectorAnswer oParentObj, "Offset", dOffset
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
                
                oSL.Add "BottomSnipeOutsideSetback"
                
            ' -------------------------------------------------------------------
            ' Pick a web cut to resolve the shape if a corner feature is not used
            ' Otherwise pick one to only trim the outside material
            ' -------------------------------------------------------------------
            ' It was hoped that we could use the snipe version even if there is a small amount of material outside, so
            ' that if conditions changed the same cut could be used and any overridden settings would not be lost
            ' But we've been unsuccessful getting the snipe case to resolve properly in all cases
            ' For now, we must switch between then based on the answer
            ' The default is Straight, and presumably the user will not want to use Sniped when there is a small amount
            ' of material
            Case gsFaceToOutsideCorner, gsFaceToOutside
                
                If sShapeOutSide = gsSniped Then
                    oSL.Add "SquareFlangeReliefWebBottomCutSnipe"
                Else
                    oSL.Add "SquareFlangeReliefBottom"
                End If
            
            Case gsInsideToOutsideCorner, gsInsideToOutside
                
                If sShapeOutSide = gsSniped Then
                    oSL.Add "InsideHookFlangeReliefSnipeBottom"
                Else
                    oSL.Add "InsideHookFlangeReliefBottom"
                End If
            
            Case gsInsideCornerToOutside
                
                If sShapeOutSide = gsSniped Then
                    oSL.Add "SquareAlongFlangeReliefWebBottomCutSnipe"
                Else
                    oSL.Add "SquareAlongFlangeReliefBottom"
                End If
            
            Case gsEdgeToOutside
                
                If sShapeOutSide = gsSniped Then
                    oSL.Add "EdgeHookFlangeReliefOutsideWebBottomSnipe"
                Else
                    oSL.Add "EdgeHookFlangeReliefOutsideWebBottom"
                End If
        End Select
    Else
        oSL.Add "BottomSnipeOutside"
    End If

    Exit Sub
    
ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
End Sub

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
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
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
