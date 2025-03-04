VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrEndCutCornerSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-14, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMCornerFeatRules
'  File        : MbrEndCutCornerSel.cls
'
'  Description :
'       It selects the list of available Items or SmartClasses in the context of
'       the smart occurrence.
'
'       - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'       - Questions are defined to parameterize the selection
'
'  Author      : Alligators
'
'  History     :
'    18/APR/2011 - Creation
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    22/Sep/2011 - svsmylav TR-202526: edge id of CF is corrected to JXSEC_BOTTOM_FLANGE_RIGHT/JXSEC_TOP_FLANGE_RIGHT in If condition
'                  in GetProjBoundingFlangeThickness method (earlier check was using bottom flange top/top flange bottom: this won't work for CF on outside corner).
'                  Also added Top/bottom port edge ids in If condition to support CF at outside top/bottom corners.
'
'   27/Feb/2012 - svsmylav/GH TR-211221: Rolled backed the temporary workaround done for CR-195749
'                 i.e., other than L and 2L cases CF would not create when bottom is face port.
'
'   22/Mar/2012 -GH TR-201127: Modified the GetInsideRatholeItemName() method to select proper Corner Feature
'                              at top when bounded has top flange and at bottom when bounded has bottom flanges.
'   13/Apr/2012 - Alligators TR214704: For Generic AC case, if TopEdge/BottomEdge Inside Corner is 'Yes': first two items of CF selection
'                 now use GetInsideRatholeItemName method call so that InsideFaceRathole/InsideFaceRathole2 and FaceToInsideRathole/FaceToInsideRathole2
'                 are listed respectively.
'
'*********************************************************************************************
Const m_SelectorProgid As String = CUSTOMERID + "CornerFeatRules.MbrEndCutCornerSel"
Const m_SelectorName As String = CUSTOMERID + "CornerFeatRules.MbrEndCutCornerSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\MbrEndCutCornerSel.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description :
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
 
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description :
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    ' Define questions
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
 
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description :
'
'*********************************************************************************************
Public Sub SelectorLogic(oSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    ' -------------------------------------
    ' Create a corner feature wrapper class
    ' -------------------------------------
    Dim oSDOCorner As New StructDetailObjects.CornerFeature
    Set oSDOCorner.object = oSL.SmartOccurrence
    
    ' -----------------------------------
    ' Get the AC object and its item name
    ' -----------------------------------
    ' Special-case generic ACs predefine the desired answer
    ' Set the values, even if the default would match what is desired by the AC
    ' 1) to guard against changes in the default logic affecting the special AC
    ' 2) to avoid recomputing a lot of information, such as where the bounding flange intersects
    Dim sACItemName As String
    Dim oACObject As IJAppConnection
    AssemblyConnection_SmartItemName oSL.SmartOccurrence, sACItemName, oACObject
    
    ' ----------------------------------------------------------------------------------
    ' If not a generic AC, call separate method to execute selection logic, and exit sub
    ' ----------------------------------------------------------------------------------
    If GetMbrAssemblyConnectionType(oACObject) = ACType_Axis Or _
       GetMbrAssemblyConnectionType(oACObject) = ACType_Bounded Then
        ' ****** Axis member end-cuts ******
        MbrAxisEndCutCF_SelectorLogic oSL
        Exit Sub
    End If
    
    ' ------------------------------------------------------------------------
    ' Continue on for generic cases by getting the port info for the two ports
    ' ------------------------------------------------------------------------
    ' Get the edge port section IDs
    Dim oEdgePort1 As IJStructPort
    Dim oEdgePort2 As IJStructPort
    
    Set oEdgePort1 = oSL.InputObject(INPUT_PORT2EDGE)
    Set oEdgePort2 = oSL.InputObject(INPUT_PORT3EDGE)
    
    ' -----------------------------------------
    ' Find any endcuts used to create the ports
    ' -----------------------------------------
    Dim oCut1 As IJStructFeature
    Dim oCut2 As IJStructFeature
    Dim cut1Type As StructFeatureTypes
    Dim cut2Type As StructFeatureTypes
    Dim oCut1BoundedPort As IJPort
    Dim oCut1BoundingPort As IJPort
    Dim oCut2BoundedPort As IJPort
    Dim oCut2BoundingPort As IJPort
    Dim strRatholeName As String
    
    cut1Type = SF_CornerFeature
    cut2Type = SF_CornerFeature
    
    Dim oDetailHelper As IJStructDetailHelper
    Set oDetailHelper = New StructDetailHelper
    
    Dim oOperation As IJStructOperation
    
    If Not oEdgePort1.ContextID And CTX_LATERAL Then
        oDetailHelper.FindOperatorForOperationInGraphByID oSDOCorner.GetPartObject, _
                                                          oEdgePort1.operationID, _
                                                          oEdgePort1.operatorID, _
                                                          oOperation, _
                                                          oCut1
    End If
    
    Set oOperation = Nothing
    
    If Not oEdgePort2.ContextID And CTX_LATERAL Then
        oDetailHelper.FindOperatorForOperationInGraphByID oSDOCorner.GetPartObject, _
                                                          oEdgePort2.operationID, _
                                                          oEdgePort2.operatorID, _
                                                          oOperation, _
                                                          oCut2
    End If
    
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
    
    If Not oCut1 Is Nothing Then
        cut1Type = oCut1.get_StructFeatureType
        If cut1Type = SF_WebCut Then
            Set oSDOWebCut.object = oCut1
            Set oCut1BoundedPort = oSDOWebCut.BoundedPort
            Set oCut1BoundingPort = oSDOWebCut.BoundingPort
        Else
            Set oSDOFlangeCut.object = oCut1
            Set oCut1BoundedPort = oSDOFlangeCut.BoundedPort
            Set oCut1BoundingPort = oSDOFlangeCut.BoundingPort
        End If
    End If
    
    If Not oCut2 Is Nothing Then
        cut2Type = oCut2.get_StructFeatureType
        If cut2Type = SF_WebCut Then
            Set oSDOWebCut.object = oCut2
            Set oCut2BoundedPort = oSDOWebCut.BoundedPort
            Set oCut2BoundingPort = oSDOWebCut.BoundingPort
        Else
            Set oSDOFlangeCut.object = oCut2
            Set oCut2BoundedPort = oSDOFlangeCut.BoundedPort
            Set oCut2BoundingPort = oSDOFlangeCut.BoundingPort
        End If
    End If
    

    ' -------------------------------------------
    ' If the corner is between two different cuts
    ' -------------------------------------------
    If Not oCut1 Is oCut2 Then
        ' -----------------------------------------------
        ' If the AC is a special type, choose accordingly
        ' -----------------------------------------------
        If (sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1") And _
          ((cut1Type = SF_WebCut And cut2Type = SF_FlangeCut) Or (cut2Type = SF_WebCut And cut1Type = SF_FlangeCut)) Then
            oSL.Add "Edge Face Rathole"
        ElseIf sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Or (cut1Type = SF_WebCut And cut2Type = SF_WebCut) Then
            oSL.Add "Member Step Snipe"
        ' -----------------------------------------------------------------
        ' If between a flange cut and its input web cut, choose accordingly
        ' -----------------------------------------------------------------
        ' if it's a web cut and the section ID is WEB_RIGHT, it is to the face
        ElseIf (cut1Type = SF_WebCut And cut2Type = SF_FlangeCut) Or (cut2Type = SF_WebCut And cut1Type = SF_FlangeCut) Then
            ' The proper helper objects happen to be set above
            oSL.Add "Edge Face Rathole"
        ' --------------------------------------------------------------------------
        ' Otherwise, offer all applicable type (those not relying on two edge faces)
        ' --------------------------------------------------------------------------
        Else
            oSL.Add "Member Snipe"
            oSL.Add "Member Scallop"
            oSL.Add "Member Snipe Along Corner"
            oSL.Add "Member Scallop Along Corner"
            oSL.Add "Elliptical Fillet"
            oSL.Add "Member Step Snipe"
        End If
        
    ElseIf sACItemName = "Gen_MultiBounding_01" Then
            MultiBoundingCF_Selection oSL, oCut1
    
    ' ----------------------------------------------------------------------
    ' If the corner is on a center cut between the bounding web and a flange
    ' ----------------------------------------------------------------------
    ElseIf (oEdgePort1.SectionID = JXSEC_WEB_RIGHT And _
           (oEdgePort2.SectionID = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or oEdgePort2.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT_TOP)) Or _
           (oEdgePort2.SectionID = JXSEC_WEB_RIGHT And _
           (oEdgePort1.SectionID = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or oEdgePort1.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT_TOP)) Then
        ' -----------------------------------------------
        ' If the AC is a special type, choose accordingly
        ' -----------------------------------------------
        If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
            oSL.Add "Member Snipe"
        ' -----------------------------------
        ' Otherwise, add all applicable types
        ' -----------------------------------
        Else
            oSL.Add "Member Snipe"
            oSL.Add "Member Scallop"
            oSL.Add "Member Snipe Along Corner"
            oSL.Add "Member Scallop Along Corner"
            oSL.Add "Elliptical Fillet"
        End If
    ' -----------------------------------------------------------------
    ' If the corner is on a center cut at the inside of a bounding edge
    ' -----------------------------------------------------------------
    ElseIf (oEdgePort1.SectionID = JXSEC_TOP_FLANGE_RIGHT_BOTTOM And oEdgePort2.SectionID = JXSEC_TOP_FLANGE_RIGHT) Or _
           (oEdgePort1.SectionID = JXSEC_TOP_FLANGE_RIGHT And oEdgePort2.SectionID = JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Or _
           (oEdgePort1.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT_TOP And oEdgePort2.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT) Or _
           (oEdgePort1.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT And oEdgePort2.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT_TOP) Then
        ' -----------------------------------------------
        ' If the AC is a special type, choose accordingly
        ' -----------------------------------------------
        If sACItemName = "WPEF_PF_1" Or sACItemName = "Generic_Stiff_WPEF_PF_1" Then
        
            strRatholeName = GetInsideRatholeItemName(oACObject, oEdgePort1.SectionID, oEdgePort2.SectionID, oCut1BoundedPort, oCut1BoundingPort, False)
        
            If Not strRatholeName = vbNullString Then
                oSL.Add strRatholeName
            End If
            
        ' -----------------------------------
        ' Otherwise, add all applicable types
        ' -----------------------------------
        Else
            strRatholeName = GetInsideRatholeItemName(oACObject, oEdgePort1.SectionID, oEdgePort2.SectionID, oCut1BoundedPort, oCut1BoundingPort, False)
            If Not strRatholeName = vbNullString Then
                oSL.Add strRatholeName
                oSL.Add GetInsideRatholeItemName(oACObject, oEdgePort1.SectionID, oEdgePort2.SectionID, oCut1BoundedPort, oCut1BoundingPort, True)
            End If
            
            oSL.Add "Edge Hook Relief"
            oSL.Add "Short Edge Rathole"
            oSL.Add "Two Edge Rathole"
            oSL.Add "Elliptical Fillet"
            oSL.Add "Edge Face Rathole"
            oSL.Add "Arc Corner Relief"
            oSL.Add "Outside Hook Relief"
            oSL.Add "Inside And Outside Hook Relief"
            oSL.Add "Corner To FarFace Edge Relief"
        End If
    ' ------------------------------------------------------------------------------------
    ' If between an idealized face and top or bottom from the same cut, choose accordingly
    ' ------------------------------------------------------------------------------------
    ElseIf (oEdgePort1.SectionID = JXSEC_IDEALIZED_BOUNDARY And (oEdgePort2.SectionID = JXSEC_TOP Or oEdgePort2.SectionID = JXSEC_BOTTOM)) Or _
           (oEdgePort2.SectionID = JXSEC_IDEALIZED_BOUNDARY And (oEdgePort1.SectionID = JXSEC_TOP Or oEdgePort1.SectionID = JXSEC_BOTTOM)) Then
           
        
        strRatholeName = GetInsideRatholeItemName(oACObject, oEdgePort1.SectionID, oEdgePort2.SectionID, oCut1BoundedPort, oCut1BoundingPort, False)
        If Not strRatholeName = vbNullString Then
            oSL.Add strRatholeName
        End If
    ElseIf (oEdgePort1.SectionID = JXSEC_TOP_FLANGE_RIGHT And oEdgePort2.SectionID = JXSEC_TOP) Or _
           (oEdgePort1.SectionID = JXSEC_TOP And oEdgePort2.SectionID = JXSEC_TOP_FLANGE_RIGHT) Or _
           (oEdgePort1.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT And oEdgePort2.SectionID = JXSEC_BOTTOM) Or _
           (oEdgePort1.SectionID = JXSEC_BOTTOM And oEdgePort2.SectionID = JXSEC_BOTTOM_FLANGE_RIGHT) Then
           
           ' Needs to be filled out
    End If
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
 
End Sub
'*********************************************************************************************
' Method      : MbrAxisEndCutCF_SelectorLogic
' Description : Selector items for Axis Cases
'
'*********************************************************************************************
Private Sub MbrAxisEndCutCF_SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim oCornerFeature As New StructDetailObjects.CornerFeature
    Set oCornerFeature.object = pSL.SmartOccurrence
        
    ' ------------------------------------
    ' Determine where the feature is going
    ' ------------------------------------
    Dim bIsOnBottomFlange As Boolean
    Dim eLocation As eEndCutCornerLocation
    
    GetCornerFeaturePositionOnEndCut pSL.SmartOccurrence, eLocation, bIsOnBottomFlange
    
    Dim isBottomEdge As Boolean
    If eLocation = BottomEdgeInside Or eLocation = BottomEdgeOutside Then
        isBottomEdge = True
    End If
    
    ' -------------
    ' Get the shape
    ' -------------
    Dim sBottomAnswerCol As String
    Dim sBottomShape As String
    Dim sTopAnswerCol As String
    Dim sTopShape As String
    Dim sFaceTopInsideCornerShape As String
    Dim sFaceTopInsideCornerCol As String
    Dim sFaceBtmInsideCornerShape As String
    Dim sFaceBtmInsideCornerCol As String
    
    Dim sShape As String
    Dim eAnswerCol As String
                
    GetMemberACTopAndBottomShape pSL.SmartOccurrence, sBottomAnswerCol, sBottomShape, sTopAnswerCol, sTopShape, _
                                 sFaceTopInsideCornerShape, sFaceTopInsideCornerCol, sFaceBtmInsideCornerShape, sFaceBtmInsideCornerCol
    
    Dim bIsToEdgeOnly As Boolean
    bIsToEdgeOnly = False
    If sBottomAnswerCol = vbNullString Or sTopAnswerCol = vbNullString Then
        bIsToEdgeOnly = True
    End If
    
    Select Case eLocation
    
        Case FaceBottomInsideCorner
            sShape = sFaceBtmInsideCornerShape
            eAnswerCol = sFaceBtmInsideCornerCol
        Case FaceTopInsideCorner
            sShape = sFaceTopInsideCornerShape
            eAnswerCol = sFaceTopInsideCornerCol
        Case Else
    
            If isBottomEdge Then
                If bIsToEdgeOnly Then
                    sShape = sTopShape
                    eAnswerCol = sTopAnswerCol
                Else
                    sShape = sBottomShape
                    eAnswerCol = sBottomAnswerCol
                End If
            Else
                If bIsToEdgeOnly Then
                    sShape = sBottomShape
                    eAnswerCol = sBottomAnswerCol
                Else
                    sShape = sTopShape
                    eAnswerCol = sTopAnswerCol
                End If
            End If
        
    End Select
    
            
    ' ---------------------------------------
    ' Based on where the feature is placed...
    ' ---------------------------------------
    Select Case eLocation
    
        ' --------------------------------------------------------------------------------------------------------
        ' If placed where the bounded web and flange both meet the face (they are trimmed flush, no "real" corner)
        ' --------------------------------------------------------------------------------------------------------
        Case WebAndFlangeCommon
            pSL.Add "Edge Face Rathole"
            
        ' ------------------------------------
        ' If placed between the web and flange
        ' ------------------------------------
        Case FaceBottomInsideCorner, FaceTopInsideCorner
             
            Select Case LCase(sShape)
        
        Case LCase(gsSnipe)
            pSL.Add "Member Snipe"
            pSL.Add "Member Snipe Along Corner"
        Case LCase(gsScallop)
            pSL.Add "Member Scallop"
            pSL.Add "Member Scallop Along Corner"
        Case LCase(gsFillet)
            pSL.Add "Elliptical Fillet"
        
        End Select

        ' ------------------------------------------
        ' If placed at edge inside or outside corner
        ' ------------------------------------------
        Case TopEdgeInside, BottomEdgeInside, TopEdgeOutside, BottomEdgeOutside

            Select Case eAnswerCol
                Case gsShapeAtEdgeOverlapCol
                    SelectCFforShapeAtEdgeOverlap pSL, sShape, isBottomEdge, bIsOnBottomFlange
                Case gsShapeAtEdgeOutsideCol, gsShapeAtEdgeCol
                    SelectCFforShapeAtEdge pSL, sShape, isBottomEdge, bIsOnBottomFlange
            End Select
    End Select
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "MbrAxisEndCutCF_SelectorLogic").Number
 
End Sub

'*********************************************************************************************
' Method      : SelectCFforShapeAtEdge
' Description :
'
'*********************************************************************************************
Private Sub SelectCFforShapeAtEdge(pSL As IJDSelectorLogic, _
                                   sShapeAtEdgeAnswer As String, _
                                   isBottomEdge As Boolean, _
                                   IsBottomFlange As Boolean)
    
    On Error GoTo ErrorHandler

    ' --------------------------------------
    ' See how much the web overlaps the edge
    ' --------------------------------------
    Dim sGrandParentItemName As String
    Dim oGrandParentObj As Object
    sGrandParentItemName = GetGrandParentName(pSL.SmartOccurrence, oGrandParentObj)
        
    Dim dInsideOverlap As Double
    Dim dOutsideOverlap As Double
    GetEdgeOverlapAndClearance oGrandParentObj, isBottomEdge, IsBottomFlange, dInsideOverlap, dOutsideOverlap
    
    ' ------------------------------------------------------------------------------------------------------------
    ' Select features based on the answer to the shape question and (in some cases) the overlap with bounding edge
    ' ------------------------------------------------------------------------------------------------------------
    ' The rathole features come in two versions.
    ' The first is intended for when the surfaces are the same, or the bounded surface is inside the bounding surface
    ' The second is intended for when the bounded surface is outside the bounding surface
    ' Because there is a 0.01mm margin, and because line lengths cannot be driven to zero or negative numbers, the
    ' threshold is an overlap of 0.02mm instead of 0.0mm (surfaces are even).
    Select Case LCase(sShapeAtEdgeAnswer)
        Case LCase(gsFaceToCorner), LCase(gsFaceToEdge)

            If dInsideOverlap >= 0# Then
                pSL.Add "Edge Hook Relief"
            Else
                pSL.Add "Face To Inside Rathole"
            End If
        
        Case LCase(gsFaceToFlange)
            pSL.Add "Face To Inside Rathole"
            
        Case LCase(gsInsideToEdge)
            If dInsideOverlap >= GetMinOverlapToEdge() Then
                pSL.Add "Arc Corner Relief"
            Else
                pSL.Add "Inside Face Rathole"
            End If
             
        Case LCase(gsInsideToFlange)
            If dInsideOverlap >= 0.00002 Then
                pSL.Add "Inside Face Rathole"
                pSL.Add "Edge Face Rathole"
            Else
                pSL.Add "Inside Face Rathole"
            End If
            
        Case LCase(gsCornerToFlange)
            pSL.Add "Short Edge Rathole"
            
        Case LCase(gsEdgeToFlange)
            pSL.Add "Short Edge Rathole"
            
        Case LCase(gsOutsideToEdge)
            If dOutsideOverlap > 0.005 Then
                pSL.Add "Arc Corner Relief"
            Else
                pSL.Add "Inside Face Rathole"
            End If
            
        Case LCase(gsOutsideToFlange)
            pSL.Add "Inside Face Rathole"
            
        Case Else
            '
    End Select
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectCFforShapeAtEdge").Number
    
End Sub

'*********************************************************************************************
' Method      : SelectCFforShapeAtEdgeOverlap
' Description :
'
'*********************************************************************************************
Private Sub SelectCFforShapeAtEdgeOverlap(pSL As IJDSelectorLogic, _
                                          sShapeAtEdgeAnswer As String, _
                                          bIsBottomEdge As Boolean, _
                                          bIsBottomFlange As Boolean)

    On Error GoTo ErrorHandler

    Select Case LCase(sShapeAtEdgeAnswer)
        Case LCase(gsFaceToInsideCorner), LCase(gsFaceToOutsideCorner)
            pSL.Add "Edge Hook Relief"
        Case LCase(gsFaceToOutside)
            pSL.Add "Outside Hook Relief"
        Case LCase(gsInsideToEdge), LCase(gsEdgeToOutside)
            pSL.Add "Arc Corner Relief"
        Case LCase(gsInsideToOutsideCorner), LCase(gsInsideCornerToOutside)
            pSL.Add "Corner To FarFace Edge Relief"
        Case LCase(gsInsideToOutside)
            pSL.Add "Inside And Outside Hook Relief"
        Case LCase(gsEdgeToFlange)
            pSL.Add "Short Edge Rathole"
        Case LCase(gsCornerToFlange)
            pSL.Add "Short Edge Rathole"
        Case LCase(gsFaceToEdge)
            pSL.Add "Edge Hook Relief"
        Case Else
            '
    End Select
 
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectCFforShapeAtEdgeOverlap").Number
       
End Sub
' ** Start CM **
'*********************************************************************************************
' If needed Add Custom Method HERE
'*********************************************************************************************
' ** End CM **

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************
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
    
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
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

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

Private Function GetInsideRatholeItemName(oACObject As Object, cornerPort1XID As JXSEC_CODE, _
                                          cornerPort2XID As JXSEC_CODE, _
                                          oECBoundedPort As IJPort, _
                                          oECBoundingPort As IJPort, _
                                          toFace As Boolean) As String

    On Error GoTo ErrorHandler
    
    ' ------------------------------------
    ' Determine position on bounded object
    ' ------------------------------------
    GetInsideRatholeItemName = vbNullString
    
    Dim isForBoundedBottom As Boolean
    
    ' If bounding object is a profile, the port reflects the position (top port for bounded top flange)
    If TypeOf oECBoundingPort.Connectable Is ISPSMemberPartPrismatic Then
        If (cornerPort1XID = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or cornerPort2XID = JXSEC_BOTTOM_FLANGE_RIGHT_TOP) Then
            isForBoundedBottom = True
        ElseIf (cornerPort1XID = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or cornerPort2XID = JXSEC_TOP_FLANGE_RIGHT_BOTTOM) Then
            isForBoundedBottom = False
        Else
            Exit Function
        End If
    ' If bounding object is a plate, the port reflect the opposite flange (top port for bounded bottom flange)
    ' Expected usage for corner feature at top of "FlushActual" web cut is when bottom flange intersects, and vice versa
    '
    '     ---------------------------------bounded top
    '     ---------------------------------
    '
    '
    '
    '             FlushActualTop
    '     ---------------------+
    '                          +----------------------
    '     ---------------------+----------------------bounded bottom
    '             FlushActualBtm
    '
    ElseIf TypeOf oECBoundingPort.Connectable Is IJPlate Then
        If (cornerPort1XID = JXSEC_TOP Or cornerPort2XID = JXSEC_TOP) Then
            isForBoundedBottom = True
        ElseIf (cornerPort1XID = JXSEC_BOTTOM Or cornerPort2XID = JXSEC_BOTTOM) Then
           isForBoundedBottom = False
        Else
            Exit Function
           
        End If
    End If
    
    ' ----------------------------------------------------
    ' Return an empty string is there is no bounded flange
    ' ----------------------------------------------------
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    CrossSection_Flanges oECBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
        
    If TypeOf oECBoundingPort.Connectable Is ISPSMemberPartPrismatic Then
        If ((Not isForBoundedBottom) And (Not bTFL) And (Not bTFR)) Or _
           (isForBoundedBottom And (Not bBFL) And (Not bBFR)) Then
            Exit Function
        End If
    End If
    
    ' -----------------------------------
    ' If the bounding object is a profile
    ' -----------------------------------
    ' add stiffener later
    If TypeOf oECBoundingPort.Connectable Is ISPSMemberPartPrismatic Then
        
        If toFace Then
            GetInsideRatholeItemName = "Face To Inside Rathole"
        Else
            GetInsideRatholeItemName = "Inside Face Rathole"
        End If

    ' ---------------------------------
    ' If the bounding object is a plate
    ' ---------------------------------
    ElseIf TypeOf oECBoundingPort.Connectable Is IJPlate Then
            GetInsideRatholeItemName = "Inside Face Rathole"
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CornerPlacement").Number
End Function

Private Sub MultiBoundingCF_Selection(oSL As IJDSelectorLogic, oCut As IJStructFeature)

 On Error GoTo ErrorHandler

    Dim oBoundedPort As IJPort
    Dim oMappedPortsColl As Collection
    Dim dAngleColl As Collection
    
    Dim oSDOWebCut As New StructDetailObjects.WebCut
    Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
    Dim eCutType As StructFeatureTypes
    
    If Not oCut Is Nothing Then
        eCutType = oCut.get_StructFeatureType
        If eCutType = SF_WebCut Then
            Set oSDOWebCut.object = oCut
            Set oBoundedPort = oSDOWebCut.BoundedPort
        Else
            Set oSDOFlangeCut.object = oCut
            Set oBoundedPort = oSDOFlangeCut.BoundedPort
        End If
    End If

    Dim oACObject As Object
    AssemblyConnection_SmartItemName oSL.SmartOccurrence, , oACObject
    
    'Both Cut1 and Cut 2 are same
    If (eCutType = SF_WebCut) Then
        GetMultiBoundingEdgeMap oACObject, WebCut, oMappedPortsColl, dAngleColl
    ElseIf (eCutType = SF_FlangeCut) Then
        If oSDOFlangeCut.IsTopFlange Then
            GetMultiBoundingEdgeMap oACObject, FlangeCutTop, oMappedPortsColl, dAngleColl
        Else
            GetMultiBoundingEdgeMap oACObject, FlangeCutBottom, oMappedPortsColl, dAngleColl
        End If
    End If

    ' Get the edge port section IDs
    Dim oEdgePort1 As IJStructPort
    Dim oEdgePort2 As IJStructPort
    
    Set oEdgePort1 = oSL.InputObject(INPUT_PORT2EDGE)
    Set oEdgePort2 = oSL.InputObject(INPUT_PORT3EDGE)
    
    If oMappedPortsColl.Count <= 1 Then Exit Sub
    
    Dim dAngle As Double
    dAngle = GetAngleBtwEdgeIDs(oEdgePort1.SectionID, oEdgePort2.SectionID, dAngleColl)
    
    Dim dAngle110 As Double
    Dim dAngle165 As Double
    
    dAngle110 = degreeToRadian(110)
    dAngle165 = degreeToRadian(165)
    
    ' When Bounding Surface Collection is two
    If oMappedPortsColl.Count = 2 Then
        If LessThanOrEqualTo(dAngle, dAngle110) Then  '110 Degrees
            oSL.Add "Member Snipe"
            oSL.Add "Member Radius Snipe"
        ElseIf GreaterThan(dAngle, dAngle110) And LessThanOrEqualTo(dAngle, dAngle165) Then   ' between 110 and 135 degrees
            oSL.Add "Member Radius Snipe"
            oSL.Add "Member Snipe"
        End If
    Else
        
        Dim oMemObjs As IJDMemberObjects
        Set oMemObjs = oCut  'oCut1 and oCut2 are same by this time
        
        Dim lCount As Long
        Dim iPhyConnCount As Integer
        Dim iFeatureCount As Integer
        Dim oMemberObject As Object
        
        iPhyConnCount = 0
        iFeatureCount = 0
        
        'To decide the Multibounding case and to select the appropriate Corner Feature
        'Get the PC and Corner Feature Members counts from the WebCut cut definition
        For lCount = 1 To oMemObjs.Count
            
            Set oMemberObject = oMemObjs.ItemByDispid(lCount)
            
            If Not oMemberObject Is Nothing Then
                If TypeOf oMemberObject Is IJStructPhysicalConnection Then
                    iPhyConnCount = iPhyConnCount + 1
                ElseIf TypeOf oMemberObject Is IJStructFeature Then
                    iFeatureCount = iFeatureCount + 1
                End If
            End If
        Next
        
        If (iPhyConnCount = 3 And iFeatureCount = 1) Or (iPhyConnCount = 5 And iFeatureCount = 2) Then
            If LessThanOrEqualTo(dAngle, dAngle110) Then    '110 Degrees
                oSL.Add "Member Snipe"
                oSL.Add "Member Radius Snipe"
            ElseIf GreaterThan(dAngle, dAngle110) And LessThanOrEqualTo(dAngle, dAngle165) Then   ' 135 degrees
                oSL.Add "Member Radius Snipe"
                oSL.Add "Member Snipe"
            Else ' For default case
                oSL.Add "Member Snipe"
            End If
            
        ElseIf (iPhyConnCount = 2 And iFeatureCount = 1) Or (iPhyConnCount = 3 And iFeatureCount = 2) Then
                oSL.Add "Member Step Snipe"
                oSL.Add "Face Rathole" '''' New Corner Feature should be added after creating
                                    
        ElseIf (iPhyConnCount = 2 And iFeatureCount = 2) Or (iPhyConnCount = 3 And iFeatureCount = 4) Then
            If LessThanOrEqualTo(dAngle, dAngle110) Then    '110 Degrees
                oSL.Add "Member Snipe"
                oSL.Add "Member Radius Snipe"
            ElseIf GreaterThan(dAngle, dAngle110) And LessThanOrEqualTo(dAngle, dAngle165) Then   ' 135 degrees
                oSL.Add "Member Radius Snipe"
                oSL.Add "Member Snipe"
            Else
                oSL.Add "Inside Face Rathole"
                oSL.Add "Member Step Snipe"
            End If
        End If
    End If

 Exit Sub
 
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "MultiBoundingCF_Selection").Number
    
End Sub
