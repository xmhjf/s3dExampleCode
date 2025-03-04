VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootCornerSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : RootCornerSel.cls
'
'  Description :
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature'
'
'  Author      : Alligators
'
'   12 Dec 2011 - pnalugol - enhancements to fix Corner feature on slot issues 205720
'   3-May-2013 - Alligators -TR-211987
'           Made changes in SelectorLogic to handle std. member as support.
'   30-01-2014  - GH      - TR-CP-228731 Made changes in SelectorLogic
'                           Added CF Face Port as optional Argument for IfSeamsExistWithInDistance() method
'*********************************************************************************************


Const m_SelectorProgid As String = CUSTOMERID + "CornerFeatRules.RootCornerSel"
Const m_SelectorName As String = CUSTOMERID + "CornerFeatRules.RootCornerSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\RootCornerSel.cls"

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
    pQH.SetQuestion gsPlacement, "Manual", "PlacementCol", "SetPlacement", CUSTOMERID + "CornerFeatRules.CornerSelCM" ' NOTE: This default is only used if not created by Assembly connection
    pQH.SetQuestion gsCrackArrest, "Yes", "BooleanCol", "SetCrackArrest", CUSTOMERID + "CornerFeatRules.CornerSelCM"
    pQH.SetQuestion gsDrainage, "Yes", "BooleanCol", "SetDrainage", CUSTOMERID + "CornerFeatRules.CornerSelCM"
    pQH.SetQuestion gsCornerFlip, "NoFlip", "FlipCol", "SetFlip", CUSTOMERID + "CornerFeatRules.CornerSelCM"     'Will use answer to this question in parameter rule, not selector rule
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol", "SetApplyTreatment", CUSTOMERID + "CornerFeatRules.CornerSelCM"
    pQH.SetQuestion gsSeamAdjustment, "Yes", "BooleanCol", "SetEnforceSeamAdjustment", CUSTOMERID + "CornerFeatRules.CornerSelCM"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
 
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description :
'
'*********************************************************************************************
Public Sub SelectorLogic(pSL As IJDSelectorLogic)
     
    On Error GoTo ErrorHandler

    ' Get corner feature
    Dim oCornerFeature As IJSDOCornerFeature
    Set oCornerFeature = New StructDetailObjectsex.CornerFeature
    Set oCornerFeature.object = pSL.SmartOccurrence
    
    'Get answers
    Dim strPlacement As String
    strPlacement = pSL.Answer(gsPlacement)
    
    Dim strDrainage As String
    strDrainage = pSL.Answer(gsDrainage)
    
    Dim strCrackArrest As String
    strCrackArrest = pSL.Answer(gsCrackArrest)
    
    Dim strCornerFlip As String
    strCornerFlip = pSL.Answer(gsCornerFlip)
    
    
    Dim strSeamAdj As String
    strSeamAdj = pSL.Answer(gsSeamAdjustment)
    
    
    'Get corner feature part object
    Dim oPart As Object
    Set oPart = oCornerFeature.GetPartObject
    
    'Check if the feature is on Slot
    Dim bFeatureOnSlot As Boolean
    bFeatureOnSlot = False
    
    Dim oCFChild As IJDesignChild
    Dim oParentSlot As Object

    Set oCFChild = oCornerFeature.object
    Set oParentSlot = oCFChild.GetParent

    If TypeOf oParentSlot Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oParentSlot

        If oFeature.get_StructFeatureType = SF_Slot Then
            bFeatureOnSlot = True
        End If
    End If
    
    If Not bFeatureOnSlot Then  'since the below 2 wrappers fail for CF on slot; a seperate TR is filed on stingrays to fix them
    'until then these will not be called for CF on slots
        'Get the location of CF
        Dim oLocation As IJDPosition
        oCornerFeature.GetLocationOfCornerFeature oLocation
        
        Dim bDirecValid As Boolean
        
        Dim oUAxis As IJDVector
        Dim oVAxis As IJDVector
        Dim dUEdgeLength As Double
        Dim dVEdgeLength As Double
        
        Set oUAxis = New dVector
        Set oVAxis = New dVector
        
        'Get UAxis and VAxis of CF
        bDirecValid = oCornerFeature.GetUnFlippedCornerFeatureData(oUAxis, oVAxis, dUEdgeLength, dVEdgeLength)
    End If
    
    Dim bSeamItemExists As Boolean
            
    Dim oEdgePort1 As IJPort, oEdgePort2 As IJPort, oOtherPort1 As IJPort, oOtherPort2 As IJPort
    Dim oOtherConnectable1 As IJConnectable, oOtherConnectable2 As IJConnectable
    
    Dim oSupportObject1 As Object, oSupportObject2 As Object, oRootParent1 As Object, oRootParent2 As Object
    
    Dim oConnPort1 As IJPort, oConnPort2 As IJPort
    
    Dim oSDFUtils As IJSDFeatureAttributes
    Set oSDFUtils = New SDFeatureUtils
    Dim oFP1 As Object, oEP1 As Object, oEP2 As Object
    
    oSDFUtils.get_CornerCutInputsEx pSL.SmartOccurrence, oFP1, oEP1, oEP2
    
'    Set oEdgePort1 = pSL.InputObject(INPUT_PORT2EDGE)
'    Set oEdgePort2 = pSL.InputObject(INPUT_PORT3EDGE)
    
    If TypeOf oEP1 Is IJPort Then
        Set oConnPort1 = GetACPortByBruteForce(oEP1)
        
        If oConnPort1 Is Nothing Then
            Set oConnPort1 = oEdgePort1
        End If
    End If
    
    If TypeOf oEP2 Is IJPort Then
        Set oConnPort2 = GetACPortByBruteForce(oEP2)
        
        If oConnPort2 Is Nothing Then
           Set oConnPort2 = oEdgePort2
        End If
    End If
    
    'determine if the feature is on slot, if yes, determine if the feature is placed at base corner,
    'inside corner, outside corner. if its a base corner, let it allow follow the normal execution
    'to allow it to adjust according to seams.
    
    Dim bIsBaseCF As Boolean, bIsInsideCF As Boolean, bOuterCF As Boolean
    bIsBaseCF = False: bIsInsideCF = False: bOuterCF = False
    
    If bFeatureOnSlot Then
        DetermineCFPositionOnSlot oEP1, oEP2, bIsBaseCF, bIsInsideCF, bOuterCF
        If bIsInsideCF Then
            pSL.Add "SmartSnipe"
            pSL.Add "SmartScallop"
            Exit Sub
        ElseIf bOuterCF Then
            If True = CheckPartClassExist("VariableEdgeCorner") Then
               pSL.Add "VariableEdgeCorner"
               Exit Sub
            End If
        ElseIf bIsBaseCF Then
            GoTo DefaultExecution
            Exit Sub
        End If
    End If
    
    bSeamItemExists = CheckPartClassExist("SeamScallop")
    If Not bSeamItemExists Then
        GoTo DefaultExecution
    End If
    
    'Get the other connectables of the part on which CF is applied
    If Not oConnPort1 Is Nothing Then
        GetOtherPortAndConnectable oConnPort1, oOtherPort1, oOtherConnectable1
    End If
    If Not oConnPort2 Is Nothing Then
        GetOtherPortAndConnectable oConnPort2, oOtherPort2, oOtherConnectable2
    End If
    
    'get the seams on object and see if something is in tolerance
    'if yes, set the variable values
    
    Dim bSeamFound As Boolean, bSeamFoundU As Boolean, bSeamFoundV As Boolean
    Dim dSeamDistance_U As Double, dSeamDistance_V As Double, dSeamDistance As Double
    Dim bSeamValid As Boolean, bSeamValid1 As Boolean, bSeamValid2 As Boolean
    
    bSeamFound = False
    bSeamValid = False
    Dim dULength As Double
    Dim dVLength As Double
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurence As IJSmartOccurrence
    Set oSmartOccurence = pSL.SmartOccurrence
    
    Dim oSymbDef As IJDSymbolDefinition
    
    If Not oSmartOccurence.Item = "" Then
        Set oSmartItem = oSmartOccurence.SmartItemObject
        Set oSymbDef = oSmartItem.SymbolDefinitionDef
    End If
    
    
    Dim sCFName As String
    sCFName = ""
    
    If Not oSmartItem Is Nothing Then
        sCFName = oSmartItem.Name
    End If
    
    'Get ULength and Vlength of default CF which is used in search distance
    GetCFlengths sCFName, oCornerFeature, strCrackArrest, strDrainage, dULength, dVLength
    
    'parameters for Corner feature with seam
    Dim nSeamFound As Integer
    Dim dDistanceToSeam As Double
    Dim nSeamFlip As Integer    'flip specific to CF with seam
    
    
    'if already a seam item exists, search distance for a seam is made from the Seam item ULength and Vlength
    If StrComp(sCFName, "LongScallopWithSeam", vbTextCompare) = 0 Or StrComp(sCFName, "LongScallopWithSeamWithCollar", vbTextCompare) = 0 Then
        GetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "SeamFound", nSeamFound
        GetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "DistToSeam", dDistanceToSeam
        GetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "sFlip", nSeamFlip
    End If
    
    If nSeamFound = 1 Then
        dULength = dDistanceToSeam + SEAM_SEARCHDISTANCE
        dVLength = 0.1
    End If
    
    Dim gsFlip As Variant
    GetSelectorAnswer pSL, gsCornerFlip, gsFlip
    
    Dim dTemp As Double
    
    If Not sCFName = "" Then
        If StrComp(sCFName, "LongScallopWithSeam", vbTextCompare) = 0 Or StrComp(sCFName, "LongScallopWithSeamWithCollar", vbTextCompare) = 0 Then
            If nSeamFlip = 1 Then
                dTemp = dULength
                dULength = dVLength
                dVLength = dTemp
            End If
        Else
            If StrComp(gsFlip, "Flip", vbTextCompare) = 0 Then
                dTemp = dULength
                dULength = dVLength
                dVLength = dTemp
            End If
        End If
    End If
    
    Dim oStructDetailHelper As New StructDetailObjects.Helper
    
    'Look for a valid seam from default CF Ulength and VLength. Search distance is 15mm in and out of the CF edge.
    If Not oOtherConnectable1 Is Nothing And oOtherConnectable2 Is Nothing Then
        Set oSupportObject1 = oOtherConnectable1
        Set oRootParent1 = oStructDetailHelper.Object_RootParentSystem(oSupportObject1)
        
        bSeamFound = IfSeamsExistWithInDistance(oLocation, oRootParent1, dULength, dSeamDistance, oUAxis, True, bSeamValid, SEAM_SEARCHDISTANCE, oFP1)
        
        If bSeamValid Then SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "sFlip", 0
            
    ElseIf Not oOtherConnectable2 Is Nothing And oOtherConnectable1 Is Nothing Then
        Set oSupportObject2 = oOtherConnectable2
        Set oRootParent2 = oStructDetailHelper.Object_RootParentSystem(oSupportObject2)
        
        bSeamFound = IfSeamsExistWithInDistance(oLocation, oRootParent2, dVLength, dSeamDistance, oVAxis, True, bSeamValid, SEAM_SEARCHDISTANCE, oFP1)
        
        If bSeamValid Then SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "sFlip", 1
    
    ElseIf Not oOtherConnectable1 Is Nothing And Not oOtherConnectable2 Is Nothing Then
        Set oSupportObject1 = oOtherConnectable1
        Set oSupportObject2 = oOtherConnectable2
        
        Dim oMemberPart As ISPSMemberPartCommon
            
        If TypeOf oSupportObject1 Is ISPSMemberPartCommon Then
            Set oMemberPart = oSupportObject1
            Set oRootParent1 = oMemberPart.MemberSystem
        Else
            Set oRootParent1 = oStructDetailHelper.Object_RootParentSystem(oSupportObject1)
        End If
        
        If TypeOf oSupportObject2 Is ISPSMemberPartCommon Then
            Set oMemberPart = oSupportObject2
            Set oRootParent2 = oMemberPart.MemberSystem
        Else
            Set oRootParent2 = oStructDetailHelper.Object_RootParentSystem(oSupportObject2)
        End If
        
        bSeamFoundU = IfSeamsExistWithInDistance(oLocation, oRootParent1, dULength, dSeamDistance_U, oUAxis, True, bSeamValid1, SEAM_SEARCHDISTANCE, oFP1)
        bSeamFoundV = IfSeamsExistWithInDistance(oLocation, oRootParent2, dVLength, dSeamDistance_V, oVAxis, True, bSeamValid2, SEAM_SEARCHDISTANCE, oFP1)
       
        If bSeamValid1 Then
            SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "sFlip", 0
            bSeamValid = True
            dSeamDistance = dSeamDistance_U
        ElseIf bSeamValid2 = True Then
            bSeamValid = True
            SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "sFlip", 1
            dSeamDistance = dSeamDistance_V
        Else
            bSeamValid = False
            SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "sFlip", 0
        End If
    End If
    
    
    'If the user enforces to look for a seam, only seam item is populated.
    If bSeamValid = True And strSeamAdj = "Yes" Then
        SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "SeamFound", 1
        SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "DistToSeam", dSeamDistance
        pSL.Add "SeamScallop"
        Exit Sub
    End If
    
    'if the user doesnt want to over write the selection even if seam is presnt, he can do so by overwriting the answer
    'for EnforceSeamAdjustment question
    If bSeamValid = True And strSeamAdj = "No" Then
        pSL.Add "SeamScallop"
        pSL.Add "SmartGeneralCorner"
        pSL.Add "SmartScallop"
        pSL.Add "SmartSnipe"
        Exit Sub
    Else
        If nSeamFound = 0 And bSeamValid = False Then
            SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "SeamFound", 0
            SetAttributeOnInterface pSL.SmartOccurrence, "IJUACornerFeatureSeam", "DistToSeam", 0
        End If
        GoTo DefaultExecution
        Exit Sub
    End If
    
DefaultExecution:
    
    'Selection code
    If TypeOf oPart Is IJPlate Then
        Dim oPlate As New StructDetailObjects.PlatePart
        Set oPlate.object = oPart
        If IsBracket(oPart) Then
        
            pSL.Add "SmartScallop"
            pSL.Add "SmartGeneralCorner"
            pSL.Add "SmartSnipe"
        ElseIf bFeatureOnSlot Then
            Select Case oPlate.Tightness
                Case NonTight
                    pSL.Add "SmartScallop"
                    pSL.Add "SmartSnipe"
                Case Else
                    pSL.Add "SmartSnipe"
                    pSL.Add "SmartScallop"
            End Select
        Else
            Select Case oPlate.Tightness
                Case NonTight
                    Select Case strCrackArrest
                        Case "Yes"
                            pSL.Add "SmartGeneralCorner"
                            pSL.Add "SmartScallop"
                            pSL.Add "SmartSnipe"
                        Case "No"
                            Select Case strDrainage
                                Case "Yes"
                                    pSL.Add "SmartScallop"
                                    pSL.Add "SmartGeneralCorner"
                                    pSL.Add "SmartSnipe"
                                Case "No"
                                    Call CornerPlacement(pSL)
                            End Select
                    End Select
                Case Else   'Plate is Watertight, airtight, or oiltight
                    Call CornerPlacement(pSL)
            End Select
        End If
        
    ElseIf TypeOf oPart Is IJProfile Or TypeOf oPart Is ISPSMemberPartPrismatic Then  'Part is a profile or member
                'see if its case of mutual bound stiffeners.
    Dim oChild As IJDesignChild
    Set oChild = oCornerFeature.object
    Dim oParent As Object
    Set oParent = oChild.GetParent
    If TypeOf oParent Is IJStructFeature Then
        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oParent
        If oStructFeature.get_StructFeatureType = SF_WebCut Then
            Dim oAC As New StructDetailObjects.AssemblyConn
            Dim oFeatureChild As IJDesignChild
            Set oFeatureChild = oStructFeature
            Dim bMutualBound As Boolean
            
            Set oAC.object = oFeatureChild.GetParent
            bMutualBound = IsMutualBound(oAC)
            If bMutualBound Then
                pSL.Add "SmartScallop"
                Exit Sub
            End If
         End If
    End If
            
        Select Case strDrainage
            Case "Yes"
                
                '**********************************************
                'new code for checking the corner gap
                If oCornerFeature.CornerGapExists(0.001) Then
                    pSL.Add "SmartGeneralCorner"
                Else 'no gap, take the defaults
                    pSL.Add "SmartScallop"
                    pSL.Add "SmartSnipe"
                    pSL.Add "SmartGeneralCorner"
                End If
                
                '************************************************
                
                
            Case "No"
                Call CornerPlacement(pSL)
        End Select
    End If
    
    If True = CheckPartClassExist("VariableEdgeCorner") Then
       pSL.Add "VariableEdgeCorner"
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
End Sub


'  ------------------------------------------------------------
' This selection rule selects a corner feature based on whether it has been
' placed manually or by an assembly connection

'*********************************************************************************************
' Method      : CornerPlacement
' Description :
'
'*********************************************************************************************
Private Sub CornerPlacement(pSL As IJDSelectorLogic)
On Error GoTo ErrorHandler

' Get corner feature
Dim oCornerFeature As IJSDOCornerFeature
Set oCornerFeature = New StructDetailObjectsex.CornerFeature
Set oCornerFeature.object = pSL.SmartOccurrence

'Get answers
Dim strPlacement As String
strPlacement = pSL.Answer(gsPlacement)

'Get corner part object
On Error Resume Next
Dim oPart As IJPlate
Set oPart.object = oCornerFeature.GetPartObject
    
Select Case strPlacement
    Case "Manual"   'Always have all corner feature items available for manual placement
        'After bug on using parms in multiple selectors is fixed (03001000?), a separate
        'Manual class will be created, containing a complete list of items (such as
        '"ManualScallop") pointing to the same symbols files and same parms as the
        'Smart items.
        pSL.Add "SmartScallop"
        pSL.Add "SmartGeneralCorner"
        pSL.Add "SmartSnipe"
    Case "AssyConn"
        'No Corner Feature
End Select

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CornerPlacement").Number
 
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

Private Function GetACPortByBruteForce(oPort As IJPort) As IJPort
    
    Set GetACPortByBruteForce = Nothing
    
    Dim oStructPort As IJStructPort
    Set oStructPort = oPort
    
    Dim nConnections As Long
    Dim connData() As ConnectionData
    
    Dim oHelper As New StructDetailObjects.Helper
    
    oHelper.Object_AppConnections oPort.Connectable, AppConnectionType_Assembly, nConnections, connData
    
    'nConnections = GetAllConnectables(oPort.Connectable, AppConnectionType_Assembly, connData)

    Dim i As Long
    For i = 1 To nConnections
    
        Dim oConnStructPort As IJStructPort
        Set oConnStructPort = connData(i).ConnectingPort
    
        If TypeOf oPort.Connectable Is IJProfile Then
            Dim portPrimaryContext As Long
            Dim connPortPrimaryContext As Long
        
            portPrimaryContext = oStructPort.ContextID And (CTX_BASE Or CTX_OFFSET Or CTX_LATERAL)
            connPortPrimaryContext = oConnStructPort.ContextID And (CTX_BASE Or CTX_OFFSET Or CTX_LATERAL)
        
            If (portPrimaryContext = connPortPrimaryContext) Then
                Set GetACPortByBruteForce = oConnStructPort
                Exit Function
            End If
        ElseIf oConnStructPort.ContextID = oStructPort.ContextID And _
               oConnStructPort.operationID = oStructPort.operationID And _
               oConnStructPort.operatorID = oStructPort.operatorID And _
               oConnStructPort.SectionID = oStructPort.SectionID Then
               
               Set GetACPortByBruteForce = oConnStructPort
               Exit Function
        End If
    Next i
End Function


