VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SingleBaseParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'   ---------------
'   It computes the item paramaters in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Outputs defined by name the collection of parameters
'
Option Explicit

Const m_ParameterRuleProgid As String = "ChamferRules.SingleBaseParm"
Const m_ParameterRuleName As String = "ChamferRules.SingleBaseParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\ChamferRules\SingleBaseParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput CHAMFER_PART
    pIH.SetInput OPPOSITE_PART
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "Depth1Base"

  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
    
    Dim lStatusCode As Long
    
    Dim bBaseToBase As Boolean

    Dim dBase As Double
    Dim dChamferBase As Double
    Dim dOppositeBase As Double
    Dim dChamferOffset As Double
    Dim dOppositeOffset As Double
    Dim dChamferThickness As Double
    
    Dim dEdgeAngle As Double
    Dim dBaseDelta As Double
    Dim dOffsetDelta As Double
    
    Dim oRootParent1 As Object
    Dim oRootParent2 As Object
    
    Dim oChamfer As StructDetailObjects.Chamfer
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oChamferPart As StructDetailObjects.PlatePart
    Dim oOppositePart As StructDetailObjects.PlatePart
    
    'Initialize Chamfer Part/Ports
    Set oChamfer = New StructDetailObjects.Chamfer
    Set oChamferPart = New StructDetailObjects.PlatePart
    Set oOppositePart = New StructDetailObjects.PlatePart
    
    Set oChamfer.object = pPRL.SmartOccurrence
    Set oChamferPart.object = oChamfer.ChamferedPart
    Set oOppositePart.object = oChamfer.DrivesChamferPart
        
    If pPRL.SmartItem.Name = "SingleSidedBaseTeePC" Then
        'this is a tee base chamfer
        dChamferThickness = pPRL.SelectorAnswer("ChamferRules.RootTeeChamferSel", "ChamferThickness")
        dBase = oChamferPart.PlateThickness - dChamferThickness
    
    ElseIf pPRL.SmartItem.Name = "SingleSidedBasePC" Or _
           pPRL.SmartItem.Name = "SingleSidedBase" Then
           
        ' check if Chamfer Plate Parts are Leaf Plate Parts from the same Root
        Set oSDO_Helper = New StructDetailObjects.Helper
        Set oRootParent1 = oSDO_Helper.Object_RootParentSystem(oChamferPart.object)
        Set oRootParent2 = oSDO_Helper.Object_RootParentSystem(oOppositePart.object)
         
        If oRootParent1 Is oRootParent2 Then
            ' Chamfer Plate Parts are Leaf Plate Parts from the same Root
            'for "SingleSidedBasePC" cases: Determine the Base depth
            'for "SingleSidedBase" cases: Determine the Base depth
            
            'we know we want a chamfer on our base side, just need to get the value
            'get the offsets on ourself from molded to base and offset
            'get the offsets on the opposite part, using us as a reference
            'compute the deltas for base and offset, relative to our base and offset
            'if the chamfer is on the base side, our base delta should be >0 (expected)
            'compute the delta on the offset side with the opposite sign so that a chamfer on
            'our part would have a delta >0
            'use the larger value, which we would expect to be the base delta
            'if this value is < .0001, use .0001
            
            'get our offsets
            dChamferBase = oChamferPart.OffsetToBaseFace
            dChamferOffset = oChamferPart.OffsetToOffsetFace
            
            'get opposite offsets, relative to ours
            dOppositeBase = oOppositePart.OffsetToBaseFace(oChamferPart)
            dOppositeOffset = oOppositePart.OffsetToOffsetFace(oChamferPart)
            
            dBaseDelta = dOppositeBase - dChamferBase       'compute this forward so that >0 means chamfer is on us
            dOffsetDelta = dChamferOffset - dOppositeOffset 'compute this backward so that >0 means chamfer is on us
            
            If dBaseDelta > dOffsetDelta Then
                dBase = dBaseDelta
            Else
                dBase = dOffsetDelta
            End If
            
            If dBase < 0.0001 Then
                dBase = 0.0001
            End If
            
        Else
            ' Chamfer Plate Parts are not from the same Root Plate System
            ' Calculate thickness differences using the Chamfer Sketch Plane
            ' (This might have an Impact on Performance)
            oChamfer.GetChamferBaseOffsetData dBaseDelta, dOffsetDelta, _
                                              bBaseToBase, dEdgeAngle, lStatusCode
            If lStatusCode <> 0 Then
                dBase = 0#
            Else
                ' The Single Sided Chamfer uses "girth" Constraints for Offsets
                ' Need to adjust the Vertical distance by the Edge Angle
                dBase = Abs(dBaseDelta / dEdgeAngle)
            End If
        End If

    Else
        'for (UnKnown) cases: Determine the Offset depth
        dBase = 0.0001
    End If
    
    'Make sure Base is not 0.  If so, this parameter is being
    'updated before the symbol will be deleted due to assoc
    'sequence (or there is an error in the rules)
    If dBase < 0.0001 Then dBase = 0.0001
    
    'Set Depth1Base parameter
    pPRL.Add "Depth1Base", dBase
    
    
  Set oChamfer = Nothing
  Set oChamferPart = Nothing
  Set oOppositePart = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
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
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
    
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation
  
  Dim pDFact As New DefinitionFactory
  pDFact.InitAbstractParameterRule pPR
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

