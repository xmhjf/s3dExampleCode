VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoubleSidedParm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "ChamferRules.DoubleSidedParm"
Const m_ParameterRuleName As String = CUSTOMERID + "ChamferRules.DoubleSidedParm"
Const m_FamilyProgid As String = ""
Const PI As Double = 3.141592654

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "ChamferRules\DoubleSidedParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  ' TODO - Add input description below
    
    ' Part to be chamfered
    ' Opposite Part - determines if a chamfer is needed and if so what size
    
    pIH.SetInput CHAMFER_PART
    pIH.SetInput OPPOSITE_PART
      
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  
    pOH.SetOutput "Depth1Base"
    pOH.SetOutput "Depth1Offset"
    pOH.SetOutput "AngleRadius1Base"
    pOH.SetOutput "AngleRadius1Offset"
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
    
    Dim lStatusCode As Long
    
    Dim dBase As Double
    Dim dOffset As Double
    
    Dim bBaseToBase As Boolean
    Dim bNormalConsistent As Boolean
    
    Dim dChamferBase As Double
    Dim dOppositeBase As Double
    Dim dChamferOffset As Double
    Dim dOppositeOffset As Double
    
    Dim dEdgeAngle As Double
    Dim dPlateThickness As Double
    Dim dChamferThickness As Double
    
    Dim oRootParent1 As Object
    Dim oRootParent2 As Object
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    Dim oConnectable1 As IJConnectable
    Dim oConnectable2 As IJConnectable
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oChamferPart As StructDetailObjects.PlatePart
    Dim oOppositePart As StructDetailObjects.PlatePart
    
    Dim bBoolean As Boolean
    Dim bChamferSlope As Double
    Dim dSlope As Double
    
    Dim oChamfer As StructDetailObjects.Chamfer
    Set oChamfer = New StructDetailObjects.Chamfer
    Set oChamfer.object = pPRL.SmartOccurrence
    
    dSlope = DegToRad(72)
    'Initialize Chamfer Part/Ports
    
    Set oPort1 = oChamfer.Chamfered
    Set oPort2 = oChamfer.DrivesChamfer
    
    ' Check for Plate Edge being "Chamfered" by Profile (Stiffener, Edge Reinforcement) Edge
    Set oConnectable1 = oPort1.Connectable
    Set oConnectable2 = oPort2.Connectable
    If Not TypeOf oConnectable1 Is IJPlatePart Then
    ElseIf Not TypeOf oConnectable2 Is IJPlatePart Then
        PlateEdgeByStiffenerEdgeChamferData oPort1, oPort2, dChamferBase, dChamferOffset
'''MsgBox "SMChamferRules.DoubleSidedParm::ParameterRuleLogic" & vbCrLf & _
'''       "dChamferBase   = " & Format(dChamferBase, "0.0000") & vbCrLf & _
'''       "dChamferOffset = " & Format(dChamferOffset, "0.0000")

        ' Know that a Chamfer Depth of 0.0 is not valid
        If dChamferBase < 0.001 Then
            dChamferBase = 0.001
        End If
        
        If dChamferOffset < 0.001 Then
            dChamferOffset = 0.001
        End If
        
        pPRL.Add "Depth1Base", dChamferBase
        pPRL.Add "AngleRadius1Base", dSlope
        
        pPRL.Add "Depth1Offset", dChamferOffset
        pPRL.Add "AngleRadius1Offset", dSlope
        Exit Sub
    End If
    
    Set oChamferPart = New StructDetailObjects.PlatePart
    Set oOppositePart = New StructDetailObjects.PlatePart
    
    Set oChamferPart.object = oChamfer.ChamferedPart
    Set oOppositePart.object = oChamfer.DrivesChamferPart
    
    Dim oChamferCondition As ChamferCondition
    
    If pPRL.SmartItem.Name = "DoubleSided" Then
' check if Chamfer Plate Parts are Leaf Plate Parts from the same Root
        Set oSDO_Helper = New StructDetailObjects.Helper
        Set oRootParent1 = oSDO_Helper.Object_RootParentSystem(oChamferPart.object)
        Set oRootParent2 = oSDO_Helper.Object_RootParentSystem(oOppositePart.object)
         
        If oRootParent1 Is oRootParent2 Then
            ' Chamfer Plate Parts are Leaf Plate Parts from the same Root
            bNormalConsistent = CompareNormalOfPlateParts(oChamferPart.object, _
                                                          oOppositePart.object)
    
            If bNormalConsistent Then
                dChamferBase = Abs(oChamferPart.OffsetToBaseFace)
                dChamferOffset = Abs(oChamferPart.OffsetToOffsetFace)
                
                dOppositeBase = Abs(oOppositePart.OffsetToBaseFace)
                dOppositeOffset = Abs(oOppositePart.OffsetToOffsetFace)
                
                'Determine the Base depth
                dBase = Abs(dChamferBase - dOppositeBase)
            
                'Determine the Offset depth
                dOffset = Abs(dChamferOffset - dOppositeOffset)
            
            Else
                dChamferBase = Abs(oChamferPart.OffsetToBaseFace)
                dChamferOffset = Abs(oChamferPart.OffsetToOffsetFace)
                
                dOppositeBase = Abs(oOppositePart.OffsetToBaseFace)
                dOppositeOffset = Abs(oOppositePart.OffsetToOffsetFace)
                
                'Determine the Base depth
                dBase = Abs(dChamferBase - dOppositeOffset)
            
                'Determine the Offset depth
                dOffset = Abs(dChamferOffset - dOppositeBase)
            End If
        
        Else
            'check if cylinder portion of CAN is involved. if yes, get the CanChamferSlope
            'for Cases when the OD of CAN and PrimaryMemberSystem are nearly equal, there will be
            'no cone. chamfer would be between TubePortion of CAN and attached designed member.
            oChamferCondition = GetChamferCondition(oChamferPart, oOppositePart, bChamferSlope)
            If oChamferCondition = CanCylinderToMember Then
                dSlope = Atn(bChamferSlope)
            Else
                dSlope = DegToRad(72)
            End If
            
            ' Chamfer Plate Parts are not from the same Root Plate System
            ' Calculate thickness differences using the Chamfer Sketch Plane
            ' (This might have an Impact on Performance)
            oChamfer.GetChamferBaseOffsetData dChamferBase, dChamferOffset, _
                                              bBaseToBase, dEdgeAngle, lStatusCode
            If lStatusCode <> 0 Then
                dBase = 0#
                dOffset = 0#
            Else
                ' The Double Sided Chamfer uses Vertical Constraints for Offsets
                dBase = Abs(dChamferBase)
                dOffset = Abs(dChamferOffset)
            End If
        End If

    Else
        'this is a tee chamfer
        'TR-169311
        GetSelectorAnswer pPRL, "ChamferThickness", dChamferThickness
        
        dPlateThickness = oChamferPart.PlateThickness
        dBase = (dPlateThickness - dChamferThickness) / 2
        dOffset = (dPlateThickness - dChamferThickness) / 2
    End If
    
    'Make sure Base is not 0.  If so, this parameter is being
    'updated before the symbol will be deleted due to assoc
    'sequence (or there is an error in the rules)
    If dBase < 0.0001 Then dBase = 0.0001
    
    'Set Depth1Base parameter
    pPRL.Add "Depth1Base", dBase
        
    'Make sure Offset is not 0.  If so, this parameter is being
    'updated before the symbol will be deleted due to assoc
    'sequence (or there is an error in the rules)
    If dOffset < 0.0001 Then dOffset = 0.0001
  
    'Set Depth1Offset parameter
    pPRL.Add "Depth1Offset", dOffset
    pPRL.Add "AngleRadius1Base", dSlope
    pPRL.Add "AngleRadius1Offset", dSlope
    
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
  pIH.definition = pPR
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

