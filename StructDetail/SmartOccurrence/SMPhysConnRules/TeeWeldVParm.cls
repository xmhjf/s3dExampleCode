VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TeeWeldVParm"
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

Const m_ParameterRuleProgid As String = CUSTOMERID + "PhysConnRules.TeeWeldVParm"
Const m_ParameterRuleName As String = CUSTOMERID + "PhysConnRules.TeeWeldVParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "PhysConnRules\TeeWeldVParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
    
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler
  
  pOH.SetOutput "MoldedFillet"
  pOH.SetOutput "AntiMoldedFillet"
  pOH.SetOutput "FilletMeasureMethod"
  pOH.SetOutput "Category"
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "RefSideFirstBevelDepth"        'User code "A"
  pOH.SetOutput "RefSideFirstBevelMethod"
  pOH.SetOutput "RefSideFirstBevelAngle"        'User code "a"
  pOH.SetOutput "AntiRefSideFirstBevelDepth"    'User code "B"
  pOH.SetOutput "AntiRefSideFirstBevelMethod"
  pOH.SetOutput "AntiRefSideFirstBevelAngle"    'User code "b"


  ' call sub to declare the computed TEE outputs for IJWeldSymbol
  AddWeldParmRuleOutputs TEE_WELD_V, pOH
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
  On Error GoTo ErrorHandler
      
    '*****************************************************
    'For SM Implementation
    'This is the Delivered TeeWeld4, which is a full penetration bevel on one side
    'For Tsuneishi, the angle was modified to 45 degrees
    'The fillets are (0.2 * thickness)
    '*******************************************************
        
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = pPRL.SmartOccurrence
    
    'Get data required for Parameter Rule
    Dim sCategory As String
    Dim sBevelMethod As String
    Dim sClassSociety As String
    Dim dThickness1 As Double
    Dim dThickness2 As Double
     
    Get_ParameterRuleData pPRL, "TeeWeldV", _
                          sClassSociety, sCategory, sBevelMethod, _
                          dThickness1, dThickness2

    Dim sFirstWeldingSide As String
    GetSelectorAnswer pPRL, "FirstWeldingSide", sFirstWeldingSide
    
    'Get the part names
    Dim sPartName1 As String
    Dim oNamedItem1 As IJNamedItem
    Set oNamedItem1 = oPhysConn.ConnectedObject1
    sPartName1 = oNamedItem1.Name
    
    Dim dMountingAngle As Double
    dMountingAngle = oPhysConn.MountingAngle
    
    pPRL.Add "RefPartName", sPartName1
  
    'This is the angle definition that will be used in the calculations below
    Dim dAngle As Double
    If TypeOf oPhysConn.ConnectedObject1 Is IJCollarPart Then
        dAngle = 0.52359878 '30 degrees if the welded part is a collar
    Else
        dAngle = 0.785398 '45 degrees
    End If
    
    'set the reference side
    Dim sRefSide As String
    Dim sMoldedSide_Profile As String
    sRefSide = GetRefSide(oPhysConn.object)
    Dim bRefIsMolded As Boolean
    If sRefSide = "Base" Then
        pPRL.Add "ReferenceSide", "molded"
        bRefIsMolded = True
    ElseIf sRefSide = "Offset" Then
        pPRL.Add "ReferenceSide", "antimolded"
        bRefIsMolded = False
    Else
        sMoldedSide_Profile = GetMoldedSide(oPhysConn.object)
        pPRL.Add "ReferenceSide", sRefSide
        bRefIsMolded = True
    End If
    
    'This is the depth of the bevel (full thickness)
    ' We are not sure which side it is on yet
    Dim dBevelDepth As Double
    dBevelDepth = dThickness1
    
    ' variables to compute the values for the angle and depth on the two sides
    Dim dRefSideFirstBevelAngle As Double
    Dim dRefSideFirstBevelDepth As Double
    Dim dAntiRefSideFirstBevelAngle As Double
    Dim dAntiRefSideFirstBevelDepth As Double
    
    'the values are set, depending on the welding side
    If sFirstWeldingSide = "Molded" Then
       
        'calculate Angle, depending on the bevel method
        Select Case sBevelMethod
          Case gsConstant
            If sRefSide = "Base" Or sMoldedSide_Profile = "WebLeft" Or sMoldedSide_Profile = "TopFlangeTopFace" Or sMoldedSide_Profile = "BottomFlangeBottomFace" Then
              dRefSideFirstBevelAngle = PI - dAngle - oPhysConn.MountingAngle
              dRefSideFirstBevelDepth = dBevelDepth
              dAntiRefSideFirstBevelAngle = 0
              dAntiRefSideFirstBevelDepth = 0
            ElseIf sRefSide = "Offset" Or sMoldedSide_Profile = "WebRight" Or _
            sMoldedSide_Profile = "TopFlangeBottomFace" Or sMoldedSide_Profile = "BottomFlangeTopFace" Then
              dRefSideFirstBevelAngle = 0
              dRefSideFirstBevelDepth = 0
              dAntiRefSideFirstBevelAngle = PI - dAngle - oPhysConn.MountingAngle
              dAntiRefSideFirstBevelDepth = dBevelDepth
            Else
              dRefSideFirstBevelAngle = PI - dAngle - oPhysConn.MountingAngle
              dRefSideFirstBevelDepth = dBevelDepth
              dAntiRefSideFirstBevelAngle = 0
              dAntiRefSideFirstBevelDepth = 0
            End If
            
            pPRL.Add "RefSideFirstBevelMethod", 65536
            pPRL.Add "AntiRefSideFirstBevelMethod", 65536
            
          Case gsVarying
            If sRefSide = "Base" Or sMoldedSide_Profile = "WebLeft" Or sMoldedSide_Profile = "TopFlangeTopFace" Or sMoldedSide_Profile = "BottomFlangeBottomFace" Then
              dRefSideFirstBevelAngle = dAngle
              dRefSideFirstBevelDepth = dBevelDepth
              dAntiRefSideFirstBevelAngle = 0
              dAntiRefSideFirstBevelDepth = 0
            ElseIf sRefSide = "Offset" Or sMoldedSide_Profile = "WebRight" Or _
            sMoldedSide_Profile = "TopFlangeBottomFace" Or sMoldedSide_Profile = "BottomFlangeTopFace" Then
              dRefSideFirstBevelAngle = 0
              dRefSideFirstBevelDepth = 0
              dAntiRefSideFirstBevelAngle = dAngle
              dAntiRefSideFirstBevelDepth = dBevelDepth
            Else
              dRefSideFirstBevelAngle = dAngle
              dRefSideFirstBevelDepth = dBevelDepth
              dAntiRefSideFirstBevelAngle = 0
              dAntiRefSideFirstBevelDepth = 0
            End If
            
            pPRL.Add "RefSideFirstBevelMethod", 65537
            pPRL.Add "AntiRefSideFirstBevelMethod", 65537
            
       End Select
        
    Else 'the bevel is on the antimolded side
        
        'calculate Angle, depending on the bevel method
        Select Case sBevelMethod
          Case gsConstant
            If sRefSide = "Base" Or sMoldedSide_Profile = "WebLeft" Or sMoldedSide_Profile = "TopFlangeTopFace" Or sMoldedSide_Profile = "BottomFlangeBottomFace" Then
              dRefSideFirstBevelAngle = 0
              dRefSideFirstBevelDepth = 0
              dAntiRefSideFirstBevelAngle = PI - dAngle - oPhysConn.MountingAngle
              dAntiRefSideFirstBevelDepth = dBevelDepth
            ElseIf sRefSide = "Offset" Or sMoldedSide_Profile = "WebRight" Or _
            sMoldedSide_Profile = "TopFlangeBottomFace" Or sMoldedSide_Profile = "BottomFlangeTopFace" Then
              dRefSideFirstBevelAngle = PI - dAngle - oPhysConn.MountingAngle
              dRefSideFirstBevelDepth = dBevelDepth
              dAntiRefSideFirstBevelAngle = 0
              dAntiRefSideFirstBevelDepth = 0
            Else
              dRefSideFirstBevelAngle = 0
              dRefSideFirstBevelDepth = 0
              dAntiRefSideFirstBevelAngle = PI - dAngle - oPhysConn.MountingAngle
              dAntiRefSideFirstBevelDepth = dBevelDepth
            End If
            
            pPRL.Add "RefSideFirstBevelMethod", 65536
            pPRL.Add "AntiRefSideFirstBevelMethod", 65536
            
          Case gsVarying
            If sRefSide = "Base" Or sMoldedSide_Profile = "WebLeft" Or sMoldedSide_Profile = "TopFlangeTopFace" Or sMoldedSide_Profile = "BottomFlangeBottomFace" Then
              dRefSideFirstBevelAngle = 0
              dRefSideFirstBevelDepth = 0
              dAntiRefSideFirstBevelAngle = dAngle
              dAntiRefSideFirstBevelDepth = dBevelDepth
            ElseIf sRefSide = "Offset" Or sMoldedSide_Profile = "WebRight" Or _
            sMoldedSide_Profile = "TopFlangeBottomFace" Or sMoldedSide_Profile = "BottomFlangeTopFace" Then
              dRefSideFirstBevelAngle = dAngle
              dRefSideFirstBevelDepth = dBevelDepth
              dAntiRefSideFirstBevelAngle = 0
              dAntiRefSideFirstBevelDepth = 0
            Else
              dRefSideFirstBevelAngle = 0
              dRefSideFirstBevelDepth = 0
              dAntiRefSideFirstBevelAngle = dAngle
              dAntiRefSideFirstBevelDepth = dBevelDepth
            End If
            
            pPRL.Add "RefSideFirstBevelMethod", 65537
            pPRL.Add "AntiRefSideFirstBevelMethod", 65537
            
        End Select
    End If
  
    ' set the actual values in the outputs
    pPRL.Add "RefSideFirstBevelDepth", dRefSideFirstBevelDepth
    pPRL.Add "AntiRefSideFirstBevelDepth", dAntiRefSideFirstBevelDepth
    pPRL.Add "RefSideFirstBevelAngle", dRefSideFirstBevelAngle
    pPRL.Add "AntiRefSideFirstBevelAngle", dAntiRefSideFirstBevelAngle
    
    'Calculate the fillet size
    Dim dFillet As Double
    
    If dThickness1 <= dThickness2 Then
      dFillet = dThickness1 * 0.2
    Else
      dFillet = dThickness2 * 0.2
    End If
    
    pPRL.Add "MoldedFillet", dFillet
    pPRL.Add "AntiMoldedFillet", dFillet
    pPRL.Add "FilletMeasureMethod", 65536
     
    If sCategory = gsNormal Then
        pPRL.Add "Category", 65537
    ElseIf sCategory = gsDeep Then
        pPRL.Add "Category", 65538
    ElseIf sCategory = gsFull Then
        pPRL.Add "Category", 65539
    ElseIf sCategory = gsChain Then
        pPRL.Add "Category", 65540
    ElseIf sCategory = "Staggered" Then
        pPRL.Add "Category", 65541
    ElseIf sCategory = "OneSidedBevel" Then
        pPRL.Add "Category", 65542
    ElseIf sCategory = "TwoSidedBevel" Then
        pPRL.Add "Category", 65543
    Else
        pPRL.Add "Category", 65536
    End If
     
    
    ' fill in the correct values for IJWeldSymbol
    ' this method will include a check for any of the input parameters from the bevel
    ' that have been overridden by the user
    SetCalculatedTeeWeldParams pPRL, TEE_WELD_V, bRefIsMolded, _
                               dRefSideFirstBevelDepth, _
                               dAntiRefSideFirstBevelDepth, _
                               dRefSideFirstBevelAngle, _
                               dAntiRefSideFirstBevelAngle, _
                               pcr_WG_Bevel, _
                               dFillet, dFillet, 0#, 0#, 0#

    Set oPhysConn = Nothing

  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleLogic").Number
End Sub


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





