VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtWeld4Parm"
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

Const m_ParameterRuleProgid As String = "PhysConnRules.ButtWeld4Parm"
Const m_ParameterRuleName As String = "PhysConnRules.ButtWeld4Parm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\ButtWeld4Parm.cls"

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
  
  pOH.SetOutput "RootGap"
  pOH.SetOutput "NRRootGap"
  pOH.SetOutput "Nose"                      'User code "N"
  pOH.SetOutput "NRNose"
  pOH.SetOutput "NoseOrientationAngle"      'User code "n"
  pOH.SetOutput "NRNoseOrientationAngle"
  pOH.SetOutput "ButtCategory"
  pOH.SetOutput "ReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "RefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "NRReferenceSide", imsARGUMENT_IS_BSTR
  pOH.SetOutput "NonRefPartName", imsARGUMENT_IS_BSTR
  pOH.SetOutput "ReferencePart"
 
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
    
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Set oPhysConn.object = pPRL.SmartOccurrence
    
    'get answer from selector
    Dim sCategory As String
    sCategory = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "Category")
    Dim sWorkcenter As String
    sWorkcenter = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "Workcenter")
    Dim sFlipBevel As String
    sFlipBevel = pPRL.SelectorAnswer("PhysConnRules.ButtWeldSel", "FlipBevel")
    
    'Get Thicknesses
    Dim dThickness As Double
    Dim dThickness2 As Double
    
    'Get the part names
    Dim sPartName1 As String
    Dim sPartName2 As String
    Dim oNamedItem1 As IJNamedItem
    Dim oNamedItem2 As IJNamedItem
    Set oNamedItem1 = oPhysConn.ConnectedObject1
    Set oNamedItem2 = oPhysConn.ConnectedObject2
    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
         sPartName1 = oNamedItem2.Name
         sPartName2 = oNamedItem1.Name
    Else
         sPartName1 = oNamedItem1.Name
         sPartName2 = oNamedItem2.Name
    End If
      
    If TypeOf oPhysConn.ConnectedObject1 Is IJPlate And _
       TypeOf oPhysConn.ConnectedObject2 Is IJPlate Then
       Dim oPlate1 As New StructDetailObjects.PlatePart
       Dim oPlate2 As New StructDetailObjects.PlatePart
       Set oPlate1.object = oPhysConn.ConnectedObject1
       Set oPlate2.object = oPhysConn.ConnectedObject2
          If oPlate1.OffsetToBaseFace(oPlate2) < oPlate2.OffsetToBaseFace And _
              oPlate1.OffsetToOffsetFace(oPlate2) < oPlate2.OffsetToOffsetFace Then
              dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace
          ElseIf oPlate1.OffsetToBaseFace(oPlate2) > oPlate2.OffsetToBaseFace And _
              oPlate1.OffsetToOffsetFace(oPlate2) > oPlate2.OffsetToOffsetFace Then
              dThickness = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2)
          Else
              If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                  dThickness = oPhysConn.Object2Thickness
                  dThickness2 = oPhysConn.Object1Thickness
              Else
                  dThickness = oPhysConn.Object1Thickness
                  dThickness2 = oPhysConn.Object2Thickness
              End If
          End If
    Else
      If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
          dThickness = oPhysConn.Object2Thickness
          dThickness2 = oPhysConn.Object1Thickness
      Else
          dThickness = oPhysConn.Object1Thickness
          dThickness2 = oPhysConn.Object2Thickness
      End If
    End If
    
    Select Case sWorkcenter
      Case gsMachine1
          pPRL.Add "RootGap", 0.0005
          pPRL.Add "NRRootGap", 0.0005
      Case gsMachine2
          pPRL.Add "RootGap", 0.0006
          pPRL.Add "NRRootGap", 0.0006
    End Select
   
    'Calculation of angles
    Dim dButtMountingAngle As Double
    Dim dBevelAngle As Double
    
    ' This used to use CornerButtMountingAngle, but this method causes some
    ' data loss.  Changing to use TeeMountingAngle (TR170710)
    'dButtMountingAngle = oPhysConn.CornerButtMountingAngle
    dButtMountingAngle = oPhysConn.TeeMountingAngle
    
    ' the code is expecting the angle to be less than 180 degrees (PI),
    ' but it could be greater than PI because of an issue in GetMountingAngle
    ' adjust it before using it  (TR170710) --
    ' **** this should be changed when DI 171304 is implemented ****
    If (dButtMountingAngle - TOL) >= PI Then
        dButtMountingAngle = (2 * PI) - dButtMountingAngle
    End If
  
    ' it is actually expecting it to be less than 90 degrees (PI/2).  I am not
    ' sure about this part of the check, but the bug above was making this value
    ' negative.
    If (dButtMountingAngle - TOL) >= 1.570796 Then
        dButtMountingAngle = PI - dButtMountingAngle
    End If
    
    If (dButtMountingAngle - TOL) >= -0.0001 And (dButtMountingAngle - TOL) <= 0.0001 Or _
        (dButtMountingAngle - TOL) >= (PI - 0.0001) And (dButtMountingAngle - TOL) <= (PI + 0.0001) Then 'no knuckle
        dBevelAngle = 0.392699
    Else 'knuckle
        dBevelAngle = 0.785398 + dButtMountingAngle
    End If
  
    'get the reference side
    Dim sRefSide As String
    sRefSide = GetRefSide(oPhysConn.ConnectedObject1)
    
    If sRefSide = "Base" Then
        pPRL.Add "ReferenceSide", "molded"
        pPRL.Add "NRReferenceSide", "molded"
    ElseIf sRefSide = "Offset" Then
        pPRL.Add "ReferenceSide", "antimolded"
        pPRL.Add "NRReferenceSide", "antimolded"
    Else
        pPRL.Add "ReferenceSide", sRefSide
        pPRL.Add "NRReferenceSide", sRefSide
    End If
                       
    If sCategory = gsOneSided Then
        pPRL.Add "ButtCategory", 65536
    Else
        pPRL.Add "ButtCategory", 65537
    End If


    dBevelAngle = 0.392699
    'Set nose and ref part values
    If sFlipBevel = "No" Then
          'set the reference part
          If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
              pPRL.Add "ReferencePart", 2
          Else
              pPRL.Add "ReferencePart", 1
          End If
          pPRL.Add "Nose", dThickness
          pPRL.Add "NoseOrientationAngle", dBevelAngle
          pPRL.Add "NRNose", dThickness2
          pPRL.Add "NRNoseOrientationAngle", 1.570796
          pPRL.Add "RefPartName", sPartName1
          pPRL.Add "NonRefPartName", sPartName2
    Else
          'set the reference part
          If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
              pPRL.Add "ReferencePart", 1
          Else
              pPRL.Add "ReferencePart", 2
          End If
          pPRL.Add "Nose", dThickness2
          pPRL.Add "NoseOrientationAngle", dBevelAngle
          pPRL.Add "NRNose", dThickness
          pPRL.Add "NRNoseOrientationAngle", 1.570796
          pPRL.Add "RefPartName", sPartName2
          pPRL.Add "NonRefPartName", sPartName1
    End If
  
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




