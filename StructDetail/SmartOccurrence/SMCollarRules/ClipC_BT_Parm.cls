VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClipC_BT_Parm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Parameter rule for ClipC_BT (Connected Clip, Bended Top)

' Modify class name to particular rule
Private Const m_sClassName As String = "ClipC_BT_Parm"

' No need to modify following variables
Private Const m_sRuleProgID As String = m_sProjectName + "." + m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgID As String = ""
Private Const MODULE = m_sProjectPath + m_sClassName + ".cls"
'
Implements IJDUserSymbolServices

Private Sub ParameterRuleInputs(oIH As IJDInputsHelper)
   On Error GoTo ErrorHandler

   oIH.SetInput INPUT_PENETRATING
   oIH.SetInput INPUT_BOUNDINGPLATE
   oIH.SetInput INPUT_SLOT
  
   Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Private Sub ParameterRuleOutputs(oOH As IJDOutputsHelper)
   On Error GoTo ErrorHandler

   oOH.SetOutput PARAM_SIDE_OF_PART
   oOH.SetOutput PARAM_LAP_DISTANCE
   oOH.SetOutput PARAM_TOP_CLEARANCE
   
   Exit Sub
  
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Private Sub ParameterRuleLogic(oPRL As IJDParameterLogic)
   On Error GoTo ErrorHandler

   'Get slot object
   Dim oSlot As New StructDetailObjects.Slot
   
   Set oSlot.object = oPRL.InputObject(INPUT_SLOT)
    
   'Get Stiffener Side
   Dim sStiffenrSide As String
   Dim dTopClearance As Double
   Dim dBottomClearance As Double
   Dim dLapDistance As Double
   Dim dPlateThickness As Double
   Dim oHelper As StructDetailObjects.Helper

   Set oHelper = New StructDetailObjects.Helper
   
   Select Case oHelper.ObjectType(oSlot.Penetrated)
      Case SDOBJECT_PLATE
         Dim oPlateWrapper As New StructDetailObjects.PlatePart
            
         Set oPlateWrapper.object = oSlot.Penetrated
         sStiffenrSide = oPlateWrapper.StiffenerSide
         Set oPlateWrapper = Nothing
         
      Case Else
         sStiffenrSide = "Base"
   End Select
    Set oHelper = Nothing
    
   'Get Collar SideOfPart Answer
   Dim sCollarSide As String
   
   'TR-169311 Trying to retrieve the question's answer
   'without harcoding the path
   GetSelectorAnswer oPRL, gsCollarSideOfPart, sCollarSide
   
   ' Parameter rule:  Set Collar SideOfPart based on Answer
   ' StiffenerSide Base and Offset are internally set and not exposed to the user.
   ' IJDParameterLogic calls PenetrationSymbol, where:
   ' 0 = BASE
   ' 1 = OFFSET
   ' 2 = PLATE_CENTERED
   Dim nSideOfPart As Integer
   
   Select Case sCollarSide
      Case gsNoFlip
         If sStiffenrSide = "Base" Then
            nSideOfPart = 1
         ElseIf sStiffenrSide = "Offset" Then
            nSideOfPart = 0
         End If
         
      Case gsFlip
         If sStiffenrSide = "Base" Then
            nSideOfPart = 0
         ElseIf sStiffenrSide = "Offset" Then
            nSideOfPart = 1
         End If
         
      Case gsCentered
         nSideOfPart = 2
         
   End Select
   
   ' Calculate other parameters based on profile height
   Dim dProfilePartHeight As Double
   Dim sXSectionType As String
  If TypeOf oSlot.Penetrating Is IJProfile Then
    'Get Penetrating object
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oSlot.Penetrating
    dProfilePartHeight = oProfilePart.Height
    sXSectionType = oProfilePart.sectionType
    Set oProfilePart = Nothing
  ElseIf TypeOf oSlot.Penetrating Is IJPlate Then
    Dim oSlotMappingRule As IJSlotMappingRule
    Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
    
    Dim oWeb As Object
    Dim oFlange As Object
    Dim o2ndWeb As Object
    Dim o2ndFlange As Object
    oSlotMappingRule.GetSectionAlias oSlot.Penetrating, oSlot.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange

    oSlotMappingRule.GetSectionDepth oSlot.Penetrating, oSlot.Penetrated, dProfilePartHeight
    Set oSlotMappingRule = Nothing
  Else
    'Unsupported Penetrating Object
    Exit Sub
  End If
  
  Set oSlot = Nothing
  dProfilePartHeight = Round(dProfilePartHeight, 3)
     
   Select Case sXSectionType
      Case "EA", "UA"
         dLapDistance = 0.045 ' FlangeWidth + 0.065 - (FlangeWidth + 0.025)
         
         If dProfilePartHeight >= 0.1 And dProfilePartHeight <= 0.125 Then
            dTopClearance = 0.035
         ElseIf dProfilePartHeight > 0.125 And dProfilePartHeight <= 0.45 Then
            dTopClearance = 0.05
         Else

            dTopClearance = 0.035
         End If
         
      Case "B"
         dLapDistance = 0.045 ' 0.12 - 0.075
         
         If dProfilePartHeight >= 0.18 Then
            dTopClearance = 0.05
         Else

            dTopClearance = 0.05
         End If
      
      Case "FB"
         ' Ref left connected slot for slot width
         If dProfilePartHeight <= 0.125 Then
            dLapDistance = 0.04 ' (0.13 - 0.09)
            dTopClearance = 0.035
            
         ElseIf dProfilePartHeight > 0.125 And dProfilePartHeight <= 0.3 Then
            dLapDistance = 0.04 ' (0.13 - 0.09)
            dTopClearance = 0.05
                        
         ElseIf dProfilePartHeight > 0.3 Then
            dLapDistance = 0.005 ' (0.13 - 0.125) ?
            dTopClearance = 0.05
            
         Else
 
            dLapDistance = 0.04
            dTopClearance = 0.035
         End If
         
      Case "BUTL2", "BUT"
         dLapDistance = 0.045 ' FlangeWidth + 0.065 - (FlangeWidth + 0.025)
         
         If dProfilePartHeight <= 0.45 Then
            dTopClearance = 0.06
            
         ElseIf dProfilePartHeight > 0.45 And dProfilePartHeight <= 0.6 Then
            dTopClearance = 0.09
            
         ElseIf dProfilePartHeight > 0.6 And dProfilePartHeight <= 1.5 Then
            dTopClearance = 0.1
         
         Else

            dTopClearance = 0.06
         End If
               
      Case Else
         MsgBox m_sClassName & " should NOT be fired for " & sXSectionType & ", check selection rule"
         dLapDistance = 0.04
         dTopClearance = 0.035
         
   End Select
   
   oPRL.Add PARAM_SIDE_OF_PART, nSideOfPart
   oPRL.Add PARAM_LAP_DISTANCE, dLapDistance
   oPRL.Add PARAM_TOP_CLEARANCE, dTopClearance

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
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal vDefinitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_sRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oPR As IJDSymbolDefinition)
  On Error Resume Next
  
   ' Remove all existing defined Input and Output (Representations)
   ' before defining the current Inputs and Outputs
   oPR.IJDInputs.RemoveAllInput
   oPR.IJDRepresentations.RemoveAllRepresentation
  
   Dim oDFact As New DefinitionFactory
   Dim oIH As IJDInputsHelper
   
   oDFact.InitAbstractParameterRule oPR
   Set oIH = New InputHelper
   oIH.definition = oPR
   oIH.InitAs m_sFamilyProgID
   ParameterRuleInputs oIH
   
   '----------- Define Output and Add it to the representation ------------------
   Dim oOH As IJDOutputsHelper
   
   Set oOH = New OutputHelper
   oOH.Representation = oPR.IJDRepresentations.Item(1)
   oOH.InitAs m_sFamilyProgID
   ParameterRuleOutputs oOH
   
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal sCB As String, _
                                                             ByVal vDP As Variant, _
                                                             ByVal oRM As Object) As Object
   Dim oDFact As New DefinitionFactory
  
   Set IJDUserSymbolServices_InstanciateDefinition = oDFact.InstanciateParameterRule( _
                                                     m_sRuleProgID, _
                                                     sCB, _
                                                     IJDUserSymbolServices_GetDefinitionName(vDP), _
                                                     oRM)
   Set oDFact = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal oSymbolOccurrence As Object, _
                                                       ByVal oRepName As String, _
                                                       ByVal oOutputColl As Object, _
                                                       vArrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(oSymbolOccurrence As Object, _
                                                     ByVal oTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(oRep As IJDRepresentation)
   Dim oPRL As IJDParameterLogic
  
   Set oPRL = New ParameterLogic
   oPRL.Representation = oRep
   ParameterRuleLogic oPRL
   Set oPRL = Nothing
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

