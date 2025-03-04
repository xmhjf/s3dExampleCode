VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarOneSnipeAParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Parameter rule for CollarOneSnipeA(B)

' Modify class name to particular rule
Private Const m_sClassName As String = "CollarOneSnipeAParm"

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
   oOH.SetOutput PARAM_INNER_CORNER_RADIUS
   oOH.SetOutput PARAM_OUTER_CORNER_RADIUS
   
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
   Dim dLapDistance As Double
   Dim dInnerCornerRadius As Double
   Dim dOuterCornerRadius As Double
   Dim dPlateThickness As Double
   Dim oHelper As StructDetailObjects.Helper
   Dim oPlateWrapper As New StructDetailObjects.PlatePart
   Dim oProfileWrapper As New StructDetailObjects.ProfilePart

   Set oHelper = New StructDetailObjects.Helper
   
   Select Case oHelper.ObjectType(oSlot.Penetrated)
      Case SDOBJECT_PLATE
         Set oPlateWrapper.object = oSlot.Penetrated
         sStiffenrSide = oPlateWrapper.StiffenerSide
         dPlateThickness = oPlateWrapper.PlateThickness
         
      Case Else
         sStiffenrSide = "Base"
         Set oProfileWrapper.object = oSlot.Penetrated
         dPlateThickness = oProfileWrapper.WebThickness
   End Select
  
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

   Dim dWebBaseAngle As Double
   Dim dProfilePenetratedAngle As Double
   Dim dTopWebAngle As Double
   Dim sCollarItemName As String
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim eWebXid As JXSEC_CODE
   
   Set oCollarWrapper.object = oPRL.SmartOccurrence
   sCollarItemName = oCollarWrapper.ItemName
   If InStr(sCollarItemName, "_B") > 0 Then
      eWebXid = JXSEC_WEB_LEFT
   Else
      eWebXid = JXSEC_WEB_RIGHT
   End If
   
   GetSlotAngles oSlot.object, _
                 eWebXid, _
                 , _
                 dProfilePenetratedAngle, _
                 dTopWebAngle, _
                 nSideOfPart, _
                 False, _
                 True, _
                 True
                 
   ' Get clip thickness
   Dim dClipThickness As Double
   Dim vResultType As EResultType
   Dim sTableName As String
   Dim nCodeListValue As Long
   Dim sShortDesc As String
   Dim sLongDesc As String
   Dim dDelta As Double
   Dim dDistance As Double
   
   oHelper.GetCustomAttributeParamValue oPRL.SmartOccurrence, _
                                        "IJUAClipProps", _
                                        "ClipThickness", _
                                        dClipThickness, _
                                        vResultType, _
                                        sTableName, _
                                        nCodeListValue, _
                                        sShortDesc, _
                                        sLongDesc
                                        
   dLapDistance = 0.055
   dInnerCornerRadius = 0.025
   dOuterCornerRadius = 0.09
   
   ' Calculate other parameters based on profile height
   ' Calculate Inner/Outer Courner Radius based on angle between profile and stiffened plate
   Dim bCalculateInnerRadius As Boolean
   Dim dProfilePartWidth As Double
   Dim sXSectionType As String
  
  If TypeOf oSlot.Penetrating Is IJProfile Then
    'Get Penetrating object
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Set oProfilePart.object = oSlot.Penetrating
    dProfilePartWidth = oProfilePart.Width
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

    oSlotMappingRule.GetSectionWidth oSlot.Penetrating, oSlot.Penetrated, dProfilePartWidth
    Set oSlotMappingRule = Nothing
  Else
    'Unsupported Penetrating Object
    Exit Sub
  End If
  Set oSlot = Nothing
         
   bCalculateInnerRadius = True
   Select Case sXSectionType
      Case "EA", "UA"
         ' 0.08 + 0.055 - 0.0015 = 0.1335
         dDistance = 0.1335 - (dPlateThickness + dClipThickness) * Sin(dProfilePenetratedAngle) / 2
         
      Case "B"
         ' 0.125 + 0.055 - 0.0015 = 0.1785
         dDistance = 0.1785 - dProfilePartWidth - (dPlateThickness + dClipThickness) * Sin(dProfilePenetratedAngle) / 2
         
      Case "FB"
         ' 0.12 + 0.055 - 0.0015 = 0.126
         dDistance = 0.126 - dProfilePartWidth - (dPlateThickness + dClipThickness) * Sin(dProfilePenetratedAngle) / 2
         
      Case "BUTL2"
         ' 0.08 + 0.055 - 0.0015 = 0.1335
         dDistance = 0.1335 - (dPlateThickness + dClipThickness) * Sin(dProfilePenetratedAngle) / 2
         
      Case "BUT"
         ' 0.08 + 0.055 - 0.0015 = 0.1335
         dDistance = 0.1335 - (dPlateThickness + dClipThickness) * Sin(dProfilePenetratedAngle) / 2

      Case Else
         MsgBox m_sClassName & " should NOT be fired for " & sXSectionType & ", check selection rule"
         dLapDistance = 0.04
         bCalculateInnerRadius = False
         
   End Select
   Set oProfileWrapper = Nothing
   
   If bCalculateInnerRadius = True Then
      dDelta = dDistance / Sin(dTopWebAngle) - _
               dInnerCornerRadius / Tan(dTopWebAngle / 2) - _
               dOuterCornerRadius * Tan(dTopWebAngle / 2)
      If dDelta < 0.003 Then
         dInnerCornerRadius = dInnerCornerRadius - 0.003 + dDelta
      End If
   End If
   
   oPRL.Add PARAM_SIDE_OF_PART, nSideOfPart
   oPRL.Add PARAM_LAP_DISTANCE, dLapDistance
   oPRL.Add PARAM_INNER_CORNER_RADIUS, dInnerCornerRadius
   oPRL.Add PARAM_OUTER_CORNER_RADIUS, dOuterCornerRadius
   
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

