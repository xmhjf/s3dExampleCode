VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bkt1SParm"
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

Const m_ParameterRuleProgid As String = "BktSelRules.Bkt1SParm"
Const m_ParameterRuleName As String = "BktSelRules.Bkt1SParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "\SharedContent\Src\StructDetail\SmartOccurrence\BktSelRules\Bkt1SParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler

    pIH.SetInput INPUT_BRACKETPLANE
    
    pIH.SetInput INPUT_SUPPORTS
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
  On Error GoTo ErrorHandler

  pOH.SetOutput "Height1"
  pOH.SetOutput "Width1"
  pOH.SetOutput "Nose1"
  pOH.SetOutput "Nose2"
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub


' Description:ne
'  Get the profile cross-section as the first support and determine its cross-section
'  IF it is a type of TEE (eg BUTT, BUTL2, T-Type, etc) set bracket width =  half flange legth minus 0.5* web thickness - 20mm for weld
'  otherwise , set bracket width = flange legth minus 0.5* web thickness - 20mm for weld
'  Use bracket height = 75% of the web length - flange thickness.
'  Keep the nose values = to the catalog value

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
     
    Dim dNose1 As Double, dNose2 As Double
    dNose1 = 20  'mm
    dNose2 = 20  'mm
    Dim oAttrHelper As IJDAttributes
    Dim oAttr As IJDAttribute
    Dim oAttributeCol As IMSAttributes.IJDAttributesCol
    Dim oBracketByPlaneSO As IJSmartOccurrence
    
    Set oBracketByPlaneSO = pPRL.SmartOccurrence
    If (oBracketByPlaneSO Is Nothing) Then Exit Sub
    
    ' The attributes Height1,Width1,Nose1 and Radius1 are availabe from
    ' the interface IJUABracketProps.Get them from catalog and set on the SmartItem
    Set oAttrHelper = oBracketByPlaneSO
 
    Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracket2Props")
    Set oAttr = oAttributeCol.Item("Nose2")
    If Not (oAttr Is Nothing) Then
        If (oAttr.Value > 0) Then dNose2 = oAttr.Value
        pPRL.Add "Nose2", oAttr.Value
        Set oAttr = Nothing
    End If
     Set oAttributeCol = oAttrHelper.CollectionOfAttributes("IJUABracketProps")
    Set oAttr = oAttributeCol.Item("Nose1")
    If Not (oAttr Is Nothing) Then
        If (oAttr.Value > 0) Then dNose1 = oAttr.Value
        pPRL.Add "Nose1", oAttr.Value
        Set oAttr = Nothing
    End If
    
    ' try to use teh profiel cross-section to determine the 1S bracket height and wirth.
    ' if there is a problem, use teh default catalog value
    On Error Resume Next
    
    Dim flangewidth As Double, webheight As Double
    flangewidth = -0.1
    webheight = -0.1
    Dim oSupports As IJElements
    Dim nSupports As Long
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupports, nSupports
    Dim oProfileObject As StructDetailObjects.ProfileSystem
    Set oProfileObject = New StructDetailObjects.ProfileSystem
    Dim oDCrossSection As IJDProfileSection
    
    
    If (oSupports Is Nothing) Then GoTo skip
    If (oSupports.Count < 1) Then GoTo skip
    If TypeOf oSupports(1) Is IJProfile Then
    
        With oProfileObject
            ' use the affected leaf system as the cross section can change
            Set .object = oSupports(1)
            Set oDCrossSection = .object
            'if the profile has a flange, it can be used for a "1S"
            If Not (oDCrossSection Is Nothing) Then
                If (oDCrossSection.SectionType And Tee) > 0 Then
                    flangewidth = .FlangeLength * 0.5
                Else
                    flangewidth = .FlangeLength
                End If
            End If
            flangewidth = flangewidth - (0.5 * .WebThickness) - dNose1
            webheight = (0.75 * .WebLength) - .FlangeThickness - dNose2
        End With
  Else
    Dim oMemberSys As ISPSMemberSystem
    Set oMemberSys = oSupports(1)
    
    Dim oMemberPart As ISPSMemberPartPrismatic
    Set oMemberPart = oMemberSys.DesignPartAtEnd(SPSMemberAxisStart)
    
    Dim oMemberSD As StructDetailObjects.MemberPart
    Set oMemberSD = New StructDetailObjects.MemberPart
    
    Dim oCrossSection As ISPSCrossSection
    
    With oMemberSD
         ' use the affected leaf system as the cross section can change
            Set .object = oMemberPart
            Set oCrossSection = oMemberPart.crossSection
            'if the profile has a flange, it can be used for a "1S"
            If Not (oCrossSection Is Nothing) Then
               
                    flangewidth = .FlangeLength / 2
               
            End If
            flangewidth = flangewidth - (0.5 * .WebThickness) - dNose1
            webheight = (0.75 * .WebLength / 2) - .FlangeThickness - dNose2
        End With
    
        Set oMemberSD = Nothing
        Set oMemberPart = Nothing
        Set oCrossSection = Nothing
    End If
skip:
    If (webheight > 0) Then
        pPRL.Add "Height1", webheight
    Else
        Set oAttr = oAttributeCol.Item("Height1")
        pPRL.Add "Height1", oAttr.Value
        Set oAttr = Nothing
    End If
    
    If (flangewidth > 0) Then
        pPRL.Add "Width1", flangewidth
    Else
        Set oAttr = oAttributeCol.Item("Width1")
        pPRL.Add "Width1", oAttr.Value
        Set oAttr = Nothing
    End If

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





