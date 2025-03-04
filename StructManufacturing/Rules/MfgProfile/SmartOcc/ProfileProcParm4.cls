VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileProcParm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : ProfileProcParm2.cls
'
'  Description  : Profile Marking Parameter rule
'
'  Author       :
'
'  History      : 2004.04.21 MJV Included the correct errorhandling procedures
'
'
'**********************************************************************************************************
Option Explicit

Const MODULE = "ProfileProcParm4"

Const m_ParameterRuleProgid As String = "StructMfgSelectors.ProfileProcParm4"
Const m_ParameterRuleName As String = "StructMfgSelectors.ProfileProcParm4"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
Const METHOD = "ParameterRuleInputs"
  On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PROFILE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pIH.ReportError
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
    Const METHOD = "ParameterRuleLogic"
    On Error GoTo ErrorHandler
    
    Dim oProfileClass As Object

    If TypeOf pPLH.InputObject(INPUT_PROFILE) Is IJStiffenerPart Then
        Set oProfileClass = New StructDetailObjects.ProfilePart
    ElseIf TypeOf pPLH.InputObject(INPUT_PROFILE) Is IJBeamPart Then
        Set oProfileClass = New StructDetailObjects.BeamPart
    End If
    Set oProfileClass.object = pPLH.InputObject(INPUT_PROFILE)

    ' check if the cross-section is a built-up or a flatbar
    Dim bIsBuiltUp As Boolean
    bIsBuiltUp = oProfileClass.IsCrossSectionABuiltUp
       
    Dim oProfilePartSupport As IJProfilePartSupport
    Set oProfilePartSupport = New ProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = pPLH.InputObject(INPUT_PROFILE)
    
    Dim oEntityHelper As New MfgEntityHelper
    Dim sProfileSectionTypeName As String
    sProfileSectionTypeName = oEntityHelper.GetProfileSectionTypeName(oProfilePartSupport)
    If sProfileSectionTypeName = "RectangleTube" Or sProfileSectionTypeName = "Profile_RoundTube" Or _
        sProfileSectionTypeName = "RoundBar" Or sProfileSectionTypeName = "TwoL" Or _
         sProfileSectionTypeName = "Member_RoundTube" Or sProfileSectionTypeName = "LBar" Then
        If oEntityHelper.CheckIfPartExistsInSectionXML(pPLH.InputObject(INPUT_PROFILE)) Then
            pPLH.Add "ProfileUpside", 2112 'WebLeftFromXML
            pPLH.Add "ProfileMfgAsPlate", 2160  ' "Plate-Plate"
        End If
    End If
    Set oEntityHelper = Nothing
    
    Dim eType As ProfileSectionType
    eType = oProfilePartSupport.SectionType
     
    If ((bIsBuiltUp = True) Or (eType = Flat_Bar)) Then
        pPLH.Add "ProfileMfgAsPlate", 2158  ' "Apply"
    Else
        pPLH.Add "ProfileMfgAsPlate", 2159  ' "Ignore"
    End If
    
    ' for BUTL3 - set upside to WEBLEFT
    Dim oXSect As IJCrossSection
    oProfilePartSupport.GetCrossSection oXSect
    Dim strType As String
    strType = oXSect.Type
        
    ' always unfold
    pPLH.Add "ProfileUnfold", 2148  ' "Apply"
    
    Dim eOrientation As IMSProfileEntity.StructMoldedOrientation
    Dim bValid As Boolean

    If TypeOf pPLH.InputObject(INPUT_PROFILE) Is IJStiffenerPart Then
        oProfileClass.Get_SecondOrientation eOrientation, bValid
    ElseIf TypeOf pPLH.InputObject(INPUT_PROFILE) Is IJBeamPart Then
        oProfileClass.Get_PrimOrientation eOrientation, bValid
    End If
    Select Case eOrientation
        Case 0 'Port/fore/above
            If (LCase(strType) = "butl3") Then
                pPLH.Add "ProfileUpside", 2110   ' WebLeft
            Else
                pPLH.Add "ProfileUpside", 2110    ' WebLeftt
            End If
            pPLH.Add "ProfileNeutralAxis", 2120  ' Calculated
            pPLH.Add "ProfileIBL", 2130         '  NumberIs1
            pPLH.Add "ProfileWeldTab", 1141     ' Ignore
            
        Case 1 'Starboard / aft / below
            pPLH.Add "ProfileUpside", 2110   ' WebLeft
            pPLH.Add "ProfileNeutralAxis", 2121    ' Fixed
            pPLH.Add "ProfileIBL", 2130           '  NumberIs1
            pPLH.Add "ProfileWeldTab", 1141       ' Ignore
            
        Case InboardOrient
            If (LCase(strType) = "butl3") Then
                pPLH.Add "ProfileUpside", 2110   ' WebLeft
            Else
                pPLH.Add "ProfileUpside", 2110    ' WebLeft
            End If
            pPLH.Add "ProfileNeutralAxis", 2121  ' Fixed
            pPLH.Add "ProfileIBL", 2130          '  NumberIs1
            pPLH.Add "ProfileWeldTab", 1141      ' Ignore
            
       Case OutboardOrient
            pPLH.Add "ProfileUpside", 2110   ' WebLeft
            pPLH.Add "ProfileNeutralAxis", 2121   ' Fixed
            pPLH.Add "ProfileIBL", 2130           '  NumberIs1
            pPLH.Add "ProfileWeldTab", 1141       ' Ignore

        Case Else
            pPLH.Add "ProfileUpside", 2110   ' WebLeft
            pPLH.Add "ProfileNeutralAxis", 2121  ' Fixed
            pPLH.Add "ProfileIBL", 2130          '  NumberIs1
            pPLH.Add "ProfileWeldTab", 1141      ' Ignore
            
    End Select
    
    ' * Get the Curvature of a Profile Part
    Dim Pl_Curvature As ProfileCurvature
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    
    Set oProfileWrapper.object = pPLH.InputObject(INPUT_PROFILE)
    Pl_Curvature = oProfileWrapper.CurvatureType
    
    Select Case Pl_Curvature
        Case PROFILE_CURVATURE_Straight
            pPLH.Add "ProfileCoamingUnwrap", 2152  ' "Ignore"
            pPLH.Add "ProfileClamp", 2162   '  "Ignore"
            
        Case PROFILE_CURVATURE_CurvedAlongWeb, _
             PROFILE_CURVATURE_CurvedAlongFlange, _
             (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange), _
             PROFILE_CURVATURE_KnuckledAlongWeb, _
             PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             PROFILE_CURVATURE_CurvedAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             PROFILE_CURVATURE_KnuckledAlongFlange, _
             PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_KnuckledAlongFlange, _
             PROFILE_CURVATURE_CurvedAlongFlange Or PROFILE_CURVATURE_KnuckledAlongFlange, _
             PROFILE_CURVATURE_CurvedAlongWeb Or (PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb), _
             PROFILE_CURVATURE_CurvedAlongFlange Or (PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb), _
             PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange) Or PROFILE_CURVATURE_KnuckledAlongWeb, _
             (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange) Or PROFILE_CURVATURE_KnuckledAlongFlange, _
             (PROFILE_CURVATURE_CurvedAlongWeb Or PROFILE_CURVATURE_CurvedAlongFlange) Or _
             (PROFILE_CURVATURE_KnuckledAlongFlange Or PROFILE_CURVATURE_KnuckledAlongWeb), _
             PROFILE_CURVATURE_BendKnuckleAlongFlange, _
             PROFILE_CURVATURE_BendKnuckleAlongWeb
            pPLH.Add "ProfileCoamingUnwrap", 2150  ' Calculated
            pPLH.Add "ProfileClamp", 2160   '  Calculated
            
        Case Else
            pPLH.Add "ProfileCoamingUnwrap", 2151  ' Fixed
            pPLH.Add "ProfileClamp", 2161   '  Fixed
            
    End Select
    
    Set oProfileWrapper = Nothing
    
    pPLH.Add "ProfileBevel", 2171 ' "Ignore"
    pPLH.Add "ProfileShrinkage", 2201 ' "Ignore"
  
    Set oProfileClass = Nothing
    Set oProfilePartSupport = Nothing
    Set oPartSupport = Nothing
    Set oXSect = Nothing
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    'pPLH.ReportError
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractParameterRule pSelector
 
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
 '----------- Define Output and Add it to the representation ------------------
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pSelector.IJDRepresentations.Item(1)
  pOH.SetOutput "ProfileUpside"
  pOH.SetOutput "ProfileNeutralAxis"
  pOH.SetOutput "ProfileIBL"
  pOH.SetOutput "ProfileWeldTab"
  pOH.SetOutput "ProfileCoamingUnwrap"
  pOH.SetOutput "ProfileClamp"
  pOH.SetOutput "ProfileBevel"
  pOH.SetOutput "ProfileMfgAsPlate"
  pOH.SetOutput "ProfileUnfold"
  pOH.SetOutput "ProfileShrinkage"
  
  DefineViewPrefix pSelector, "INGR"
  Set pOH = Nothing

End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)

  Dim pPLH As IJDParameterLogic
  Set pPLH = New ParameterLogic
  pPLH.Representation = pRep
  ParameterRuleLogic pPLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


 


