VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrktReinforceSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------------
'  Copyright (C) 2008, Intergraph Corporation.  All rights reserved.
'
'  FILE:  BrktReinforceSel.cls
'
'  DESCRIPTION:
'  Rule to create buckling stiffener or ER reinforcement
'  on a bracket (tripping or by plane)
'
'  ORIGINAL AUTHOR:  Bill Cox
'
'-------------------------------------------------------------------------------

Implements IJBracketReinforcementSelectionRule

Private Const E_FAIL = -2147467259
Private Const MODULE = "S:\ShipStructure\Data\ReinforceBracketRules\BrktReinforceSel.cls"
Private Const BUCKLING_STIFF_NAMING_CATEGORY = "Buckling Stiffener"
Private Const STIFFENER_NAMING_CATEGORY = "Stiffener"
Private sError As String
Private sMETHOD As String

Private m_oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes

Private Sub IJBracketReinforcementSelectionRule_CreateReinforcementOnBracket( _
    ByVal pBracketPlate As IJPlate, _
    ByVal pBracketSymOrSO As Object, _
    ByVal pResourceManager As Object)
    
    ' Create helper for reinforcement creation:
    Dim oBracketUtils As IJBracketAttributes
    Set oBracketUtils = GetBracketUtils
    
    Dim oBracketRule As IJReinforceBracketByRule
    Set oBracketRule = pBracketPlate
    
    Dim eReinforcementType As ShpStrBracketReinforcementType
    eReinforcementType = oBracketRule.BracketReinforcementType
    
    Dim oReinforcement As Object
    Dim existingType As ShpStrBracketReinforcementType
    Dim oBracketByRule As IJBracketReinforcementByRule
    Dim oBracketReinforcementRule As IJBracketReinforcementXSecRule
    Dim oSupports As Collection
    Dim nSupportCount As Long
        
    ' delete old reinforcement if changed
    Dim oOldStiff As IJStiffener
    Dim bReinforcementExists As Boolean
    Dim oProfile As IJProfile
   
    bReinforcementExists = ExistingReinforcementFound(pBracketPlate, existingType, oOldStiff)
    If bReinforcementExists Then
    
        If Not (existingType = eReinforcementType) Then
        
            If existingType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement Or _
               existingType = BRACKETREINFORCEMENTTYPE_BucklingStiffener Then
                    
                DeleteOldReinforcement pBracketPlate, oOldStiff
                Set oOldStiff = Nothing
            End If
        End If
    End If
    
    If eReinforcementType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement Then

        ' create ER on bracket
        Dim oIJERDefinition As IJEdgeReinforcementDefinition
        Set oIJERDefinition = New GSCADCreateModifyUtilities.CEdgeReinforcementUtils
    
        On Error GoTo ErrorHandler ' reset to handle placement errors
        Set oReinforcement = oIJERDefinition.PlaceEdgeReinforcementOnBracketByRule( _
                                         pResourceManager, _
                                         pBracketPlate, _
                                         pBracketSymOrSO, _
                                         oOldStiff)
    
        SetDefaultNamingRule oReinforcement
        
        ' only set the category the first time
        If oOldStiff Is Nothing Then
            Set oProfile = oReinforcement
            oProfile.NamingCategory = GetNamingCategoryPriority(oProfile, STIFFENER_NAMING_CATEGORY, pResourceManager)
        End If

    ElseIf eReinforcementType = BRACKETREINFORCEMENTTYPE_BucklingStiffener Then

        Dim oProfileDefn As IJProfileDefinition
        Dim oBucklingStiffener As IJStiffener
        Dim eProfileType As StructProfileType
        
        eProfileType = GetProfileTypeFromSymbol(pBracketPlate, pBracketSymOrSO)
            
        Set oProfileDefn = New GSCADCreateModifyUtilities.ProfileUtils
        Set oBucklingStiffener = oProfileDefn.PlaceBucklingStiffener( _
                                    eProfileType, _
                                    pResourceManager, _
                                    pBracketPlate, _
                                    pBracketSymOrSO, _
                                    oOldStiff)
        Set oProfileDefn = Nothing
        If Not oBucklingStiffener Is Nothing Then

            ' set up default naming rule
            SetDefaultNamingRule oBucklingStiffener
            
            ' add code to calculate primary orientation here
        End If
        
        ' only set the category the first time
        If oOldStiff Is Nothing Then
            Set oProfile = oBucklingStiffener
            oProfile.NamingCategory = GetNamingCategoryPriority(oProfile, BUCKLING_STIFF_NAMING_CATEGORY, pResourceManager)
        End If
    End If

    Exit Sub
    
ErrorHandler:
    ' do some error handling here
End Sub

Private Sub IJBracketReinforcementSelectionRule_DeleteReinforcementOnBracket( _
    ByVal pBracketPlate As IJPlate, _
    ByVal pBracketSymOrSO As Object)
    
    Dim oOldReinforcement As Object
    Dim existingType As ShpStrBracketReinforcementType
    
    If ExistingReinforcementFound(pBracketPlate, existingType, oOldReinforcement) Then _
        DeleteOldReinforcement pBracketPlate, oOldReinforcement
        
End Sub

Private Function IJBracketReinforcementSelectionRule_IsReinforcementOnBracketNeeded( _
    ByVal pBracketPlate As IJPlate, _
    ByVal pBracketSymOrSO As Object) As Boolean

    Dim oBracketRule As IJReinforceBracketByRule
    Dim ReinforcementType As ShpStrBracketReinforcementType
    Dim bIsNeeded As Boolean
    
    
    bIsNeeded = False
    Set oBracketRule = pBracketPlate
    
    If SymbolSupportsReinforcement(pBracketSymOrSO) Then
        
        ' if override provided, return answer according to override value
        ' otherwise calculate L/T and create by rule
        If oBracketRule.OverrideReinforceBracketRule Then
        
            ReinforcementType = oBracketRule.BracketReinforcementType
        
            If ReinforcementType = BRACKETREINFORCEMENTTYPE_BucklingStiffener Or _
               ReinforcementType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement Then
               
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If
        Else
            Dim oBracketUtils As IJBracketAttributes
            Dim dLTValue As Double
            
            Set oBracketUtils = GetBracketUtils
            oBracketUtils.CalculateBracketLTValue pBracketPlate, pBracketSymOrSO, dLTValue
            Set oBracketUtils = Nothing
            
            ReinforcementType = BRACKETREINFORCEMENTTYPE_None ' default
            
            If (dLTValue >= 50) And (dLTValue < 200) Then
                ReinforcementType = BRACKETREINFORCEMENTTYPE_BucklingStiffener
                bIsNeeded = True
                
            ElseIf dLTValue >= 200 Then
                ReinforcementType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement
                bIsNeeded = True
                
            Else ' otherwise do nothing
            End If
            
            oBracketRule.BracketReinforcementType = ReinforcementType
            
        End If
        
    Else ' symbol does not support reinforcement, so force the value
        oBracketRule.BracketReinforcementType = BRACKETREINFORCEMENTTYPE_None
    End If
    
    IJBracketReinforcementSelectionRule_IsReinforcementOnBracketNeeded = bIsNeeded

End Function

' checks for existing reinforcement, returning the type found and stiffener object
' if its a buckling stiffener or ER
Private Function ExistingReinforcementFound(oBracket As Object, _
                                            existingType As ShpStrBracketReinforcementType, _
                                            oOldStiff As IJStiffener) As Boolean

    On Error Resume Next
    Dim oBracketUtils As IJBracketAttributes
    Dim oReinforcementUnk As IUnknown
    Dim exists As Boolean
    exists = False
    
    Set oBracketUtils = GetBracketUtils
    oBracketUtils.GetBracketReinforcementCreatedByRule oBracket, existingType, oReinforcementUnk
    Set oBracketUtils = Nothing
    
    If Not oReinforcementUnk Is Nothing Then
        exists = True
        Set oOldStiff = oReinforcementUnk
    Else
        exists = False
    End If
    
    ExistingReinforcementFound = exists
End Function

Private Sub SetDefaultNamingRule(oReinforcement As Object)

    Const METHOD = "SetDefaultNamingRule"
    
    Dim oNamingUtils As IJNamingUtils2
    Dim oStructEntityNaming As IJDStructEntityNaming
    
    ' don't fail out if naming rule not set correctly
    On Error Resume Next
    Set oNamingUtils = New GSCADCreateModifyUtilities.StructEntityUtils
    
    Set oStructEntityNaming = oReinforcement
    oStructEntityNaming.NamingRule = _
        oNamingUtils.GetDefaultNamingRule(StructObjectType_ProfileSystem)
           
    Exit Sub
End Sub

Private Function GetBracketUtils() As IJBracketAttributes
    If m_oBracketUtils Is Nothing Then
        Set m_oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    End If
    
    Set GetBracketUtils = m_oBracketUtils
End Function

Private Sub DeleteOldReinforcement(oBracketPlate As IJPlate, oOldReinforcement As Object)
    
    Dim oBracketUtils As IJBracketAttributes
    Set oBracketUtils = GetBracketUtils
    
    ' delete the ER or buckling stiffener
    oBracketUtils.DeleteBracketReinforcementCreatedByRule oBracketPlate, oOldReinforcement

End Sub

Private Function GetProfileTypeFromSymbol(oPlate As Object, oSymbol As Object) As StructProfileType
    Dim oPlateUtil As IJPlateAttributes
    Dim oProfUtil As IJProfileAttributes
    Dim eProfileType As StructProfileType
        
    'Evaluate the Profile type that could be created with this combination
    Set oPlateUtil = New PlateUtils
    Set oProfUtil = New ProfileUtils
    Dim oWireBody As IJWireBody
    
    eProfileType = sptUnknown
    
    On Error Resume Next
    oPlateUtil.GetWireBodyFromSymbol oSymbol, oWireBody
    If Not oWireBody Is Nothing Then
        eProfileType = oProfUtil.GetProfileTypeFromCurveAndPlateType(oPlate, oWireBody)
    Else
        eProfileType = sptUnknown
    End If
    
    GetProfileTypeFromSymbol = eProfileType

End Function

Private Function SymbolSupportsReinforcement(oSymbol As Object) As Boolean
    Dim oBSContour As IJDStructContour
    Dim dLength As Double
    Dim oBracketUtils As IJBracketAttributes
    Set oBracketUtils = GetBracketUtils
    
    On Error Resume Next
    oBracketUtils.GetBucklingStiffenerContourFromSymbol oSymbol, oBSContour
    oBracketUtils.GetUnsupportedEdgeLengthFromSymbol Nothing, oSymbol, dLength
    
    If (dLength > 0) And (Not oBSContour Is Nothing) Then
        SymbolSupportsReinforcement = True
    Else
        SymbolSupportsReinforcement = False
    End If
End Function

Public Function GetNamingCategoryPriority(oReinforce As IJProfile, _
                                          sNamingCategory As String, _
                                          oResourceManager As Object) As Long
                                          
    Const PROCNAME = "GetNamingCategoryPriority"
    
    Dim oUtil As IJMetaDataCategoryQuery
    Dim arrLongNames() As String
    Dim arrShortNames() As String
    Dim arrPriority() As Long
    Dim lLbound As Long
    Dim lUbound As Long
    Dim lIndex As Long
    
    On Error Resume Next
    
    Set oUtil = New CMetaDataCategoryQuery
    If TypeOf oReinforce Is IJERSystem Then
        oUtil.GetEdgeReinforcementCategoryInfo oResourceManager, _
                                               arrLongNames, _
                                               arrShortNames, _
                                               arrPriority
    Else ' just standard profile, check its type
        Dim oStiffener As IJStiffener
        Set oStiffener = oReinforce
        Select Case oStiffener.pType
            Case StructProfileType.sptLongitudinal
                oUtil.GetLongitudinalProfileCategoryInfo oResourceManager, _
                                                         arrLongNames, _
                                                         arrShortNames, _
                                                         arrPriority
            Case StructProfileType.sptTransversal
                oUtil.GetTransverseProfileCategoryInfo oResourceManager, _
                                                       arrLongNames, _
                                                       arrShortNames, _
                                                       arrPriority
            Case Else ' just use vertical
                oUtil.GetVerticalProfileCategoryInfo oResourceManager, _
                                                     arrLongNames, _
                                                     arrShortNames, _
                                                     arrPriority
        End Select
    End If
    
    lLbound = LBound(arrLongNames)
    lUbound = UBound(arrLongNames)
    
    For lIndex = lLbound To lUbound
        If arrLongNames(lIndex) = sNamingCategory Then
            GetNamingCategoryPriority = arrPriority(lIndex)
            Exit For

        End If
    Next lIndex
    
    Set oUtil = Nothing
    
    Erase arrLongNames
    Erase arrShortNames
    Erase arrPriority
    
End Function


