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
'  Copyright (C) 2008-2010, Intergraph Corporation.  All rights reserved.
'
'  FILE:  BrktReinforceSel.cls
'
'  DESCRIPTION:
'  Rule to create buckling stiffener or ER reinforcement
'  on a bracket (tripping or by plane)
'
'  ORIGINAL AUTHOR:  Bill Cox
'  DI-CP-235957  StrDet: Invalid DB Col Name issues (31 Char limit) view gen. on Oracle DB 
'-------------------------------------------------------------------------------

Implements IJBracketReinforcementSelectionRule

Private Const E_FAIL = -2147467259
Private Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SMBracketReinforcementRules\BrktReinforceSel.cls"
Private Const BUCKLING_STIFF_NAMING_CATEGORY = "Buckling Stiffener"
Private Const STIFFENER_NAMING_CATEGORY = "Stiffener"
Private Const FLANGE_ROOT_SELECTOR = "BracketPltSysFlgContr"
Private sError As String
Private sMETHOD As String

Private m_bIsBPS As Boolean

Private m_oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes

Private Sub IJBracketReinforcementSelectionRule_CreateReinforcementOnBracket( _
    ByVal pBracketPlate As IJPlate, _
    ByVal pBracketSymOrSO As Object, _
    ByVal pResourceManager As Object)
    
    Const METHOD = "IJBracketReinforcementSelectionRule_CreateReinforcementOnBracket"
    If TypeOf pBracketPlate Is IJBracketPlateSystem Then m_bIsBPS = True Else m_bIsBPS = False
    
    ' Create helper for reinforcement creation:
    Dim oBracketUtils As IJBracketAttributes
    Set oBracketUtils = GetBracketUtils
    
    Dim oBracketRule As IJReinforceBracketByRule
    Dim oBPS As IJBracketPlateSystem
    Dim lReinforcementType As Long
    
    If m_bIsBPS Then
        Set oBPS = pBracketPlate
        lReinforcementType = oBPS.GetReinforcementType(BktReinforcementType_Mask)
    Else
        Set oBracketRule = pBracketPlate

        lReinforcementType = oBracketRule.BracketReinforcementType
        
        ' translate old types to new ones
        If (lReinforcementType = BRACKETREINFORCEMENTTYPE_BucklingStiffener) Then
            lReinforcementType = BktReinforcementType_BucklingStiffener
        ElseIf (lReinforcementType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement) Then
            lReinforcementType = BktReinforcementType_EdgeReinforcement
        Else
            lReinforcementType = BktReinforcementType_None
        End If
    End If
    
    Dim oReinforcement As Object
    Dim existingType As Long
    Dim oBracketReinforcementRule As IJBracketReinforcementXSecRule
        
    ' delete old reinforcement if changed
    Dim oOldStiff As IJStiffener
    Dim bReinforcementExists As Boolean
    Dim oProfile As IJProfile
    
    bReinforcementExists = ExistingReinforcementFound(pBracketPlate, existingType, oOldStiff)
    If bReinforcementExists Then
    
        If Not (existingType = lReinforcementType) Then
        
            If (existingType And BktReinforcementType_EdgeReinforcement) Or _
               (existingType And BktReinforcementType_BucklingStiffener) Then
                    
                DeleteOldReinforcement pBracketPlate, oOldStiff
                Set oOldStiff = Nothing
            ElseIf existingType And BktReinforcementType_Flange Then
                DeleteOldReinforcement pBracketPlate, Nothing
            End If
        End If
    End If
    
    If lReinforcementType And BktReinforcementType_EdgeReinforcement Then

        ' create ER on bracket
        Dim oIJERDefinition As IJEdgeReinforcementDefinition
        Set oIJERDefinition = New GSCADCreateModifyUtilities.CEdgeReinforcementUtils
    
        On Error GoTo ErrorHandler ' reset to handle placement errors
        Set oReinforcement = oIJERDefinition.PlaceEdgeReinforcementOnBracketByRule( _
                                         pResourceManager, _
                                         pBracketPlate, _
                                         pBracketSymOrSO, _
                                         oOldStiff)
    
        If Not oReinforcement Is Nothing Then
            SetDefaultNamingRule oReinforcement
    
            ' only set the category the first time
            If oOldStiff Is Nothing Then
                Set oProfile = oReinforcement
                oProfile.NamingCategory = GetNamingCategoryPriority(oProfile, STIFFENER_NAMING_CATEGORY, pResourceManager)
            End If
        End If
        
    ElseIf lReinforcementType And BktReinforcementType_BucklingStiffener Then

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
            
            ' only set the category the first time
            If oOldStiff Is Nothing Then
                On Error Resume Next
                Set oProfile = oBucklingStiffener
                If Not oProfile Is Nothing Then
                    oProfile.NamingCategory = GetNamingCategoryPriority(oProfile, BUCKLING_STIFF_NAMING_CATEGORY, pResourceManager)
                End If
                On Error GoTo ErrorHandler
            End If
            
            ' add code to calculate primary orientation here
        End If
    ElseIf lReinforcementType And BktReinforcementType_Flange Then
        Dim bFlangeSOExists As Boolean
        bFlangeSOExists = False

        Dim oFlange As Object
        bFlangeSOExists = CheckFlangeSOExists(pBracketPlate)
        If bFlangeSOExists = False Then
            oBracketUtils.CreateFlange pBracketSymOrSO, _
                                       FLANGE_ROOT_SELECTOR, _
                                       oFlange
        End If
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Private Sub IJBracketReinforcementSelectionRule_DeleteReinforcementOnBracket( _
    ByVal pBracketPlate As IJPlate, _
    ByVal pBracketSymOrSO As Object)
    Const METHOD = "IJBracketReinforcementSelectionRule_DeleteReinforcementOnBracket"
    On Error GoTo ErrorHandler

    If TypeOf pBracketPlate Is IJBracketPlateSystem Then m_bIsBPS = True Else m_bIsBPS = False
    
    Dim oOldReinforcement As Object
    Dim existingType As Long
    
    If ExistingReinforcementFound(pBracketPlate, existingType, oOldReinforcement) Then _
        DeleteOldReinforcement pBracketPlate, oOldReinforcement
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
        
End Sub

Private Function IJBracketReinforcementSelectionRule_IsReinforcementOnBracketNeeded( _
    ByVal pBracketPlate As IJPlate, _
    ByVal pBracketSymOrSO As Object) As Boolean
    
    If TypeOf pBracketPlate Is IJBracketPlateSystem Then m_bIsBPS = True Else m_bIsBPS = False


    Const METHOD = "IJBracketReinforcementSelectionRule_IsReinforcementOnBracketNeeded"
    On Error GoTo ErrorHandler
    
    Dim oBracketRule As IJReinforceBracketByRule
    Dim oBPS As IJBracketPlateSystem
    Dim ReinforcementType As Long
    Dim bIsNeeded As Boolean
    
    ReinforcementType = BktReinforcementType_None
    bIsNeeded = False
    Dim bOverrideRule As Boolean
    bOverrideRule = False
    
    If Not m_bIsBPS Then
        Set oBracketRule = pBracketPlate
        bOverrideRule = oBracketRule.OverrideReinforceBracketRule
        ReinforcementType = oBracketRule.BracketReinforcementType
        
        ' translate old types to new ones
        If (ReinforcementType = BRACKETREINFORCEMENTTYPE_BucklingStiffener) Then
            ReinforcementType = BktReinforcementType_BucklingStiffener
        ElseIf (ReinforcementType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement) Then
            ReinforcementType = BktReinforcementType_EdgeReinforcement
        Else
            ReinforcementType = BktReinforcementType_None
        End If
    Else
        Dim lOverride As Long
        Set oBPS = pBracketPlate
        lOverride = oBPS.GetOverrideType(BktOverride_ReinforcementRule)
        If (lOverride And BktOverride_ReinforcementRule) Then
            bOverrideRule = True
            ReinforcementType = oBPS.GetReinforcementType(BktReinforcementType_Mask)
        End If
    End If
    
    bIsNeeded = False
    
    If SymbolSupportsReinforcement(pBracketSymOrSO) Or _
       (bOverrideRule And ReinforcementType = BktReinforcementType_Flange) Then
        
        ' if override provided, return answer according to override value
        ' otherwise calculate L/T and create by rule
                
        If bOverrideRule Then
        
            If (ReinforcementType And BktReinforcementType_BucklingStiffener) Or _
               (ReinforcementType And BktReinforcementType_EdgeReinforcement) Or _
               (ReinforcementType And BktReinforcementType_Flange) Then
               
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
            
            ReinforcementType = BktReinforcementType_None ' default
            
            If (dLTValue >= 50) And (dLTValue < 200) Then
                ReinforcementType = BktReinforcementType_BucklingStiffener
                bIsNeeded = True
                
            ElseIf dLTValue >= 200 Then
                ReinforcementType = BktReinforcementType_EdgeReinforcement
                bIsNeeded = True
                
            Else ' otherwise do nothing
            End If
            
            If Not m_bIsBPS Then
                
                ' translate back to old ones
                If (ReinforcementType And BktReinforcementType_BucklingStiffener) Then
                    ReinforcementType = BRACKETREINFORCEMENTTYPE_BucklingStiffener
                ElseIf (ReinforcementType And BktReinforcementType_EdgeReinforcement) Then
                    ReinforcementType = BRACKETREINFORCEMENTTYPE_EdgeReinforcement
                Else
                    ReinforcementType = BRACKETREINFORCEMENTTYPE_None
                End If
                oBracketRule.BracketReinforcementType = ReinforcementType
            Else
                oBPS.SetReinforcementType BktReinforcementType_Mask, ReinforcementType
            End If
            
        End If
                
    Else ' symbol does not support reinforcement, so force the value
        If Not m_bIsBPS Then
            oBracketRule.BracketReinforcementType = BRACKETREINFORCEMENTTYPE_None
        Else
            oBPS.SetReinforcementType BktReinforcementType_Mask, BktReinforcementType_None
        End If
    End If
    
    IJBracketReinforcementSelectionRule_IsReinforcementOnBracketNeeded = bIsNeeded
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function

' checks for existing reinforcement, returning the type found and stiffener object
' if its a buckling stiffener or ER
Private Function ExistingReinforcementFound(oBracket As Object, _
                                            existingType As Long, _
                                            oOldStiff As IJStiffener) As Boolean

    Const METHOD = "ExistingReinforcementFound"
    On Error GoTo ErrorHandler

    On Error Resume Next
    Dim oBracketUtils As IJBracketAttributes
    Dim oReinforcementUnk As IUnknown
    Dim exists As Boolean
    exists = False
    
    Set oBracketUtils = GetBracketUtils
    oBracketUtils.GetBktReinforcementCreatedByRule oBracket, existingType, oReinforcementUnk
    
    ' only BPS supports flange
    If Not m_bIsBPS And (existingType And BktReinforcementType_Flange) Then
        existingType = BktReinforcementType_None
        Set oReinforcementUnk = Nothing
    End If
    
    Set oBracketUtils = Nothing
    'DI-CP-238709  Improve Bracket Resymbolization Large Edit Performance  
    If existingType <> 0 Then
        exists = True
        Set oOldStiff = oReinforcementUnk
    Else
        exists = False
    End If
    
    ExistingReinforcementFound = exists
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Private Sub SetDefaultNamingRule(oReinforcement As Object)

    Const METHOD = "SetDefaultNamingRule"
    On Error GoTo ErrorHandler
    
    Dim oNamingUtils As IJNamingUtils2
    Dim oStructEntityNaming As IJDStructEntityNaming
    
    ' don't fail out if naming rule not set correctly
    On Error Resume Next
    Set oNamingUtils = New GSCADCreateModifyUtilities.StructEntityUtils
    
    Set oStructEntityNaming = oReinforcement
    oStructEntityNaming.NamingRule = _
        oNamingUtils.GetDefaultNamingRule(StructObjectType_ProfileSystem)
       
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Private Function GetBracketUtils() As IJBracketAttributes
    If m_oBracketUtils Is Nothing Then
        Set m_oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    End If
    
    Set GetBracketUtils = m_oBracketUtils
End Function

Private Sub DeleteOldReinforcement(oBracketPlate As IJPlate, oOldReinforcement As Object)
    
    Const METHOD = "DeleteOldReinforcement"
    On Error GoTo ErrorHandler
    
    Dim oBracketUtils As IJBracketAttributes
    Set oBracketUtils = GetBracketUtils
    
    ' delete the ER or buckling stiffener
    oBracketUtils.DeleteBracketReinforcementCreatedByRule oBracketPlate, oOldReinforcement
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Sub

Private Function GetProfileTypeFromSymbol(oPlate As Object, oSymbol As Object) As StructProfileType



    Const METHOD = "GetProfileTypeFromSymbol"
    On Error GoTo ErrorHandler
    
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
    'GetProfileTypeFromCurveAndPlateType will complain if the bracket's plate type is not yet defined.
    If Not oWireBody Is Nothing And Not oPlate.plateType = UntypedPlate Then
        eProfileType = oProfUtil.GetProfileTypeFromCurveAndPlateType(oPlate, oWireBody)
    Else
        eProfileType = sptUnknown
    End If
     
    GetProfileTypeFromSymbol = eProfileType
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function

Private Function SymbolSupportsReinforcement(oSymbol As Object) As Boolean


    Const METHOD = "SymbolSupportsReinforcement"
    On Error GoTo ErrorHandler
    
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
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Public Function GetNamingCategoryPriority(oReinforce As IJProfile, _
                                          sNamingCategory As String, _
                                          oResourceManager As Object) As Long
                                          
    Const METHOD = "GetNamingCategoryPriority"
    
    On Error GoTo ErrorHandler
    
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
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    
End Function

Private Function CheckFlangeSOExists(oBracketPlateSys As Object) As Boolean
    Const sMETHOD = "CheckFlangeSOExists"
    On Error Resume Next

    'Initialize the function output
    CheckFlangeSOExists = False

    Dim oBracketPlate As IJPlateSystem
    Dim oFlangeAE As IJPlateFlange_AE
    Dim oFlange As Object
    Dim oFlangeSymbol As Object

    Set oBracketPlate = oBracketPlateSys

    Set oFlange = oBracketPlate.FlangeActiveEntity(Nothing)
    If Not oFlange Is Nothing Then
        Set oFlangeAE = oFlange
        Set oFlangeSymbol = oFlangeAE.FlangeSymbol
    End If

    If Not oFlangeSymbol Is Nothing Then
        'Flange already exists
        CheckFlangeSOExists = True
    End If

cleanup:
    Set oFlangeSymbol = Nothing
    Set oFlangeAE = Nothing
    Set oFlange = Nothing
    Set oBracketPlate = Nothing

End Function

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Private Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

