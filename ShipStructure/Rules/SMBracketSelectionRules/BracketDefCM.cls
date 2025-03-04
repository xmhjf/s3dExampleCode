VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SMBracketSelectionRules\BracketDefCM"


Private sError As String
Private sMETHOD As String


'***********************************************************************
' METHOD:  CMBracketReinforcementCondition
'
' DESCRIPTION:  Conditional that determines if there is a bracket flange
'
'***********************************************************************

Public Sub CMBracketReinforcementCondition(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    ' This CM method is no more used.
    
    Const sMETHOD = "CMBracketReinforcementCondition"
    
    IsBracketReinforcementNeeded pMD, bIsNeeded
    
    Exit Sub
    
ErrorHandler:
    'If reinforcement creation failed then it is a warning
    'Raise warning so that bound happens.
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Sub IsBracketReinforcementNeeded(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    Const sMETHOD = "IsBracketReinforcementNeeded"
    
    Dim oBracketPlate As IJPlate
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Dim sPlaceBracketFlange As String
    Dim pSupportColl As IJElements
    Dim oSelectionRule As IJBracketReinforcementSelectionRule
    
    bIsNeeded = False
    GetBracketSOAndPlateFromMD pMD, oBracketUtils, oBracketByPlaneSO, oBracketPlate
    
    'Trim is moved to CMEvaluateBracketboundaries

    ' see if rule-based reinforcement is required
    Set oSelectionRule = oBracketUtils.GetBracketReinforcementSelectionRule
    If Not oSelectionRule Is Nothing Then
        bIsNeeded = oSelectionRule.IsReinforcementOnBracketNeeded(oBracketPlate, oBracketByPlaneSO)
    Else
        bIsNeeded = False
    End If
    
    ' to get the actual final geometry when creating brackets without flange.
    ' If it is a flanged bracket(bIsNeeded = True) it is done after creation of Flange.
    If bIsNeeded = False Then
        
        'remove any existing reinforcement that may have been created previously
        If Not oSelectionRule Is Nothing Then _
            oSelectionRule.DeleteReinforcementOnBracket oBracketPlate, oBracketByPlaneSO

    Else
        ' Flange/reinforcement is needed.
        If Not oSelectionRule Is Nothing Then
            Dim oBracketObj As IJDObject
            Set oBracketObj = oBracketPlate
            
            oSelectionRule.CreateReinforcementOnBracket oBracketPlate, oBracketByPlaneSO, oBracketObj.ResourceManager

            Set oBracketObj = Nothing
        End If

    End If
    
    Set oBracketByPlaneSO = Nothing
    Set oBracketUtils = Nothing
    Set pSupportColl = Nothing
    
    Exit Sub
    
ErrorHandler:
    'If reinforcement creation failed then it is a warning
    'Raise warning so that bound happens.
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Sub ConstructBracketReinforcement( _
    ByVal oMemberDescription As IJDMemberDescription, _
    ByVal oResourceManager As IUnknown, _
    ByRef oFlange As Object)

    On Error GoTo ErrorHandler
    Const sMETHOD = "ConstructBracketReinforcement"

    Dim oBracketPlate As IJPlate
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Dim oBracketByPlaneSO As IJSmartOccurrence
    
    GetBracketSOAndPlateFromMD oMemberDescription, oBracketUtils, oBracketByPlaneSO, oBracketPlate
        
    Dim sPlaceBracketFlange As String
    Dim pSupportColl As IJElements
'    sPlaceBracketFlange = GetBracketFlangeAnswer(oBracketPlate, oBracketUtils, oBracketByPlaneSO, pSupportColl)
'
'    If sPlaceBracketFlange = "Yes" Then
'
'        'First check whether the Flange is already there.
'        'If available no need to create another.
'        Dim bFlangeSOExists As Boolean
'        bFlangeSOExists = False
'
'        bFlangeSOExists = CheckFlangeSOExists(oBracketPlate)
'        If bFlangeSOExists = False Then
'            oBracketUtils.CreateFlange oBracketByPlaneSO, _
'                                       "BracketByPlaneFlangeContour", _
'                                       oFlange
'        End If
'    Else ' must be a buckling stiffener or ER
    
        Dim oSelectionRule As IJBracketReinforcementSelectionRule
        Set oSelectionRule = oBracketUtils.GetBracketReinforcementSelectionRule
        
        If Not oSelectionRule Is Nothing Then
            oSelectionRule.CreateReinforcementOnBracket oBracketPlate, oBracketByPlaneSO, oResourceManager
            
            ' After creating the reinforcement on the bracket we have to return it as the flange
            ' object in the return argument of this function
            
            Dim lRfmtType As Long
            
            If TypeOf oBracketPlate Is IJBracketPlateSystem Then
                oBracketUtils.GetBktReinforcementCreatedByRule oBracketPlate, lRfmtType, oFlange
            Else
                oBracketUtils.GetBracketReinforcementCreatedByRule oBracketPlate, lRfmtType, oFlange
            End If
            
        End If
'    End If
    
    Set oBracketUtils = Nothing
    Set oBracketByPlaneSO = Nothing
    Set oBracketPlate = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

Public Sub DeleteBracketReinforcement(ByVal oFlangeMember As IJDMemberDescription)

    On Error GoTo ErrorHandler
    'Get Plate System, then RemoveFlange.
    Const sMETHOD = "DeleteBracketReinforcement"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub
Public Sub FinalConstructBracketReinforcement(pMemberDesc As IJDMemberDescription)
    On Error GoTo ErrorHandler
    Const sMETHOD = "FinalConstructBracket"
    
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Private Function GetBracketSOAndPlateFromMD(oMD As IJDMemberDescription, _
                                            oBracketUtils As IJBracketAttributes, _
                                            oBracketByPlaneSO As IJSmartOccurrence, _
                                            oBracketPlate As IJPlate)
    
    ' Get Bracket-by-Plane Smart Occurrence:
    Set oBracketByPlaneSO = oMD.CAO
    
    If oBracketUtils Is Nothing Then _
        Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils

    Set oBracketPlate = oBracketUtils.GetBracketByPlaneFromBracketContour(oBracketByPlaneSO)

End Function

'Private Function CheckFlangeSOExists(oBracketPlateSys As Object) As Boolean
'    Const sMETHOD = "CheckFlangeSOExists"
'    On Error GoTo ErrorHandler
'
'    'Initialize the function output
'    CheckFlangeSOExists = False
'
'    Dim oBracketPlate As IJPlateSystem
'    Dim oFlangeAE As IJPlateFlange_AE
'    Dim oFlange As Object
'    Dim oFlangeSymbol As Object
'
'    Set oBracketPlate = oBracketPlateSys
'
'    Set oFlange = oBracketPlate.FlangeActiveEntity(Nothing)
'    If Not oFlange Is Nothing Then
'        Set oFlangeAE = oFlange
'        Set oFlangeSymbol = oFlangeAE.FlangeSymbol
'    End If
'
'    If Not oFlangeSymbol Is Nothing Then
'        'Flange already exists
'        CheckFlangeSOExists = True
'    End If
'
'cleanup:
'    Set oFlangeSymbol = Nothing
'    Set oFlangeAE = Nothing
'    Set oFlange = Nothing
'    Set oBracketPlate = Nothing
'
'    Exit Function
'
'ErrorHandler:
'    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
'    GoTo cleanup
'End Function

'Private Function GetBracketFlangeAnswer(oBracketPlate As Object, _
'                                        oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes, _
'                                        oBracketByPlaneSO As IJSmartOccurrence, _
'                                        pSupportColl As IJElements) As String
'
'    Dim oBracketByPlane As New StructDetailObjects.BracketByPlane
'    Set oBracketByPlane.object = oBracketPlate
'
'    'Get Supports
'    Dim pRefPlane As IJPlane
'    Dim oUPoint As IJPoint
'    Dim oVPoint As IJPoint
'    Dim strRootSelector As String
'
'    oBracketByPlane.GetInputs pRefPlane, oUPoint, oVPoint, strRootSelector, pSupportColl
'
'    Set pRefPlane = Nothing
'    Set oUPoint = Nothing
'    Set oVPoint = Nothing
'
'    Dim strIASelectionRuleRoot As String '= "IABracketPlateSystemRules_RootBktSel"
'
'    strIASelectionRuleRoot = "IABracketPlateSystemRules_RootBktSel"
'    'define interface from which to get the question
'
'    Dim pHelper As New StructDetailObjects.Helper
'    Dim retValue As Variant
'    Dim retValueType As EResultType
'    Dim sTableName As String
'    Dim lCodeList As Long
'    Dim sShortDesc As String
'    Dim sLongDesc As String
'
'    'use helper to get the answer from the question
'    Call pHelper.GetCustomAttributeParamValue(oBracketByPlaneSO, strIASelectionRuleRoot, "PlaceBracketFlange", _
'            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)
'
'    'Yes/No are 65536 and 65537.  If the returned value is zero, assume that the
'    'SmartOccurrence answer hasn't yet been defined, in which case return "No"
'    '(the presumed default)
'
'    If retValue = 0 Then
'        GetBracketFlangeAnswer = "No"
'    Else
'        'check if answer is from code list; if so, get short desc, if not, use retValue
'        If Len(Trim(sTableName)) > 0 Then
'            GetBracketFlangeAnswer = sShortDesc
'        Else
'            GetBracketFlangeAnswer = retValue
'        End If
'    End If
'End Function
