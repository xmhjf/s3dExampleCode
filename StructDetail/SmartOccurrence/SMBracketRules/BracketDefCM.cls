' DI-CP-235957  StrDet: Invalid DB Col Name issues (31 Char limit) view gen. on Oracle DB  
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

Const MODULE = "\StructDetail\SmartOccurrence\BracketRules\" & CUSTOMERID & "BracketDefCM"


Private sError As String
Private sMETHOD As String

Public Sub Bracket_FinalConstruct( _
    ByVal oAggregatorDescription As IJDAggregatorDescription)

    On Error GoTo ErrorHandler
    sMETHOD = "Bracket_FinalConstruct"
    
    Dim oCAO As Object
    Set oCAO = oAggregatorDescription.CAO
    
    If TypeOf oCAO Is IJSmartPlate Then
    
        ' For Bracket Plate Part:
        ' =======================
        
        ' Get bracket plate
        Dim pBracket As IJSmartPlate
        Set pBracket = oCAO
        
        ' Create the SDCreateModifyUtilities and place bracket
        Dim oSmartPlateUtils As New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        Dim oBracketOps As GSCADSDCreateModifyUtilities.IJSDSmartPlateOperations
        
        Set oBracketOps = oSmartPlateUtils
        
        oBracketOps.PlaceBracket pBracket
        
        Set pBracket = Nothing
        Set oSmartPlateUtils = Nothing
      
    Else
        sError = "Unexpected Custom Assembly Object"
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

'***********************************************************************
' METHOD:  CMCreateBracketFlange
'
' DESCRIPTION:  Conditional that determines if there is a bracket flange
'
'***********************************************************************
Public Sub CMCreateBracketFlange(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    sMETHOD = "CMCreateBracketFlange"

    bIsNeeded = True

    Dim oBracketPlate As IJPlate
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Dim sPlaceBracketFlange As String
    Dim pSupportColl As IJElements
    Dim oSelectionRule As IJBracketReinforcementSelectionRule
    ' 2 Supports only
    Dim pHelper As New StructDetailObjects.Helper
    Dim strBracketType As String
    Dim strIASelectionRule As String
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String
    
    GetBracketSOAndPlateFromMD pMD, oBracketUtils, oBracketByPlaneSO, oBracketPlate
        
    sPlaceBracketFlange = GetBracketFlangeAnswer(oBracketPlate, oBracketUtils, oBracketByPlaneSO, pSupportColl)
    
    If sPlaceBracketFlange = "Yes" Then
    
        If pSupportColl.Count >= 3 Then
            'Currently not supporting Flange creation for 3S, 4S and 5S brackets
            strIASelectionRule = "IA" & CUSTOMERID & "BracketRules_BracketByPlaneMT3SSel"
            Call pHelper.GetCustomAttributeParamValue(oBracketByPlaneSO, strIASelectionRule, "BracketByPlaneType", _
                    retValue, retValueType, sTableName, lCodeList, strBracketType, sLongDesc)
            'Create Flange only if Bracket Contour has a linear free edge...
            If strBracketType = "3SLinearTrim" Then
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If
            
        Else
            strIASelectionRule = "IA" & CUSTOMERID & "BracketRules_BracketByPlaneMT2SSel"
            Call pHelper.GetCustomAttributeParamValue(oBracketByPlaneSO, strIASelectionRule, "BracketByPlaneType", _
                    retValue, retValueType, sTableName, lCodeList, strBracketType, sLongDesc)
                
            strBracketType = Mid(strBracketType, 3, 2)
            'Create Flange only if Bracket Contour has a linear free edge...
            If strBracketType = "LT" Then
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If
        End If
    Else

        ' if no flange, see if buckling stiffener or ER is required
        Set oSelectionRule = oBracketUtils.GetBracketReinforcementSelectionRule
        If Not oSelectionRule Is Nothing Then
            bIsNeeded = oSelectionRule.IsReinforcementOnBracketNeeded(oBracketPlate, oBracketByPlaneSO)
        Else
            bIsNeeded = False
        End If
    End If

    ' Trim the Plate system after resymbolization of the bracket contour
    ' to get the actual final geometry when creating brackets without flange.
    ' If it is a flanged bracket(bIsNeeded = True) it is done after creation of Flange.
    If bIsNeeded = False Then
        
        'If the Bracket Plate is already flanged, remove the flange.
        Dim oBracketPlateSystem As IJPlateSystem
        Dim oFlangeAE As Object
        
        Set oBracketPlateSystem = oBracketPlate
        
        If Not oBracketPlateSystem Is Nothing Then
            Set oFlangeAE = oBracketPlateSystem.FlangeActiveEntity(Nothing)
            
            If Not oFlangeAE Is Nothing Then
                oBracketPlateSystem.RemoveFlange Nothing
                Set oFlangeAE = Nothing
            End If
            
            Set oBracketPlateSystem = Nothing
        End If
        
        'also remove any existing reinforcement that may have been created previously
        If oSelectionRule Is Nothing Then _
            Set oSelectionRule = oBracketUtils.GetBracketReinforcementSelectionRule
            
        If Not oSelectionRule Is Nothing Then _
            oSelectionRule.DeleteReinforcementOnBracket oBracketPlate, oBracketByPlaneSO
                
        'Trim the plate system with the bracket contour.
        'Since the Bracket-by-Plane Smart Occurrence has access to all
        'the data (plate system), only it as input is needed.
        oBracketUtils.TrimPlateSystemIntoBracket oBracketByPlaneSO

    Else
        ' Flange/reinforcement is needed.
        If sPlaceBracketFlange = "No" Then

            ' If it is not a flange proper it is assumed that it is bracket reinforcement by rule.
            If oSelectionRule Is Nothing Then
                Set oSelectionRule = oBracketUtils.GetBracketReinforcementSelectionRule
            End If

            If Not oSelectionRule Is Nothing Then
                Dim oBracketObj As IJDObject
                Set oBracketObj = oBracketPlate
                
                oSelectionRule.CreateReinforcementOnBracket oBracketPlate, oBracketByPlaneSO, oBracketObj.ResourceManager

                Set oBracketObj = Nothing
            End If

        End If

    End If
    
    Set oBracketByPlaneSO = Nothing
    Set oBracketUtils = Nothing
    Set pSupportColl = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub ConstructBracketFlange( _
    ByVal oMemberDescription As IJDMemberDescription, _
    ByVal oResourceManager As IUnknown, _
    ByRef oFlange As Object)

    On Error GoTo ErrorHandler
    sMETHOD = "ConstructBracketFlange"
    
    Dim oBracketPlate As IJPlate
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Dim oBracketByPlaneSO As IJSmartOccurrence
    
    GetBracketSOAndPlateFromMD oMemberDescription, oBracketUtils, oBracketByPlaneSO, oBracketPlate
        
    Dim sPlaceBracketFlange As String
    Dim pSupportColl As IJElements
    sPlaceBracketFlange = GetBracketFlangeAnswer(oBracketPlate, oBracketUtils, oBracketByPlaneSO, pSupportColl)
    
    If sPlaceBracketFlange = "Yes" Then
    
        'First check whether the Flange is already there.
        'If available no need to create another.
        Dim bFlangeSOExists As Boolean
        bFlangeSOExists = False
        
        bFlangeSOExists = CheckFlangeSOExists(oBracketPlate)
        
        If bFlangeSOExists = False Then
            oBracketUtils.CreateFlange oBracketByPlaneSO, _
                                       "BracketByPlaneFlgContr", _
                                       oFlange
        End If
    Else ' must be a buckling stiffener or ER
    
        Dim oSelectionRule As IJBracketReinforcementSelectionRule
        Set oSelectionRule = oBracketUtils.GetBracketReinforcementSelectionRule
        
        If Not oSelectionRule Is Nothing Then
            oSelectionRule.CreateReinforcementOnBracket oBracketPlate, oBracketByPlaneSO, oResourceManager
            
            ' After creating the reinforcement on the bracket we have to return it as the flange
            ' object in the return argument of this function.
            
            Dim eRfmtType As ShpStrBracketReinforcementType
            
            oBracketUtils.GetBracketReinforcementCreatedByRule oBracketPlate, eRfmtType, oFlange
            
        End If
    End If
    
    Set oBracketUtils = Nothing
    Set oBracketByPlaneSO = Nothing
    Set oBracketPlate = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

Public Sub DeleteBracketFlange( _
    ByVal oFlangeMember As IJDMemberDescription _
)

    On Error GoTo ErrorHandler
    'Get Plate System, then RemoveFlange.
    sMETHOD = "DeleteBracketFlange"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

Private Function CheckFlangeSOExists(oBracketPlateSys As Object) As Boolean
    Const sMETHOD = "CheckFlangeSOExists"
    On Error GoTo ErrorHandler
    
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
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    GoTo cleanup
End Function

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

Private Function GetBracketFlangeAnswer(oBracketPlate As Object, _
                                        oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes, _
                                        oBracketByPlaneSO As IJSmartOccurrence, _
                                        pSupportColl As IJElements) As String
    
    Dim oBracketByPlane As New StructDetailObjects.BracketByPlane
    Set oBracketByPlane.object = oBracketPlate

    'Get Supports
    Dim pRefPlane As IJPlane
    Dim oUPoint As IJPoint
    Dim oVPoint As IJPoint
    Dim strRootSelector As String

    oBracketByPlane.GetInputs pRefPlane, oUPoint, oVPoint, strRootSelector, pSupportColl

    Set pRefPlane = Nothing
    Set oUPoint = Nothing
    Set oVPoint = Nothing
    
    Dim strIASelectionRuleRoot As String '= "IABracketRules_RootBracketByPlaneSel"

    strIASelectionRuleRoot = "IA" & CUSTOMERID & "BracketRules_RootBracketByPlaneMTSel"
    'define interface from which to get the question
    
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String

    'use helper to get the answer from the question
    Call pHelper.GetCustomAttributeParamValue(oBracketByPlaneSO, strIASelectionRuleRoot, "PlaceBracketFlange", _
            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)

    'check if answer is from code list; if so, get short desc, if not, use retValue
    If Len(Trim(sTableName)) > 0 Then
        GetBracketFlangeAnswer = sShortDesc
    Else
        GetBracketFlangeAnswer = retValue
    End If

End Function



