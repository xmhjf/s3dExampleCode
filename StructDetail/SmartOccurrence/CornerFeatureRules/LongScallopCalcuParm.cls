VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongScallopCalcuParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Parameter Rule:
'
Option Explicit

Const m_ParameterRuleProgid As String = "CornerFeatureRules.LongScallopCalcuParm"
Const m_ParameterRuleName As String = "CornerFeatureRules.LongScallopCalcuParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CornerFeatureRules\LongScallopCalcuParm.cls"

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
  
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
End Sub

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
  
    pOH.SetOutput "Ulength"
    pOH.SetOutput "Radius"
    pOH.SetOutput "Flip"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
  
    ' Create CornerFeature wrapper
    Dim oCornerFeatureWrapper As New Structdetailobjects.CornerFeature
    Dim oPart As Object
    Dim strCornerFlip As String
    
    Set oCornerFeatureWrapper.object = pPRL.SmartOccurrence
    Set oPart = oCornerFeatureWrapper.GetPartObject

    ' Get Corner Feature Orientation Answer
    strCornerFlip = pPRL.SelectorAnswer("CornerFeatureRules.RootCornerSel", gsCornerFlip)

    Dim dULength As Double
    Dim dRadius As Double
    Dim oGapLengthGap As Double
    Dim oGapHeightGap As Double
    Dim oGapTolerance As Double
    
    ' Set Ulength and Radius
    oGapTolerance = 0.001 'will not find gaps < 1 mm
    If oCornerFeatureWrapper.MeasureCornerGap(oGapTolerance, _
                                                                       oGapLengthGap, _
                                                                       oGapHeightGap) Then
        If oGapLengthGap > 0# Then
            dULength = oGapLengthGap + 0.025
        ElseIf oGapHeightGap > 0# Then
            dULength = oGapHeightGap + 0.025
        End If
        
        dRadius = 0.035  'default to 35 mm
    Else
        If TypeOf oPart Is IJPlate Then
            dULength = 0.1
            dRadius = 0.05
        ElseIf TypeOf oPart Is IJStiffener Then
            Dim oProfilePartWrapper As New Structdetailobjects.ProfilePart
            
            Set oProfilePartWrapper.object = oPart
            Select Case oProfilePartWrapper.WebLength
                Case 0 To 0.2
                    dULength = 0.07
                    dRadius = 0.035
                Case Is > 0.2, Is < 0.4
                    dULength = 0.1
                    dRadius = 0.05
                Case Is >= 0.4
                    dULength = 0.15
                    dRadius = 0.075
            End Select
            Set oProfilePartWrapper = Nothing
        ElseIf TypeOf oPart Is IJBeam Then
            Dim oBeamPartWrapper As New Structdetailobjects.BeamPart
            
            Set oBeamPartWrapper.object = oPart
            Select Case oBeamPartWrapper.WebLength
                Case 0 To 0.2
                    dULength = 0.07
                    dRadius = 0.035
                Case Is > 0.2, Is < 0.4
                    dULength = 0.1
                    dRadius = 0.05
                Case Is >= 0.4
                    dULength = 0.15
                    dRadius = 0.075
            End Select
            Set oBeamPartWrapper = Nothing
        End If
    End If
    
    pPRL.Add "Ulength", dULength
    pPRL.Add "Radius", dRadius
    
    ' Set Flip flag
    Dim bIsTBH As Boolean
    
    bIsTBH = False
    If TypeOf oPart Is IJPlate Then
        Dim oPlatePartWrapper As New Structdetailobjects.PlatePart
        
        Set oPlatePartWrapper.object = oPart
        If oPlatePartWrapper.PlateType = TBulkheadPlate Then
            bIsTBH = True
        End If
        Set oPlatePartWrapper = Nothing
    End If
    
    If bIsTBH = True Then
        '
        ' If answer to Flip question is "NoFlip"(default),which means placing the elongation(U) along the most "horizontal" edge
        '    If U vector is close to Z,flip   ---- Initial U direction is NOT what the user wants
        '    else No flip                         ---- Initial U direction is what the user wants
        ' If answer to Flip question is "Flip",which means NOT placing the elongation along the most "horizontal" edge
        '   If U vector is close to Z, No flip ---- Initial U direction is what the user wants
        '   else flip                                  ---- Initial U direction is NOT what the user wants
        '
        Dim oUAxis As IJDVector
        Dim oVAxis As IJDVector
        Dim dUEdgeLength As Double
        Dim dVEdgeLength As Double
        Dim oZAxis As IJDVector
        
        If oCornerFeatureWrapper.GetUnFlippedCornerFeatureData(oUAxis, oVAxis, dUEdgeLength, dVEdgeLength) = True Then
            Dim dDot As Double
            
            Set oZAxis = New DVector
            oZAxis.Set 0#, 0#, 1#
            dDot = Abs(oUAxis.Dot(oZAxis))
            
            Set oUAxis = Nothing
            Set oVAxis = Nothing
            Set oZAxis = Nothing
            
            If dDot > 0.71 Then
                ' The angle between U and Z is greater than 45 degrees
                If strCornerFlip = "NoFlip" Then
                    pPRL.Add "Flip", 1
                Else
                    pPRL.Add "Flip", 0
                End If
            Else
                ' The angle between U and Z is queal to or less than 45 degrees
                If strCornerFlip = "NoFlip" Then
                    pPRL.Add "Flip", 0
                Else
                    pPRL.Add "Flip", 1
                End If
            End If
        Else
            ' Failed to get U direction,set flip flag based on the answer
            If strCornerFlip = "NoFlip" Then
                pPRL.Add "Flip", 0
            Else
                pPRL.Add "Flip", 1
            End If
        End If
    Else
        ' For non-TBH,set flip flag based on the answer
        If strCornerFlip = "NoFlip" Then
            pPRL.Add "Flip", 0
        Else
            pPRL.Add "Flip", 1
        End If
    End If
    
    Set oPart = Nothing
    Set oCornerFeatureWrapper = Nothing
    
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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSymDef As IJDSymbolDefinition)
    On Error Resume Next
    
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    oSymDef.IJDInputs.RemoveAllInput
    oSymDef.IJDRepresentations.RemoveAllRepresentation
    
    Dim oDefFact As New DefinitionFactory
    Dim oIH As IJDInputsHelper
    Dim oOH As IJDOutputsHelper
    
    oDefFact.InitAbstractParameterRule oSymDef
    Set oDefFact = Nothing
    
    Set oIH = New InputHelper
    oIH.Definition = oSymDef
    oIH.InitAs m_FamilyProgid
    ParameterRuleInputs oIH
    Set oIH = Nothing
    
    Set oOH = New OutputHelper
    oOH.Representation = oSymDef.IJDRepresentations.Item(1)
    oOH.InitAs m_FamilyProgid
    ParameterRuleOutputs oOH
    Set oOH = Nothing
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim oDefFact As New DefinitionFactory
    
    Set IJDUserSymbolServices_InstanciateDefinition = oDefFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
    Set oDefFact = Nothing
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
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
