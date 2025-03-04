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
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : LongScallopCalcuParm.cls
'
'  Description :
'
'  Author      : Alligators
'
'   pnalugol - Nov 1 2011 - 202656 - Changes made to support Corner features on Slots.
'                         - Used StructDetailObjectsEx to create corner feature

Const m_ParameterRuleProgid As String = CUSTOMERID + "CornerFeatRules.LongScallopCalcuParm"
Const m_ParameterRuleName As String = CUSTOMERID + "CornerFeatRules.LongScallopCalcuParm"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "CornerFeatRules\LongScallopCalcuParm.cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
  
    pIH.SetInput INPUT_PORT1FACE
    pIH.SetInput INPUT_PORT2EDGE
    pIH.SetInput INPUT_PORT3EDGE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleInputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    On Error GoTo ErrorHandler
  
    pOH.SetOutput "Ulength"
    pOH.SetOutput "Radius"
    pOH.SetOutput "Flip"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ParameterRuleOutputs").Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    On Error GoTo ErrorHandler
  
    ' Create CornerFeature wrapper
    Dim oCornerFeatureWrapper As IJSDOCornerFeature
    Set oCornerFeatureWrapper = New StructDetailObjectsex.CornerFeature
    Set oCornerFeatureWrapper.object = pPRL.SmartOccurrence

    Dim oPart As Object
    Dim strCornerFlip As String
        Dim SOCorner As IJSmartOccurrence
    Set SOCorner = pPRL.SmartOccurrence
    
    Dim sItem As String
    sItem = SOCorner.Item
    
    Set oPart = oCornerFeatureWrapper.GetPartObject
    
    Dim dDistanceToSeam As Double
    Dim bSeamFound As Boolean
    
    Dim dRadius As Double
    dRadius = 0.05   'default to 50 mm
    
    GetAttributeOnInterface pPRL.SmartOccurrence, "IJUACornerFeatureSeam", "SeamFound", bSeamFound
    GetAttributeOnInterface pPRL.SmartOccurrence, "IJUACornerFeatureSeam", "DistToSeam", dDistanceToSeam
    
    If bSeamFound Then
        pPRL.Add "Flip", 0
        pPRL.Add "Ulength", dDistanceToSeam + SEAM_SEARCHDISTANCE
        pPRL.Add "Radius", dRadius
                 
    Exit Sub
    End If

    ' Get Corner Feature Orientation Answer

    GetSelectorAnswer pPRL, gsCornerFlip, strCornerFlip
    
    Dim dULength As Double
   
    Dim oGapLengthGap As Double
    Dim oGapHeightGap As Double
    Dim oGapTolerance As Double
    
    Dim sName As String
    sName = "LongScallop 50x100 Calculate"
    
    Dim sNameAlongCorner As String
    sNameAlongCorner = "LongScallop Along Corner 50x100 Calculate"

    ' Set Ulength and Radius
    oGapTolerance = 0.001 'will not find gaps < 1 mm
    If oCornerFeatureWrapper.MeasureCornerGap(oGapTolerance, _
                                                                       oGapLengthGap, _
                                                                       oGapHeightGap) And sItem <> sName And sItem <> sNameAlongCorner Then
        If oGapLengthGap > 0# Then
            dULength = oGapLengthGap + 0.015
        ElseIf oGapHeightGap > 0# Then
            dULength = oGapHeightGap + 0.015
        End If
        
        dRadius = 0.05   'default to 50 mm
    Else
        If TypeOf oPart Is IJPlate Then
            dULength = 0.1
            dRadius = 0.05
            
        ElseIf TypeOf oPart Is IJStiffener Then
            Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
            
            Set oProfilePartWrapper.object = oPart
            
        If sName = sItem Or sNameAlongCorner = sItem Then
         dULength = 0.1
         dRadius = 0.05
        Else
            Select Case oProfilePartWrapper.WebLength
                Case 0 To 0.2
                    dULength = 0.07
                    dRadius = 0.035
                Case 0.2 To 0.4
                    dULength = 0.1
                    dRadius = 0.05
                Case Is >= 0.4
                    dULength = 0.15
                    dRadius = 0.075
            End Select
            Set oProfilePartWrapper = Nothing
           End If
        ElseIf TypeOf oPart Is IJBeam Then
            Dim oBeamPartWrapper As New StructDetailObjects.BeamPart
            
            Set oBeamPartWrapper.object = oPart
            Select Case oBeamPartWrapper.WebLength
                Case 0 To 0.2
                    dULength = 0.07
                    dRadius = 0.035
                Case 0.2 To 0.4
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
        Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
        
        Set oPlatePartWrapper.object = oPart
        If oPlatePartWrapper.plateType = TBulkheadPlate Then
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
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
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
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim oDefFact As New DefinitionFactory
    
    Set IJDUserSymbolServices_InstanciateDefinition = oDefFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
    Set oDefFact = Nothing
End Function
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
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
