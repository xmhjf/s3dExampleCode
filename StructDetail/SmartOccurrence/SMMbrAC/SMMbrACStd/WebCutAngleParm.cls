VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutAngleParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2014, Intergraph Corporation.  All rights reserved.
'
'  File        : WebCutOffestParm.cls
'
'  Description : Parmameter rule for end cuts that need to define the cutting depth and Extension
'
'  Author      : Alligators
'
'*********************************************************************************************

Const m_sClassName As String = "WebCutAngleParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    pOH.SetOutput "CutDepth"
    pOH.SetOutput "TopAngle"
    pOH.SetOutput "BottomAngle"
    pOH.SetOutput "TopOffset"
    pOH.SetOutput "BottomOffset"
    pOH.SetOutput "MinimumGap"
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(oPRL As IJDParameterLogic)

    Const METHOD = m_DefinitionProgid & "::ParameterRuleLogic"
    
    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    Dim dHalfNose As Double
    Dim dChmfer As Double
    Dim dTopChamfer As Double
    Dim dBtmChamfer As Double
    Dim dBdgDepth As Double
    Dim dBddDepth As Double
    Dim dHalfBdgDepth As Double
    Dim dCuttingDepth As Double
    
    Dim dTopSnipeAngle As Double
    Dim dBottomSnipeAngle As Double
    Dim dTopSnipeLength As Double
    Dim dBottomSnipeLength As Double
    Dim dFilletGap As Double
    

    Dim oPortBounding As IJPort
    Dim oPortBounded As IJPort
    Dim oBoundingObject As Object
    Dim oBoundedObject As Object
    Dim OSDO_BdngPart As New StructDetailObjects.MemberPart
    Dim OSDO_BdedPart As New StructDetailObjects.MemberPart
    Dim bIsBoundingTube As Boolean
    Dim ModValueCham As Double
    Dim dTopModValue As Double
    Dim dBtmModValue As Double
    Dim vAttributeValue As Variant

    Set oPortBounding = oPRL.InputObject(INPUT_BOUNDING)
    Set oPortBounded = oPRL.InputObject(INPUT_BOUNDED)
    Set oBoundingObject = oPortBounding.Connectable
    Set oBoundedObject = oPortBounded.Connectable
    bIsBoundingTube = IsTubularMember(oPortBounding.Connectable)
    Dim dDistFromTop As Double
    Dim dDistFromBtm As Double
    Dim dBndgThickness As Double
    Dim dBddThickness As Double
    'based on these values, decision is taken whether it is symmetrically or non-symmetrical case.
    'distance from bounding to bounded top and bounded bottom are calculated
    
    dDistFromTop = GetDistanceFromBounding(oBoundingObject, oBoundedObject, JXSEC_TOP)
    dDistFromBtm = GetDistanceFromBounding(oBoundingObject, oBoundedObject, JXSEC_BOTTOM)
    Set OSDO_BdngPart.object = oBoundingObject
    dBdgDepth = OSDO_BdngPart.Height
    dBndgThickness = OSDO_BdngPart.webThickness
    
    Set OSDO_BdedPart.object = oBoundedObject
    If IsWebPenetrated(oPortBounding, oPortBounded) Then
        dBddDepth = OSDO_BdedPart.Height
        dBddThickness = OSDO_BdedPart.webThickness
    Else
        dBddDepth = GetWidthFromStructDetailObjects(oBoundedObject)
        dBddThickness = OSDO_BdedPart.flangeThickness
    End If
    If Equal(dBndgThickness, 0) Then
        If Equal(dBddThickness, 0) Then
            dFilletGap = GetWidthFromStructDetailObjects(oBoundedObject)
        Else
            dFilletGap = dBddThickness
        End If
    ElseIf Equal(dBddThickness, 0) Then
        dFilletGap = dBndgThickness
    ElseIf LessThan(dBddThickness, dBndgThickness) Then
        dFilletGap = dBddThickness
    Else
        dFilletGap = dBndgThickness
    End If
    dHalfBdgDepth = dBdgDepth / 2
    dHalfNose = Sqr(Abs(((dHalfBdgDepth * dHalfBdgDepth) - ((dHalfBdgDepth - dFilletGap) * (dHalfBdgDepth - dFilletGap)))))
    dHalfNose = dHalfNose + 0.005
    
    If Equal(dDistFromTop, dDistFromBtm) Then
        dChmfer = (dBddDepth / 2) - dHalfNose
        dChmfer = dChmfer * 1000
        ModValueCham = CInt(dChmfer) Mod 5
        ModValueCham = ModValueCham / 1000
        dChmfer = CInt(dChmfer) / 1000
        If bIsBoundingTube Then
            If LessThan(dChmfer, 0.005) Or LessThan(dChmfer, OSDO_BdedPart.flangeThickness) Then
                dTopSnipeLength = 0.01
                dBottomSnipeLength = 0.01
                dTopSnipeAngle = Atn(1) * 2
                dBottomSnipeAngle = Atn(1) * 2
            Else
                If ModValueCham = 0 Then  'divided by 5
                    dTopSnipeLength = dChmfer
                    dBottomSnipeLength = dChmfer
                Else
                    dTopSnipeLength = dChmfer - ModValueCham 'roundoff to 5 multiples
                    dBottomSnipeLength = dChmfer - ModValueCham
                End If
                dTopSnipeAngle = Atn(1)
                dBottomSnipeAngle = Atn(1)
            End If
        End If
    Else
        dTopChamfer = ((dBddDepth - dHalfNose) - dHalfBdgDepth) - dDistFromBtm
        dBtmChamfer = ((dBddDepth - dHalfNose) - dHalfBdgDepth) - dDistFromTop
        
        dTopChamfer = dTopChamfer * 1000
        dBtmChamfer = dBtmChamfer * 1000
        
        dTopModValue = CInt(dTopChamfer) Mod 5
        dTopModValue = dTopModValue / 1000
        dBtmModValue = CInt(dBtmChamfer) Mod 5
        dBtmModValue = dBtmModValue / 1000
        
        dTopChamfer = dTopChamfer / 1000
        dBtmChamfer = dBtmChamfer / 1000
        
        If bIsBoundingTube Then
            If ((LessThan(dBtmChamfer, 0.005) Or LessThan(dBtmChamfer, OSDO_BdedPart.flangeThickness)) And _
                Not (LessThan(dTopChamfer, 0.005) Or LessThan(dTopChamfer, OSDO_BdedPart.flangeThickness))) Then
                dBottomSnipeLength = 0.01
                dBottomSnipeAngle = Atn(1) * 2
                If dTopModValue = 0 Then  'divided by 5
                    dTopSnipeLength = dTopChamfer
                Else
                    dTopSnipeLength = dTopChamfer - dTopModValue 'roundoff to 5 multiples
                End If
                dTopSnipeAngle = Atn(1)
                
            ElseIf ((LessThan(dTopChamfer, 0.005) Or LessThan(dTopChamfer, OSDO_BdedPart.flangeThickness)) And _
                Not (LessThan(dBtmChamfer, 0.005) Or LessThan(dBtmChamfer, OSDO_BdedPart.flangeThickness))) Then
                dTopSnipeLength = 0.01
                dTopSnipeAngle = Atn(1) * 2
                If dBtmModValue = 0 Then  'divided by 5
                    dBottomSnipeLength = dBtmChamfer
                Else
                    dBottomSnipeLength = dBtmChamfer - dBtmModValue  'roundoff to 5 multiples
                End If
                dBottomSnipeAngle = Atn(1)
                
            Else
                If Equal(dTopModValue, 0) Then   'divided by 5
                    dTopSnipeLength = dTopChamfer
                Else
                    dTopSnipeLength = dTopChamfer - dTopModValue 'roundoff to 5 multiples
                End If
                
                If Equal(dBtmModValue, 0) Then   'divided by 5
                    dBottomSnipeLength = dBtmChamfer
                Else
                    dBottomSnipeLength = dBtmChamfer - dBtmModValue  'roundoff to 5 multiples
                End If
                dTopSnipeAngle = Atn(1)
                dBottomSnipeAngle = Atn(1)
            End If
        End If
        
    End If
    
    ' setting cutdepth as( (width of the section) + (tolerance) )
    
    dCuttingDepth = GetWidthFromStructDetailObjects(oBoundedObject)
    dCuttingDepth = dCuttingDepth + 0.001
    
    oPRL.Add "CutDepth", dCuttingDepth
    oPRL.Add "MinimumGap", dFilletGap
    oPRL.Add "TopAngle", dTopSnipeAngle
    oPRL.Add "BottomAngle", dBottomSnipeAngle
    oPRL.Add "TopOffset", dTopSnipeLength
    oPRL.Add "BottomOffset", dBottomSnipeLength
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
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
  IJDUserSymbolServices_GetDefinitionName = m_DefinitionProgid
End Function

' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
  On Error Resume Next
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  pPR.IJDInputs.RemoveAllInput
  pPR.IJDRepresentations.RemoveAllRepresentation

  Dim pDFact As New DefinitionFactory
  Dim pIH As IJDInputsHelper
  Dim pOH As IJDOutputsHelper
  
  pDFact.InitAbstractParameterRule pPR
  
  Set pIH = New InputHelper
  pIH.definition = pPR
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
  Set pOH = New OutputHelper
  pOH.Representation = pPR.IJDRepresentations.Item(1)
  pOH.InitAs m_FamilyProgid
  ParameterRuleOutputs pOH
End Sub

' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim oPRL As IJDParameterLogic
  Set oPRL = New ParameterLogic
  oPRL.Representation = pRep
  ParameterRuleLogic oPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************




