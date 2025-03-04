VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BearingPlateParm_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2014 Intergraph Corporation. All rights reserved.
'
'File : BearingPlateParm_4.cls
'
'Author : Alligators
'
'Description :
'   Control Parameters of the Member Items for the MbrBearingEndSel Selector Class
'
'
'History:
'
'********************************************************************
Option Explicit

Const m_sClassName As String = "BearingPlateParm_4"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

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

Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"

    pOH.SetOutput "Offset3"

  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
        
    sMsg = "Defining ParameterRule Outputs"
    
    Dim oAttributes             As IJDAttributes
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Set oSmartOccurrence = pPRL.SmartOccurrence
    Set oAttributes = oSmartOccurrence
    Dim dNoseOffset_Test As Double
    dNoseOffset_Test = oAttributes.CollectionOfAttributes("IJUANoseOffset").Item("NoseOffsetDriving").value
         
    Dim dOffset3 As Double
    Dim sACName As String
    Dim oACObject As Object
    AssemblyConnection_SmartItemName pPRL.SmartOccurrence, sACName, oACObject

    Dim oAppConn As IJAppConnection
    Set oAppConn = oACObject
    
    Dim oBoundingPort As IJPort
    GetAssemblyConnectionInputs oAppConn, Nothing, oBoundingPort
        
        
    ' ------------------------------------
    ' Get information about the connection
    ' ------------------------------------
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    Dim bPenetratesWeb As Boolean
    
    InitMemberConnectionData oACObject, oBoundedData, oBoundingData, lStatus, sMsg
    
    Dim oBoundedAxis_Vector As IJDVector
    Set oBoundedAxis_Vector = New dVector
    If oBoundedData.ePortId = SPSMemberAxisStart Then
        oBoundedAxis_Vector.Set oBoundedData.Matrix.IndexValue(0), _
                                oBoundedData.Matrix.IndexValue(1), _
                                oBoundedData.Matrix.IndexValue(2)
    ElseIf oBoundedData.ePortId = SPSMemberAxisEnd Then
        oBoundedAxis_Vector.Set -oBoundedData.Matrix.IndexValue(0), _
                                -oBoundedData.Matrix.IndexValue(1), _
                                -oBoundedData.Matrix.IndexValue(2)
    Else
        Exit Sub
    End If
    oBoundedAxis_Vector.Length = 1
    Dim oBoundedLocation As IJDPosition
    Set oBoundedLocation = New DPosition
    
    oBoundedLocation.Set oBoundedData.Matrix.IndexValue(12), _
                            oBoundedData.Matrix.IndexValue(13), _
                                oBoundedData.Matrix.IndexValue(14)
        
    Dim oSurface  As IJSurfaceBody
    
    If TypeOf oBoundingPort.Geometry Is IJSurfaceBody Then
    
        Set oSurface = oBoundingPort.Geometry
        Dim oPortNormal As New dVector
        
        oSurface.GetNormalFromPosition oBoundedLocation, oPortNormal
        oPortNormal.Length = 1
        
        'Get Sine
        Dim oSineVec As IJDVector
        Set oSineVec = oBoundedAxis_Vector.Cross(oPortNormal)
        
        'Get Cosine
        Dim dCosine As Double
        dCosine = Abs(oBoundedAxis_Vector.Dot(oPortNormal))
        If dCosine < 0.2 Then
            sMsg = "If the inclination is more than 78 degrees then bearing plate fails"
            GoTo ErrorHandler
        End If
        
        Dim eResultType As eResultType
        Dim dLadderInclination As Double
        dLadderInclination = GetPI / 2 - Atn(oSineVec.Length / dCosine)
        dOffset3 = dNoseOffset_Test / Tan(dLadderInclination)
    Else
        dOffset3 = 0.0001  '0.1mm
    End If
        
    pPRL.Add "Offset3", dOffset3
    
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
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
End Function

' ********************************************************************************************
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
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pDFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

' ********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

' ********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
  Dim pPRL As IJDParameterLogic
  Set pPRL = New ParameterLogic
  pPRL.Representation = pRep
  ParameterRuleLogic pPRL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
