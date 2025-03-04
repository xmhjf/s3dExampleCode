VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootAssyConnSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : RootAssyConnSel.cls
'
'  Description :
'       Root Default Selector for selecting EndCut Assembly Connection with SPS Members as inputs
'       Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'           one Port will be SPSMemberAxisAlong (Bounding Member)
'           one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "RootAssyConnSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    ' TODO - Add your question below
    '===============================
    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the Basic Type of Member Assembly Connection here
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sSelector As String
    Dim lStatus As Long
    
    Dim bColinear As Boolean
    Dim bEndToEnd As Boolean
    Dim bRightAngle As Boolean
    
    Dim oAppConnection As IJAppConnection
    Dim oEditJDArgument As IJDEditJDArgument
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    
    sMsg = "Unknown Error"
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    Set oAppConnection = pSelectorLogic.SmartOccurrence
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' To Handle Insert Plate AC creating a stub AC - CR 219586
    '   To Figure out the AC is of Penetration AC kind need to check
    '   if one of the App Connection Port(i.e Penetrating Port) is PLATE PART's Port if yes then
    '   it is PENETRATION AC
    If Not oAppConnection Is Nothing Then
        Dim oElements_Ports As IJElements
        ' Get the Assembly Connection Ports from the IJAppConnection
        oAppConnection.enumPorts oElements_Ports
        
        ' for Member AC, require two(2) Ports, especially for Insert Plate AC
        '   one Port will be PortId type of SPSMemberAxisAlong (Penetrated Member's Port)
        '   other Port will be of Plate Port (Penetrating Part)

        Dim oConnectedPort1 As IJPort
        Dim oConnectedPort2 As IJPort
        
        If TypeOf oElements_Ports.Item(1) Is IJPort Then
            Set oConnectedPort1 = oElements_Ports.Item(1)
            If TypeOf oConnectedPort1.Connectable Is IJPlate Then
                pSelectorLogic.Add "PenetrationACSel"
                Exit Sub
            End If
        End If
        
        If TypeOf oElements_Ports.Item(2) Is IJPort Then
            Set oConnectedPort2 = oElements_Ports.Item(2)
            If TypeOf oConnectedPort2.Connectable Is IJPlate Then
                pSelectorLogic.Add "PenetrationACSel"
                Exit Sub
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    If lStatus <> 0 Then
        ' If not a supported Member By Member Case
        ' Do Not select a current Item
        Exit Sub
    End If
    
    ' if Supported/Bounded Member Part and Supporting/Bounding MemberPart are same
    ' then this is Muiltple General/Generic Supporting/Bounding objects
    ' Plate Part Edge/Face Ports, Ref. Planes, Points, IJSurfaceBody
    If oBoundedData.MemberPart Is oBoundingData.MemberPart Then
        ' Case: Multiple Generic Member End to non-MemberPart
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        If oEditJDArgument.GetCount > 0 Then
            sSelector = "GenAssyConnSel"
            pSelectorLogic.Add sSelector
        Else
            ' Case: Free Member End Cut
            sSelector = "MBR_FreeEndSel"
            pSelectorLogic.Add sSelector
        End If
        
        Exit Sub
    End If
    
    ' Check if Assembly Connection is End To End Type
    sSelector = ""
    CheckEndToEndConnection oBoundedData.MemberPart, oBoundingData.MemberPart, bEndToEnd, bColinear, bRightAngle
    If bEndToEnd Then
        ' Assembly Connection is End To End Type
        ' Check if Axis are Colinear
        If Not oBoundingData.MemberPart.IsPrismatic Then
            ' Bounding Member is NOT a MemberPartPrismatic
            ' Initial implementation will not support:
            ' Member Part Prismatic bounded by Designed Member Part
            Exit Sub
            
        ElseIf bColinear Then
            ' Case: End To End / Colinear Axis
            sSelector = "MbrSplitSel"
            pSelectorLogic.Add "MbrSplitSel"
        Else
            ' Case: End To End / Non-Colinear Axis
            sSelector = "MbrMiterSel"
            pSelectorLogic.Add sSelector
        End If
    Else
        ' Default Assembly Connection is Along Axis Type
        sSelector = "MbrAxisSel"
        pSelectorLogic.Add sSelector
    End If
    
    ' Add the Item to the current Selection
    If Len(Trim(sSelector)) < 1 Then
        ' Current Configuration is not valid for Member By Member End Cuts
        sMsg = "Could not configure correct Assembly connection"
        GoTo ErrorHandler
    End If
    
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub
  
  
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
' ** End CM **


' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'     - Following Code Should not be edited
'     - It exposes the Selector as a regular symbol definition
'
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorProgid
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.Definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object

    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
    
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
' Method      : CMSelector
' Description :
'
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)

    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
    
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
