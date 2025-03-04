VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenTubeByTubeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenTubeByTubeSel.cls
'
'Description :
'   Selector for Selection of Member bounded by Generic Port/Object End Cuts
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port/Object will be: (Bounding Object)
'           Plate Base/Offset/Lateral Face Port
'           Profile Base/Offset/ Lateral SubPort
'           MemberPart Base/Offset/ Lateral SubPort
'           Reference(Grid) Plane (IJPlane)
'           Point (IJPoint)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_sClassName As String = "GenTubeByTubeSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
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

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED

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

    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the appropriate standard/normal cut
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    sMsg = "Selection logic"
    
    'For tubebytube cases
    Dim bIsBoxCut As Boolean
    
    'determine if the current cut is a part of box cut
    'get the bounding object and see if it has long box item and see if it's bounding
    'is same as bounded of current object.
    
    Dim oBoundedObject As Object, oBoundingObject As Object
    If TypeOf pSelectorLogic.InputObject(INPUT_BOUNDED) Is IJPort Then
        Dim oBoundedPort As IJPort
        Set oBoundedPort = pSelectorLogic.InputObject(INPUT_BOUNDED)
        Set oBoundedObject = oBoundedPort.Connectable
    Else
        Set oBoundedObject = pSelectorLogic.InputObject(INPUT_BOUNDED)
    End If
    
    Dim oSDO_Bounded As New StructDetailObjects.MemberPart
    Set oSDO_Bounded.object = oBoundedPort.Connectable
    
    Dim oDesignParent As IJDesignParent
    Dim ppChildren As IJDObjectCollection
    
    If TypeOf pSelectorLogic.InputObject(INPUT_BOUNDING) Is IJPort Then
        Dim oBoundingPort As IJPort
        Set oBoundingPort = pSelectorLogic.InputObject(INPUT_BOUNDING)
        Set oDesignParent = oBoundingPort.Connectable
        Set oBoundingObject = oDesignParent
    Else
        Set oDesignParent = pSelectorLogic.InputObject(INPUT_BOUNDING)
    End If
    oDesignParent.GetChildren ppChildren
    
    Dim bIsBoundedTube As Boolean
    Dim bIsBoundingTube As Boolean
    
    Dim selString As String
    
    bIsBoundedTube = IsTubularMember(oBoundedPort.Connectable)
    bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
    
    If Not bIsBoundedTube And Not bIsBoundingTube Then
        Exit Sub
    End If
       
    Dim oObject As Object
    If Not ppChildren.Count = 0 Then
        For Each oObject In ppChildren
        
            Dim oOtherSmartOcc As IJSmartOccurrence
            Dim oOtherSmartitem As IJSmartItem
            Dim sOtherItemName As String
            
            If TypeOf oObject Is IJSmartOccurrence Then
                Set oOtherSmartOcc = oObject
                Set oOtherSmartitem = oOtherSmartOcc.SmartItemObject
                sOtherItemName = oOtherSmartitem.Name
            End If
            
            If InStr(1, sOtherItemName, "LongBox", vbTextCompare) > 0 Then
                If TypeOf oObject Is IJStructAssemblyConnection Then
                    Dim oACParent As IJDesignParent
                    Set oACParent = oObject
                    Dim oChildren As IJDObjectCollection
                    oACParent.GetChildren oChildren
                    Dim oFeatureObject As Object
                    For Each oFeatureObject In oChildren
                        If TypeOf oFeatureObject Is IJStructFeature Then
                            Dim oFeature As GSCADSDCreateModifyUtilities.IJStructFeature
                            Set oFeature = oFeatureObject
                            If oFeature.get_StructFeatureType = SF_WebCut Then
                                Dim oOtherWebCut As New StructDetailObjects.WebCut
                                Set oOtherWebCut.object = oFeature
                                If oOtherWebCut.Bounding Is oBoundedObject Then
                                    bIsBoxCut = True
                                Exit For
                                End If
                            End If
                        End If
                    Next oFeatureObject
                    If bIsBoxCut = True Then Exit For
                End If
            End If
        Next oObject
    End If
    
    If bIsBoxCut Then
        pSelectorLogic.Add "M_Web_ShortBox_Tube"
    Else
        pSelectorLogic.Add "Generic_tube_01"
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
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
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
    pIH.definition = pSelector
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
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CMSelector
' Description :
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


