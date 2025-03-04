VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenMBWebSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2014, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenMBWebSel.cls
'
'*********************************************************************************************

Const m_sClassName As String = "GenMBWebSel"
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
    Dim oEndCutObject As Object
    
    Dim oBoundedPort As Object
    Dim oBoundingPort As Object

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    ' Get the WebCut Inputs from the StructFeature object
    Set oEndCutObject = pSelectorLogic.SmartOccurrence
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject

    sMsg = "Set the Bounded object"
    Set oBoundedPort = oSDO_WebCut.BoundedPort
    If (oBoundedPort Is Nothing) Then
        sMsg = "Bounded Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    sMsg = "Set the Bounding object"
    Set oBoundingPort = oSDO_WebCut.BoundingPort
    If (oBoundingPort Is Nothing) Then
        sMsg = "Bounding Port Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    Dim oBoundingPortColl As New Collection
    Dim dAngleColl As New Collection
    Dim dLengthColl As New Collection
    Dim pPosColl As New Collection
    
    Dim oMappedPortsColl As Collection
    Dim oPos1 As IJDPosition
    Dim oPos2 As IJDPosition
        
    GetMultiBoundingEdgeMap oEndCutObject, WebCut, oMappedPortsColl, dAngleColl, pPosColl
        
    If oMappedPortsColl Is Nothing Then
        Exit Sub
    End If
                
    Dim i As Integer
    For i = 2 To pPosColl.Count
        Set oPos1 = pPosColl.Item(i)
        Set oPos2 = pPosColl.Item(i - 1)
        dLengthColl.Add Round(oPos1.DistPt(oPos2), 6)
    Next

    '*********************************************************************************************
    '******************************Get Angles, Lenghths using Mapping ****************************
    '*********************************************************************************************
    Dim bBtmLeft As Boolean
    Dim bBtmRight As Boolean
    Dim bTopLeft As Boolean
    Dim bTopRight As Boolean
        
    Dim dAngle195 As Double
    Dim dAngle165 As Double
    Dim dPI As Double
    
    dPI = GetPI()
    dAngle195 = degreeToRadian(195)
    dAngle165 = degreeToRadian(165)
    
    Select Case oMappedPortsColl.Count
        Case 1:
            pSelectorLogic.Add "W_1S_1PC_01"
    
        Case 2: 'Two Bounding Surfaces
            
            If Equal(dAngleColl.Item(1), dPI) Then
                pSelectorLogic.Add "W_2S_2PC_00" ' Coplanar case
            ElseIf dAngleColl.Item(1) >= degreeToRadian(135) Then  ' more than 135 degrees
                pSelectorLogic.Add "W_2S_2PC_01" 'Smart item without CF
                pSelectorLogic.Add "W_2S_2PC_1CF_01" 'Smart item with CF
            Else
                pSelectorLogic.Add "W_2S_2PC_1CF_01"
                pSelectorLogic.Add "W_2S_2PC_01"
            End If
        
        Case 3: 'Three Bounding Surfaces
            'Symbol with middle facing down
            If GreaterThanOrEqualTo(dAngleColl.Item(1), dAngle195) And _
                LessThanOrEqualTo(dAngleColl.Item(2), dAngle165) Then
                
                If GreaterThan(dLengthColl.Item(2), dLengthColl.Item(1)) Then
                
                    CrossSection_Flanges oBoundedPort.Connectable, bTopLeft, bBtmLeft, bTopRight, bBtmRight
                    
                    If bTopLeft Or bTopRight Then
                        pSelectorLogic.Add "W_3S_2PC_2CF_01_B" 'PC2+PC3+C12+C23 (CF12 removes all of port 1)
                    Else
                        pSelectorLogic.Add "W_3S_3PC_1CF_01_B"
                    End If
                    pSelectorLogic.Add "W_3S_3PC_01_B" 'PC1+PC2+PC3
                    
                Else
                    pSelectorLogic.Add "W_3S_2PC_1CF_01_B" 'PC1+PC3+CF23 (feature removes all of port 2)
                    pSelectorLogic.Add "W_3S_3PC_1CF_01_B" 'PC1+PC2+PC3+C23 (feature does not remove all of port 2)"
                    pSelectorLogic.Add "W_3S_3PC_01_B" 'PC1+PC2+PC3
                End If
                
            'Symbol with middle facing up
            ElseIf LessThanOrEqualTo(dAngleColl.Item(1), dAngle165) And _
                GreaterThanOrEqualTo(dAngleColl.Item(2), dAngle195) Then
                
                If GreaterThan(dLengthColl.Item(2), dLengthColl.Item(3)) Then
                
                    CrossSection_Flanges oBoundedPort.Connectable, bTopLeft, bBtmLeft, bTopRight, bBtmRight
                    
                    If bBtmLeft Or bBtmRight Then
                        pSelectorLogic.Add "W_3S_2PC_2CF_01_A" 'PC1+PC2+C12+C23 (CF23 removes all of port 3)
                    Else
                        pSelectorLogic.Add "W_3S_3PC_1CF_01_A"
                    End If
                    
                    pSelectorLogic.Add "W_3S_3PC_01_A" 'PC1+PC2+PC3
                    
                Else
                    pSelectorLogic.Add "W_3S_2PC_1CF_01_A" 'PC1+PC3+CF12 (feature removes all of port 2)
                    pSelectorLogic.Add "W_3S_3PC_1CF_01_A" 'PC1+PC2+PC3+C12 (feature does not remove all of port 2)
                    pSelectorLogic.Add "W_3S_3PC_01_A" 'PC1+PC2+ PC3
                End If
            Else
                If Equal(dAngleColl.Item(1), dPI) And Equal(dAngleColl.Item(2), dPI) Then
                    pSelectorLogic.Add "W_3S_3PC_00" 'Coplanar cases
                Else
                    pSelectorLogic.Add "W_3S_3PC_01_A" 'PC1+PC2+ PC3
                End If
            End If
            
        Case 5: 'Five Bounding Surfaces
            
            'Symbol with middle facing down
            If GreaterThanOrEqualTo(dAngleColl.Item(1), dAngle195) And LessThanOrEqualTo(dAngleColl.Item(2), dAngle165) And _
                LessThanOrEqualTo(dAngleColl.Item(3), dAngle165) And GreaterThanOrEqualTo(dAngleColl.Item(4), dAngle195) Then
                
                If GreaterThan(dLengthColl.Item(2), dLengthColl.Item(1)) And GreaterThan(dLengthColl.Item(4), dLengthColl.Item(5)) Then
                    
                    CrossSection_Flanges oBoundedPort.Connectable, bTopLeft, bTopRight, bBtmLeft, bBtmRight
                    
                    If (bTopLeft Or bTopRight) And (bBtmLeft Or bBtmRight) Then
                        pSelectorLogic.Add "W_5S_3PC_4CF_01_B" 'PC2+PC3+PC4+C12+C23+CF34+CF45 (CF12 and CF45 removes all of port 1 and port5 respectively)
                    End If
                    
                    pSelectorLogic.Add "W_5S_5PC_01_B" 'PC1+PC2+PC3+PC4+PC5
                
                ElseIf GreaterThan(dLengthColl.Item(1), dLengthColl.Item(2)) And GreaterThan(dLengthColl.Item(5), dLengthColl.Item(4)) Then
                    pSelectorLogic.Add "W_5S_3PC_2CF_01_B" 'PC1+PC3+PC5+CF23+CF34 (CF removes all of port 2 and Port 4)
                    pSelectorLogic.Add "W_5S_5PC_2CF_01_B" 'PC1+PC2+PC3+PC4+PC5+CF23+CF34 (feature does not remove all of port 2 and port4)
                    pSelectorLogic.Add "W_5S_5PC_01_B" 'PC1+PC2+PC3+PC4+PC5
                Else
                    pSelectorLogic.Add "W_5S_5PC_01_B" 'PC1+PC2+PC3+PC4+PC5"
                End If
            
            'Symbol with middle facing down
            ElseIf LessThanOrEqualTo(dAngleColl.Item(1), dAngle165) And GreaterThanOrEqualTo(dAngleColl.Item(2), dAngle195) And _
                GreaterThanOrEqualTo(dAngleColl.Item(3), dAngle195) And LessThanOrEqualTo(dAngleColl.Item(4), dAngle165) Then
                
                If GreaterThan(dLengthColl.Item(1), dLengthColl.Item(2)) And GreaterThan(dLengthColl.Item(5), dLengthColl.Item(4)) Then
                    pSelectorLogic.Add "W_5S_3PC_2CF_01_A" 'PC1+PC3+PC5+CF12+CF45 (CF removes all of port 2 and Port 4)
                    pSelectorLogic.Add "W_5S_5PC_2CF_01_A" 'PC1+PC2+PC3+PC4+PC5+CF12+CF45 (feature does not remove all of port 2 and port4)
                    pSelectorLogic.Add "W_5S_5PC_01_A" 'PC1+PC2+PC3+PC4+PC5
                Else
                    pSelectorLogic.Add "W_5S_5PC_01_A" 'PC1+PC2+PC3+PC4+PC5
                End If
            Else
                pSelectorLogic.Add "W_5S_5PC_01_A" 'PC1+PC2+PC3+PC4+PC5
            End If
            
        End Select

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
