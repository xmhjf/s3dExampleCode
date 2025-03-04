VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\WebCutSel.cls"

'------------- Give the Selector Name and ProgId
Const m_SelectorProgid As String = "EndCutRules.WebCutSel"   ' - TODO
Const m_SelectorName As String = "EndCutRules.WebCutSel"                   ' - TODO
Const m_FamilyProgid As String = ""       ' TODO


Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    'Define questions.  !! Note !!  This default is only used if web cut is NOT created by an assembly connection.
    strError = "Setting questions."
    SetQuestionEndCutType pQH
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol"
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", strError).Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    Dim strError As String
    
    Set oBoundedObject = pSLH.InputObject(INPUT_BOUNDED)
    If (oBoundedObject Is Nothing) Then
        strError = "pSLH.InputObject(INPUT_BOUNDED) is NOTHING"
        GoTo ErrorHandler
    End If
        
    Set oBoundingObject = pSLH.InputObject(INPUT_BOUNDING)
    If (oBoundingObject Is Nothing) Then
        strError = " pSLH.InputObject(INPUT_BOUNDING) is NOTHING"
        GoTo ErrorHandler
    End If
    
    'get the end cut type
    Dim strEndCutType As String
    strEndCutType = pSLH.Answer(QUES_ENDCUTTYPE)
    
    '$$$ Free End functionality
    'Get WebCut smart occurence to determine if this is a FreeEnd WebCut
    Dim bFreeEndCut As Boolean
    Dim oWebCut As New Structdetailobjects.WebCut
    Set oWebCut.object = pSLH.SmartOccurrence
    bFreeEndCut = oWebCut.IsFreeEndCut
    
    ' get Bounded Part
    Dim oBoundedPart As Object
    Dim oBoundedPort As IJPort
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
        Set oBoundedPart = oBoundedPort.Connectable
    Else
        ' error, BoundedObject MUST be a IJPort object
        strError = "BoundedObject MUST be a IJPort object"
        GoTo ErrorHandler
    End If
    
    Dim dPortValue As Long
    Dim strPortName As String
    Dim sTypeObject As String
    Dim sObjectType As String
    Dim phelper As New Structdetailobjects.Helper
    
    'Is the profile end is a Free End
    '$$$ Free End functionality, FreeEnd WebCut
    If bFreeEndCut Then
        'Is the End Cut Type = F (free)
        Select Case strEndCutType
            Case gsF, gsFV, gsS
    
                ' get the Bounded Part's CrossSection type
                Dim sSectionType As String
                If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
                    Dim oProfile As New Structdetailobjects.ProfilePart
                    Set oProfile.object = oBoundedPart
                    sSectionType = oProfile.SectionType
                ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
                    Dim oBeam As New Structdetailobjects.BeamPart
                    Set oBeam.object = oBoundedPart
                    sSectionType = oBeam.SectionType
                Else
                    ' Bounded Part is Not Profile or Beam
                    Exit Sub
                End If
            
                'If the profile is a Flat Bar
                If sSectionType = "FB" Then
                        'Is the Profile an Edge Reinforcement
                        phelper.GetObjectTypeData oBoundedPart, _
                                                  sTypeObject, sObjectType
                        If sTypeObject = "EdgeReinforcement" Then
                            'If the Edge Reinforcement mounted On Face
                            'then treat it as a normal stiffener
                            Dim oEdgeReinforcement As New Structdetailobjects.EdgeReinforcement
                            Set oEdgeReinforcement.object = oBoundedPart
                            If Not oEdgeReinforcement.ReinforcementPosition = "OnFace" Then  'Edge Reinforcement is on edge
                                pSLH.Add "FreeEndWebCut_F1_ER"
                            Else
                                'Edge Reinforcement is on face
                                'Snipe profile endcut
                                pSLH.Add "WebCut_F1B"
                            End If
                        Else
                            'Bounded Profile is Not Edge Reinforcement
                            'Snipe profile endcut
                            pSLH.Add "WebCut_F1B"
                        End If
                        
                                                                    
                    'TR-78614 the symbol is changed

              ElseIf sSectionType = "R" Or sSectionType = "SB" Or sSectionType = "P" Or sSectionType = "sqTu" Then
                    pSLH.Add "WebCut_F1B"

                   'TR-78614 the symbol is changed
                        
                      
                Else
                    'Snipe profile endcut
                    pSLH.Add "WebCut_F1"
                End If
                
            Case gsW, gsC
                'Cut web normal with no physical connection
                pSLH.Add "FreeEndWebCut_C1Spline"
        End Select
        
        Exit Sub
    End If

    ' get Bounding Part if not a FreeEnd WebCut
    Dim oBoundingPart As Object
    Dim oBoundingPort As IJPort
    If TypeOf oBoundingObject Is IJPort Then
        Set oBoundingPort = oBoundingObject
        Set oBoundingPart = oBoundingPort.Connectable
    Else
        ' error, BoundingObject MUST be a IJPort object if not FreeEnd WebCut
        strError = "BoundingObject MUST be a IJPort object"
        GoTo ErrorHandler
    End If
    
    ' check if the bounding object is a plate or a profile
    If TypeOf oBoundingPart Is IJPlate Then
        Call WebCutNoFlange(oBoundedPart, True, strEndCutType, pSLH)
        
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oBounding As New Structdetailobjects.ProfilePart
        Set oBounding.object = oBoundingPart
            
        Select Case oBounding.SectionType
            Case "T_XType", "TSType", "BUT", "BUTL2", "I", "I_SType", "ISType"
                Call WebCutFlanged(oBoundedPart, strEndCutType, pSLH)
                
            Case "FB", "P", "R", "SB", "SqTu"
                Call WebCutNoFlange(oBoundedPart, False, strEndCutType, pSLH)
                
            Case "B", "BUTL3", "C", "C_SType", "EA", "UA"
            
                'Get the name of the port, used to determine side for the symbol
                dPortValue = phelper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
                If dPortValue = JXSEC_WEB_RIGHT Or _
                    dPortValue = JXSEC_TOP Or _
                    dPortValue = JXSEC_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                    dPortValue = JXSEC_FLANGE_TOP Or _
                    dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                    dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                    dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                    dPortValue = JXSEC_WEB_LEFT_TOP Or _
                    dPortValue = JXSEC_FLANGE_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                    dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                    dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                    Call WebCutFlanged(oBoundedPart, strEndCutType, pSLH)
                Else
                    Call WebCutNoFlange(oBoundedPart, False, strEndCutType, pSLH)
                End If
        End Select
        
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBoundingBeam As New Structdetailobjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        
        Select Case oBoundingBeam.SectionType
            Case "T_XType", "TSType", "BUT", "BUTL2", "I", "I_SType", "ISType"
                Call WebCutFlanged(oBoundedPart, strEndCutType, pSLH)
                
            Case "FB", "P", "R", "SB", "SqTu"
                Call WebCutNoFlange(oBoundedPart, False, strEndCutType, pSLH)
                
            Case "B", "BUTL3", "C", "C_SType", "EA", "UA"
            
                'Get the name of the port, used to determine side for the symbol
                dPortValue = phelper.GetBoundingProfileFace(oBoundedPort, oBoundingPort)
          
                If dPortValue = JXSEC_WEB_RIGHT Or _
                    dPortValue = JXSEC_TOP Or _
                    dPortValue = JXSEC_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT Or _
                    dPortValue = JXSEC_FLANGE_TOP Or _
                    dPortValue = JXSEC_TOP_FLANGE_TOP Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_TOP Or _
                    dPortValue = JXSEC_TOP_FLANGE_LEFT_TOP Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_TOP Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_TOP Or _
                    dPortValue = JXSEC_WEB_RIGHT_TOP Or _
                    dPortValue = JXSEC_WEB_LEFT_TOP Or _
                    dPortValue = JXSEC_FLANGE_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Or _
                    dPortValue = JXSEC_TOP_FLANGE_LEFT_BOTTOM Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_BOTTOM Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM Or _
                    dPortValue = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM Or _
                    dPortValue = JXSEC_WEB_RIGHT_BOTTOM Or _
                    dPortValue = JXSEC_WEB_LEFT_BOTTOM Then
                    Call WebCutFlanged(oBoundedPart, strEndCutType, pSLH)
                Else
                    Call WebCutNoFlange(oBoundedPart, False, strEndCutType, pSLH)
                End If
        End Select
    End If
          
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", strError).Number
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

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

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSLH As IJDSelectorLogic
    Set pSLH = New SelectorLogic
    pSLH.Representation = pRep
    SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

