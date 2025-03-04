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
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMEndCutRules
'  File        : WebCutDefCM1.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    22/Feb/2012 - svsmylav
'           CR-174918: Updated 'SetFeatureType' method so that CF is Not created
'           by default for WebCut which is created from 'Extend' Mfg. method option.
'*********************************************************************************************
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\WebCutSel.cls"

'------------- Give the Selector Name and ProgId
Const m_SelectorProgid As String = CUSTOMERID + "EndCutRules.WebCutSel"   ' - TODO
Const m_SelectorName As String = CUSTOMERID + "EndCutRules.WebCutSel"                   ' - TODO
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
    
    'this question comes from the assembly conn rules
    pQH.SetQuestion "EndCutType", "Welded", "EndCutTypeCodeList"
    
    'these questions are set at the web cut level
    pQH.SetQuestion gsApplyTreatment, "No", "BooleanCol", "ApplyTreatmentCM", CUSTOMERID + "EndCutRules.WebCutSel"
    pQH.SetQuestion "Theta", "60", , "SetTheta", CUSTOMERID + "EndCutRules.WebCutSel"
    pQH.SetQuestion "FeatureType", "Radius", "FeatureTypeCodeList", "SetFeatureType", CUSTOMERID + "EndCutRules.WebCutSel"
    
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
    Dim oWebCut As New StructDetailObjects.WebCut
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
    Dim phelper As New StructDetailObjects.Helper
    
    'get the answer for theta
    Dim sTheta As String
    sTheta = pSLH.SelectorAnswer(CUSTOMERID + "EndCutRules.WebCutSel", "Theta")
    Dim dTheta As Double
    dTheta = Val(sTheta)
    dTheta = DegToRad(dTheta)

    
    'Is the profile end is a Free End
    '$$$ Free End functionality, FreeEnd WebCut
    If bFreeEndCut Then
        'Is the End Cut Type = F (free)
        Select Case strEndCutType
            Case "Snip"
    
                ' get the Bounded Part's CrossSection type
                Dim sSectionType As String
                If phelper.ObjectType(oBoundedPart) = SDOBJECT_STIFFENER Then
                    Dim oProfile As New StructDetailObjects.ProfilePart
                    Set oProfile.object = oBoundedPart
                    sSectionType = oProfile.SectionType
                ElseIf phelper.ObjectType(oBoundedPart) = SDOBJECT_BEAM Then
                    Dim oBeam As New StructDetailObjects.BeamPart
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
                            Dim oEdgeReinforcement As New StructDetailObjects.EdgeReinforcement
                            Set oEdgeReinforcement.object = oBoundedPart
                            If Not oEdgeReinforcement.ReinforcementPosition = "OnFace" Then  'Edge Reinforcement is on edge
                                pSLH.Add "FreeEndWebCut_F1_ER"
                            Else
                                'Edge Reinforcement is on face
                                'Snipe profile endcut
                                pSLH.Add "Snip"
                            End If
                        Else
                            'Bounded Profile is Not Edge Reinforcement
                            'Snipe profile endcut
                            'check theta angle to get type
                            If dTheta >= DegToRad(65) And dTheta < DegToRad(115) Then
                                pSLH.Add "SnipNoNose"
                            ElseIf dTheta >= DegToRad(115) Then
                                'the value if out of range, so set a default
                                pSLH.Add "SnipNoNoseRadius"
                            Else
                                'this is the default case, and theta is less than 65 degrees
                                'add the snip by default
                                pSLH.Add "Snip"
                            End If
                                            
                        End If
                                                                                            

                ElseIf sSectionType = "R" Or sSectionType = "SB" Or sSectionType = "P" Or sSectionType = "sqTu" Or sSectionType = "HalfR" Then
                    pSLH.Add "Snip"


                Else
                
                    'check theta angle to get type
                    If dTheta >= DegToRad(65) And dTheta < DegToRad(115) Then
                        pSLH.Add "SnipNoNose"
                    ElseIf dTheta >= DegToRad(115) Then
                        'the value if out of range, so set a default
                        pSLH.Add "SnipNoNoseRadius"
                    Else
                        'this is the default case, and theta is less than 65 degrees
                        'add the snip by default
                        If sSectionType = gsB Then
                            pSLH.Add "Snip"
                        Else
                            pSLH.Add "SnipWithFlange"
                        End If
                    End If
                
                End If
                
                pSLH.Add "SnipBeyondBoundary"
                
            Case "Welded", "Cutback", "Bracketed"
                'Cut web normal with no physical connection
                pSLH.Add "StraightSnip"
                
             Case "Clip"
                pSLH.Add "Clip"
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
        Dim oBounding As New StructDetailObjects.ProfilePart
        Set oBounding.object = oBoundingPart
            
        Select Case oBounding.SectionType
            Case "T_XType", "TSType", "BUT", "BUTL2", "I", "I_SType", "ISType", "H"
                Call WebCutFlanged(oBoundedPart, strEndCutType, pSLH)
                
            Case "FB", "P", "R", "SB", "SqTu"
                Call WebCutNoFlange(oBoundedPart, False, strEndCutType, pSLH)
                
            Case gsB, "BUTL3", "C", "C_SType", "EA", "UA"
            
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
        Dim oBoundingBeam As New StructDetailObjects.BeamPart
        Set oBoundingBeam.object = oBoundingPart
        
        Select Case oBoundingBeam.SectionType
            Case "T_XType", "TSType", "BUT", "BUTL2", "I", "I_SType", "ISType", "H"
                Call WebCutFlanged(oBoundedPart, strEndCutType, pSLH)
                
            Case "FB", "P", "R", "SB", "SqTu"
                Call WebCutNoFlange(oBoundedPart, False, strEndCutType, pSLH)
                
            Case gsB, "BUTL3", "C", "C_SType", "EA", "UA"
            
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
    pIH.definition = pSelector
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
Public Sub ApplyTreatmentCM(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer("ApplyTreatment") = "No"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ApplyTreatment").Number
End Sub


Public Sub SetTheta(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    pSL.Answer("Theta") = "60"
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetTheta").Number
End Sub
Public Sub SetFeatureType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    On Error GoTo ErrorHandler
    
    'From the QA related input get the current Symbol Def.
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    'Get EndCutType to drive feature placement
    Dim sEndCutType As String
    sEndCutType = pSL.SelectorAnswer(CUSTOMERID + "EndCutRules.WebCutSel", "EndCutType")
    
    Dim bMutuallybounded As Boolean
    Dim oAC As New StructDetailObjects.AssemblyConn
    Dim pWebCut As New StructDetailObjects.WebCut
    Set pWebCut.object = pSL.SmartOccurrence
    
    Dim oChild As IJDesignChild
    Set oChild = pWebCut.object
    
    If TypeOf oChild.GetParent Is IJAssemblyConnection Then
    
        Set oAC.object = oChild.GetParent
        bMutuallybounded = IsMutualBound(oAC)
        If bMutuallybounded Then
            pSL.Answer("FeatureType") = "None"
            Exit Sub
        End If
    
    End If
    
    Dim oWebCut As New StructDetailObjects.WebCut
    Set oWebCut.object = pSL.SmartOccurrence
    Dim phelper As New StructDetailObjects.Helper
    Dim dPortValue As Long
    Dim bExcluded As Boolean
    
    bExcluded = False
    If TypeOf oWebCut.Bounded Is IJBeam Then
        bExcluded = True
    ElseIf TypeOf oWebCut.Bounded Is IJStiffener Then
        Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
        
        Set oProfilePartWrapper.object = oWebCut.Bounded
        If oProfilePartWrapper.SectionType = "H" Or _
           oProfilePartWrapper.SectionType = "I" Or _
           oProfilePartWrapper.SectionType = "HalfR" Or _
           oProfilePartWrapper.SectionType = "R" Then
            bExcluded = True
        End If
        Set oProfilePartWrapper = Nothing
        
        'Check for Convex profile knuckle and if so set default answer to 'None'
        Dim oProfileKnuckle As Object
        Dim iKnuckleMfgMethod As Integer
        Dim bIsConvex As Boolean
    
        bIsConvex = False ' Initialize
        GetProfileKnuckleType oWebCut.object, iKnuckleMfgMethod, oProfileKnuckle, bIsConvex
        If iKnuckleMfgMethod = pkmmSplitAndExtend And bIsConvex Then
            pSL.Answer("FeatureType") = "None"
            Exit Sub
        End If
    End If
    
    If sEndCutType = "Welded" And bExcluded = False Then
       If TypeOf oWebCut.Bounded Is IJProfileER Then
            Dim oProfileER As IJProfileER
            Set oProfileER = oWebCut.Bounded
                        
            If oProfileER.Position = ER_OnEdgeCentered Or oProfileER.Position = ER_OnEdgeOffset Then
               pSL.Answer("FeatureType") = "None"
               Exit Sub
            ElseIf oProfileER.Position = ER_OnFace Then
               pSL.Answer("FeatureType") = "Radius"
               Exit Sub
            Else
               pSL.Answer("FeatureType") = "None"
            End If
        ElseIf TypeOf oWebCut.Bounding Is IJPlate Then
            pSL.Answer("FeatureType") = "Radius"
        Else
            'check the connected port
            dPortValue = phelper.GetBoundingProfileFace(oWebCut.BoundedPort, oWebCut.BoundingPort)
            If dPortValue = JXSEC_TOP Or _
                dPortValue = JXSEC_BOTTOM Then
                pSL.Answer("FeatureType") = "Radius"
            ElseIf dPortValue = JXSEC_WEB_RIGHT Or _
                    dPortValue = JXSEC_WEB_LEFT Then
                'there seems to be a problem with the feature placement against the web.
                'default to "none" for now.
                pSL.Answer("FeatureType") = "None"
'                pSL.Answer("FeatureType") = "Snipe"
            Else
                pSL.Answer("FeatureType") = "None"
            End If
        End If
    Else
        pSL.Answer("FeatureType") = "None"
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetFeatureType").Number
End Sub

