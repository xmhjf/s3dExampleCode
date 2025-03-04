VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEndCutSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMFreeEndCutRules
'  File        : FreeEndCutSel.cls
'
'  Description :
'       It selects the list of available Items or SmartClasses in the context of
'       the smart occurrence.
'
'       - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'       - Questions are defined to parameterize the selection
'
'
'  Author      : Alligators
'
'  History     :
'    6/Mar/2012 - svsmylav
'       CR-CP-174918: 'EndCutType' answer is set to 'Welded' for 'Extend' Mfg. and convex knuckle case.
'       CR-CP-273575:  Profile Connection type is used to set 'EndCutType' answer
'*********************************************************************************************

Const m_SelectorProgid As String = CUSTOMERID + "FreeEndCutRules.FreeEndCutSel"
Const m_SelectorName As String = CUSTOMERID + "FreeEndCutRules.FreeEndCutSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "FreeEndCutRules\FreeEndCutSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput FE_INPUT_BOUNDED
    pIH.SetInput FE_INPUT_BOUNDING
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    Dim strError As String
    strError = "Setting questions."
    
    'Define questions
    pQH.SetQuestion "EndCutType", "Welded", "EndCutTypeCodeList", "SetEndCutType", CUSTOMERID + "FreeEndCutRules.FreeEndCutSel"
    pQH.SetQuestion "ChamferType", "None", , "ChamferQuestionOnER", CUSTOMERID + "FreeEndCutRules.FreeEndCutSel"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    Dim strError As String
    
    On Error GoTo ErrorHandler
    strError = "Querying Free End Cut type."
    
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pSL.SmartOccurrence
    
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject

    If TypeOf oBoundingObject Is IJPort Then
'    If oBoundedObject Is oBoundingObject Then
        ' Bounded Free End Cut is NOT bounded by Ref. Plane
        Dim oPort As IJPort
        Set oPort = oBoundedObject
        
        If TypeOf oPort.Connectable Is ISPSMemberPartPrismatic Then
            'If it is member use following smart item because this handles Free EndCut migration if the member splits
            pSL.Add "MbrFreeEndCut_Default"
        Else
             pSL.Add "FreeEndCut_Default"
        End If
        
    Else
        ' Bounded Free End Cut is bounded by Ref. Plane
        pSL.Add "FreeEndCut_Default"
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic").Number
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
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
Public Sub SetEndCutType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object)
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
    
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    Set oFreeEndCut = pSL.SmartOccurrence
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
        
    If TypeOf oBoundedObject Is IJPort Then
        
        Dim oTempPort As IJPort
        Set oTempPort = oBoundedObject
        
        Dim oBoundedPart As Object
        Set oBoundedPart = oTempPort.Connectable
        
        If TypeOf oBoundedPart Is IJEndCutAttributes Then
        
            Dim oEndCutAttributes As IJEndCutAttributes
            Set oEndCutAttributes = oBoundedPart
        
            Dim oStructPort As IJStructPort
            Set oStructPort = oBoundedObject
            
            Dim lCodeListValue As Long
                
            'Get Stiffener connection type
            If oStructPort.ContextID = CTX_BASE Then
                lCodeListValue = oEndCutAttributes.StartConnectionType
            Else    'If oStructPort.ContextID = CTX_OFFSET Then
                lCodeListValue = oEndCutAttributes.EndConnectionType
            End If
                        
            'Zero indicates "By Rule"
            If lCodeListValue > 0 Then
            
                Dim oSD_Helper As StructDetailObjects.Helper
                Set oSD_Helper = New StructDetailObjects.Helper
                
                Dim sLongDesc As String
                Dim sShotDesc As String

                oSD_Helper.GetCustomAttributeCodeListDescription "StiffenerConnectionType", lCodeListValue, sShotDesc, sLongDesc
                
                pSL.Answer("EndCutType") = sShotDesc
                Exit Sub
            End If
        End If
    End If
    
 'If Free EndCut is from profile pseudo knuckle, need to use
 ' plate port as bounding
     Dim bSetEndCutType As Boolean
     Dim oPort As IJPort
     
     bSetEndCutType = False 'Initialize
     'Check if this Free end cut is because of 'Extend' otion on knuckle
     If IsFreeEndCutFromExtend_ConvexKnuckle(oBoundedObject) Then
         'Profile is on Convex side, so it intersects base plate
         Set oPort = oBoundedObject
         
         'NOTE:- user needs Welded endcut for profile height less than or equal SwageHeight,
         'otherwise welded-upto-Swageheight and from there a Sniped webcut
         'to get SwageHeight need an API from detailing.
         Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
         Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
         Set oSDO_ProfilePart.object = oPort.Connectable
         Dim dSwageHeight As Double
         
         'Assumption:- Swage is considered to be high enough to have a completely welded webcut
         dSwageHeight = oSDO_ProfilePart.Height 'this value needs to be obtained from new API call.
         If LessThanOrEqualTo(oSDO_ProfilePart.Height, dSwageHeight) Then
             pSL.Answer("EndCutType") = "Welded" 'EndCut needs PC, to be welded to plate
             bSetEndCutType = True
         End If
     End If
     If bSetEndCutType = False Then
         pSL.Answer("EndCutType") = "Snip"
     End If
 
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SetEndCutType").Number
End Sub


Public Sub ChamferQuestionOnER(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
    ByRef ppArgument As Object)
    'This procedure is invoked before the SelectorLogic procedure in the Selector Class
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    Dim oInput As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oInput = pInput
    Set oSymbolDefinition = oInput.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = oSymbolDefinition.IJDRepresentations(1)
    pSL.Answer("ChamferType") = "None"
        
End Sub

Public Function ChamferSelectorOnER(oFreeEndCut_Object As Object) As String
    On Error GoTo ErrorHandler

    Dim sChamferType As String
    Dim sSectionType As String
    
    Dim dThicknessDiff As Double
    
    sChamferType = "None"
    ChamferSelectorOnER = sChamferType
    
    If oFreeEndCut_Object Is Nothing Then
     Exit Function
     
    ElseIf TypeOf oFreeEndCut_Object Is IJFreeEndCut Then
        
        Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
        Set oFreeEndCut = oFreeEndCut_Object
        
        Dim oBoundingObject As Object
        Dim oBoundedObject As Object
        oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
        
        If TypeOf oBoundedObject Is IJPort Then
            Dim oPort As IJPort
            Set oPort = oBoundedObject
        Else
            Exit Function
        End If
        'If the connected object is an Edge Reinforcement then returns the Chamfer type
        'based on the selection otherwise returns the default value "None".
        If TypeOf oPort.Connectable Is IJProfileER Then
            Dim oBoundedPart As IJProfileER
            Set oBoundedPart = oPort.Connectable
            
            Dim oSDO_ProfilePart As New StructDetailObjects.ProfilePart
            Set oSDO_ProfilePart.object = oPort.Connectable
            
            sSectionType = oSDO_ProfilePart.sectionType
            
            If (sSectionType = "FB") Then
                Dim oERPort As IJPort
                Set oERPort = oSDO_ProfilePart.MountingFacePort
                
                Dim oSDO_StructPort As IJStructPort
                Set oSDO_StructPort = oERPort
                
                Dim oWeldProfileSecond As StructMoldedOrientation
                Dim bFlag As Boolean
                oSDO_ProfilePart.Get_SecondOrientation oWeldProfileSecond, bFlag
                'Get the Profile secondary orientation and returns Chamfer type based on Mounting port.
                If oWeldProfileSecond = ER_FromPlate Then
                    If oSDO_StructPort.OperatorID = JXSEC_WEB_RIGHT Then
                        sChamferType = "Obj1Base"
                    ElseIf oSDO_StructPort.OperatorID = JXSEC_WEB_LEFT Then
                        sChamferType = "Obj1Offset"
                    Else
                        sChamferType = "None"
                    End If
                ElseIf oWeldProfileSecond = ER_TowardPlate Then
                    If oSDO_StructPort.OperatorID = JXSEC_WEB_RIGHT Then
                        sChamferType = "Obj1Offset"
                    ElseIf oSDO_StructPort.OperatorID = JXSEC_WEB_LEFT Then
                        sChamferType = "Obj1Base"
                    Else
                        sChamferType = "None"
                    End If
                Else
                    'None For any other orientations.
                End If
            Else
            'Presently not handling other type of cross sections.
            End If
        Else
        'Presently handling for Edge Reinforcements only.
        End If
    End If
        'Return the chamfer type that is required.
    ChamferSelectorOnER = sChamferType
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ChamferSelectorOnER").Number

End Function


