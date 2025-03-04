VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffEndByKnuckleSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMAssyConRul
'  File        : StiffEndByKnuckleSel.cls
'
'  Description :
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
'
'  Author      : Alligators
'
'  History     :
'    12/Jan/2012 - svsmylav
'           CR-174918: Updated code to select items for profile pseudo-knuckle.
'    8/Mar/2012 - svsmylav
'           CR-174918: Used LessThanOrEqualTo method to compare dLength1, dLength2.
'    17/04/2012 - GH -CR-174918
'           Updated the Init_KnuckleEndCutType() method, to create Long/Short box cuts when Profile Knuckle Mfg method is 'SplitAndExtend' on convex side
'           If the angle is more than 150 degrees than Miter Endcut is selected as default.
'*********************************************************************************************

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.StiffEndByKnuckleSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.StiffEndByKnuckleSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffEndByKnuckleSel.cls"

Implements IJDUserSymbolServices
'

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    sError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs", sError).Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    'Define questions
    sError = "Setting questions."
    pQH.SetQuestion "EndCutType", "Welded", "EndCutTypeCodeList", "SetEndCutType", CUSTOMERID + "AssyConRul.AssyConnSelCM"
    pQH.SetQuestion "WeldPart", "First"
    ' From the Knuckle Point used as the Splitter
    ' Determine if this is defaulted to a "Miter" or "Boxed" End Cut
    '   default is "Boxed" : if 88 degrees <= Knuckle Angle <= 92 degrees
    '       Determine which Profile Part should be the "Long Boxed " End Cut
    '           default "ShortBox" : if Part 2 Length is greater then Part 1 Length
    '           else default "LongBox"
    '   else default "Miter"
    pQH.SetQuestion "KnuckleEndCut", "Miter", _
                    "ProfileKnuckleEndCutTypes", _
                    "Init_KnuckleEndCutType", _
                    CUSTOMERID + "AssyConRul.StiffEndByKnuckleSel"
    
    pQH.SetQuestion "ChamferType", "None", , "ChamferQuestionOnER", CMLIBRARY_ASSYCONNSEL
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions", sError).Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    
    Dim sError As String
    Dim sKnuckleEndCut As String
    
    'Select the Defintion Item based on the "KnuckleEndCut" question:
    sKnuckleEndCut = pSL.Answer("KnuckleEndCut")
    If LCase(Trim(sKnuckleEndCut)) = LCase("Miter") Then
        pSL.Add "StiffEndToMiter"

    ElseIf LCase(Trim(sKnuckleEndCut)) = LCase("ShortBoxed") Then
        pSL.Add "StiffEndToShortBox"

    ElseIf LCase(Trim(sKnuckleEndCut)) = LCase("LongBoxed") Then
        pSL.Add "StiffEndToLongBox"

    Else
        sError = "invalid KnuckleEndCut: " & sKnuckleEndCut
        GoTo ErrorHandler
    End If
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorLogic", sError).Number
End Sub
  
' *******************************************************************************************
' *******************************************************************************************
' *******************************************************************************************
'   Custom Methods for Settting Default Questions
' *******************************************************************************************
Public Sub Init_KnuckleEndCutType(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                                  ByRef ppArgument As Object)
    'This is invoked before the SelectorLogic procedure in the Selector Class
    On Error GoTo ErrorHandler
    
    Dim sError As String
    Dim sKnuckleEndCut As String
    
    Dim dLength1 As Double
    Dim dLength2 As Double
    Dim dKnuckleAngle As Double
    
    Dim oInputDG As IMSSymbolEntities.IJDInputDuringGame
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    
    Dim oSelectorLogic As IJDSelectorLogic
    
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn
    
    ' Get Symbol Rep so Selector Logic object can be created
    ' Get Symbol Definition
    sError = "get Symbol Definition"
    Set oInputDG = pInput
    Set oSymbolDefinition = oInputDG.definition
    
    ' Create/Initialize the Selector Logic Object from the symbol definition
    sError = "Set Symbol SelectorLogic"
    Set oSelectorLogic = New SelectorLogic
    oSelectorLogic.Representation = oSymbolDefinition.IJDRepresentations(1)
    
    ' Create Assembly Connection Object
    sError = "get Bounding/Bounded Ports from Assembly Connection "
    Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssemblyConn.object = oSelectorLogic.SmartOccurrence
 
    Set oBoundedObject = oSDO_AssemblyConn.Port1
    Set oBoundingObject = oSDO_AssemblyConn.Port2
        
    ' From the Input Objects (IJPorts)
    ' Retrieve the Knuckle Point object used to Split the Profile
    Dim dPI As Double
    Dim bFromKnuckle As Boolean
    Dim oProfileKnuckle As IJProfileKnuckle
    bFromKnuckle = oSDO_AssemblyConn.FromKnuckle(oProfileKnuckle)
    If oProfileKnuckle Is Nothing Then
        dKnuckleAngle = 90#
    Else
        ' From the Knuckle Point,
        ' Retrieve the Knuckle Angle
        ' Convert from Radians to Degrees
        dPI = 3.14159265358979
        dKnuckleAngle = (oProfileKnuckle.Angle / dPI) * 180#
    End If
    
    ' From the Knuckle Point Angle used as the Splitter
    ' Determine if this is defaulted to a "Miter" or "Boxed" End Cut
    '   default "Boxed" : if 88 degrees <= Knuckle Angle <= 92 degrees
    '   else default "Miter"
    sError = "set Default KnuckleEndCut answer "
    
    'Get Profile Lengths
    Set oBoundedObject = oSDO_AssemblyConn.ConnectedObject1
    Set oBoundingObject = oSDO_AssemblyConn.ConnectedObject2
    
    ' From the Input Objects (Profile Parts)
    ' Retrieve each Profile Part's Length
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_ProfilePart.object = oBoundedObject
    dLength1 = oSDO_ProfilePart.MoldedLength

    Set oSDO_ProfilePart.object = oBoundingObject
    dLength2 = oSDO_ProfilePart.MoldedLength
            
    If dKnuckleAngle < 88# Then
        ' Knuckle Angle less then 88 degrees: default to "Miter"
        sKnuckleEndCut = "Miter"
        
    ElseIf dKnuckleAngle > 92# Then
        ' Knuckle Angle greater then 92 degrees: default to "Miter"
        sKnuckleEndCut = "Miter"
        
    Else
        ' Knuckle Angle close to 90 degrees: default to "ShortBoxed" or "LongBoxed"

        sError = "set Default for Boxed End answer "
        If dLength1 <= dLength2 Then
            sKnuckleEndCut = "ShortBoxed"
        Else
            sKnuckleEndCut = "LongBoxed"
        End If
    End If
    
    'For Knuckle type is Split And Extend and is Convex: Default to "ShortBoxed" or "LongBoxed"
    Dim oProfileKnuHelper As IJProfileKnuckleHelper
    Dim oProfileKnuckleMfg As IJProfileKnuckleMfg
    Dim dRequiredAngle As Double
    Dim dAngleInDeg As Double
    Dim bIsConvex As Boolean
    
    If Not oProfileKnuckle Is Nothing Then
        Set oProfileKnuHelper = New IMSProfileKnuckleEntity.ProfileKnuckleHelper
        bIsConvex = oProfileKnuHelper.IsConvexKnuckle(oProfileKnuckle)
        Set oProfileKnuckleMfg = oProfileKnuckle
       
        If bIsConvex And oProfileKnuckleMfg.ManufacturingMethod = pkmmSplitAndExtend Then
            
            'Get Angle Betweeen the two stiffeners
            Dim oFreeeEndCutUtil As New GSCADStructDetailUtil.StructDetailFreeEndUtil
            
            oFreeeEndCutUtil.GetAngleAtSplitManualKnuckle oBoundedObject, oBoundingObject, oProfileKnuckle, dKnuckleAngle
            dAngleInDeg = (dKnuckleAngle / GetPI) * 180#
            
            If dAngleInDeg < 180 Then
                dRequiredAngle = dAngleInDeg
            Else
                dRequiredAngle = 360 - dAngleInDeg
            End If
            
            'If the Angle between the two stiffeners is more than 150 degrees then we are setting Miter as default
            If dRequiredAngle > 150 Then
                sKnuckleEndCut = "Miter"
                Exit Sub
            End If
            
            If LessThanOrEqualTo(dLength1, dLength2) Then
                sKnuckleEndCut = "ShortBoxed"
            Else
                sKnuckleEndCut = "LongBoxed"
            End If
        End If
    End If
        
    oSelectorLogic.Answer("KnuckleEndCut") = sKnuckleEndCut
    
  Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Init_KnuckleEndCutType", sError).Number
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
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgID, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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




