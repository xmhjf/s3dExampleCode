VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketByPlaneMT4SSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'   Selector:
'   ---------
'   It selects the list of available Items or SmartClasses in the context of
'   the smart occurrence.
'
'   - Inputs can be provided explicitly, by default they are identical to the inputs of the  family
'   - Questions are defined to parameterize the selection
'
Option Explicit

Const m_SelectorProgid As String = "BracketRules.BracketByPlaneMT4SSel"
Const m_SelectorName As String = "BracketRules.BracketByPlaneMT4SSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\BracketRules\BracketByPlaneMT4SSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    ' inputs:
    '   Plane that bracket is on
    '   Bracket Plate System
    '   U Point
    '   V Point
    '   Group of supports
    
    pIH.SetInput INPUT_BRACKETPLANE
    pIH.SetInput INPUT_BRACKETPLATE
    pIH.SetInput INPUT_UPOINT
    pIH.SetInput INPUT_VPOINT
    pIH.SetInput INPUT_SUPPORTS
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub

Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    On Error GoTo ErrorHandler
    ' TODO - Add your question below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
    On Error GoTo ErrorHandler
    'get the number of supports selected
    Dim nSupports As Long
    Dim oSupportsColl As IJElements
    
    ' Get Bracket-by-Plane Smart Occurrence:
    Dim oBracketByPlaneSO As IJSmartOccurrence
    Set oBracketByPlaneSO = pSL.SmartOccurrence
      
    ' Create helper for final construction of bracket:
    Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    
    'Get the supports from BracketContourSO.
    oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupportsColl, nSupports


    Set oBracketUtils = Nothing
    
    'Get Supports
    Dim oSupport3 As Object
    Dim oSupport4 As Object

    If Not oSupportsColl Is Nothing Then
        If oSupportsColl.Count > 3 Then
            Set oSupport3 = oSupportsColl.Item(3)
            Set oSupport4 = oSupportsColl.Item(4)
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
        
'************************************************************************* New. rule selector logic...
        'Check support  4 and 5 is plate or profile
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
    Dim oProfileSup3 As StructDetailObjects.ProfileSystem
    Dim oProfileSup4 As StructDetailObjects.ProfileSystem
    
    Select Case pHelper.ObjectType(oSupport3)
        Case SDOBJECT_PLATESYSTEM
            Select Case pHelper.ObjectType(oSupport4)
                Case SDOBJECT_PLATESYSTEM
                    ' No 4S brackets exist for plate by plate systems...
                Case SDOBJECT_STIFFENERSYSTEM
                     pSL.Add "BPMT_4SBracketPlateProfile"
            End Select
        Case SDOBJECT_STIFFENERSYSTEM
            Select Case pHelper.ObjectType(oSupport4)
                Case SDOBJECT_PLATESYSTEM
                    pSL.Add "BPMT_4SBracketProfilePlate"
                Case SDOBJECT_STIFFENERSYSTEM
                     Dim uBracketPoint As IJPoint
                     Dim vBracketPoint As IJPoint
                     Set uBracketPoint = pSL.InputObject(INPUT_UPOINT)
                     Set vBracketPoint = pSL.InputObject(INPUT_VPOINT)
                     
                    ' Get the profiles for support three and four...
                     Set oProfileSup3 = New StructDetailObjects.ProfileSystem
                     Set oProfileSup3.object = oSupportsColl.Item(3)
                     
                     Set oProfileSup4 = New StructDetailObjects.ProfileSystem
                     Set oProfileSup4.object = oSupportsColl.Item(4)
                    
                    ' find out if one if one of the profiles is a FB...
                     If oProfileSup3.SectionType = "FB" Then
                              ' If support three is a flat bar then check to see if support 4 is flange in...
                              If IsFlangeTowardsBracket(oProfileSup4.object, oBracketByPlaneSO, uBracketPoint, vBracketPoint) Then
                                    pSL.Add "4SCLTOffSetFromSup4FlangeInSc"
                              Else
                              
                                    'No defined bracket exist for flange out and one flat bar or if both supports are flat bars...
                                    ' Added default deleveried symbol...
                                    pSL.Add "BPMT_4SBracketProfileProfile"
                              End If
                              
                     ElseIf oProfileSup4.SectionType = "FB" Then
                             ' If support four is a flat bar then check to see if support 3 is flange in...
                             If IsFlangeTowardsBracket(oProfileSup3.object, oBracketByPlaneSO, uBracketPoint, vBracketPoint) Then
                                    pSL.Add "4SCLTOffSetFromSup3FlangeInSc"
                            Else
                                  'No defined bracket exist for flange out and one flat bar or if both supports are flat bars...
                                  ' Added default deleveried symbol...
                                    pSL.Add "BPMT_4SBracketProfileProfile"
                            End If

                     ElseIf Not IsFlangeTowardsBracket(oProfileSup3.object, oBracketByPlaneSO, uBracketPoint, vBracketPoint) Then
                              ' If support three is flange out, so check to see if support 4 is flange in...
                              If IsFlangeTowardsBracket(oProfileSup4.object, oBracketByPlaneSO, uBracketPoint, vBracketPoint) Then
                                    pSL.Add "4SCLTOffSetFromSup4FlangeInSc"
                              Else
                                    'No defined bracket exist for flange out and one flat bar or if both supports are flat bars...
                                    ' Added default deleveried symbol...
                                    pSL.Add "BPMT_4SBracketProfileProfile"
                              End If
                              
                     ElseIf Not IsFlangeTowardsBracket(oProfileSup4.object, oBracketByPlaneSO, uBracketPoint, vBracketPoint) Then
                              ' If support four is flange out, so check to see if support 3 is flange in...
                              If IsFlangeTowardsBracket(oProfileSup3.object, oBracketByPlaneSO, uBracketPoint, vBracketPoint) Then
                                    pSL.Add "4SCLTOffSetFromSup3FlangeInSc"
                              Else
                                  'No defined bracket exist for flange out and one flat bar or if both supports are flat bars...
                                  ' Added default deleveried symbol...
                                    pSL.Add "BPMT_4SBracketProfileProfile"
                              End If
                              
                     Else 'both flanges are flange in...
                              pSL.Add "4SCLTOffSetFromFlangeScallop"
                     End If
                 
                 Case Else
                    'Support 3 or 4 is not a plate or profile
            End Select
            
    End Select
    
        
'************************************************************************* End New. rule selector logic...

'************************************************************************* Orig. rule selector logic...
    'Check support  4 and 5 is plate or profile
'    Dim pHelper As StructDetailObjects.Helper
'    Set pHelper = New StructDetailObjects.Helper
'
'    Select Case pHelper.ObjectType(oSupport3)
'        Case SDOBJECT_PLATESYSTEM
'            Select Case pHelper.ObjectType(oSupport4)
'                Case SDOBJECT_PLATESYSTEM
'                    pSL.Add "BPMT_4SBracketPlates"
'                Case SDOBJECT_STIFFENERSYSTEM
'                    pSL.Add "BPMT_4SBracketPlateProfile"
'            End Select
'        Case SDOBJECT_STIFFENERSYSTEM
'            Select Case pHelper.ObjectType(oSupport4)
'                Case SDOBJECT_PLATESYSTEM
'                    pSL.Add "BPMT_4SBracketProfilePlate"
'                Case SDOBJECT_STIFFENERSYSTEM
'                    pSL.Add "BPMT_4SBracketProfileProfile"
'            End Select
'        Case Else
'            'Support 3 or 4 is not a plate or profile
'    End Select
'************************************************************************* End Orig. rule selector logic...

    Set oSupportsColl = Nothing
    Set oSupport3 = Nothing
    Set oSupport4 = Nothing
    Set oBracketByPlaneSO = Nothing
    
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
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub

' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

