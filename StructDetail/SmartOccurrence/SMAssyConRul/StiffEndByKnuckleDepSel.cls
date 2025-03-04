VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffEndByKnuckleDepSel"
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

Const m_SelectorProgID As String = CUSTOMERID + "AssyConRul.StiffEndByKnuckleDepSel"
Const m_SelectorName As String = CUSTOMERID + "AssyConRul.StiffEndByKnuckleDepSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffEndByKnuckleDepSel.cls"

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
    Dim oAC As New StructDetailObjects.AssemblyConn
    Set oAC.object = pSL.SmartOccurrence
    
    Dim oSecondAC As Object
    Dim oSeconACName As String
    Dim bMutualBounded As Boolean
    
    IsMutuallyBoundedCase oAC.ConnectedObject1, oAC.ConnectedObject2, bMutualBounded
    If Not bMutualBounded Then
        Exit Sub
    End If
    
    GetOtherACInMutualBoundedCase oAC, oSecondAC, oSeconACName
    
    If oSeconACName = "StiffEndToMiter" Then
        pSL.Add "StiffEndToMiter_Dep"
    ElseIf oSeconACName = "StiffEndToShortBox" Then
        pSL.Add "StiffEndToLongBox_Dep"
    ElseIf oSeconACName = "StiffEndToLongBox" Then
        pSL.Add "StiffEndToShortBox_Dep"
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
    
    Dim bMutuallyBounded As Boolean
    Dim oSecondAC As Object
    Dim oSecondACName As String
    
    IsMutuallyBoundedCase oSDO_AssemblyConn.ConnectedObject1, oSDO_AssemblyConn.ConnectedObject2, bMutuallyBounded
    If bMutuallyBounded Then
         GetOtherACInMutualBoundedCase oSDO_AssemblyConn, oSecondAC, oSecondACName   'ACs to support Mutually bounded cases for SB/LB end cuts.
         If oSecondACName = "StiffEndToShortBox" Then
             sKnuckleEndCut = "LongBoxed"
         ElseIf oSecondACName = "StiffEndToLongBox" Then
             sKnuckleEndCut = "ShortBoxed"
         ElseIf oSecondACName = "StiffEndToMiter" Then
             sKnuckleEndCut = "Miter"
         End If
    End If
'    MsgBox sKnuckleEndCut
         
    If sKnuckleEndCut = "" Then
    
    ' From the Knuckle Point Angle used as the Splitter
    ' Determine if this is defaulted to a "Miter" or "Boxed" End Cut
    '   default "Boxed" : if 88 degrees <= Knuckle Angle <= 92 degrees
    '   else default "Miter"
    sError = "set Default KnuckleEndCut answer "
    If dKnuckleAngle < 88# Then
        ' Knuckle Angle less then 88 degrees: default to "Miter"
        sKnuckleEndCut = "Miter"
        
    ElseIf dKnuckleAngle > 92# Then
        ' Knuckle Angle greater then 92 degrees: default to "Miter"
        sKnuckleEndCut = "Miter"
        
    Else
        sKnuckleEndCut = "Miter"
'        ' Knuckle Angle close to 90 degrees: default to "ShortBoxed" or "LongBoxed"
'        Set oBoundedObject = oSDO_AssemblyConn.ConnectedObject1
'        Set oBoundingObject = oSDO_AssemblyConn.ConnectedObject2
'
'        ' From the Input Objects (Profile Parts)
'        ' Retrieve each Profile Part's Length
'        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
'        Set oSDO_ProfilePart.object = oBoundedObject
'        dLength1 = oSDO_ProfilePart.MoldedLength
'
'        Set oSDO_ProfilePart.object = oBoundingObject
'        dLength2 = oSDO_ProfilePart.MoldedLength
'
'        sError = "set Default for Boxed End answer "
'
'        If dLength1 <= dLength2 Then
'            sKnuckleEndCut = "ShortBoxed"
'        Else
'            sKnuckleEndCut = "LongBoxed"
'        End If
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






