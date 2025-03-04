VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtWeldSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = "PhysConnRules.ButtWeldSel"
Const m_SelectorName As String = "PhysConnRules.ButtWeldSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\ButtWeldSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  
    pQH.SetQuestion "Category", gsOneSided, "ButtWeldCategory"
    pQH.SetQuestion "Workcenter", gsMachine1, "WorkcenterCol"
    pQH.SetQuestion "UpSide", gsReference, "UpsideCol"
    pQH.SetQuestion "ChamferType", "None"
    pQH.SetQuestion "FlipBevel", "No", "BooleanCol"
    pQH.SetQuestion "FirstWeldingSide", "Molded", "FirstWeldingSideCol"

  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)

  Dim sERROR As String

  On Error GoTo ErrorHandler
  
  sERROR = ""
  
  ' Get Class Arguments
  Dim oPhysConn As New StructDetailObjects.PhysicalConn
  Dim dThickness As Double
  Set oPhysConn.object = pSLH.SmartOccurrence
  
  ' get thicknesses
'  If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
'    dThickness = oPhysConn.Object1Thickness
'  Else
'    dThickness = oPhysConn.Object2Thickness
'  End If
  If TypeOf oPhysConn.ConnectedObject1 Is IJPlate And _
     TypeOf oPhysConn.ConnectedObject2 Is IJPlate Then
     Dim oPlate1 As New StructDetailObjects.PlatePart
     Dim oPlate2 As New StructDetailObjects.PlatePart
     Set oPlate1.object = oPhysConn.ConnectedObject1
     Set oPlate2.object = oPhysConn.ConnectedObject2
        If oPlate1.OffsetToBaseFace(oPlate2) < oPlate2.OffsetToBaseFace And _
            oPlate1.OffsetToOffsetFace(oPlate2) < oPlate2.OffsetToOffsetFace Then
            dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace
        ElseIf oPlate1.OffsetToBaseFace(oPlate2) > oPlate2.OffsetToBaseFace And _
            oPlate1.OffsetToOffsetFace(oPlate2) > oPlate2.OffsetToOffsetFace Then
            dThickness = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2)
        Else
            If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
                dThickness = oPhysConn.Object2Thickness
            Else
                dThickness = oPhysConn.Object1Thickness
            End If
        End If
  Else
    If oPhysConn.Object1Thickness > oPhysConn.Object2Thickness Then
        dThickness = oPhysConn.Object2Thickness
    Else
        dThickness = oPhysConn.Object1Thickness
    End If
  End If

   ' Get the Question answers
   Dim sSided As String
   sSided = pSLH.Answer("Category")

  ' ********************* Selection code - TODO  ****************
  
  Select Case sSided
     Case gsOneSided
        If dThickness <= 0.012 Then
            pSLH.Add "ButtWeld1"
        ElseIf dThickness > 0.012 And dThickness <= 0.03 Then
            pSLH.Add "ButtWeld2"
            pSLH.Add "ButtWeld4"
        ElseIf dThickness > 0.03 Then
            pSLH.Add "ButtWeld3"
        End If
        
    Case gsTwoSided
        If dThickness <= 0.014 Then
            pSLH.Add "ButtWeld1"
        ElseIf dThickness > 0.014 And dThickness <= 0.02 Then
            pSLH.Add "ButtWeld2"
            pSLH.Add "ButtWeld4"
        ElseIf dThickness > 0.02 Then
            pSLH.Add "ButtWeld3"
        End If
  End Select
  
  ' *********************************************************
  
  Set oPhysConn = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic", sERROR).Number
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
  
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractSelector pSelector
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
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
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

