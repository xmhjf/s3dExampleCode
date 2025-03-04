VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckleTabSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "KnuckleTabSel"

Private Const PI = 3.14159265358979
Const m_SelectorProgid As String = "StrMfgTabSelRules.KnuckleTabSel"
Const m_SelectorName As String = "StrMfgTabSelRules.KnuckleTabSel"
Const m_FamilyProgid As String = ""

Private m_oCatalogConnection As IJDPOM

Implements IJDUserSymbolServices
Implements IJMfgTabSelectionHelper

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  Const METHOD = "SelectorsInputs"
  
  On Error GoTo ErrorHandler
  
  pIH.SetInput INPUT_PLATE
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pIH.ReportError
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  Const METHOD = "SelectorQuestions"
  On Error GoTo ErrorHandler
  
  'No questions defined
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pQH.ReportError
End Sub

Public Sub SelectorLogic(pSL As IJDSelectorLogic)
  Const METHOD = "SelectorLogic"
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  'pSL.ReportError sERROR
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
  DefineViewPrefix pSelector, "INGR"
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

Private Function IJMfgTabSelectionHelper_SelectTabCatalogPart(ByVal oPartClassObj As Object, ByVal oTabCandidateObj As Object, ByVal lTabType As Long, ByVal oUnfoldedGeom2dColl As Object) As Object
    Const METHOD = "IJMfgTabSelectionHelper_SelectTabCatalogPart"
    On Error GoTo ErrorHandler
    
    Dim oTabCandidate As IJMfgTabCandidate
    Set oTabCandidate = oTabCandidateObj
    
    If oTabCandidate.NumberOfKnuckles < 1 Then
        GoTo CleanUp
    End If
    
    Dim oSymbolHandling As Object
    Set oSymbolHandling = SP3DCreateObject("MfgPlateTabHelper.SymbolHandling")
    
    'Get the tabs of specified type.
    Dim oTabsColl As IJDCollection
    Set oTabsColl = oSymbolHandling.GetTabsFromCatalog("JUASMPlateKnuckleTabCriteria")
    If oTabsColl Is Nothing Then
        GoTo CleanUp
    End If

    Dim lKnuckleTabCount As Long
    lKnuckleTabCount = oTabsColl.Size
    If lKnuckleTabCount < 1 Then
        GoTo CleanUp
    End If
    
    'Propose Knuckle tab only when sweep angle is less than 175 degrees and intersection angle with plate
    'edge is less than 45 degrees
    Dim oPlatePart As IJPlatePart
    Dim dKnuckleSweepAngle As Double, dPlateEdgeAngle As Double
    Dim oPort As IJPort
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
        
    dKnuckleSweepAngle = oTabCandidate.KnuckleAngle(1)

    Dim dPlateEdgeAngle1 As Double
    dPlateEdgeAngle1 = oTabCandidate.SmallestAngleBetweenKnuckleAndFirstContour(1)
    If (dPlateEdgeAngle1 < 0.01) Then dPlateEdgeAngle1 = PI - dPlateEdgeAngle1
        
    Dim dPlateEdgeAngle2 As Double
    If Not oTabCandidate.SecondContour Is Nothing Then
        dPlateEdgeAngle2 = oTabCandidate.SmallestAngleBetweenKnuckleAndSecondContour(1)
    End If
    If (dPlateEdgeAngle2 < 0.01) Then dPlateEdgeAngle2 = PI - dPlateEdgeAngle2
        
    If (dPlateEdgeAngle1 < dPlateEdgeAngle2) Then
        dPlateEdgeAngle = dPlateEdgeAngle1
    Else
        dPlateEdgeAngle = dPlateEdgeAngle2
    End If
        
    'Get the plate part
    Dim oPortMoniker As IMoniker
    Set oPortMoniker = oTabCandidate.FirstContour.GetMoniker
    If Not oPortMoniker Is Nothing Then
        Set oPort = oMfgRuleHelper.BindToObject(GetActiveConnection.GetResourceManager(GetActiveConnectionName), oPortMoniker)
        If Not oPort Is Nothing Then
            Set oPlatePart = oPort.Connectable
        End If
    End If
        
    Dim oIJDPart As IJDPart
    
    'Find the best tab symbol.
    Dim lIndex As Long
    For lIndex = 1 To lKnuckleTabCount
        Set oIJDPart = oTabsColl.Item(lIndex)
        
        Dim oIJDTabTypeAttr As IJDAttribute
        Set oIJDTabTypeAttr = GetAttribute(oIJDPart, "IJUASMPlateTabType", "TabType")
        
        If lTabType = oIJDTabTypeAttr.Value Then
                  
            Dim oIJDTabLengthAttr As IJDAttribute
            Set oIJDTabLengthAttr = GetAttribute(oIJDPart, "IJUASMPlateKnuckleTab", "Length1")
                
            'Check if the plate is corrugate plate
            If IsCorrugated(oPlatePart) = True Then
               If (dKnuckleSweepAngle < ((175 * PI) / 180) And dPlateEdgeAngle < ((45 * PI) / 180)) Or _
                  IsTabCandidateEligibleForKnuckleTab(oTabCandidate, oIJDTabLengthAttr.Value) = True Then
                  
                  If oIJDPart.PartNumber = "MfgKnuckleTab_230" Then 'Propose the sym for corrugate plate
                      Set IJMfgTabSelectionHelper_SelectTabCatalogPart = oIJDPart
                      oTabCandidate.TabType = lTabType
                      Exit For
                  End If
                  
               End If
            Else
                                  
                If (dKnuckleSweepAngle < ((175 * PI) / 180) And dPlateEdgeAngle < ((45 * PI) / 180)) Or _
                    IsTabCandidateEligibleForKnuckleTab(oTabCandidate, oIJDTabLengthAttr.Value) = True Then
                    
                    If oIJDPart.PartNumber = "MfgKnuckleTabLength_150" Then 'propose symbol suitable for the specific case
                        Set IJMfgTabSelectionHelper_SelectTabCatalogPart = oIJDPart
                        oTabCandidate.TabType = lTabType
                        Exit For
                    End If
                End If
                               
            End If 'End Of Plate Type Check
            
            Set oIJDTabLengthAttr = Nothing
                 
        End If 'lTabType = oIJDAttr.Value
            
        Set oIJDPart = Nothing
        Set oIJDTabTypeAttr = Nothing
        
    Next lIndex
 
CleanUp:
    Set oSymbolHandling = Nothing
    Set oTabsColl = Nothing
        
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
'This method checks if the plate type is "Undefined" or the plate type is not defined
Private Function IsUndefinedPlateType(oPlatePart As IJPlatePart) As Boolean
    Const METHOD = "IsUndefinedPlateType"
    On Error GoTo ErrorHandler
    
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlatePart, "IJUAProductionPlateInfo", "CustomPlate")
    
    If Not oAttr Is Nothing Then
        If oAttr.Value = vbEmpty Then
            IsUndefinedPlateType = True
        ElseIf oAttr.Value = 1 Then ' 1 is value for "UnDefined" in codelist "CustomProdInfoPlateType"
            IsUndefinedPlateType = True
        Else
            IsUndefinedPlateType = False
        End If
    Else
        IsUndefinedPlateType = True 'Return true even if the plate type is not defined
    End If
    
    Set oAttr = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IsTabCandidateEligibleForKnuckleTab(ByVal oTabCandidate As IJMfgTabCandidate, ByVal TabLength As Double) As Boolean
    Const METHOD = "IsTabCandidateEligibleForKnuckleTab"
    On Error GoTo ErrorHandler
    
    Dim oDrivenContourCS As IJComplexString
    Set oDrivenContourCS = oTabCandidate.ContourGeometryPortionNearTab(0.5)
        
    Dim dX As Double, dY As Double, dZ As Double
    oTabCandidate.GetTabLocation dX, dY, dZ
    
    Dim oTabLocation As IJDPosition
    Set oTabLocation = New DPosition
    oTabLocation.x = dX
    oTabLocation.y = dY
    oTabLocation.z = dZ
    
    Dim oMfgEntityHelper As New MfgEntityHelper
    Dim dLen             As Double
    
    dLen = oMfgEntityHelper.GetMinLengthOfGeometryPortionNearPos(oDrivenContourCS, oTabLocation, TabLength) 'Tab length
    
    If dLen < TabLength Then
        IsTabCandidateEligibleForKnuckleTab = True
    Else
        IsTabCandidateEligibleForKnuckleTab = False
    End If
   
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
