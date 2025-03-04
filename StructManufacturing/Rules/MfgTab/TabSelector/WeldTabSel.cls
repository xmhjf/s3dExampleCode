VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldTabSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "StrMfgTabSelRules.WeldTabSel"

Private Const m_SelectorProgid As String = "StrMfgTabSelRules.WeldTabSel"
Private Const m_SelectorName As String = "StrMfgTabSelRules.WeldTabSel"
Private Const m_FamilyProgid As String = ""

Private Const m_ViewName As String = "JUASMCornerTabCriteria"

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

Private Function GetPlatePortGivenMfgContour(oContour2d As IJMfgGeom2d) As IJPort
    Const METHOD = "GetPlatePortGivenMfgContour"
    On Error GoTo ErrorHandler
    
    Set GetPlatePortGivenMfgContour = Nothing
    If oContour2d Is Nothing Then Exit Function

    Dim oPortMnkr As IMoniker
    Set oPortMnkr = oContour2d.GetMoniker
    If oPortMnkr Is Nothing Then Exit Function
    
    Dim oPOM As IJDObject
    Set oPOM = oContour2d
    If oPOM Is Nothing Then Exit Function
    
    Dim oUtil As New MfgMGHelper
    
    Dim oPort As IJPort
    oUtil.BindMoniker oPOM.ResourceManager, oPortMnkr, oPort
    
    Set GetPlatePortGivenMfgContour = oPort

CleanUp:
    Set oPort = Nothing
    Set oPortMnkr = Nothing
    Set oPOM = Nothing
    Set oUtil = Nothing
    
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  GoTo CleanUp
End Function

Private Function GetPlateGivenMfgContour(oContour2d As IJMfgGeom2d) As IJPlatePart
    Const METHOD = "GetPlateGivenMfgContour"
    On Error GoTo ErrorHandler
    
    Set GetPlateGivenMfgContour = Nothing
    
    Dim oPort As IJPort
    Set oPort = GetPlatePortGivenMfgContour(oContour2d)
    If Not oPort Is Nothing Then Set GetPlateGivenMfgContour = oPort.Connectable
    
CleanUp:
    Set oPort = Nothing
    
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  GoTo CleanUp
End Function

Private Function GetConnectedPlateGivenMfgContour(oContour2d As IJMfgGeom2d) As IJPlatePart
    Const METHOD = "GetConnectedPlateGivenMfgContour"
    On Error GoTo ErrorHandler

    If oContour2d Is Nothing Then Exit Function
    
    Dim oMfgBevel As IJMfgBevel
    Set oMfgBevel = oContour2d.GetBevel
    If oMfgBevel Is Nothing Then Exit Function
    
    Dim oAppConn As IJAppConnection
    Set oAppConn = oMfgBevel.PhysicalConnection
    If oAppConn Is Nothing Then Exit Function
    
    Dim oConnPorts As IJElements
    oAppConn.enumPorts oConnPorts
    
    If oConnPorts Is Nothing Then Exit Function
    If oConnPorts.Count <> 2 Then Exit Function
    
    Dim MyPort As IJPort
    Set MyPort = GetPlatePortGivenMfgContour(oContour2d)

    Dim OtherPort As IJPort
    If MyPort Is oConnPorts.Item(1) Then
        Set OtherPort = oConnPorts.Item(2)
    ElseIf MyPort Is oConnPorts.Item(2) Then
        Set OtherPort = oConnPorts.Item(1)
    Else
        Exit Function
    End If

    If Not OtherPort Is Nothing Then Set GetConnectedPlateGivenMfgContour = OtherPort.Connectable
    
CleanUp:
    Set oMfgBevel = Nothing
    Set oAppConn = Nothing
    Set oConnPorts = Nothing
    Set MyPort = Nothing
    Set OtherPort = Nothing
    
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
  GoTo CleanUp
End Function

Private Function PlatesAreInSameAssembly(oPlate1 As IJPlatePart, oPlate2 As IJPlatePart) As Boolean
    Const METHOD = "PlatesAreInSameAssembly"
    On Error GoTo ErrorHandler

    PlatesAreInSameAssembly = False
    
    If oPlate1 Is Nothing Or oPlate2 Is Nothing Then Exit Function
    
    Dim oAssChild1 As IJAssemblyChild
    Set oAssChild1 = oPlate1
    
    Dim oAssChild2 As IJAssemblyChild
    Set oAssChild2 = oPlate2
    
    If oAssChild1.Parent Is oAssChild2.Parent Then
        PlatesAreInSameAssembly = True
    End If
    
    Set oAssChild1 = Nothing
    Set oAssChild2 = Nothing
    
    Exit Function

ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Private Function GetLengthOfGeom2dSegment(oGeom2d As IJMfgGeom2d) As Double
    Const METHOD = "GetLengthOfGeom2dSegment"
    On Error GoTo ErrorHandler

    Dim oCS As IJComplexString
    Set oCS = oGeom2d.GetGeometry
    
    Dim oCrv As IJCurve
    Set oCrv = oCS
    
    GetLengthOfGeom2dSegment = oCrv.Length
    
    Set oCS = Nothing
    Set oCrv = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    Exit Function
End Function


' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************
Private Function IJMfgTabSelectionHelper_SelectTabCatalogPart(ByVal oPartClassObj As Object, ByVal oTabCandidateObj As Object, ByVal lTabType As Long, ByVal oUnfoldedGeom2dColl As Object) As Object
    Const METHOD = "IJMfgTabSelectionHelper_SelectTabCatalogPart"
    On Error GoTo ErrorHandler
    
    Dim oTabCandidate As IJMfgTabCandidate
    Set oTabCandidate = oTabCandidateObj
    
    If oTabCandidate.NumberOfKnuckles > 0 Then
        GoTo CleanUp
    End If
    
    Dim oSymbolHandling As Object
    Set oSymbolHandling = SP3DCreateObject("MfgPlateTabHelper.SymbolHandling")
    
    'Get the tabs of specified type.
    Dim oOpeningTabs As IJDCollection
    Set oOpeningTabs = oSymbolHandling.GetTabsFromCatalog(m_ViewName)
    If oOpeningTabs Is Nothing Then
        GoTo CleanUp
    End If
    
    Dim lOpeningTabCount As Long, lIndex As Long
    lOpeningTabCount = oOpeningTabs.Size
    If lOpeningTabCount < 1 Then
        GoTo CleanUp
    End If
    
    Dim oPlatePart As IJPlatePart
    Set oPlatePart = GetPlateGivenMfgContour(oTabCandidate.FirstContour)
    If oPlatePart Is Nothing Then
        Set oPlatePart = GetPlateGivenMfgContour(oTabCandidate.SecondContour)
        If oPlatePart Is Nothing Then GoTo CleanUp
    End If

    Dim oHelper As IJPartSupport
    Set oHelper = New PlatePartSupport
    Set oHelper.Part = oPlatePart
    
    'Need to propose Feature tab when the opening/feature is split by a seam
    
    'Check if any of the contours is a feature or an opening
    Dim Operator(1 To 2) As Object
    On Error Resume Next
    oHelper.GetPortOperator GetPlatePortGivenMfgContour(oTabCandidate.FirstContour), Operator(1)
    oHelper.GetPortOperator GetPlatePortGivenMfgContour(oTabCandidate.SecondContour), Operator(2)
    
    Dim oOpeningObj As Object
    Dim bIsFirstOpening   As Boolean
    Dim bIsSecondOpening  As Boolean
    
    If (TypeOf Operator(1) Is IJOpeningEntity) Or (TypeOf Operator(1) Is IJStructFeature) Or (TypeOf Operator(1) Is IJSketchFeature) Then
                     
        bIsFirstOpening = True
        Set oOpeningObj = Operator(1)
                
    ElseIf (TypeOf Operator(2) Is IJOpeningEntity) Or (TypeOf Operator(2) Is IJStructFeature) Or (TypeOf Operator(2) Is IJSketchFeature) Then
                
        bIsSecondOpening = True
        Set oOpeningObj = Operator(2)
                
    End If
    
    If (TypeOf Operator(1) Is IJSeam) Or (TypeOf Operator(2) Is IJSeam) Then 'End of If bIsFirstOpening = True Or bIsSecondOpening = True
        Dim oSeam As IJSeam
        Dim oContourObj       As Object
        Dim oPort             As IJStructPort

        Dim eSeamType         As StructSeamType
        Dim oSeamType         As IJDSeamType
        
        If TypeOf Operator(1) Is IJSeam And TypeOf Operator(2) Is IJSeam Then
            Dim oSeamType1 As IJDSeamType
            Dim oSeamType2 As IJDSeamType
            
            Set oSeamType1 = Operator(1)
            Set oSeamType2 = Operator(2)
            
            Dim eSeamType1 As StructSeamType
            Dim eSeamType2 As StructSeamType
            
            eSeamType1 = oSeamType1.SeamType
            eSeamType2 = oSeamType2.SeamType
            
            If eSeamType1 = sstPlanningSeam Then
                Set oContourObj = GetPlatePortGivenMfgContour(oTabCandidate.FirstContour)
                Set oSeam = Operator(1)
            Else 'If eSeamType2 = sstPlanningSeam Then
                Set oContourObj = GetPlatePortGivenMfgContour(oTabCandidate.SecondContour)
                Set oSeam = Operator(2)
            End If
            
        Else
        
            If TypeOf Operator(1) Is IJSeam Then
                Set oContourObj = GetPlatePortGivenMfgContour(oTabCandidate.FirstContour)
                Set oSeam = Operator(1)
            Else
                Set oContourObj = GetPlatePortGivenMfgContour(oTabCandidate.SecondContour)
                Set oSeam = Operator(2)
            End If
        End If
        
        Dim oSeamTypeWeldTab As IJDSeamType
        Set oSeamTypeWeldTab = oSeam
        
        Dim eSeamTypeWeldTab As StructSeamType
        eSeamTypeWeldTab = oSeamTypeWeldTab.SeamType
        
        If eSeamTypeWeldTab <> sstPlanningSeam Then Exit Function
        
        If TypeOf oContourObj Is IJPort Then
            Set oPort = oContourObj
            Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
            Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
            Set oSDPartSupport.Part = oPlatePart

            Dim oConnectedObjects        As Collection
            Dim oConnections             As Collection
            Dim oThisPartPort            As Collection
            Dim oOtherPartPort           As Collection

            oSDPartSupport.GetConnectedObjects ConnectionPhysical, oConnectedObjects, oConnections, _
                oThisPartPort, oOtherPartPort
            
            Dim lPortCount As Long
            For lPortCount = 1 To oThisPartPort.Count
                'compare thisPartPort and oPort
                Dim oTempPort As IJPort
                Set oTempPort = oThisPartPort.Item(lPortCount)
                Dim oTempOperator As Object
                
                
                
                oHelper.GetPortOperator oTempPort, oTempOperator
                'if oTempOperator is oSeam
                If oTempOperator Is oSeam Then
                    For lIndex = 1 To lOpeningTabCount
                        Dim oPart As IJDPart
                        Set oPart = oOpeningTabs.Item(lIndex)
                        'get tab type attribute
                        Dim oAttr As IJDAttribute
                        Set oAttr = GetAttribute(oPart, "IJUASMPlateTabType", "TabType")
                        If oAttr.Value = 3 And lTabType = 3 Then  'WeldTab
                            If oPart.PartNumber = "MfgTabWeldType4_100_10" Then 'Return the symbol for the testcase
                                Set IJMfgTabSelectionHelper_SelectTabCatalogPart = oPart
                                oTabCandidate.TabType = lTabType
                                Exit Function
                            End If
                        End If
                        Set oPart = Nothing
                        Set oAttr = Nothing
                    Next
                End If
            Next
        End If
        Set oPort = Nothing
        Set oContourObj = Nothing

    End If
    
CleanUp:
    Set oSymbolHandling = Nothing
    Set oOpeningTabs = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
Private Function GetAssemblyStageCode(ByVal oThisPart As Object, ByVal oOtherPart As Object) As Long
    Const METHOD = "GetAssemblyStageCode"
    On Error GoTo ErrorHandler
    
    GetAssemblyStageCode = -1
    
    Dim oHelper As StructDetailObjects.Helper
    Set oHelper = New StructDetailObjects.Helper
    
    On Error Resume Next
    Dim oFirstMeet  As IJAssembly
    Set oFirstMeet = oHelper.FirstMeet(oThisPart, oOtherPart)
    On Error GoTo ErrorHandler
    
    If Not oFirstMeet Is Nothing Then
        Dim oAssyBase  As IJAssemblyBase
        Set oAssyBase = oFirstMeet
        
        GetAssemblyStageCode = oAssyBase.Stage
    End If
            
CleanUp:
    Set oHelper = Nothing
    Set oFirstMeet = Nothing
    Set oAssyBase = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
Private Function GetConnectedPart(ByVal oContourPort As IJStructPort) As Object
    Const METHOD = "GetConnectedPart"
    On Error GoTo ErrorHandler
        
    Dim oInputPort As IJPort
    Set oInputPort = oContourPort
        
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PlatePartSupport
    Set oPartSupport.Part = oInputPort.Connectable
    
    Dim ThisPortColl As Collection
    Dim OtherPortColl As Collection
    Dim ConnectedObjColl As Collection
    Dim ConnectionsColl As Collection
    oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                         ConnectedObjColl, ConnectionsColl, _
                                         ThisPortColl, OtherPortColl
    
    Dim iCnt As Integer
    For iCnt = 1 To ThisPortColl.Count
        Dim oThisPort As IJStructPort
        Set oThisPort = ThisPortColl.Item(iCnt)
        
        If (oThisPort.ContextID = oContourPort.ContextID) And (oThisPort.operatorID = oContourPort.operatorID) And (oThisPort.operationID = oContourPort.operationID) Then
            'return the connected part
            Set GetConnectedPart = ConnectedObjColl.Item(iCnt)
            Exit For
        End If
        
        Set oThisPort = Nothing
    Next iCnt
    
CleanUp:
    
    Set oPartSupport = Nothing
    Set ThisPortColl = Nothing
    Set OtherPortColl = Nothing
    Set ConnectedObjColl = Nothing
    Set ConnectionsColl = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Function GetOpeningGeometry(ByVal oTabCandidateGeom2d As IJMfgGeom2d, ByVal oUnfoldedGeom2dColl As Object) As IJComplexString
    Const METHOD = "GetOpeningGeometry"
    On Error GoTo ErrorHandler
    
    Dim oContourMoniker As IMoniker
    Set oContourMoniker = oTabCandidateGeom2d.GetMoniker
    
    Dim oMoniker1 As JDMoniker
    Set oMoniker1 = New JDMoniker
    oMoniker1.JMoniker = oContourMoniker
    
    Dim oGeom2dColl As IJElements
    Set oGeom2dColl = New JObjectCollection
    
    Dim oMfgColl2dInput As IJMfgGeomCol2d
    Set oMfgColl2dInput = oUnfoldedGeom2dColl
    
    Dim oGeom2d        As IJMfgGeom2d
    Dim oGeom2dMoniker As IMoniker
    Dim oMoniker2      As JDMoniker
    Dim oMfgGeomHelper As New MfgGeomHelper
    Dim idx As Long
    For idx = 1 To oMfgColl2dInput.GetCount
    
        'Get all geom2d objects with that moniker
        Set oGeom2d = oMfgColl2dInput.GetGeometry(idx)
        Set oGeom2dMoniker = oGeom2d.GetMoniker
        
        If Not oGeom2dMoniker Is Nothing Then
            Set oMoniker2 = New JDMoniker
            oMoniker2.JMoniker = oGeom2dMoniker
            
            If oMoniker1.IsEqual(oMoniker2) Then
                oGeom2dColl.Add oGeom2d
            End If
                        
            Set oMoniker2 = Nothing
        End If
        
        Set oGeom2dMoniker = Nothing
        Set oGeom2d = Nothing
        
    Next
        
    'Merge the geom2ds if there are multiple geom2d objects
    If oGeom2dColl.Count > 1 Then
        'Merge the complex strings
        Dim oMergedElems As IJElements
        Set oMergedElems = New JObjectCollection
        Set oMergedElems = oMfgGeomHelper.OptimizedMergingOfInputCurves(oGeom2dColl)
        
        Set GetOpeningGeometry = oMergedElems.Item(1)
    Else
        Set GetOpeningGeometry = oGeom2dColl.Item(1).GetGeometry
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
