VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaterStopDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "EdgeFeatureRules.WaterStopDef"
Const m_ItemName As String = CUSTOMERID + "EdgeFeatureRules.WaterStopDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\" + CUSTOMERID + "EdgeFeatureRules\WaterStopDef.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler

    pIH.SetInput INPUT_EDGE
    pIH.SetInput INPUT_POINT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler

    pAD.SetCMFinalConstruct CUSTOMERID + "EdgeFeatureRules.WaterStopDef", "CMFinalConstructEF"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim oMD As IJDMemberDescription
    
    Set oMD = pMDs.AddMember("PC", 1, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    oMD.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMIsPCNeeded"
    
    Set oMD = Nothing
     
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub
Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                         ByVal pResourceManager As IUnknown, _
                         ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim oEdgeFeatureWrapper As StructDetailObjects.EdgeFeature
    Dim oEdgeFeature As Object
    Dim oPartWithFeature As Object
    
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    
    ' Get edge feature
    Set oEdgeFeature = pMemberDescription.CAO
    
    ' Set up edge feature wrapper
    Set oEdgeFeatureWrapper = New StructDetailObjects.EdgeFeature
    Set oEdgeFeatureWrapper.object = oEdgeFeature
    
    ' Get the cut out part
    Set oPartWithFeature = oEdgeFeatureWrapper.GetPartObject
        
    ' For stiffer part, physical connection is between face port created
    ' by edge feature and face port of the stiffened plate
    
    ' For plate part, physical connection is between face port created by
    ' edge feature and face port of the bounding object ( Plate or Stiffener part? )
        
    ' Get port created by edge cut
    If TypeOf oPartWithFeature Is IJProfilePart Then
        Dim oProfilePart As IJProfilePart
        Dim oProfilePartWrapper As StructDetailObjects.ProfilePart
        
        Set oProfilePart = oPartWithFeature
        Set oProfilePartWrapper = New StructDetailObjects.ProfilePart
        Set oProfilePartWrapper.object = oProfilePart
        Set oPort1 = oProfilePartWrapper.CutoutPort(oEdgeFeature)
        
        Set oProfilePartWrapper = Nothing
        Set oProfilePart = Nothing
        
    ElseIf TypeOf oPartWithFeature Is IJPlatePart Then
        Dim oPlatePart As IJPlatePart
        Dim oPlatePartWrapper As StructDetailObjects.PlatePart
    
        Set oPlatePart = oPartWithFeature
        Set oPlatePartWrapper = New StructDetailObjects.PlatePart
        Set oPlatePartWrapper.object = oPlatePart
        Set oPort1 = oPlatePartWrapper.CutoutPort(oEdgeFeature)
        
        Set oPlatePartWrapper = Nothing
        Set oPlatePart = Nothing
    End If
    Set oPartWithFeature = Nothing
    
    ' Get connected port2
    Set oPort2 = oEdgeFeatureWrapper.GetConnectedPortAtEdgeFeature
        
    If Not oPort1 Is Nothing And Not oPort2 Is Nothing Then
        ' Create physical connection
        Dim strStartClass As String
        Dim oPhysicalConnectionWrapper As New StructDetailObjects.PhysicalConn
        
        strStartClass = "TeeWeld" ' "TeeWeldWaterStop" ?
        oPhysicalConnectionWrapper.Create pResourceManager, _
                                          oPort1, _
                                          oPort2, _
                                          strStartClass, _
                                          oEdgeFeature, _
                                          ConnectionPhysical
        Set pObject = oPhysicalConnectionWrapper.object
        Set oPhysicalConnectionWrapper = Nothing
    End If
    
    Set oPort1 = Nothing
    Set oPort2 = Nothing
    Set oEdgeFeatureWrapper = Nothing
    Set oEdgeFeature = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC").Number
    
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructEF(ByVal pAggregatorDescription As IJDAggregatorDescription)

  Dim oEdgeFeatureWrapper As New StructDetailObjects.EdgeFeature
  
  Set oEdgeFeatureWrapper.object = pAggregatorDescription.CAO
  
  oEdgeFeatureWrapper.AddFeatureToCut
    
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation

    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous member descriptions
    ItemMembers pMDs
  
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_ItemProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

'********************************************************************************************
' Method      : CMIsPCNeeded
' Description :
'********************************************************************************************
Public Sub CMIsPCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
  Const METHOD = "::CMIsPCNeeded"
  On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub





