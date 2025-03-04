VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndEndAlongAxisFPCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID
Private Const m_sClassName As String = "EndEndAlongAxisFPCDef"
'
Private Const m_sRuleProgID As String = PROJECT_NAME & "." & m_sClassName
Private Const m_sRuleName As String = m_sRuleProgID
Private Const m_sFamilyProgid As String = ""
Private Const MODULE = PROJECT_PATH & m_sClassName & ".cls"

'User names for item members
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

Private Sub ItemInputs(oIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    oIH.SetInput INPUT_BOUNDING
    oIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Private Sub ItemAggregator(oAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    oAD.UserTypeClsid = CA_WEBCUT       '
    oAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    oAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructWebCut"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Private Sub ItemMembers(oMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim oMemDesc As IJDMemberDescription
    
    Set oMemDesc = oMDs.AddMember("PC1", _
                                  1, _
                                  "CMConstructPC1", _
                                  imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsPC1Needed"
    Set oMemDesc = Nothing
    
    Set oMemDesc = oMDs.AddMember("EndToEndPhyConnFFET", _
                                   2, _
                                   "CMConstructFET", _
                                   imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional m_sRuleProgID, "CMIsFETNeeded"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructWebCut(ByVal oAggregatorDescription As IJDAggregatorDescription)
    Dim oWebCutCM As New WebCutDefCM
    
    oWebCutCM.WebCut_FinalConstruct oAggregatorDescription
    
    Set oWebCutCM = Nothing
End Sub

Public Sub CMIsPC1Needed(ByVal oMD As IJDMemberDescription, _
                         ByRef bIsNeeded As Boolean)
   On Error GoTo ErrorHandler
   Dim sError As String
   
   ' Bottom flange PC is always needed,
   ' Top flange PC is dependent on cross section size
   bIsNeeded = True
   
   Dim oCommonHelper As New DefinitionHlprs.CommonHelper
   Dim sAnswerToBottomFlange As String
   
   sError = "Get bottom flange"
   'TR-169311 Trying to retrieve question's answer without
   'hardcoding the path
   GetSelectorAnswer oMD.CAO, gsBottomFlange, sAnswerToBottomFlange
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If

   Set oCommonHelper = Nothing
   If sAnswerToBottomFlange = "Yes" Then
      Exit Sub
   End If
   
   ' Top flange cut,PC is conditional
   Dim oFlangeCutWrapper As New StructDetailObjects.FlangeCut
   Dim oBoundedPart As Object
   Dim oBoundingPart As Object
   Dim oBoundingObject As Object
   
   sError = "Get flange cut inputs"
   Set oFlangeCutWrapper.object = oMD.CAO
   Set oBoundedPart = oFlangeCutWrapper.Bounded
   Set oBoundingObject = oFlangeCutWrapper.BoundingPort
   If TypeOf oBoundingObject Is IJPort Then
      Set oBoundingPart = oFlangeCutWrapper.Bounding
   ElseIf TypeOf oBoundingObject Is IJPlane Then
      Dim oDesignChild As IJDesignChild
      Dim oAC As Object
      Dim oACWrapper As New StructDetailObjects.AssemblyConn
      
      Set oDesignChild = oMD.CAO
      Set oAC = oDesignChild.GetParent
      Set oACWrapper.object = oAC
      Set oBoundingPart = oACWrapper.ConnectedObject2
   End If
   
   Dim dBoundedHeight As Double
   Dim dBoundingHeight As Double
   Dim dBoundedFT As Double
   Dim dBoundingFT As Double
   Dim sSectionType As String
   
   If TypeOf oBoundedPart Is IJStiffenerPart Then
      Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
      
      Set oProfilePartWrapper.object = oBoundedPart
      dBoundedHeight = Round(oProfilePartWrapper.Height, 3)
      dBoundedFT = Round(oProfilePartWrapper.FlangeThickness, 3)
   
      Set oProfilePartWrapper.object = oBoundingPart
      dBoundingHeight = Round(oProfilePartWrapper.Height, 3)
      dBoundingFT = Round(oProfilePartWrapper.FlangeThickness, 3)
      
   ElseIf TypeOf oBoundedPart Is IJBeamPart Then
      Dim oBeamPartWrapper As New StructDetailObjects.BeamPart
   
      Set oBeamPartWrapper.object = oBoundedPart
      dBoundedHeight = Round(oBeamPartWrapper.Height, 3)
      dBoundedFT = Round(oBeamPartWrapper.FlangeThickness, 3)
   
      Set oProfilePartWrapper.object = oBoundingPart
      dBoundingHeight = Round(oBeamPartWrapper.Height, 3)
      dBoundingFT = Round(oBeamPartWrapper.FlangeThickness, 3)
      
   End If
   
   If dBoundedHeight > dBoundingHeight Then
      If dBoundedHeight - dBoundedFT > dBoundingHeight Then
         bIsNeeded = False
      Else
         bIsNeeded = True
      End If
      
   ElseIf dBoundedHeight < dBoundingHeight Then
      If dBoundedHeight < dBoundingHeight - dBoundingFT Then
         bIsNeeded = False
      Else
         bIsNeeded = True
      End If
   Else
      bIsNeeded = True
   End If
   
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, MODULE, "CMIsPC1Needed", sError).Number
   
End Sub

Public Sub CMConstructPC1(ByVal oMD As IJDMemberDescription, _
                          ByVal oResourceManager As IUnknown, _
                          ByRef oObject As Object)
    
   Set oObject = CreatePCForAlongAxisEndCut( _
                          oMD, _
                          oResourceManager, _
                          "ButtWeld", _
                          JXSEC_IDEALIZED_BOUNDARY, _
                          JXSEC_IDEALIZED_BOUNDARY)
End Sub

Public Sub CMIsFETNeeded(ByVal oMD As IJDMemberDescription, _
                         ByRef bIsNeeded As Boolean)
   ' Forward to WebCutDefCM
   Dim oWebCutDefCM As New WebCutDefCM
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_FreeEdgeTreatment) Then
        bIsNeeded = False
        Exit Sub
    End If
   
   oWebCutDefCM.CMGetEdgeTreatment oMD, bIsNeeded
   Set oWebCutDefCM = Nothing
End Sub

Public Sub CMConstructFET(ByVal oMD As IJDMemberDescription, _
                          ByVal oResourceManager As IUnknown, _
                          ByRef oObject As Object)
   ' For ward to WebCutDefCM
   Dim oWebCutDefCM As New WebCutDefCM
   
   oWebCutDefCM.CMConstructFET oMD, _
                               oResourceManager, _
                               oObject
   Set oWebCutDefCM = Nothing
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal vDefinitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_sRuleProgID
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(oDefinition As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    oDefinition.IJDInputs.RemoveAllInput
    oDefinition.IJDRepresentations.RemoveAllRepresentation
    
    oDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    oDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' Define the inputs
    Dim oIH As IJDInputsHelper
    
    Set oIH = New InputHelper
    oIH.Definition = oDefinition
    oIH.InitAs m_sFamilyProgid
    ItemInputs oIH
    
    Dim oPDs As IJDPropertyDescriptions
    
    ' Define the aggregator
    Dim oAD As IJDAggregatorDescription
    Set oAD = oDefinition
    Set oPDs = oAD
    oPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator oAD
     
    ' Define the members
    Dim oMDs As IJDMemberDescriptions
    Set oMDs = oDefinition
    oMDs.RemoveAll ' Remove all the previous member descriptions
    ItemMembers oMDs
  
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition( _
          ByVal sCodeBase As String, _
          ByVal vDefParams As Variant, _
          ByVal ActiveConnection As Object) As Object
    Dim oDefinition As IJDSymbolDefinition
    Dim oCAFactory As New CAFactory
    
    Set oDefinition = oCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    oDefinition.ProgId = m_sRuleProgID
    oDefinition.CodeBase = sCodeBase
    oDefinition.Name = IJDUserSymbolServices_GetDefinitionName(vDefParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition oDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = oDefinition
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation( _
          ByVal oSblOcc As Object, _
          ByVal sRepName As String, _
          ByVal oOutputColl As Object, _
          ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence( _
          ByRef oSymbolOccurence As Object, _
          ByVal oTransactionMgr As Object) As Boolean
End Function

