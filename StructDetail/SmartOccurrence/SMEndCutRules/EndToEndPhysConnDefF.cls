VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToEndPhysConnDefF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID

Const m_CAName As String = "EndToEndPhysConnDefF"                   ' - TODO
Const m_CAProjectName As String = CUSTOMERID + "EndCutRules"            ' - TODO

Const m_CAProgid As String = m_CAProjectName + "." + m_CAName
Const m_FamilyProgid As String = ""

'User names for item members
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

' Special user defined Edge ID used in the WebCut Symbol files
Const JXSEC_FLANGE_CUT_BY_WEBCUT = 8194

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\EndToEndPhysConnDefF.cls"

'Newly created member objects
Private m_oPhysConn As Object

'Error management
Private sError As String
Private sMETHOD As String

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    pAD.UserTypeClsid = CA_WEBCUT       '
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructWebCut"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim pMemDesc As IJDMemberDescription
    Set pMemDesc = pMDs.AddMember("PC1", 1, "CMConstructPC1", _
                                  imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional m_CAProgid, "CMIsPC1Needed"
    
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("EndToEndPhyConnFFET", 2, "CMConstructFET", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMGetEdgeTreatment"
    
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructWebCut(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Dim clsWebCutCM As New WebCutDefCM
    
    clsWebCutCM.WebCut_FinalConstruct pAggregatorDescription
    
    Set clsWebCutCM = Nothing
    Set m_oPhysConn = Nothing
End Sub

Public Sub CMConstructPC1(ByVal pMemberDescription As IJDMemberDescription, _
                          ByVal pResourceManager As IUnknown, _
                          ByRef pObject As Object)
    
    Dim clsWebCutCM As WebCutDefCM
    Set clsWebCutCM = New WebCutDefCM
            
zMsgBox "EndToEndPhysConnDef::CMConstructPC1"

    Set m_oPhysConn = clsWebCutCM.CAConstruct_PhysConn(pMemberDescription, _
                                                pResourceManager, _
                                                "ButtWeld", _
                                                JXSEC_IDEALIZED_BOUNDARY, _
                                                JXSEC_IDEALIZED_BOUNDARY)
    Set pObject = m_oPhysConn
    Set clsWebCutCM = Nothing
End Sub



' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_CAProgid
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
    pIH.Definition = pDefinition
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
    pDefinition.ProgId = m_CAProgid
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

Public Sub CMIsPC1Needed(ByVal oMD As IJDMemberDescription, _
                         ByRef bIsNeeded As Boolean)
   ' Bottom flange PC is always needed,
   ' Top flange PC is dependent on cross section size
   bIsNeeded = True
   
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
   
   Dim oCommonHelper As New DefinitionHlprs.CommonHelper
   Dim sAnswerToBottomFlange As String
   
   'TR-169311 Trying to retrieve question's answer without
   'hardcoding the path
   GetSelectorAnswer oMD.CAO, gsBottomFlange, sAnswerToBottomFlange
   
   Set oCommonHelper = Nothing
   If sAnswerToBottomFlange = "Yes" Then
      Exit Sub
   End If
   
   ' Top flange cut,PC is conditional
   Dim oFlangeCutWrapper As New StructDetailObjects.FlangeCut
   Dim oBoundedPart As Object
   Dim oBoundingPart As Object
   
   Set oFlangeCutWrapper.object = oMD.CAO
   Set oBoundedPart = oFlangeCutWrapper.Bounded
   Set oBoundingPart = oFlangeCutWrapper.Bounding
   
   Dim dBoundedHeight As Double
   Dim dBoundingHeight As Double
   Dim dBoundedFT As Double
   Dim dBoundingFT As Double
   
   'In miter, incase of mutual bound we pass miter plane as bounding. so check if Bounding is a valid port or not
   If TypeOf oBoundingPart Is IJPlane Then
    Dim oAC As New StructDetailObjects.AssemblyConn
    Dim oChild As IJDesignChild
    Set oChild = oFlangeCutWrapper.object
    
    Set oAC.object = oChild.GetParent
    Set oBoundingPart = oAC.ConnectedObject2
   End If

   
   GetHeightAndFlangeThickness oBoundedPart, dBoundedHeight, dBoundedFT
   GetHeightAndFlangeThickness oBoundingPart, dBoundingHeight, dBoundingFT
   
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
   
End Sub







