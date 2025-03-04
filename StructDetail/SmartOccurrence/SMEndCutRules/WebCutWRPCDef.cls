VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutWRPCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID

Const m_CAName As String = "WebCutWRPCDef"                   ' - TODO
Const m_CAProjectName As String = CUSTOMERID + "EndCutRules"            ' - TODO

Const m_CAProgid As String = m_CAProjectName + "." + m_CAName
Const m_FamilyProgid As String = ""

'User names for item members
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\WebCutWRPCDef.cls"

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
    
    'Dim pMemDesc As IJDMemberDescription
    'Set pMemDesc = pMDs.AddMember("PC1", 1, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    
    Dim pMemDesc As IJDMemberDescription

'  Following members are Tee welds and are conditional based on profile type.
    Set pMemDesc = pMDs.AddMember("WEBCUTRPC1", 1, "CMWebRightPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMWWebRightCon"
    Set pMemDesc = Nothing
            
    Set pMemDesc = pMDs.AddMember("WEBCUTRPC2", 2, "CMTopFlangeRightBottomPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMTopFlangeRightBottomCon"
    Set pMemDesc = Nothing

    Set pMemDesc = pMDs.AddMember("WEBCUTRPC3", 3, "CMTopFlangeRightPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMTopFlangeRightCon"
    Set pMemDesc = Nothing

    Set pMemDesc = pMDs.AddMember("WEBCUTRPC4", 4, "CMWebRightTopCornerPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMWebRightTopCornerCon"
    Set pMemDesc = Nothing
    
    Set pMemDesc = pMDs.AddMember("WEBCUTRPC5", 5, "CMTopFlangeRightBottomCornerPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMTopFlangeRightBottomCornerCon"
    Set pMemDesc = Nothing
    
    Set pMemDesc = pMDs.AddMember("WEBCUTRPC6", 6, "CMWebRightBottomCornerPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMWebRightBottomCornerCon"
    Set pMemDesc = Nothing
    
    Set pMemDesc = pMDs.AddMember("WEBCUTRPC7", 7, "CMBottomFlangeRightTopPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMBottomFlangeRightTopCon"
    Set pMemDesc = Nothing
    
    Set pMemDesc = pMDs.AddMember("WEBCUTRPC8", 8, "CMBottomFlangeRightTopCornerPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMBottomFlangeRightTopCornerCon"
    Set pMemDesc = Nothing
    
    Set pMemDesc = pMDs.AddMember("WEBCUTRPC9", 9, "CMBottomFlangeRightPC", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMBottomFlangeRightCon"
    Set pMemDesc = Nothing
       
    Set pMemDesc = pMDs.AddMember("WebCutsWRightFET", 10, "CMConstructFET", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMGetEdgeTreatment"
    Set pMemDesc = Nothing
       
    'place a corner feature as a child of the web cut
    Set pMemDesc = pMDs.AddMember("CF1", 11, "CMCornerFeatureSnipe", CUSTOMERID + "EndCutRules.WebCutDefCM1")
    pMemDesc.SetCMConditional CUSTOMERID + "EndCutRules.WebCutDefCM1", "CMCornerFeatureSnipeCon"
    Set pMemDesc = Nothing
       
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

Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    Dim clsWebCutCM As New WebCutDefCM
            
    Set m_oPhysConn = clsWebCutCM.CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                          JXSEC_IDEALIZED_BOUNDARY, JXSEC_IDEALIZED_BOUNDARY)
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










