VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CopeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : CopeDef.cls
'
'Author : RP
'
'Description :
'    Module for coping the member part
'
'History:
'
' 06/09/02  CE    Added IsCopeNeededByShapeAndIncidence() to determine cope req based on member incidence
' 05/04/03  RP    Removed check for Zero values for all attributes for now. Zero is OK for clearances.
'                 It is not OK for cope length and cope depth. We need to look for these attributes when they
'                 are enabled.
' 02/02/04  JS    Added code to check if the copes/cutback planes need to be added to the geometry during a
'                 copy operation since the copy copies the assembly connection with the feature copes and
'                 the cutback plane but the part does not copy these features so they need to be re-applied.
'                 TR#52040.
' 02/16/04  RP    Added a parameter, resource manager, to the  cope and cutback creation methods. When resouce
'                 manager is nothing, the plane or wire body is not persisted. Moved the code that hides the
'                 plane/wirebody to the CMConstruct().
' 08/27/04  MH    CMMigrate
' 06/14/05  JS    Changed CMConditionalCutback to return errors as warnings so the part geometry gets computed(TR#49217)
' 06/13/06  RP    Changes due to impact from curved members. The cutback and cope ouputs IJSurface and IJModelBody now.
'                 Cutback creates surface by projection or revolution. Cope creates a solid cutter (DI#84001)
' 27/Sep/06 AS   TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
' 11/02/06  RP              TR#108733 and DI#84749 - Copelength and depth not computed. Also
'                           moved compute of these properties of aggregator section to avoid
'                           extra compute
' 07/06/07  MH   TR 121116  added  CheckForUndefinedValueAndRaiseError
'
' 12/02/14  SK   TR 240064  Split migration should be handled by the feature
'**********************************************************************************************************************

Option Explicit


Private Const MODULE = "CopeDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFeatureMacros.CopeDef"
Const m_ItemName As String = "SPSFeatureMacros.CopeDef"
Private Const strSourceFile As String = "CopeDef.cls"

Dim bOnPreLoad As Boolean
Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1
Private m_oLocalizer As IJLocalizer 'DI 80096

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJSymbolVersion

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  Dim oInput As IJDInput
  Dim oInputs As IJDInputs
  
  
  pIH.SetInput "Supported"
  pIH.SetInput "Supporting"
  
  Set oInputs = pIH.definition 'optional plate or Plane, which will be used for the cutback
  Set oInput = New DInput
  oInput.Name = "Plane"
  oInput.index = 3
  oInput.Properties = igDESCRIPTION_OPTIONAL
  oInputs.Add oInput
  
  Set oInput = Nothing
  Set oInputs = Nothing
  
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub



Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler

  Dim strIIDCope As String
  Dim oMetaData As IJDAttributeMetaData
  Dim oObject As IJDObject

  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
  
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pDefinition
  DefinitionInputs pIH

  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{CA64DDA7-EB18-4346-89C1-21C45FA8B836}"     'CStructFeature2
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  Set oObject = pDefinition
  Set oMetaData = oObject.ResourceManager
  strIIDCope = oMetaData.IID("IJUASPSCope")
  pAPDs.AddProperty "IJUASPSCope", 1, strIIDCope, "CMEvaluateIJUASPSCope", imsCOOKIE_ID_USS_LIB ' added CMEvaluate for TR#75329
  Set oMetaData = Nothing
'  pAPDs.AddProperty "IJUASPSCope", 1, "{DCD766F3-5653-45fe-B6C4-9CB760565BA0}"
'  pAPDs.AddProperty "IJUASPSCope", 1, "{87CBC1CD-8B47-4EAA-0089-F468511F0EAC}"
  Set pAPDs = Nothing

  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
   

  Set pMemberDescription = pMemberDescriptions.AddMember("CopeCutback", 1, "CMConstructCutback", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputCutback"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseCutback"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalCutback"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateCutback"
  
  Set pPropertyDescriptions = pMemberDescription
 'ouputs IJSurface = {7D82F810-D270-11D1-9558-0060973D4824}
  pPropertyDescriptions.AddProperty "ComputeCutback", 1, "{7D82F810-D270-11D1-9558-0060973D4824}", "CMComputeCutback", imsCOOKIE_ID_USS_LIB
  
  
  
  Set pMemberDescription = pMemberDescriptions.AddMember("Cope1", 2, "CMConstructCope1", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputCope1"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalCope1"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseCope1"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateCope"

  Set pPropertyDescriptions = pMemberDescription
  'outputs IJDModelBody
  pPropertyDescriptions.AddProperty "ComputeCope1", 1, "{9628D999-CEC5-11D1-B3BE-080036D85603}", "CMComputeCope1", imsCOOKIE_ID_USS_LIB
  Set pMemberDescription = Nothing
  
  Set pMemberDescription = pMemberDescriptions.AddMember("Cope2", 3, "CMConstructCope2", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputCope2"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalCope2"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseCope2"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateCope"

  Set pPropertyDescriptions = pMemberDescription
  'outputs IJDModelBody
  pPropertyDescriptions.AddProperty "ComputeCope2", 1, "{9628D999-CEC5-11D1-B3BE-080036D85603}", "CMComputeCope2", imsCOOKIE_ID_USS_LIB
  Set pMemberDescription = Nothing
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  Set pMemberDescriptions = Nothing
  Exit Sub
  
ErrorHandler:  HandleError MODULE, MT
  
End Sub


Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler
    
  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateIJUASPSCope(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateIJUASPSCope"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim oAttribsCAO As IJDAttributes
    Dim dblLength As Double, dblDepth As Double
    Dim lngRule As Long
    
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set oAttribsCAO = oSmartOcc
    
    lngRule = oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("SizingRule").Value
    
    CheckForUndefinedValueAndRaiseError oSmartOcc, lngRule, "StructACSizingRule", 20
 
    If lngRule <> 1 Then ' we set the cope length and depth only when it is "ByRule"
    'that is cope is driven by clearance
    
        'so exit if it is not driven by clearance but by explicit length and width
        Exit Sub
    End If
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    GetCopeDimensions oSuppingPart, oSuppedPart, oSuppedPort.portIndex, oAttribsCAO, dblLength, dblDepth
    'from cutback plane and cope point figure out cope length and depth
    
    'set cope length and depth
    oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Length").Value = dblLength
    oAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("Depth").Value = dblDepth
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub

' check isPlanar property to see if cope is needed.
' if not isPlanar, then check supporting member and incidence angle.

Public Sub CMConditionalCope1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalCope1"
  On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim pIJAttribsCAO As IJDAttributes
    Dim bAlwaysPlanar As Boolean

    Dim iQuadrant As Integer
    Dim strError As String
    Set oSmartOcc = pMemberDescription.CAO
    Set pIJAttribsCAO = oSmartOcc
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    
    bAlwaysPlanar = pIJAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("AlwaysPlanar").Value
    
    If bAlwaysPlanar Then
        bIsNeeded = False
    Else
        Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
        Set oSuppedPart = oSuppedPort.Part
        Set oSuppingPart = oSuppingPort.Part

        iQuadrant = GetIncidentMemberQuadrant(oSuppingPart, oSuppedPart, oSuppedPort.portIndex)
        bIsNeeded = IsCope1NeededByShapeAndIncidence(oSuppingPart, iQuadrant)
    End If

    If IsReadOnlyObject(oSuppedPart) Then
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_READONLY_OBJECT, oSmartOcc, Nothing
        Err.Raise SPS_MACRO_WARNING
    End If

    Exit Sub
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If
End Sub

Public Sub CMConstructCope1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const MT = "CMConstructCope1"
    On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim iQuadrant As Integer
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern
    Dim pIJAttribsCAO As IJDAttributes
    Dim oControlFlags As IJControlFlags
    Dim oCutter As IJDModelBody


    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set pIJAttribsCAO = oSmartOcc
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    'Used for any shape. The CMConstruct() for cope uses this. The CMEvaluate() will call shape specific cope
    'creation method. The CMConstruct calling this lighter version makes it faster. Also reduces error
    'as the specific cope creation method may require the cutback and it may not be computed when the cope
    'is constructed
    CreateSimpleTopCopeforShape oSuppingPart, oSuppedPart, oSuppedPort.portIndex, pIJAttribsCAO, pResourceManager, oCutter
    
    Set pObj = oCutter
    
    'hide the cutter
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
    
    'get the collection of operators
    Set oStructOperationPattern = oSuppedPart
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    If oCollectionOfOperators Is Nothing Then
        Set oCollectionOfOperators = New JObjectCollection
    End If

    oCollectionOfOperators.Add oCutter
    oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    

    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
Public Sub CMSetInputCope1(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputCope1"
     On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeCope1(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeCope1"
    On Error GoTo ErrorHandler

    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim iQuadrant As Integer
    Dim pIJAttribsCAO As IJDAttributes
    Dim oCutter As IJDModelBody


    Set oCutter = pObject
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOcc
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    If IsNewObject(oCutter) Or IsAttributeModified(oSmartOcc) Or _
        IsItModified(oSuppedPart) Or IsItModified(oSuppingPart) Then
        
        CreateTopCopeforShape oSuppingPart, oSuppedPart, oSuppedPort.portIndex, pIJAttribsCAO, Nothing, oCutter
    End If
        
    
        
    ' The following block of code is necessary for a copy operation where the wirebody exists without
    '   the the wirebody existing in the operation pattern collection since the CMConstruct method
    '   does not get called on a copy. So... verify the provided wirebody exists in the collection
    '   and if not set the wirebody as an operation pattern.
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern
    Set oStructOperationPattern = oSuppedPart
    Dim oExistingWire As Object
    Dim bFoundWire As Boolean
    Dim cnt
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    If oCollectionOfOperators Is Nothing Then ' Collection must exist so create it if it does not
        Set oCollectionOfOperators = New JObjectCollection
    End If
    bFoundWire = False
    For cnt = 1 To oCollectionOfOperators.count
        Set oExistingWire = oCollectionOfOperators.Item(cnt)
        If oExistingWire Is oCutter Then
            bFoundWire = True
            Exit For
        End If
    Next cnt
    If Not bFoundWire Then
        oCollectionOfOperators.Add oCutter
        oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    End If
    
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub

Public Sub CMReleaseCope1(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseCope1"
On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
Public Sub CMMigrateCope(ByVal pMemberDescription As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateCope"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oRefColl As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern
    Dim oModelBody As IJDModelBody
    Dim oExistingWire As Object
    Dim bFoundWire As Boolean
    Dim cnt

    'MsgBox METHOD
    
    Set oModelBody = pMemberDescription.object
    
    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part

    Set oStructOperationPattern = oSuppedPart
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    
    If oCollectionOfOperators Is Nothing Then ' Collection must exist so create it if it does not
        Set oCollectionOfOperators = New JObjectCollection
    End If
    bFoundWire = False
    For cnt = 1 To oCollectionOfOperators.count
        Set oExistingWire = oCollectionOfOperators.Item(cnt)
        If oExistingWire Is oModelBody Then
            bFoundWire = True
            Exit For
        End If
    Next cnt
    If Not bFoundWire Then
        oCollectionOfOperators.Add oModelBody
        oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' check isPlanar property to see if cope is needed.
' if not isPlanar, then check supporting member and incidence angle.

Public Sub CMConditionalCope2(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalCope2"
  On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim iQuadrant As Integer
    Dim strError As String
    Dim pIJAttribsCAO As IJDAttributes
    Dim bAlwaysPlanar As Boolean

    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set pIJAttribsCAO = oSmartOcc

    bAlwaysPlanar = pIJAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("AlwaysPlanar").Value
    
    If bAlwaysPlanar Then
        bIsNeeded = False
    Else
        Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
        Set oSuppedPart = oSuppedPort.Part
        Set oSuppingPart = oSuppingPort.Part
        iQuadrant = GetIncidentMemberQuadrant(oSuppingPart, oSuppedPart, oSuppedPort.portIndex)
        bIsNeeded = IsCope2NeededByShapeAndIncidence(oSuppingPart, iQuadrant)
    End If
    
    If IsReadOnlyObject(oSuppedPart) Then
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_READONLY_OBJECT, oSmartOcc, Nothing
        Err.Raise SPS_MACRO_WARNING
    End If

    Exit Sub
  
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If
End Sub

Public Sub CMConstructCope2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const MT = "CMConstructCope2"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim iQuadrant As Integer
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern
    Dim pIJAttribsCAO As IJDAttributes
    Dim oControlFlags As IJControlFlags
    Dim oCutter As IJDModelBody
    
    
    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set pIJAttribsCAO = oSmartOcc
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part

    'Used for any shape. The CMConstruct() for cope uses this. The CMEvaluate() will call shape specific cope
    'creation method. The CMConstruct calling this lighter version makes it faster. Also reduces error
    'as the specific cope creation method may require the cutback and it may not be computed when the cope
    'is constructed
    CreateSimpleBottomCopeforShape oSuppingPart, oSuppedPart, oSuppedPort.portIndex, pIJAttribsCAO, pResourceManager, oCutter
    
    
    Set pObj = oCutter
    
    'hide the cutter
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
    
    'create the collection of operators
    
    Set oStructOperationPattern = oSuppedPart
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    If oCollectionOfOperators Is Nothing Then
        Set oCollectionOfOperators = New JObjectCollection
    End If

    oCollectionOfOperators.Add oCutter
    oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    
 
  Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
Public Sub CMSetInputCope2(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputCope2"
     On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeCope2(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMComputeCope2"
    On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim iQuadrant As Integer
    Dim pIJAttribsCAO As IJDAttributes
    Dim jobject As IJDObject

    Dim oCutter As IJDModelBody

    Set oCutter = pObject
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOcc
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)

    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part

    If IsNewObject(oCutter) Or IsAttributeModified(oSmartOcc) Or _
        IsItModified(oSuppedPart) Or IsItModified(oSuppingPart) Then

        CreateBottomCopeforShape oSuppingPart, oSuppedPart, oSuppedPort.portIndex, pIJAttribsCAO, Nothing, oCutter
    End If

   
    ' The following block of code is necessary for a copy operation where the wirebody exists without
    '   the the wirebody existing in the operation pattern collection since the CMConstruct method
    '   does not get called on a copy. So... verify the provided wirebody exists in the collection
    '   and if not set the wirebody as an operation pattern.
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern
    Set oStructOperationPattern = oSuppedPart
    Dim oExistingWire As Object
    Dim bFoundWire As Boolean
    Dim cnt
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    If oCollectionOfOperators Is Nothing Then ' Collection must exist so create it if it does not
        Set oCollectionOfOperators = New JObjectCollection
    End If
    bFoundWire = False
    For cnt = 1 To oCollectionOfOperators.count
        Set oExistingWire = oCollectionOfOperators.Item(cnt)
        If oExistingWire Is oCutter Then
            bFoundWire = True
            Exit For
        End If
    Next cnt
    If Not bFoundWire Then
        oCollectionOfOperators.Add oCutter
        oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    End If

    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub

Public Sub CMReleaseCope2(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseCope2"
On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub

Public Sub CMMigrateCutback(ByVal pMemberDescription As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateCutback"
    On Error GoTo ErrorHandler
    
    'MsgBox METHOD
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Set pSmartOccurrence = pMemberDescription.CAO
    
    Dim pReferencesCollection As IJDReferencesCollection
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)
    
    Dim pSplitAxisPort As ISPSSplitAxisPort
    Set pSplitAxisPort = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(1)
    
    Dim pMemberPartPrismatic As ISPSMemberPartPrismatic
    Set pMemberPartPrismatic = pSplitAxisPort.Part
    
    Call pMemberPartPrismatic.AddCutbackSurface(pSplitAxisPort.portIndex, pMemberDescription.object)

    Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMConditionalCutback(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalCutback"
  On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim strError As String
    Dim pIJAttribsCAO As IJDAttributes
    Dim bAlwaysPlanar As Boolean
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim iPort As SPSMemberAxisPortIndex
    Dim x#, y#, z#, nx#, ny#, nz#
    
    
    Set oSmartOcc = pMemberDescription.CAO
    Set pIJAttribsCAO = oSmartOcc
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    bAlwaysPlanar = pIJAttribsCAO.CollectionOfAttributes("IJUASPSCope").Item("AlwaysPlanar").Value
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part

    If Not bAlwaysPlanar Then
        bIsNeeded = IsCutbackNeededByShape(oSuppingPort.Part)
    Else
        bIsNeeded = True
    End If
    
    If IsReadOnlyObject(oSuppedPart) Then
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_READONLY_OBJECT, oSmartOcc, Nothing
        Err.Raise SPS_MACRO_WARNING
    End If
    
    'do the check below only if both members are linear
    'basically we are checking the supported is not close to parallell with the supporting
    'such that the cope at the current end cuts the other end also. This is not valid if
    'one of the members is curved

    Dim oPosAlong As IJDPosition
    Dim oMat As IJDT4x4
    Dim oVec1 As New DVector, oVec2 As New DVector
    Dim dotProd As Double
    Set oPosAlong = GetConnectionPositionOnSupping(oSuppingPart, oSuppedPart, oSuppedPort.portIndex)
    
    oPosAlong.Get x, y, z
    
    oSuppingPart.Rotation.GetTransformAtPosition x, y, z, oMat, Nothing
    'tangent to axis at connection point
    oVec1.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    oVec1.length = 1
    
    oSuppedPart.PointAtEnd(oSuppedPort.portIndex).GetPoint x, y, z
    oSuppedPart.Rotation.GetTransformAtPosition x, y, z, oMat, Nothing
    'tangent to axis at connection point
    oVec2.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    oVec2.length = 1
    
    dotProd = Abs(oVec2.Dot(oVec1))
    'check if the axes are close to parellel
    If Abs(1 - dotProd) < angleTol Then
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_SUPPORTED_PARALELTO_SUPPORTING, oSmartOcc, Nothing
        Err.Raise SPS_MACRO_WARNING
    End If

  
  Exit Sub
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If
End Sub

Public Sub CMConstructCutback(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const MT = "CMConstructCutback"
   ' Create the SmartOccurrence

    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim oPlane As IJPlane
    Dim oControlFlags As IJControlFlags
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSurface As IJSurface
    Dim oSurfaceBody  As Object
    Dim oGeomMisc As New DGeomOpsMisc

    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set pIJAttribsCAO = oSmartOcc
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part
    
    'create the Cutback surface
    CreateCutbackForShape oSuppingPart, oSuppedPart, oSuppedPort.portIndex, pIJAttribsCAO, Nothing, oSurface
    
    'create a model surface body from the surface
    oGeomMisc.CreateModelGeometryFromGType pResourceManager, oSurface, Nothing, oSurfaceBody
    
    oSuppedPart.AddCutbackSurface oSuppedPort.portIndex, oSurfaceBody
    
    'hide the cutback plane
    Set oControlFlags = oSurfaceBody
    oControlFlags.ControlFlags(&H4) = &H4

    'returns Cutback Plane
    Set pObj = oSurfaceBody
  Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
Public Sub CMSetInputCutback(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputCutback"
     On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


Public Sub CMComputeCutback(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMComputeCutback"
    On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMembFactory As New SPSMemberFactory
    Dim x#, y#, z#, nx#, ny#, nz#
    Dim oPlane As IJPlane
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSurface As IJSurface
    Dim oSurfaceBody  As Object
    Dim oGeomMisc As New DGeomOpsMisc
    
    Dim oSuppedPort As ISPSSplitAxisPort, oSuppingPort As ISPSSplitAxisPort
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic, oSuppingPart As ISPSMemberPartPrismatic
    Dim oAxisCurve As IJCurve
    
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOcc
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppingPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPart = oSuppingPort.Part

    Set oSurfaceBody = pObject
    
    If Not TypeOf oSurfaceBody Is IJSurfaceBody Then
        'cutback created in pre v7 model
        Set oSurface = oSurfaceBody
    End If
    If IsNewObject(oSurfaceBody) Or IsAttributeModified(oSmartOcc) Or _
        IsItModified(oSuppedPart) Or IsItModified(oSuppingPart) Then
        CreateCutbackForShape oSuppingPart, oSuppedPart, oSuppedPort.portIndex, pIJAttribsCAO, Nothing, oSurface
    End If
    
    If TypeOf oSurfaceBody Is IJSurfaceBody Then
        'create a model surface body from the surface
        oGeomMisc.ModifyModelGeometryFromGType oSurface, oSurfaceBody
    End If
    ' If the assembly connection has been copied then cutback count will be 0
    If oSuppedPart.Cutbacks(oSuppedPort.portIndex).count <= 0 Then
        oSuppedPart.AddCutbackSurface oSuppedPort.portIndex, oSurfaceBody
    End If
    
    
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub


Public Sub CMReleaseCutback(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseCutback"
On Error GoTo ErrorHandler
    

  Exit Sub
ErrorHandler: HandleError MODULE, MT

End Sub





' The following methods are generic for all the Custom assembly
'
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
  Const MT = "CMFinalConstructCutout"
  On Error GoTo ErrorHandler
  
  
  Dim pDefinition As IJDSymbolDefinition
  Dim pFact As IJCAFactory
  Set pFact = New CAFactory
  Set pDefinition = pFact.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_ItemProgId
  pDefinition.CodeBase = CodeBase
    
  ' Initialize the definition
  IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
  
  ' Persistence behavior
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  'returned symbol definition
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function
ErrorHandler:      HandleError MODULE, MT
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
  IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
 ' Obsolete method.
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
 IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJSymbolVersion_GetSymbolVersion() As String
    IJSymbolVersion_GetSymbolVersion = "4.2.0.0"
End Function

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FEATUREMACROS_ERROR, "ERROR")
    
    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim j As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim pAttrDescr1 As IJAttributeDescriptor
    Dim NonStateRO As Long
            
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
'            IJUserAttributeMgmt_OnAttributeChange = "ERROR::Bad Value"
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
        End If
    End If

    ' when we change an attribute, we set the AttributeDescriptor_Changed flag
    ' This flag is supposed to be cleared on the client side after updating GOPC
    ' with the changes
    '
    ' We also set the AttributeDescriptor_ChangeAtCommit flag; this flag remains
    ' once set, to give us an idea of the attribute set that changed in this transaction
    
    pAttrToChange.AttrValue = varNewAttrValue
    If (pAttrToChange.attrName = "SizingRule") Then
        If (varNewAttrValue <> 2) Then 'By rule option for sizing rule
            'gray out the block length and width on the GOPC
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = "Length") Or (pAttrDescr.attrName = "Depth")) Then
                    pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                End If
                If ((pAttrDescr.attrName = "FlangeClearance1") Or (pAttrDescr.attrName = "Increment")) Then
                    NonStateRO = Not (AttributeDescriptor_ReadOnly)
                    pAttrDescr.AttrState = pAttrDescr.AttrState And NonStateRO
                End If
                If (pAttrDescr.attrName = "FlangeClearance2") Then
                    For j = 1 To pColl.count
                        Set pAttrDescr1 = pColl.Item(j)
                        If (pAttrDescr1.attrName = "AlwaysPlanar") Then
                            If (pAttrDescr1.AttrValue = False) Then
                                NonStateRO = Not (AttributeDescriptor_ReadOnly)
                                pAttrDescr1.AttrState = pAttrDescr.AttrState And NonStateRO
                            End If
                        End If
                    Next
                End If
            Next
        Else
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = "Length") Or (pAttrDescr.attrName = "Depth")) Then
                    NonStateRO = Not (AttributeDescriptor_ReadOnly)
                    pAttrDescr.AttrState = pAttrDescr.AttrState And NonStateRO
                End If
                If ((pAttrDescr.attrName = "FlangeClearance1") Or (pAttrDescr.attrName = "FlangeClearance2") Or (pAttrDescr.attrName = "Increment")) Then
                    pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                End If
            Next
        End If
    End If
    
    Set pColl = Nothing
    
    ' If the always planar bit is set to true, then the flange inside clearance
    ' and the web clearance are to be grayed out; they should be editable otherwise
    
    If (pAttrToChange.attrName = "AlwaysPlanar") Then
        If (varNewAttrValue = True) Then
            'gray out the flange inside clearance and web clearance
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = "FlangeClearance2") Or (pAttrDescr.attrName = "WebClearance")) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                    Else
                        pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                    End If
                End If
            Next
        Else
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = "FlangeClearance2") Or (pAttrDescr.attrName = "WebClearance")) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                        NonStateRO = Not (AttributeDescriptor_ReadOnly)
                        pAttrDescr.AttrState = pAttrDescr.AttrState And NonStateRO
                    End If
                End If
            Next
        End If
    End If
    
    IJUserAttributeMgmt_OnAttributeChange = ""
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_FEATUREMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String

    
    ' Other than Web clearance, Flange Side clearance and Flange Inside clearance,
    ' the current need is to gray out the properties on the GOPC
    ' the following code does exactly that
    
    
    Set pAttrColl = CollAllDisplayedValues

    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
        If ((pAttrDescr.attrName = "Radius")) Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        End If
    Next
        
    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
    Next
    
    bOnPreLoad = False

    IJUserAttributeMgmt_OnPreLoad = ""
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_ERROR, "ERROR")

    Dim dInputs As IJDInputs
    Dim CurrentInput As IJDInput
    Dim oAttribute As IJDAttribute
    Dim PC As DParameterContent
    Dim bvalid As Boolean
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = pIJDAttrs
    Dim oSymbolDef As IJDSymbolDefinition
    Dim ErrMessage As String
    Set oSymbolDef = oSymbolOcc.IJDSymbolDefinition(2)
    Set dInputs = oSymbolDef.IJDInputs
    Set PC = New DParameterContent
    
    Set oAttribute = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item(sAttributeName)
    If oAttribute.Value <> "" Then
        If oAttribute.AttributeInfo.Type = igString Then    ' check for string type here
        Else
            PC.UomValue = oAttribute.Value
            Set CurrentInput = Nothing
            bvalid = True
            On Error Resume Next
            Set CurrentInput = dInputs.GetInputByName(oAttribute.AttributeInfo.Name)
            If Not CurrentInput Is Nothing Then
                CurrentInput.IJDInputDuringGame.definition = oSymbolDef
                CurrentInput.IJDInputStdCustomMethod.InvokeCMCheck PC, bvalid, ErrMessage
                CurrentInput.IJDInputDuringGame.definition = Nothing
                Set oSymbolOcc = Nothing
                Set oSymbolDef = Nothing
                If bvalid = False Then
'                    UserAttributeMgmt_Validate = "Symbol CMCheck Failed"
                    UserAttributeMgmt_Validate = ErrMessage
                    Exit Function
                Else
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
' get the list of interfaces implemented by the schema from IJDAttributes
' make sure that you are not looking into a system interface
' from the input interfaceName and propertyName, get the property type from catalog info
' select case on the property types, and in there, mention the valid attribute values for each propertyName
    Dim InterfaceID As Variant
'    Dim oAttrObj As IJDAttribute
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
'    Dim oAttrCol As IJDAttributesCol
    Dim oAttrCol As IJDInfosCol
    Dim IsInterfaceFound As Boolean
    Dim AttrCount As Long
    Dim AttrType As Long
    
    Set oAttributeMetaData = pIJDAttrs
    IsInterfaceFound = False
    For Each InterfaceID In pIJDAttrs
        Set oInterfaceInfo = Nothing
        Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(InterfaceID)
        If (oInterfaceInfo.IsHardCoded = False) Then
            If (oInterfaceInfo.Name = sInterfaceName) Then
                IsInterfaceFound = True
                Exit For
            End If
        End If
    Next
    
'    Set oAttributeMetaData = Nothing
    Set oInterfaceInfo = Nothing
    
    If IsInterfaceFound = False Then
        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
    
    If sAttributeName = "FlangeClearance2" Then
        Dim pStructConn As IJAppConnection
        Dim depth As Double
        Dim SecWidth As Double
        Dim colPorts As IJElements
        Dim oCrossAttrs As IJDAttributes
        Dim pPort As IJPort
        Dim pDesignChild As IJDesignChild
        Dim pDesignParent As IJDesignParent
        Dim pMemberPart As ISPSMemberPartPrismatic
    
        Set pDesignChild = pIJDAttrs
        If Not pDesignChild Is Nothing Then
            Set pDesignParent = pDesignChild.GetParent
            If Not pDesignParent Is Nothing Then
                Set pStructConn = pDesignParent
                If Not pStructConn Is Nothing Then
                    pStructConn.enumPorts colPorts
                    ' get the cross section depth from the member part
                    If colPorts.count > 0 Then
                        Set pPort = colPorts.Item(1)
                        If Not pPort Is Nothing Then
                            Set pMemberPart = pPort.Connectable
                            If Not pMemberPart Is Nothing Then
                                Set oCrossAttrs = pMemberPart.CrossSection.definition
                                If Not oCrossAttrs Is Nothing Then
                                    depth = oCrossAttrs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If varAttributeValue > (0.5 * depth) Then
            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_FLNGCLEARANCE_MEMBDEPTH, "FlangeInsideClearance > half of Member depth")
            Exit Function
        End If
    End If
'    Set oAttrCol = pIJDAttrs.CollectionOfAttributes(InterfaceID)
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.Name = sAttributeName Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (varAttributeValue < 0#) Then
'                            UserAttributeMgmt_Validate = sAttributeName
                            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_VALUE_MUSTBE_POSITIVE, "Value must be equal to or greater than Zero")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
            End Select
        End If
    Next
    
    UserAttributeMgmt_Validate = ""
    Set oAttributeMetaData = Nothing

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Dim pReferencesCollection As IJDReferencesCollection
    Dim bIsInputMigrated As Boolean
    Dim oPoint As IJPoint
    Dim ii As Integer, eleCount As Integer
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    
    Dim oOld As Object
    Dim oReplacing() As Object

    Set pSmartOccurrence = pAggregatorDescription.CAO
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)
    
    GetPositionFromRefColl pReferencesCollection, oPoint

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    ReDim oReplacing(1 To eleCount)

    For ii = 1 To eleCount

        Set oOld = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        
        Call pMigrateHelper.ObjectsReplacing(oOld, pObjectCollectionReplacing, bIsDeleted)
    
        If Not pObjectCollectionReplacing Is Nothing Then
            bIsInputMigrated = True
            SelectReplacingObject pObjectCollectionReplacing, oPoint, oReplacing(ii)
        Else
            Set oReplacing(ii) = oOld
        End If
        
        Set oOld = Nothing
        Set pObjectCollectionReplacing = Nothing
        
    Next ii
     
    If bIsInputMigrated Then
        Call pReferencesCollection.IJDEditJDArgument.RemoveAll
        pReferencesCollection.IJDEditJDArgument.SetEntity 1, oReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), CONST_ISPSSplitAxisAlongPort, "RefColl"

    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub
