VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SurfaceTrimDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'File : SurfaceTrimDef.cls
'
'Author : Structure Team
'
'Description :
'    Module for surface trim of member part
'
'History:
' Sep07,2005  AS    Rearranged the logic in CMCOnstruct, CMEvaluateAggregateProps and CMCOmpute to
'                        avoid multiple computes. Used the new methods in FeatureServices for validation,
'                        properties calculation and geometry compute. Fixed following TRs
'                        83756,80583, 78706, 79949, 56771, 80842, 84019, 83634, 84835
' Apr07,2006 AS     Replaced oStablePort type from IJPort to Object to fix TR # 94592. Made some other changes
'                   as per coding standard and code review guidelines
' May17,2006 AS     Removed code to make the feature child of the Member Part during compute. Done in SetParents anyway. TR#98119
' Sep27,2006 AS     TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
' Nov16,2006 AS     TR#109982  Added logic to place the Cutback operator in the correct PG.
'
' Mar16,2008 SS     DI#134831  Changed the code from CreateObject() to SP3DCreateObject()
'                              as the symbol is no longer registered.
' 05/19/08    AS     TR-132553  Changes to make the standalone feature return OK, even if the supg surf (optional) is missing.
'                    Add to TDL (with warning) instead of deleting the surface
'
'**********************************************************************************************************************

Option Explicit


Private Const MODULE = "SurfaceTrimDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFeatureMacros.SurfaceTrimDef"
Const m_ItemName As String = "SPSFeatureMacros.SurfaceTrimDef"
Private Const strSourceFile As String = "SurfaceTrimDef.cls"

Dim bOnPreLoad As Boolean
Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1
Private m_oLocalizer As IJLocalizer
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructFeatureSOHelper

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
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
  pIH.Definition = pDefinition
  DefinitionInputs pIH

  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{CA64DDA7-EB18-4346-89C1-21C45FA8B836}"     'CStructFeature2
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  
  pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructAggregator"
  pAD.SetCMAdaptClone imsCOOKIE_ID_USS_LIB, "CMAdaptCloneAggregator"
  pAD.SetCMIdentifyClone imsCOOKIE_ID_USS_LIB, "CMIdentifyCloneAggregator"
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  Set oObject = pDefinition
  Set oMetaData = oObject.ResourceManager
  strIIDCope = oMetaData.IID("IJUASPSSurfaceTrim")
  pAPDs.AddProperty "IJUASPSSurfaceTrim", 1, strIIDCope, "CMEvaluateAggregatorProperties", imsCOOKIE_ID_USS_LIB
  Set oMetaData = Nothing
  Set pAPDs = Nothing

  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll

  Set pMemberDescription = pMemberDescriptions.AddMember("SurfaceTrimCutback", 1, "CMConstructSurfaceTrimCutback", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSurfaceTrimCutback"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseSurfaceTrimCutback"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSurfaceTrimCutback"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSurfaceTrimCutback"
  
  pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructSurfaceTrim"
  pMemberDescription.SetCMCount imsCOOKIE_ID_USS_LIB, "CMSurfaceTrimCount"
  pMemberDescription.SetCMIdentifyClone imsCOOKIE_ID_USS_LIB, "CMIdentifyCloneSurfaceTrim"
  pMemberDescription.SetCMAdaptClone imsCOOKIE_ID_USS_LIB, "CMAdaptOnPasteSurfaceTrim"
  
  Set pPropertyDescriptions = pMemberDescription
 'outputs IJSurface = "{7D82F810-D270-11D1-9558-0060973D4824}"
  pPropertyDescriptions.AddProperty "ComputeSurfaceTrimCutback", 1, ConstIJSurface, "CMComputeSurfaceTrimCutback", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  Set pMemberDescriptions = Nothing
  Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMEvaluateAggregatorProperties(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateAggregatorProperties"
    On Error GoTo ErrorHandler
        
    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object
    Dim oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper
    Dim oSmartOcc As IJSmartOccurrence

   'Get smartoccurence, helper and parents
    Set oSmartOcc = pPropertyDescription.CAO
    Set oFeatureHelper = Me
    helperStat = oFeatureHelper.GetParents(oSmartOcc, oOperand, oOperators)
    If oOperand Is Nothing Then
        'we are definitely bad, beyond repair. delete now.
        Dim iIJDObject As IJDObject
        Set iIJDObject = oSmartOcc
        iIJDObject.Remove
    End If
    
    'early exit if inputs are not good. Don't raise E_FAIL as we are checking and creating a warning later in CMconditional
    If oOperators Is Nothing Then
        Exit Sub
    End If
    If oOperators.count = 0 Then
        Exit Sub
    End If
    
'    If helperStat <> StructSOCInputHelper_Ok Then
'        If helperStat = StructSOCInputHelper_BadNumberOfObjects Then
'            If oOperand Is Nothing Then
'                Dim iIJDObject As IJDObject
'                Set iIJDObject = oSmartOcc
''                iIJDObject.Remove
'            End If
'        End If
'
'        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_INVALIDINPUTS_TRIMCUTBACK, oSmartOcc, Nothing
'        Err.Raise E_FAIL
'    End If
    
    'Everything OK, inputs are good.
    'feature services
    Dim oMemberFactory As New SPSMemberFactory
    Dim oFeatureServices As ISPSMemberFeatureServices
    
    Dim oPart As ISPSMemberPartPrismatic
    Dim oPort As ISPSSplitAxisPort
    Dim WhichEnd As SPSMemberAxisPortIndex
    Dim oOperator As Object
    
    Dim oAttribsCAO As IJDAttributes
    Dim isSquareEnd As Boolean
    Dim WebAngle As Double
    Dim FlangeAngle As Double
    Dim oAtrCollCAO As IJDAttributesCol
    
    Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices
    Set oPort = oOperand
    Set oPart = oPort.Part
    WhichEnd = oPort.portIndex
    Set oOperator = oOperators.Item(1)
    
    Set oAttribsCAO = oSmartOcc
    Set oAtrCollCAO = oAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim")
    isSquareEnd = oAtrCollCAO.Item("SquaredEnd").Value
    
    oFeatureServices.GetTrimSurfacePartAngles oPart, oOperator, WhichEnd, isSquareEnd, WebAngle, FlangeAngle
  
    oAtrCollCAO.Item("FlangeAngle").Value = FlangeAngle
    oAtrCollCAO.Item("WebAngle").Value = WebAngle
    oAtrCollCAO.Item("TrimEnd").Value = WhichEnd
    
    Set oOperators = Nothing
    Set oOperand = Nothing
    Set oFeatureHelper = Nothing
    Set oSmartOcc = Nothing
    Set oMemberFactory = Nothing
    Set oFeatureServices = Nothing
    Set oPart = Nothing
    Set oPort = Nothing
    Set oOperator = Nothing
    Set oAttribsCAO = Nothing
    Exit Sub

ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub



Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMMigrateSurfaceTrimCutback(ByVal pMemberDescription As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateSurfaceTrimCutback"
    On Error GoTo ErrorHandler
    
    'MsgBox METHOD
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Set pSmartOccurrence = pMemberDescription.CAO
    
    Dim pMemberPartPrismatic As ISPSMemberPartPrismatic
    Dim oSuppingSurf As Object    'IJSurface
    Dim oSuppedPort As ISPSSplitAxisPort
    
    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object
    Dim oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper
    Set oFeatureHelper = Me
    helperStat = oFeatureHelper.GetParents(pSmartOccurrence, oOperand, oOperators)
    If helperStat <> StructSOCInputHelper_Ok Then
        GoTo ErrorHandler
    End If
   
    Set oSuppedPort = oOperand
    Set oSuppingSurf = oOperators.Item(1)
    Set pMemberPartPrismatic = oSuppedPort.Part
    
'    Call pMemberPartPrismatic.AddCutbackSurface(GetCutbackEnd(pMemberPartPrismatic, oSuppingSurf), pMemberDescription.object) TR#75102
    Call pMemberPartPrismatic.AddCutbackSurface(oSuppedPort.portIndex, pMemberDescription.object)

    Set oOperand = Nothing
    Set oOperators = Nothing
    Set oSuppingSurf = Nothing
    Set pMemberPartPrismatic = Nothing
    Set oSuppedPort = Nothing
    Set pSmartOccurrence = Nothing
    
    Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMConditionalSurfaceTrimCutback(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSurfaceTrimCutback"
  On Error GoTo ErrorHandler

    Dim strError As String
    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object
    Dim oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper
    Dim oSmartOcc As IJSmartOccurrence

    bIsNeeded = True
    
    Set oSmartOcc = pMemberDescription.CAO
    Set oFeatureHelper = Me
    helperStat = oFeatureHelper.GetParents(oSmartOcc, oOperand, oOperators)
    If oOperand Is Nothing Then
        bIsNeeded = False
        PostMissingInputsError oSmartOcc, strSourceFile, strError
    ElseIf oOperators Is Nothing Then
        bIsNeeded = False
        PostMissingInputsError oSmartOcc, strSourceFile, strError
    ElseIf oOperators.count = 0 Then
        bIsNeeded = False
        PostMissingInputsError oSmartOcc, strSourceFile, strError
    End If
    
  Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, strSourceFile, MT, strError).Number
End Sub

Public Sub CMConstructSurfaceTrimCutback(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const MT = "CMConstructSurfaceTrimCutback"
   ' Create the SmartOccurrence

    Dim oControlFlags As IJControlFlags
    
    'create the Cutback Plane
    Dim oPlane As IJPlane
    Dim oGeomFactory As New GeometryFactory
    Dim oIJGeometryMisc As IJGeometryMisc
    Dim oBodySheet As Object
    Dim oSmartOcc As IJSmartOccurrence
    Dim oObjFeature As IJDObject
    Dim oObjCutback As IJDObject

    Set oSmartOcc = pMemberDescription.CAO
    
    'MsgBox "CMConstructSurfaceTrimCutback"

    Set oIJGeometryMisc = New DGeomOpsMisc
    Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, 0#, 0#, 0#, 1#, 0#, 0#)
    oIJGeometryMisc.CreateModelGeometryFromGType pResourceManager, oPlane, Nothing, oBodySheet
    
    'hide the cutback plane
    Set oControlFlags = oBodySheet
    oControlFlags.ControlFlags(&H4) = &H4

    'TR#109982  All Trim objects are not placed in the correct PG
    'Now set the PG of the Cutback operator same as the Feature
    Set oObjFeature = oSmartOcc
    Set oObjCutback = oBodySheet
    oObjCutback.PermissionGroup = oObjFeature.PermissionGroup
    
    'returns Cutback Plane
    Set pObj = oBodySheet
    
    Set oPlane = Nothing
    Set oBodySheet = Nothing
    Set oGeomFactory = Nothing
    Set oIJGeometryMisc = Nothing
    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub
Public Sub CMSetInputSurfaceTrimCutback(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputSurfaceTrimCutback"
     On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSurfaceTrimCutback(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMComputeSurfaceTrimCutback"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingSurf As Object  'IJSurface
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim bIsNeeded As Boolean

    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOcc
    
   
    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object
    Dim oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper
    Dim count As Integer

   ' MsgBox "CMComputeSurfaceTrimCutback914"
     
    Set oFeatureHelper = Me
    helperStat = oFeatureHelper.GetParents(oSmartOcc, oOperand, oOperators)
    If helperStat <> StructSOCInputHelper_Ok Then
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_INVALIDINPUTS_TRIMCUTBACK, oSmartOcc, Nothing
        Err.Raise SPS_MACRO_WARNING
    End If
    
    count = oOperators.count
    Set oSuppedPort = oOperand
    Set oSuppingSurf = oOperators.Item(count)
    Set oSuppedPart = oSuppedPort.Part

    ' set the feature as a child of operand
    'MakeFeatureChildOfPart oSuppedPart, oSmartOcc 'now done in command DI:98119

    Dim oFeatureServices As ISPSMemberFeatureServices
    Dim oMemberFactory As New SPSMemberFactory
    Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices
   
    Dim WhichEnd As SPSMemberAxisPortIndex
    Dim Clearance As Double
    Dim isSquareEnd As Boolean
    Dim allOK As Boolean
    
    Clearance = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("Offset").Value
    isSquareEnd = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("SquaredEnd").Value
    WhichEnd = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSurfaceTrim").Item("TrimEnd").Value
   
    If IsReadOnlyObject(oSuppedPart) Then
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_READONLY_OBJECT, oSmartOcc, Nothing
        Err.Raise E_FAIL
    Else
        oFeatureServices.GetTrimSurfacePartGeometricCondition oSuppedPart, oSuppingSurf, WhichEnd, allOK
        If Not allOK Then
            SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_ERROR_GEOMETRYCREATION, oSmartOcc, Nothing
            Err.Raise SPS_MACRO_WARNING
        End If
    End If

    'now update pObject
    oFeatureServices.UpdateTrimmingSurface oSuppedPart, oSuppingSurf, WhichEnd, isSquareEnd, Clearance, pObject, allOK
    
    If allOK Then
        Dim oTmpObj As Object
        oSuppedPart.AddCutbackSurface WhichEnd, pObject
    Else
            SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_ERROR_GEOMETRYCREATION, oSmartOcc, Nothing
            Err.Raise SPS_MACRO_WARNING
        
    End If

    Set oOperand = Nothing
    Set oOperators = Nothing
    Set oFeatureHelper = Nothing
    Set oSuppingSurf = Nothing
    Set oSuppedPart = Nothing
    Set oSuppedPort = Nothing
    Set oSmartOcc = Nothing
    Set pIJAttribsCAO = Nothing
    Set oFeatureServices = Nothing
    Set oMemberFactory = Nothing
    
  Exit Sub

ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If
End Sub
Public Sub CMReleaseSurfaceTrimCutback(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseSurfaceTrimCutback"
On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
' The following methods are generic for all the Custom assembly
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

Public Function IJStructFeatureSOHelper_GetParents(ByVal Feature As SPSMembers.IJSmartOccurrence, Operand As Object, Operators As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus
    Const METHOD = "IJStructFeatureSOHelper_GetParents"
    On Error GoTo ErrorHandler
    
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim i As Integer
    Dim tmpOperators As IJElements
    
    Set tmpOperators = New JObjectCollection
    Set oRefColl = GetRefCollFromSmartOccurrence(Feature)
    
    If oRefColl.IJDEditJDArgument.GetCount < 1 Then
        IJStructFeatureSOHelper_GetParents = StructSOCInputHelper_BadNumberOfObjects
        Exit Function
    End If

    Set Operand = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    
    If oRefColl.IJDEditJDArgument.GetCount > 1 Then
        tmpOperators.Add oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
        Set Operators = tmpOperators
    End If
    
    ' add validation code here to check whether the parents are valid for the feature or not
    
    Set tmpOperators = Nothing
    Set oRefColl = Nothing
    IJStructFeatureSOHelper_GetParents = StructSOCInputHelper_Ok
    Exit Function

ErrorHandler:
    IJStructFeatureSOHelper_GetParents = StructSOCInputHelper_UnexpectedError
    HandleError MODULE, METHOD
End Function

Public Function IJStructFeatureSOHelper_OperandPrompt(ByVal Feature As SPSMembers.IJSmartOccurrence) As String
    IJStructFeatureSOHelper_OperandPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECTMEMBERS, "Select member(s)") ' TR#73170
End Function

Public Function IJStructFeatureSOHelper_SetParents(ByVal Feature As SPSMembers.IJSmartOccurrence, ByVal Operand As Object, ByVal Operators As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus
    Const METHOD = "IJStructFeatureSOHelper_SetParents"
    On Error GoTo ErrorHandler
    
    Dim status As StructSOCInputHelperStatus

    Dim oRefColl As IJDReferencesCollection
    Dim oSymFactory As IJDSymbolEntitiesFactory
    Dim iDesignParent As IJDesignParent
    Dim iSuppedPort As ISPSSplitAxisPort
    Dim iObject As IJDObject
    Dim count As Integer
    
    status = StructSOCInputHelper_Ok
    count = Operators.count

    If Operand Is Nothing Then
        status = StructSOCInputHelper_BadNumberOfObjects
    ElseIf Not TypeOf Operand Is ISPSSplitAxisEndPort Then
        status = StructSOCInputHelper_InvalidTypeOfObject
    ElseIf Operators Is Nothing Then
        status = StructSOCInputHelper_BadNumberOfObjects
    ElseIf count < 1 Then
        status = StructSOCInputHelper_BadNumberOfObjects
    
    Else
        ' set the feature as a child of operand
        Set iSuppedPort = Operand
        Set iDesignParent = iSuppedPort.Part
        iDesignParent.AddChild Feature
        AddPartFeatureRln iDesignParent, Feature
        
        AssignPG Feature, iDesignParent
        
        Set iDesignParent = Nothing
        Set iSuppedPort = Nothing

        Set oRefColl = GetRefCollFromSmartOccurrence(Feature)
    
        If oRefColl Is Nothing Then
            Set iObject = Feature
            Set oSymFactory = New DSymbolEntitiesFactory
            Set oRefColl = oSymFactory.CreateEntity(ReferencesCollection, iObject.ResourceManager)
            Set iObject = Nothing
            Set oSymFactory = Nothing
            ConnectSmartOccurrence Feature, oRefColl
        Else
            oRefColl.IJDEditJDArgument.RemoveAll
        End If
        
        oRefColl.IJDEditJDArgument.SetEntity 1, Operand, CONST_ISPSSplitAxisEndPort, "RefColl"
        oRefColl.IJDEditJDArgument.SetEntity 2, Operators.Item(1), ConstIJSurface, "SurfaceToRC_DEST"

        Set oRefColl = Nothing
    End If

    IJStructFeatureSOHelper_SetParents = status
    
    Exit Function

ErrorHandler:
    IJStructFeatureSOHelper_SetParents = StructSOCInputHelper_UnexpectedError
    HandleError MODULE, METHOD
End Function

Public Function IJStructFeatureSOHelper_SupportsVariableOperatorCount(ByVal Feature As SPSMembers.IJSmartOccurrence) As Boolean
    IJStructFeatureSOHelper_SupportsVariableOperatorCount = False ' for V6
End Function

Public Function IJStructFeatureSOHelper_ValidateParents(ByVal Feature As SPSMembers.IJSmartOccurrence, ByVal Operand As Object, ByVal Operators As SPSMembers.IJElements, ByVal ValMode As Long, RelationOperand As Object, RelationOperators As SPSMembers.IJElements, NextOperatorPrompt As String) As SPSMembers.StructSOCInputHelperStatus
    Const METHOD = "IJStructFeatureSOHelper_ValidateParents"
    On Error GoTo ErrorHandler
    
    Dim oOperator As Object
    Dim oPart As ISPSMemberPartPrismatic
    Dim WhichEnd As SPSMemberAxisPortIndex
    Dim oStablePort As Object 'IJPort Replaced IJPort with Object to make sure obejcts which do not implement stable port also work

    
    'MsgBox "Inside Validate parents"
    
    IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_UnexpectedError
    
    NextOperatorPrompt = ""
        
    If Operand Is Nothing Then
        IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_BadNumberOfObjects
        NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECTMEMB_TRIMMED, "Select Member to be Trimmed")
    ElseIf Not (TypeOf Operand Is ISPSMemberPartPrismatic) Then
        IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
        NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECTMEMB_TRIMMED, "Select Member to be Trimmed")
    ElseIf IsReadOnlyObject(Operand) Then
        IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_NoWriteAccess
        NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECTMEMB_TRIMMED, "Select Member to be Trimmed")
    Else
        'Now check the operators also
        If Operators Is Nothing Then
            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_BadNumberOfObjects
            NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_CUTSURFACE, "Select Cutting Surface")
        Else
            'one and and only one operator. We can proceed
            If Operators.count = 0 Then
                IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_BadNumberOfObjects
            Else
                Set oOperator = Operators(Operators.count)    'only consider the first surface
                If OwnedByConnectable(Operand, oOperator) Then
                    IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
                Else
                    'Selected surface is a valid operator.
                    Dim oMemberFactory As New SPSMemberFactory
                    Dim oFeatureServices As ISPSMemberFeatureServices
                    Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices
                    Dim allOK As Boolean
                    
                    'First do the validation by type
                    If Not IsOperatorTypeValidForTrim(oOperator) Then
                        IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
                        GoTo AfterGeometryValidation
                    End If
  
                    'Now do the geometric validation
                    oFeatureServices.GetTrimSurfacePartGeometricCondition Operand, oOperator, WhichEnd, allOK
                    
                    If allOK Then
                        Set oStablePort = GetStablePort(oOperator)
                        
                        'Check if the parents participate in another Trim
                        Dim oMemberChildren As IJDObjectCollection
                        Dim oParentSys As IJDesignParent
                        Dim oMemberChild As Object
                        Dim oOperatorTemp As SPSMembers.IJElements
                        Dim oOperandTemp As ISPSSplitAxisPort
                        Dim oFeatureHelperTemp As IJStructFeatureSOHelper
                        Dim oSmartOccurenceTemp As IJSmartOccurrence
                        Dim oSmartItemTemp As IJSmartItem
                        Dim helpStat As SPSMembers.StructSOCInputHelperStatus
                        Dim bFeatureExists As Boolean
                        Dim strDefProgId As String
                        
                        bFeatureExists = False
                        
                        Set oParentSys = Operand
                        oParentSys.GetChildren oMemberChildren, ""
                        
                        For Each oMemberChild In oMemberChildren
                            If TypeOf oMemberChild Is IJStructFeature2 Then
                                'If there is another feature on the same part.
                                Set oSmartOccurenceTemp = oMemberChild
                                Set oSmartItemTemp = oSmartOccurenceTemp.ItemObject
                                strDefProgId = oSmartItemTemp.Definition
                                Set oFeatureHelperTemp = SP3DCreateObject(strDefProgId)
                                       
                                helpStat = oFeatureHelperTemp.GetParents(oSmartOccurenceTemp, oOperandTemp, oOperatorTemp)
                                
                                If helpStat = StructSOCInputHelper_Ok Then
                                    'Now check if the other feature's parents are same, its not allowed
                                    If oOperandTemp.portIndex = WhichEnd Then
                                        If Not oOperatorTemp Is Nothing Then
                                            If oOperatorTemp.count = 1 Then
                                                If oOperatorTemp.Item(1) Is oStablePort Then
                                                    'oops. another trim already exists
                                                    bFeatureExists = True
                                                End If
                                            End If 'operator count
                                        End If 'operato nothing
                                    End If ' operand end
                                End If ' helpstat
                            End If 'structfeature
                        Next
                        
                        If bFeatureExists Then
                            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_DuplicateObject
                        Else
                            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_Ok
                        End If
                        
                        Set oMemberChildren = Nothing
                        Set oParentSys = Nothing
                        Set oMemberChild = Nothing
                        Set oOperatorTemp = Nothing
                        Set oOperandTemp = Nothing
                        Set oMemberChild = Nothing
                        Set oFeatureHelperTemp = Nothing
                        Set oSmartOccurenceTemp = Nothing
                        Set oSmartItemTemp = Nothing
                       
                    End If  'allOK
                    
                    Set oMemberFactory = Nothing
                    Set oFeatureServices = Nothing
                End If
            End If 'operators loop
        End If 'operators count loop
        
AfterGeometryValidation:
        If IJStructFeatureSOHelper_ValidateParents <> StructSOCInputHelper_Ok Then
            NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_CUTSURFACE, "Select Cutting Surface")
        End If
    End If 'main validation loop

    'Now add to the relation
    If IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_Ok Then
        If RelationOperators Is Nothing Then
            Set RelationOperators = New JObjectCollection
        Else
            RelationOperators.Clear
        End If
        RelationOperators.Add GetStablePort(oOperator)
        ' now populate the RelationOperand as a port
        Set oPart = Operand
        Set RelationOperand = oPart.AxisPort(WhichEnd)
    End If
        
    Set oOperator = Nothing
    Set oPart = Nothing
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler

    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
            
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
        End If
    End If

    IJUserAttributeMgmt_OnAttributeChange = ""
Exit Function
ErrorHandler:
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FEATUREMACROS_ERROR, "ERROR")
    HandleError MODULE, METHOD
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

    ' Other than Clearance, the current need is to gray out the properties
    ' on the GOPC. the following code does exactly that
    
    Set pAttrColl = CollAllDisplayedValues

    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
        If ((pAttrDescr.attrName = "FlangeAngle") Or (pAttrDescr.attrName = "WebAngle") Or (pAttrDescr.attrName = "TrimEnd")) Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        End If
        Set pAttrDescr = Nothing
    Next
        
    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
        Set pAttrDescr = Nothing
    Next
    
    Set pAttrColl = Nothing
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
                CurrentInput.IJDInputDuringGame.Definition = oSymbolDef
                CurrentInput.IJDInputStdCustomMethod.InvokeCMCheck PC, bvalid, ErrMessage
                CurrentInput.IJDInputDuringGame.Definition = Nothing
                Set oSymbolOcc = Nothing
                Set oSymbolDef = Nothing
                If bvalid = False Then
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
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
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
    
    Set oInterfaceInfo = Nothing
    
    If IsInterfaceFound = False Then
        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.Name = sAttributeName Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (varAttributeValue < 0#) Then
                            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_VALUE_BE_GREATERTHAN_ZERO, "Value must be >= 0")
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

' dummy CM-functions
Public Sub CMConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMConstructAggregator"
On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMAdaptCloneAggregator(ByVal pAD As IJDAggregatorDescription)
  Const METHOD = "CMAdaptCloneAggregator"
  On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

Public Sub CMIdentifyCloneAggregator(ByVal pAD As IJDAggregatorDescription, ByRef pObject As Object)
  Const METHOD = "CMIdentifyCloneAggregator"
  On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateAggregator(oAggregatorDesc As IJDAggregatorDescription, oMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateAggregator"
  On Error GoTo ErrorHandler
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oReferencesCollection As IJDReferencesCollection
    
    Dim oObjectsReplacing() As Object
    Dim bIsInputMigrated As Boolean
    
    Dim oOldMemberPart As ISPSMemberPartPrismatic
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic
    Dim oSplitAxisPort As SPSMembers.ISPSSplitAxisPort
    Dim oDesignParent As IJDesignParent
    Dim oDesignChild As IJDesignChild
    
    Set oSmartOccurrence = oAggregatorDesc.CAO
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oSmartOccurrence)
    
    MigrateRefColl oReferencesCollection, oMigrateHelper, oObjectsReplacing, bIsInputMigrated
    
    If bIsInputMigrated Then
        'If any of the inputs are indeed migrated, reset them on the ref coll
        Set oSplitAxisPort = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(1)
        Set oOldMemberPart = oSplitAxisPort.Part
        Set oSplitAxisPort = Nothing
        
        Call oReferencesCollection.IJDEditJDArgument.RemoveAll
        oReferencesCollection.IJDEditJDArgument.SetEntity 1, oObjectsReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        oReferencesCollection.IJDEditJDArgument.SetEntity 2, oObjectsReplacing(2), ConstIJSurface, "SurfaceToRC_DEST"
        
        'Now make sure surface trim is a design child of the new Member Part in case the Member Part is not the same as the old
        Set oDesignChild = oSmartOccurrence
        Set oDesignParent = oDesignChild.GetParent
        If Not oDesignParent Is Nothing Then
            If TypeOf oDesignParent Is ISPSMemberPartPrismatic Then
                Set oSplitAxisPort = oObjectsReplacing(1)
                Set oMemberPartPrismatic = oSplitAxisPort.Part
                If Not oOldMemberPart Is oMemberPartPrismatic Then
                    'Only if new memebr part is not the same
                    MakeFeatureChildOfPart oMemberPartPrismatic, oSmartOccurrence
                End If
            End If
        End If
        
    End If
    
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructSurfaceTrim(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructSurfaceTrim"
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMSurfaceTrimCount(ByVal pMemberDescription As IJDMemberDescription, ByRef lCount As Long)
    Const METHOD = "CMSurfaceTrimCount"
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMIdentifyCloneSurfaceTrim(ByVal pMemberDescription As IJDMemberDescription, ByRef pObject As Object)
  Const METHOD = "CMIdentifyCloneSurfaceTrim"
  On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

Public Sub CMAdaptOnPasteSurfaceTrim(pMD As IJDMemberDescription)
  Const METHOD = "CMAdaptOnPasteSurfaceTrim"
  On Error GoTo ErrorHandler
  
  AdaptOnPasteCutbackToPart pMD
    
Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

