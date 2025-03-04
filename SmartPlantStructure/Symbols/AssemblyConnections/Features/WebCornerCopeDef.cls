VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCornerCopeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'File : WebCornerCopeDef.cls
'
'Author : PN
'
'Description :
'    Module for corner cope of member part
'
'History:
' 02/14/05    MH     adapted to call FeatureServices for Validate and Construct
' 27/Sep/06   AS     TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
' 31/Oct/06   MH     TR 108899 in Validate  use GetTrimSurfacePartGeometricCondition to determine which end
' 07/06/07    MH     TR 121116  added  CheckForUndefinedValueAndRaiseError
' 05/19/08    AS     TR-132553  Changes to make the standalone feature return OK, even if the supg surf (optional) is missing.
'                    Add to TDL (with warning) instead of deleting the surface
'**********************************************************************************************************************

Option Explicit
Private Const MODULE = "WebCornerCopeDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFeatureMacros.WebCornerCopeDef"
Const m_ItemName As String = "SPSFeatureMacros.WebCornerCopeDef"
Private Const strSourceFile As String = "WebCornerCopeDef.cls"

Dim bOnPreLoad As Boolean
Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1
Private m_oLocalizer As IJLocalizer 'DI 80096
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
  strIIDCope = oMetaData.IID("IJUASPSRectangularWebCope")
  pAPDs.AddProperty "IJUASPSRectangularWebCope", 1, strIIDCope, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
  Set oMetaData = Nothing
  Set pAPDs = Nothing

  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll

  Set pMemberDescription = pMemberDescriptions.AddMember("CornerCope", 1, "CMConstructCornerCope", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputCornerCope"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalCornerCope"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseCornerCope"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateCornerCope"

  pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructCornerCope"
  pMemberDescription.SetCMCount imsCOOKIE_ID_USS_LIB, "CMCornerCopeCount"
  pMemberDescription.SetCMIdentifyClone imsCOOKIE_ID_USS_LIB, "CMIdentifyCloneCornerCope"
  pMemberDescription.SetCMAdaptClone imsCOOKIE_ID_USS_LIB, "CMAdaptOnPasteCornerCope"
  
  Set pPropertyDescriptions = pMemberDescription
  'outputs ijwirebody
  pPropertyDescriptions.AddProperty "ComputeCornerCope", 1, "{4BB74DC3-D601-11D2-B50F-080036D85603}", "CMComputeCornerCope", imsCOOKIE_ID_USS_LIB
  Set pMemberDescription = Nothing
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  Set pMemberDescriptions = Nothing
  Exit Sub
  
ErrorHandler:  HandleError MODULE, MT
  
End Sub
Public Sub CMEvaluateCAO(pPropertyDescription As IJDPropertyDescription, pObject As Object)
 
 'Compute the cope depth and width properties on IJUASPSRectangularWebCope
 'These are used later by the member wire body computation.
 
    Const METHOD = "CMEvaluateCAO"
    On Error GoTo ErrorHandler
  
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingSurface1 As Object, oSuppingSurface2 As Object

    Dim pIJAttribsCAO As IJDAttributes
    Dim oSuppedPort As ISPSSplitAxisPort

    Dim oCollProxy As CollectionProxy
    Dim oMemberFactory As SPSMemberFactory
    Dim oFeatureServices As ISPSMemberFeatureServices

    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object, oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper

    Dim LengthClr As Double, DepthClr As Double
    Dim RoundIncrement As Double
    Dim CopeEdge As Long, CopeLength As Double, CopeDepth As Double
    Dim ComputeOk As Boolean
    
    Set oSmartOcc = pObject     ' output "object" is the Custom Occurrence
    Set oFeatureHelper = Me
    helperStat = oFeatureHelper.GetParents(oSmartOcc, oOperand, oOperators)
    If oOperand Is Nothing Then
        'we are definitely bad, beyond repair. delete now.
        Dim iIJDObject As IJDObject
        Set iIJDObject = oSmartOcc
        iIJDObject.Remove
    End If
    
    'early exit if inputs are not good. Don't raise E_FAIL as we are checking and creating a wanriing later in CMconditional
    If oOperators Is Nothing Then
        Exit Sub
    End If
    If oOperators.count = 0 Then
        Exit Sub
    End If
    
'    If helperStat <> StructSOCInputHelper_Ok Then
'        If helperStat = StructSOCInputHelper_BadNumberOfObjects Then
'            Dim iIJDObject As IJDObject
'            Set iIJDObject = oSmartOcc
'            iIJDObject.Remove
'        End If
'
'        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_INVALIDINPUTS_TRIMCUTBACK, oSmartOcc, Nothing
'        Err.Raise E_FAIL
'    End If

    Set pIJAttribsCAO = oSmartOcc

    Set oSuppedPort = oOperand
    Set oSuppingSurface1 = oOperators.Item(1)
    If oOperators.count > 1 Then
        Set oSuppingSurface2 = oOperators.Item(2)
    End If
    Set oSuppedPart = oSuppedPort.Part

'    If IsNewObject(oWireBody) Or IsAttributeModified(oSmartOcc) Or _
'        IsItModified(oSuppedPart) Or IsItModified(oSuppingSurface1) Or IsItModified(oSuppingSurface2) Then

        Set oCollProxy = pIJAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope")

        LengthClr = oCollProxy.Item("CopeLengthClearance").Value
        DepthClr = oCollProxy.Item("CopeDepthClearance").Value
        RoundIncrement = oCollProxy.Item("RoundingIncrement").Value

        Set oMemberFactory = New SPSMemberFactory
        Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices

        oFeatureServices.ComputeRectangularWebCornerCopeProperties oSuppedPart, oSuppedPort.portIndex, _
                    oSuppingSurface1, LengthClr, oSuppingSurface2, DepthClr, RoundIncrement, _
                    CopeEdge, CopeLength, CopeDepth, ComputeOk

        If Not ComputeOk Then
            Err.Raise SPS_MACRO_WARNING
        End If

        oCollProxy.Item("CopeLength").Value = CopeLength
        oCollProxy.Item("CopeDepth").Value = CopeDepth
        oCollProxy.Item("CopeEdge").Value = CopeEdge
        oCollProxy.Item("CopeEnd").Value = oSuppedPort.portIndex

'    End If

    Set oSuppingSurface1 = Nothing
    Set oSuppingSurface2 = Nothing
    Set oSuppedPart = Nothing
    Set oSuppedPort = Nothing
    Set oSmartOcc = Nothing
    Set pIJAttribsCAO = Nothing
    Set oOperand = Nothing
    Set oOperators = Nothing
    Set oCollProxy = Nothing

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalCornerCope(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalCornerCope"
    On Error GoTo ErrorHandler

    Dim strError As String

    'add the TDL record here. This way we can make the feature operator not needed. SO member compute will go through W/o this,
    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object
    Dim oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper
    Dim oSmartOcc As IJSmartOccurrence

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
    bIsNeeded = True

    Exit Sub
   
ErrorHandler:
    Err.Raise ReportError(Err, strSourceFile, MT, strError).Number
End Sub

Public Sub CMConstructCornerCope(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructCornerCope"
    On Error GoTo ErrorHandler

    Dim oLine As IJLine ' use the simplest geometry possible just to construct a WireBody
    Dim pGeometryFactory As New GeometryFactory
    Dim oIJGeometryMisc As IJGeometryMisc
    Dim oWireBody As Object

    Set oLine = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, 0#, 0#, 0#, 1#, 0#, 0#)

    Set oIJGeometryMisc = New DGeomOpsMisc
    oIJGeometryMisc.CreateModelGeometryFromGType pResourceManager, oLine, Nothing, oWireBody
    
    Set pObj = oWireBody
    
    Exit Sub

ErrorHandler:      HandleError MODULE, MT
End Sub
Public Sub CMSetInputCornerCope(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputCornerCope"
     On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeCornerCope(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeCornerCope"
    On Error GoTo ErrorHandler

    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPart As ISPSMemberPartPrismatic

    Dim pIJAttribsCAO As IJDAttributes
    Dim oSuppedPort As ISPSSplitAxisPort

    Dim oCollProxy As CollectionProxy
    Dim oMemberFactory As SPSMemberFactory
    Dim oFeatureServices As ISPSMemberFeatureServices

    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object, oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper

    Dim Radius As Double, RadiusType As Long
    Dim WebRadius As Double, WebRadiusType As Long
    Dim CopeEdge As Long, CopeLength As Double, CopeDepth As Double
    ' TR#71317
    Dim LengthClr As Double, DepthClr As Double, RoundIncrement As Double, ComputeOk As Boolean
    Dim oSuppingSurface1 As Object, oSuppingSurface2 As Object
    
    Dim errNumber As Long

    Set oSmartOcc = pPropertyDescriptions.CAO
    Set oFeatureHelper = Me
    helperStat = oFeatureHelper.GetParents(oSmartOcc, oOperand, oOperators)
    If helperStat <> StructSOCInputHelper_Ok Then
'        GoTo ErrorHandler
        SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_INVALID_CONFIG, oSmartOcc, Nothing
        Err.Raise SPS_MACRO_WARNING
    End If

    Set pIJAttribsCAO = oSmartOcc

    Set oSuppedPort = oOperand
    Set oSuppedPart = oSuppedPort.Part
    
    Set oSuppingSurface1 = oOperators.Item(1)
    If oOperators.count > 1 Then
        Set oSuppingSurface2 = oOperators.Item(2)
    End If

'    If IsNewObject(oWireBody) Or IsAttributeModified(oSmartOcc) Or _
'        IsItModified(oSuppedPart) Or IsItModified(oSuppingSurface1) Or IsItModified(oSuppingSurface2) Then

        Set oCollProxy = pIJAttribsCAO.CollectionOfAttributes("IJUASPSRectangularWebCope")

        Radius = oCollProxy.Item("CopeRadius").Value
        RadiusType = oCollProxy.Item("CopeRadiusType").Value
        WebRadius = oCollProxy.Item("WebRadius").Value
        WebRadiusType = oCollProxy.Item("WebRadiusType").Value
    
        CheckForUndefinedValueAndRaiseError pIJAttribsCAO, RadiusType, "StructFeatureRadiusType", 30
        CheckForUndefinedValueAndRaiseError pIJAttribsCAO, WebRadiusType, "StructFeatureRadiusType", 31

        CopeLength = oCollProxy.Item("CopeLength").Value
        CopeDepth = oCollProxy.Item("CopeDepth").Value
        CopeEdge = oCollProxy.Item("CopeEdge").Value

        ' TR#71317 ; get the inputs to be used for validation
        LengthClr = oCollProxy.Item("CopeLengthClearance").Value
        DepthClr = oCollProxy.Item("CopeDepthClearance").Value
        RoundIncrement = oCollProxy.Item("RoundingIncrement").Value
        ' TR#71317 ends here
        
        Set oMemberFactory = New SPSMemberFactory
        Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices

        oFeatureServices.ComputeRectangularWebCornerCopeProperties oSuppedPart, oSuppedPort.portIndex, _
                    oSuppingSurface1, LengthClr, oSuppingSurface2, DepthClr, RoundIncrement, _
                    CopeEdge, CopeLength, CopeDepth, ComputeOk
        If Not ComputeOk Then
            SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_INVALID_CONFIG, oSmartOcc, Nothing
            Err.Raise SPS_MACRO_WARNING
        End If
        
        On Error Resume Next
    
        oFeatureServices.ComputeRectangularWebCornerCopeGeometry oSuppedPart, oSuppedPort.portIndex, _
                    CopeEdge, CopeLength, CopeDepth, _
                    RadiusType, Radius, WebRadiusType, WebRadius, pObject

        errNumber = Err.Number
        On Error GoTo ErrorHandler
        If errNumber <> 0 Then
            SPSToDoErrorNotify FeatureToDoMsgCodelist, TDL_FEATUREMACROS_INVALID_CONFIG, oSmartOcc, Nothing
            Err.Raise SPS_MACRO_WARNING
        End If

'    End If

    MakeFeatureChildOfPart oSuppedPart, oSmartOcc
    AddStructCutoutOperator pObject, oSuppedPart

    Set oSuppedPart = Nothing
    Set oSuppedPort = Nothing
    Set oSmartOcc = Nothing
    Set pIJAttribsCAO = Nothing
    Set oOperand = Nothing
    Set oOperators = Nothing
    Set oCollProxy = Nothing

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

Public Sub CMReleaseCornerCope(ByVal pMD As IJDMemberDescription)
    Const MT = "CMReleaseCornerCope"
    On Error GoTo ErrorHandler
    ' nothing to do since the wire will be deleted which removes it from the cutout list.

    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMMigrateCornerCope(ByVal pMemberDescription As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateCornerCope"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern
    Dim oWireBody As IJWireBody
    Dim oExistingWire As Object
    Dim bFoundWire As Boolean
    Dim cnt
    Dim oSuppedPort As ISPSSplitAxisPort

    'MsgBox METHOD
    
    Set oWireBody = pMemberDescription.object
    
    Set oSmartOcc = pMemberDescription.CAO
    
    Dim helperStat As StructSOCInputHelperStatus
    Dim oOperand As Object, oOperators As IJElements
    Dim oFeatureHelper As IJStructFeatureSOHelper
    Set oFeatureHelper = Me
    helperStat = oFeatureHelper.GetParents(oSmartOcc, oOperand, oOperators)
    If helperStat <> StructSOCInputHelper_Ok Then
        GoTo ErrorHandler
    End If

    Set oSuppedPort = oOperand
    Set oSuppedPart = oSuppedPort.Part

    Set oStructOperationPattern = oSuppedPart
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    
    If oCollectionOfOperators Is Nothing Then ' Collection must exist so create it if it does not
        Set oCollectionOfOperators = New JObjectCollection
    End If
    bFoundWire = False
    For cnt = 1 To oCollectionOfOperators.count
        Set oExistingWire = oCollectionOfOperators.Item(cnt)
        If oExistingWire Is oWireBody Then
            bFoundWire = True
            Exit For
        End If
    Next cnt
    If Not bFoundWire Then
        oCollectionOfOperators.Add oWireBody
        oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    End If

    Set oSuppedPart = Nothing
    Set oSuppedPort = Nothing
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
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

Public Function IJStructFeatureSOHelper_GetParents(ByVal Feature As SPSMembers.IJSmartOccurrence, Operand As Object, Operators As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus
    Const METHOD = "IJStructFeatureSOHelper_GetParents"
    On Error GoTo ErrorHandler
    
    Dim status As StructSOCInputHelperStatus
    Dim obj1 As Object, obj2 As Object
    Dim oRefColl As IMSSymbolEntities.IJDReferencesCollection
    Dim i As Integer
    Dim tmpOperators As IJElements
    
    status = StructSOCInputHelper_Ok

    Set oRefColl = GetRefCollFromSmartOccurrence(Feature)
    
    If oRefColl Is Nothing Then
        status = StructSOCInputHelper_UnexpectedError
    
    'We allow less than 2 args in the ref coll (surface could be missing)
    ElseIf oRefColl.IJDEditJDArgument.GetCount < 1 Then
        status = StructSOCInputHelper_BadNumberOfObjects

    Else
        Set Operand = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        
        If Not TypeOf Operand Is ISPSSplitAxisEndPort Then
            status = StructSOCInputHelper_InconsistentRelations
        
        Else
            Set tmpOperators = New JObjectCollection
            'if we got surfaces too, then return
            If oRefColl.IJDEditJDArgument.GetCount > 1 Then
                For i = 2 To oRefColl.IJDEditJDArgument.GetCount
                    Set obj1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(i)
                    Set obj2 = GetStablePort(obj1)
                    If Not obj1 Is obj2 Then
                        oRefColl.IJDEditJDArgument.SetEntity i, obj2, ConstIJSurface, "SurfaceToRC_DEST"
                    End If
                    tmpOperators.Add obj2
                Next i
                Set Operators = tmpOperators
            End If
        End If
    End If

    Set tmpOperators = Nothing
    Set oRefColl = Nothing
    IJStructFeatureSOHelper_GetParents = status
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
    Dim oObject As IJDObject
    Dim oSymFactory  As IJDSymbolEntitiesFactory

    Dim i As Integer
    Dim oDesignParent As IJDesignParent
    Dim oSuppedPort As ISPSSplitAxisPort

    status = StructSOCInputHelper_Ok
    
    If Operand Is Nothing Then
        status = StructSOCInputHelper_BadNumberOfObjects
    ElseIf Not TypeOf Operand Is ISPSSplitAxisEndPort Then
        status = StructSOCInputHelper_InvalidTypeOfObject
    ElseIf Operators Is Nothing Then
        status = StructSOCInputHelper_BadNumberOfObjects
    ElseIf Operators.count < 1 Then
        status = StructSOCInputHelper_BadNumberOfObjects
    Else
        
        ' set the feature as a child of operand
        Set oSuppedPort = Operand
        Set oDesignParent = oSuppedPort.Part
        oDesignParent.AddChild Feature
        AddPartFeatureRln oDesignParent, Feature

        AssignPG Feature, oSuppedPort.Part

        Set oRefColl = GetRefCollFromSmartOccurrence(Feature)

        If oRefColl Is Nothing Then
            Set oObject = Feature
            Set oSymFactory = New DSymbolEntitiesFactory
            Set oRefColl = oSymFactory.CreateEntity(ReferencesCollection, oObject.ResourceManager)
            ConnectSmartOccurrence Feature, oRefColl
        Else
            oRefColl.IJDEditJDArgument.RemoveAll
        End If

        oRefColl.IJDEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
        For i = 1 To Operators.count
             oRefColl.IJDEditJDArgument.SetEntity i + 1, Operators.Item(i), ConstIJSurface, "SurfaceToRC_DEST"
        Next i

        Set oRefColl = Nothing
        Set oSymFactory = Nothing
        Set oDesignParent = Nothing
    
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
    
    Dim status As StructSOCInputHelperStatus
    Dim objSurface1 As Object, objSurface2 As Object
    Dim PartOperand As ISPSMemberPartPrismatic
    Dim portIndex As SPSMemberAxisPortIndex

    Dim oMemberFactory As SPSMemberFactory
    Dim oFeatureServices As ISPSMemberFeatureServices
    Dim CopeEdge As Long, CopeLength As Double, CopeDepth As Double
    Dim ComputeOk As Boolean

    NextOperatorPrompt = ""
    
    status = StructSOCInputHelper_Ok

    If Operand Is Nothing Then
        status = StructSOCInputHelper_BadNumberOfObjects
    ElseIf Not TypeOf Operand Is ISPSMemberPartPrismatic Then
        status = StructSOCInputHelper_InvalidTypeOfObject
    ElseIf IsReadOnlyObject(Operand) Then
        status = StructSOCInputHelper_NoWriteAccess
    ElseIf Operators Is Nothing Then
        status = StructSOCInputHelper_BadNumberOfObjects
        NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPELENGTH, "Select Surface Defining Cope Length")
    ElseIf Operators.count < 1 Then
        status = StructSOCInputHelper_BadNumberOfObjects
        NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPELENGTH, "Select Surface Defining Cope Length")
    ElseIf Operators.count > 2 Then
        status = StructSOCInputHelper_UnexpectedError

    Else

        Set objSurface1 = Operators(1)
        If Not (TypeOf objSurface1 Is IJPlane Or TypeOf objSurface1 Is IJSurface) Then
            status = StructSOCInputHelper_InvalidTypeOfObject
            NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPELENGTH, "Select Surface Defining Cope Length")
              
        ElseIf OwnedByConnectable(Operand, objSurface1) = True Then
            status = StructSOCInputHelper_InvalidTypeOfObject
            NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPELENGTH, "Select Surface Defining Cope Length")
        
        ElseIf Operators.count > 1 Then
            Set objSurface2 = Operators(2)
            
            If Not (TypeOf objSurface2 Is IJPlane Or TypeOf objSurface2 Is IJSurface) Then
                status = StructSOCInputHelper_InvalidTypeOfObject
                NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPEDEPTH, "Select Surface Defining Cope Depth")
              
            ElseIf OwnedByConnectable(Operand, objSurface2) = True Then
                status = StructSOCInputHelper_InvalidTypeOfObject
                NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPEDEPTH, "Select Surface Defining Cope Depth")
        
            End If
        End If

        If status = StructSOCInputHelper_Ok Then

            Set oMemberFactory = New SPSMemberFactory
            Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices

            oFeatureServices.GetTrimSurfacePartGeometricCondition Operand, objSurface1, portIndex, ComputeOk
            
            If ComputeOk Then
            
                oFeatureServices.ComputeRectangularWebCornerCopeProperties Operand, portIndex, _
                    objSurface1, 0#, objSurface2, 0#, 0#, CopeEdge, CopeLength, CopeDepth, ComputeOk
                
            End If
        
            If Not ComputeOk Then
                status = StructSOCInputHelper_BadGeometry
                If objSurface2 Is Nothing Then
                    NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPELENGTH, "Select Surface Defining Cope Length")
                Else
                    NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPEDEPTH, "Select Surface Defining Cope Depth")
                End If

            ElseIf objSurface2 Is Nothing Then
                status = StructSOCInputHelper_AdditionalObjectsOptional
                NextOperatorPrompt = m_oLocalizer.GetString(IDS_FEATUREMACROS_SELECT_COPEDEPTH, "Select Surface Defining Cope Depth")
            End If
        End If
    End If

    If status = StructSOCInputHelper_AdditionalObjectsOptional Or status = StructSOCInputHelper_Ok Then
    
        Set PartOperand = Operand
        Set RelationOperand = PartOperand.AxisPort(portIndex)
        
        If RelationOperators Is Nothing Then
            Set RelationOperators = New JObjectCollection
        Else
            RelationOperators.Clear
        End If

        RelationOperators.Add GetStablePort(objSurface1)
        If Not objSurface2 Is Nothing Then
            RelationOperators.Add GetStablePort(objSurface2)
        End If
    
    End If
        
    IJStructFeatureSOHelper_ValidateParents = status

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
    
    ' Other than Side clearance, Inside clearance,
    ' the current need is to gray out the properties on the GOPC
    ' the following code does exactly that
    
    Set pAttrColl = CollAllDisplayedValues

    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
        ' all outputs should be read-only
        If ((pAttrDescr.attrName = "CopeEnd") Or (pAttrDescr.attrName = "CopeEdge") Or (pAttrDescr.attrName = "CopeLength") Or (pAttrDescr.attrName = "CopeDepth")) Then
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
    
    ' TR#73177 ; validate the min. length, depth clearance fields
    If sAttributeName = "CopeLengthClearance" Or sAttributeName = "CopeDepthClearance" Then
        Dim SecDepth As Double, SecWidth As Double, MemLen As Double
        Dim oCrossAttrs As IJDAttributes, pDesignChild As IJDesignChild
        Dim pMemberPart As ISPSMemberPartPrismatic, iLine As IJLine

        SecDepth = 0#
        SecWidth = 0#
        MemLen = 0#
        Set pDesignChild = pIJDAttrs
        If Not pDesignChild Is Nothing Then
            Set pMemberPart = pDesignChild.GetParent
            If Not pMemberPart Is Nothing Then
                Set iLine = pMemberPart
                MemLen = iLine.length
                Set iLine = Nothing
                Set oCrossAttrs = pMemberPart.CrossSection.Definition
                If Not oCrossAttrs Is Nothing Then
                    SecDepth = oCrossAttrs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
                    SecWidth = oCrossAttrs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
                End If
            End If
        End If
        If sAttributeName = "CopeDepthClearance" And varAttributeValue > SecDepth Then
            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_MINCOPEDEPTH_MEMBDEPTH, "Min. CopeDepthClearance > Member depth")
            Exit Function
        ElseIf sAttributeName = "CopeLengthClearance" And varAttributeValue > MemLen Then
            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FEATUREMACROS_MINCOPELENGTH_MEMBLENGTH, "Min. CopeLengthClearance > Member length")
            Exit Function
        End If
    End If
    ' TR#73177 ends here
    
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

Public Sub CMMigrateAggregator(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateAggregator"
  On Error GoTo ErrorHandler

   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructCornerCope(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructCornerCope"
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMCornerCopeCount(ByVal pMemberDescription As IJDMemberDescription, ByRef lCount As Long)
Const METHOD = "CMCornerCopeCount"
On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Public Sub CMIdentifyCloneCornerCope(ByVal pMemberDescription As IJDMemberDescription, ByRef pObject As Object)
  Const METHOD = "CMIdentifyCloneCornerCope"
  On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

Public Sub CMAdaptOnPasteCornerCope(pMD As IJDMemberDescription)
  Const METHOD = "CMAdaptOnPasteCornerCope"
  On Error GoTo ErrorHandler
  
  AdaptOnPasteFeatureToPart (pMD)
    
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

