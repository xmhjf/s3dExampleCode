VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CornerCopeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : CornerCopeDef.cls
'
'Author : Prakash
'
'Description :
'    Module for corner cope of member part
'
'History:
'
'**********************************************************************************************************************

Option Explicit
Private Const MODULE = "CornerCopeDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFeatureMacros.CornerCopeDef"
Const m_ItemName As String = "SPSFeatureMacros.CornerCopeDef"
Private Const strSourceFile As String = "CornerCopeDef.cls"

Dim bOnPreLoad As Boolean
Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1

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
  pAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAggregator"
  pAD.SetCMRemoveInputs imsCOOKIE_ID_USS_LIB, "CMRemoveInputAggregator"
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  Set pAD = Nothing

  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  Set oObject = pDefinition
  Set oMetaData = oObject.ResourceManager
  strIIDCope = oMetaData.IID("IJUASPSCornerCope")
  pAPDs.AddProperty "IJUASPSCornerCope", 1, strIIDCope
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
Public Sub CMSetInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMSetInputAggregator"
On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
End Sub
Public Sub CMRemoveInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMRemoveInputAggregator"
On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
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
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSuppingPlane1 As IJPlane, oSuppingPlane2 As IJPlane
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppedPort As ISPSSplitAxisPort

    Dim iQuadrant As Integer
    Dim strError As String
    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)

    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPlane1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppingPlane2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(3)

    bIsNeeded = False
    If IsReadOnlyObject(oSuppedPart) Then
        strError = "Read only object encountered"
        GoTo ErrorHandler
    End If
    
    bIsNeeded = IsCornerCopeNeeded(oSuppedPart, oSuppingPlane1, oSuppingPlane2)

    Set oSuppingPlane1 = Nothing
    Set oSuppingPlane2 = Nothing
    Set oSuppedPart = Nothing
    Set oSuppedPort = Nothing
    Set oRefColl = Nothing
    Set oSmartOcc = Nothing
  
  Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, strSourceFile, MT, strError).Number
End Sub

Public Sub CMConstructCornerCope(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const MT = "CMConstructCornerCope"
    On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
     Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingPlane1 As IJPlane, oSuppingPlane2 As IJPlane
    Dim iQuadrant As Integer
    Dim oCollectionOfOperators As IJElements
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim oStructOperationPattern As IJStructOperationPattern
    Dim oWireBody As IJWireBody
    Dim pIJAttribsCAO As IJDAttributes
    Dim oControlFlags As IJControlFlags
    Dim oSuppedPort As ISPSSplitAxisPort

    Set oSmartOcc = pMemberDescription.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    Set pIJAttribsCAO = oSmartOcc
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPlane1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppingPlane2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(3)
    
    CreateCornerCope oSuppedPart, oSuppingPlane1, oSuppingPlane2, pIJAttribsCAO, pResourceManager, oWireBody
    
    'hide the wire body
    Set oControlFlags = oWireBody
    oControlFlags.ControlFlags(&H4) = &H4
    
    Set pObj = oWireBody
    'create the collection of operators
    
    Set oStructOperationPattern = oSuppedPart
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    If oCollectionOfOperators Is Nothing Then
        Set oCollectionOfOperators = New JObjectCollection
    End If

    oCollectionOfOperators.Add oWireBody
    oStructOperationPattern.SetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    
    Set oSuppingPlane1 = Nothing
    Set oSuppingPlane2 = Nothing
    Set oSuppedPart = Nothing
    Set oSuppedPort = Nothing
    Set oRefColl = Nothing
    Set oSmartOcc = Nothing
    Set oWireBody = Nothing
    Set oStructCutoutOperationAE = Nothing
    Set oStructOperationPattern = Nothing
    Set oCollectionOfOperators = Nothing

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
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSuppingPlane1 As IJPlane, oSuppingPlane2 As IJPlane
    Dim iQuadrant As Integer
    Dim oWireBody As IJWireBody
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSuppedPort As ISPSSplitAxisPort

    Set oWireBody = pObject
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOcc
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part
    Set oSuppingPlane1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
    Set oSuppingPlane2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(3)

    If IsNewObject(oWireBody) Or IsAttributeModified(oSmartOcc) Or _
        IsItModified(oSuppedPart) Or IsItModified(oSuppingPlane1) Or IsItModified(oSuppingPlane2) Then
        CreateCornerCope oSuppedPart, oSuppingPlane1, oSuppingPlane2, pIJAttribsCAO, Nothing, oWireBody
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
    For cnt = 1 To oCollectionOfOperators.Count
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
    
    Set oSuppingPlane1 = Nothing
    Set oSuppingPlane2 = Nothing
    Set oSuppedPart = Nothing
    Set oSuppedPort = Nothing
    Set oRefColl = Nothing
    Set oSmartOcc = Nothing
    Set oWireBody = Nothing
    Set pIJAttribsCAO = Nothing

    Exit Sub
ErrorHandler:      HandleError MODULE, MT
End Sub

Public Sub CMReleaseCornerCope(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseCornerCope"
On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
Public Sub CMMigrateCornerCope(ByVal pMemberDescription As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateCornerCope"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IMSSymbolEntities.IJDReferencesCollection
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
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOcc)
    
    Set oSuppedPort = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set oSuppedPart = oSuppedPort.Part

    Set oStructOperationPattern = oSuppedPart
    oStructOperationPattern.GetOperationPattern "StructGeneric.StructCutoutOperationAE.1", oCollectionOfOperators, oStructCutoutOperationAE
    
    If oCollectionOfOperators Is Nothing Then ' Collection must exist so create it if it does not
        Set oCollectionOfOperators = New JObjectCollection
    End If
    bFoundWire = False
    For cnt = 1 To oCollectionOfOperators.Count
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
    Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
    Dim i As Integer
    Dim tmpOperators As IJElements
    
    Set tmpOperators = New JObjectCollection
    Set oRefColl = GetRefCollFromSmartOccurrence(Feature)
    Set Operand = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    For i = 2 To oRefColl.IJDEditJDArgument.GetCount
        tmpOperators.Add oRefColl.IJDEditJDArgument.GetEntityByIndex(i)
    Next i
    Set Operators = tmpOperators
    
    Set tmpOperators = Nothing
    Set oRefColl = Nothing
    IJStructFeatureSOHelper_GetParents = StructSOCInputHelper_Ok
    Exit Function

ErrorHandler:
    IJStructFeatureSOHelper_GetParents = StructSOCInputHelper_UnexpectedError
    HandleError MODULE, METHOD
End Function

Public Function IJStructFeatureSOHelper_OperandPrompt(ByVal Feature As SPSMembers.IJSmartOccurrence) As String
    IJStructFeatureSOHelper_OperandPrompt = "Select Members"
End Function

Public Function IJStructFeatureSOHelper_SetParents(ByVal Feature As SPSMembers.IJSmartOccurrence, ByVal Operand As Object, ByVal Operators As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus
    Const METHOD = "IJStructFeatureSOHelper_SetParents"
    On Error GoTo ErrorHandler
    
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oDesignParent As IJDesignParent
    Dim i As Integer
    Dim oSuppedPort As ISPSSplitAxisPort
    
    ' set the feature as a child of operand
    Set oSuppedPort = Operand
    Set oDesignParent = oSuppedPort.Part
    Set oSuppedPort = Nothing
    oDesignParent.AddChild Feature

    ' now set the reference collection
    Dim oObject As IJDObject
    Set oObject = Feature
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, oObject.ResourceManager)
    Set oObject = Nothing
    oRefCollChild.IJDEditJDArgument.SetEntity 1, Operand, CONST_ISPSSplitAxisEndPort, "RefColl"
    For i = 1 To Operators.Count
        oRefCollChild.IJDEditJDArgument.SetEntity i + 1, Operators.Item(i), IJPlane, "SPSSuppPlaneToRC_DEST"
    Next i

    ' Connect the SO to its model arguments
    ConnectSmartOccurrence Feature, oRefCollChild

    Set oRefCollChild = Nothing
    Set oSymFactory = Nothing
    Set oDesignParent = Nothing
    IJStructFeatureSOHelper_SetParents = StructSOCInputHelper_Ok
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
    Dim i As Integer
    Dim operator As Object, myPartOperand As ISPSMemberPartPrismatic, CutbackPort As SPSMemberAxisPortIndex
    
    NextOperatorPrompt = ""
    ' ValMode determines the type of validation to be done. Its valid values are:
    ' 1 - called from smartstep-1 ; validate only the operand
    ' 2 - called from smartstep-2 ; validate only the operators
    ' 3- called from compute mode ;  validate both operand and operators for their validity to the corner-cope feature
    If ValMode = 1 Or ValMode = 3 Then
        If Operand Is Nothing Or ((Not Operand Is Nothing) And (Not (TypeOf Operand Is ISPSMemberPartPrismatic))) Then
            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
        ElseIf IsReadOnlyObject(Operand) Then
            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_NoWriteAccess
        Else
            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_Ok
        End If
    End If
    If ValMode = 2 Or ValMode = 3 Then
        If Operators Is Nothing Then
            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
        Else ' it contains some elements
            IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_Ok
            For i = 1 To Operators.Count
                If Not TypeOf Operators.Item(i) Is IJPlane Then
                    IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
                    Exit For
                ElseIf IsOperandOperatorIntersect(Operand, Operators.Item(i)) Then ' it is IJPlane ; now check for intersection
                        If i = 2 Then ' additional check for second plane is it should intersect 1st plane also
                            If Operators(2) Is Operators(1) Then ' same plane should not be allowed to be selected
                                IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
                            Else
                                Dim oSurf1 As IJSurface, oSurf2 As IJSurface
                                Dim code As Geom3dIntersectConstants, intElms As IJElements
                                Set oSurf1 = Operators.Item(1)
                                Set oSurf2 = Operators.Item(2)
                                oSurf1.Intersect oSurf2, intElms, code
                                If intElms Is Nothing Or ((Not intElms Is Nothing) And intElms.Count <= 0) Then ' don't intersect
                                    IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
                                Else
                                    If IsSurfaceTrimCutbackNeeded(Operand, Operators.Item(1)) = True Or IsSurfaceTrimCutbackNeeded(Operand, Operators.Item(2)) = True Then
                                    Else
                                        IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                Else
                    IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_InvalidTypeOfObject
                    Exit For
                End If
            Next i
        End If
    End If
    If ValMode = 3 Then
            If IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_Ok Then
                If Operators.Count <> 2 Then
                    IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_BadNumberOfObjects
                    NextOperatorPrompt = "Select Cutters"
                ElseIf IsCornerCopeNeeded(Operand, Operators.Item(1), Operators.Item(2)) = False Then
                    IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_BadGeometry
                    NextOperatorPrompt = "Select Cutters"
                End If
            End If
    End If
    
    If ValMode = 3 And IJStructFeatureSOHelper_ValidateParents = StructSOCInputHelper_Ok Then
        ' now populate the RelationOperators
        If RelationOperators Is Nothing Then
            Set RelationOperators = New JObjectCollection
        End If
        RelationOperators.Clear
        For i = 1 To Operators.Count
            Set operator = Operators.Item(i)
            RelationOperators.Add operator
            Set operator = Nothing
        Next i
        ' now populate the RelationOperand as a port
        Set myPartOperand = Operand
        For i = 1 To Operators.Count
            If IsSurfaceTrimCutbackNeeded(Operand, Operators.Item(i)) Then ' member axis and plane intersect
                CutbackPort = GetCutbackEnd(myPartOperand, Operators.Item(i))
                Set RelationOperand = myPartOperand.AxisPort(CutbackPort)
                Exit For
            End If
        Next i
    End If
    
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
    IJUserAttributeMgmt_OnAttributeChange = "ERROR"
    HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = "ERROR"
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

    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
        If ((pAttrDescr.attrName = "Radius") Or (pAttrDescr.attrName = "Increment") Or (pAttrDescr.attrName = "SquaredEnd")) Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        End If
    Next
        
    For i = 1 To pAttrColl.Count
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
    UserAttributeMgmt_Validate = "ERROR"

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
        UserAttributeMgmt_Validate = "SchemaERROR"
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.Name = sAttributeName Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (varAttributeValue < 0#) Then
                            UserAttributeMgmt_Validate = "Value must be >= 0"
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

