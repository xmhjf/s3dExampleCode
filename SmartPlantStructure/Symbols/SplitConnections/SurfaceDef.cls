VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SurfaceDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : SurfaceDef.cls
'
'Author : M. Holderer
'
'Description :
'    SmartPlant Structural Interference Connection Definition file
'
'History:
'
' 06-jul-2004    Struct Dev    created.
' 02-May-2005    Struct Dev    V6 - restrict surface geom to GridPlanes.  See IsAllowedSurface.
' 31-Jan-2007    Struct dev    TR 112605 set CTL_FLAG_NOT_IN_SPATIAL_INDEX
' 30-Jun-2010    Struct dev    DI 133354 Planning Surface is allowed surface now
'
'********************************************************************

Private Const MODULE = "SurfaceDef"
Private Const m_strSourceFile = "SurfaceDef.def"
Private m_oLocalizer As IJLocalizer

Private Const PlateSystemClassIDStr = "0010A1F2"
Private Const PlanningSurfaceClassIDStr = "00113E1C"

Const m_ItemProgId As String = "SPSSplitMacros.SurfaceDef"
Const m_ItemName As String = "SPSSplitMacros.SurfaceDef"
Dim bOnPreLoad As Boolean

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements SP3DStructInterfaces.IJUserAttributeMgmt
'Implements SP3DStructInterfaces.ISPSMirrorHelper
Implements SPSMembers.IJStructILCHelper
Implements SPSMembers.IJStructSOSelectRule


'*************************************************************************
'Function
'DefinitionInputs
'
'Abstract
'Sets any required inputs for symbol evaluation
'
'Arguments
'IJDInputsHelper defined in CommonApp
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo errorHandler
  
  pIH.SetInput "SplitParentsRefColl"
  
  Exit Sub
errorHandler:
  pIH.ReportError
End Sub




'*************************************************************************
'Function
'IJDUserSymbolServices_InitializeSymbolDefinition
'
'Abstract
'Cleans the previous definition up and initializes the new one (Input, Output, Representation,
'RepresenationEvaluation, ...) constructing the symbol definition by (re)defining the inputs, outputs,
'representations, and representation.
'Note:  The previous setting of the definition must be reset in this method before constructing the new definition.
'
'Arguments:
'pDefinition  Symbol definition passed by reference that will be initialized in this method.
'
'Return
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions
'
'***************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo errorHandler

  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs -
  ' They are identical to the class inputs (i.e. penetrated and Penetrating)
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pDefinition
  DefinitionInputs pIH

  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{06AB53A9-9E14-4EDE-84DD-FFF477721904}"     'SPSSplitMemberConn
  pAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAggregator"
  pAD.SetCMRemoveInputs imsCOOKIE_ID_USS_LIB, "CMRemoveInputAggregator"
 
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  'Add following interfaces as inputs so as to get notified when they are modified.
  'These interfaces carry properties that drive position of the intersection point
  'pAPDs.AddProperty name, dispid, interfaceID, CMEvaluate, libCookie, PDProcessTime )
  
 ' pAPDs.AddProperty "ReferenceInit", 1, "IJUASPSSplitOffset", "CMInit"
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 
  'Point1 splits neither.  Evaluates split rule and stores it as a hidden "SplitRuleResult" parameter.
  Set pMemberDescription = pMemberDescriptions.AddMember("SurfacePoint1", 1, "CMConstructPoint1", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputPoint1"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalPoint1"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleasePoint1"
  'pMemberDescription.IsDeletedWithAggregator = False   'default is True
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJPoint {EA4AC251-14E7-11d2-B6F3-0060974CAEC9}
  pPropertyDescriptions.AddProperty "ComputePoint1", 1, "{EA4AC251-14E7-11d2-B6F3-0060974CAEC9}", "CMUpdatePoint", imsCOOKIE_ID_USS_LIB
  
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescriptions = Nothing
  
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMSetInputAggregator
'
'Abstract
'Can be used to set inputs on the parent SmartOccurrence prior to each CM evaluate
'
'Arguments
'IJDAggregatorDescription interface of the CustomAssemblyDefinition wrapper
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMSetInputAggregator"
On Error GoTo errorHandler

  
Exit Sub
errorHandler:     HandleError MODULE, METHOD
End Sub
'*************************************************************************
'Function
'CMRemoveInputAggregator
'
'Abstract
'Can be used to remove inputs on the parent SmartOccurrence.
'
'Arguments
'IJDAggregatorDescription interface of the CustomAssemblyDefinition wrapper
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CMRemoveInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMRemoveInputAggregator"
On Error GoTo errorHandler

  
Exit Sub
errorHandler:     HandleError MODULE, METHOD
End Sub

'*************************************************************************
'Function
'CMConditionalPoint1
'
'Abstract
'Determines whether the current member is needed as an output.
'
'Arguments
'IJDMemberDescription interface of the member
'Boolean set to True if the Member is needed.
'
'Return
'
'Exceptions
'
'***************************************************************************

Public Sub CMConditionalPoint1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalPoint1"
    On Error GoTo errorHandler
    
    bIsNeeded = True

    Exit Sub

errorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL    ' Force the ILC to the ToDo List
End Sub

'*************************************************************************
'Function
'CMConstructPoint1
'
'Abstract
'Creates the output member object
'
'Arguments
'IJDMemberDescription interface of the member
'pResourceManager used to construct the member
'pObject is the constructed object
'
'Return
'
'Exceptions
'
'***************************************************************************

Public Sub CMConstructPoint1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructPoint1"
  On Error GoTo errorHandler
  'Create a point output object, set its properties to hidden.  Not connected as input to any parent system.

  Dim ISOC As SPSMembers.IJStructSmartOccurrenceConnection
  Dim iSplitConnection As ISPSSplitMemberConnection
  Dim parents As IJElements
  Dim parent1 As Object
  Dim socStatus As SPSMembers.StructSOCInputHelperStatus
  Dim iILCHelper As SPSMembers.IJStructILCHelper

  Dim geomFactory As New GeometryFactory
  Dim iControlFlags As IJControlFlags
  Dim point As IJPoint

  Set ISOC = pMemberDescription.CAO
  Set iILCHelper = Me       'ISOC.Helper
  socStatus = iILCHelper.GetParents(ISOC, parent1, parents)
  
  Set point = geomFactory.Points3d.CreateByPoint(pResourceManager, 0#, 0#, 0#)
  Set iControlFlags = point
  iControlFlags.ControlFlags(CTL_FLAG_NO_DISPLAY) = CTL_FLAG_NO_DISPLAY
  iControlFlags.ControlFlags(CTL_FLAG_NOT_IN_SPATIAL_INDEX) = CTL_FLAG_NOT_IN_SPATIAL_INDEX
  
  If TypeOf parent1 Is IJStructSystemParent Then        'status = splitFirst
    InstallSplitPoint 0, 0, ISOC, parent1, point
  ElseIf TypeOf parents(1) Is IJStructSystemParent Then     'status = splitSecond
    InstallSplitPoint 0, 0, ISOC, parents(1), point
  Else
    Err.Raise E_FAIL
  End If

  Set iSplitConnection = ISOC
  iSplitConnection.UpdateSplitNonePorts Nothing, Nothing
  
  Set pObj = point

  Set ISOC = Nothing
  Set iILCHelper = Nothing
  Set parent1 = Nothing
  Set parents = Nothing
  Set point = Nothing
  Set geomFactory = Nothing
  Set iControlFlags = Nothing

  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMSetInputPoint1
'
'Abstract
'Used to set inputs on the child SmartOccurrence prior to CM evaluate
'
'Arguments
'IJDMemberDescription interface of the wrapped output member
'
'Return
'
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputPoint1(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputPoint1"
  On Error GoTo errorHandler
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMReleasePoint1
'
'Abstract
'Used to clear/release child SmartOccurrence
'
'Arguments
'IJDMemberDescription interface of the wrapped output member
'
'Return
'
'
'Exceptions
'
'***************************************************************************
Public Sub CMReleasePoint1(ByVal pMD As IJDMemberDescription)
  Const MT = "CMReleasePoint1"
  On Error GoTo errorHandler
  'Disconnect the point as splitter from parent1
  
  Dim ISOC As SPSMembers.IJStructSmartOccurrenceConnection
  Dim parents As IJElements
  Dim parent1 As Object
  Dim socStatus As SPSMembers.StructSOCInputHelperStatus
  Dim iILCHelper As SPSMembers.IJStructILCHelper
  Dim iIJStructSystemParent As IJStructSystemParent
  
  Set ISOC = pMD.CAO
  Set iILCHelper = Me       'ISOC.Helper
  
  socStatus = iILCHelper.GetParents(ISOC, parent1, parents)
  If socStatus = StructSOCInputHelper_Ok Then
    If TypeOf parent1 Is IJStructSystemParent Then
       Set iIJStructSystemParent = parent1
    ElseIf TypeOf parents(1) Is IJStructSystemParent Then
       Set iIJStructSystemParent = parents(1)
    End If
    iIJStructSystemParent.RemoveSplit pMD.object
  End If
  
  Set ISOC = Nothing
  Set iILCHelper = Nothing
  Set parents = Nothing
  Set parent1 = Nothing
  Set iIJStructSystemParent = Nothing

  Exit Sub
  
errorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
'*************************************************************************
'Function
'CMUpdatePoint
'
'Abstract
'Evaluates the member property
'
'Arguments
'IJDPropertyDescription interface describing the property to be evaluated
'pObject is the object whose property is being computed
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CMUpdatePoint(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const MT = "CMUpdatePoint"
    On Error GoTo errorHandler

    Dim strErrMsg As String
    Dim ISOC As IJStructSmartOccurrenceConnection
    Dim iILCHelper As IJStructILCHelper
    Dim IHStatus As StructSOCInputHelperStatus
    Dim parent1 As Object, parents As IJElements
    Dim ok As Boolean
    Dim x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double
    Dim pIJAttribs As IJDAttributes
    Dim distanceAlong As Double
    Dim splitStat As splitStatus
    Dim iSplitResult As IJStructILCConnectionResult
    
    Set ISOC = pPropertyDescription.CAO
    Set iILCHelper = ISOC.Helper
    
    IHStatus = iILCHelper.GetParents(ISOC, parent1, parents)
    If IHStatus <> StructSOCInputHelper_Ok Then
        SPSToDoErrorNotify SCToDoMsgCodelist, TDL_SPLITMACROS_UNEXPECTED_ERROR, ISOC, Nothing
        Err.Raise E_FAIL
    End If

    Set iSplitResult = ISOC
    splitStat = iSplitResult.SplitParentStatusResult

    Set pIJAttribs = ISOC
    distanceAlong = pIJAttribs.CollectionOfAttributes("IJUASPSSplitOffset").Item("Offset").Value
    
    ComputeIntersection parent1, parents(1), splitStat, distanceAlong, x1, y1, z1, x2, y2, z2, ok
    
    SetTDLErrorFlag ISOC, ok
    SetIsUnaryFlag ISOC, False

    If Not ok Then
        Dim oStructSystemParent As IJStructSystemParent
        
        If TypeOf parent1 Is IJStructSystemParent Then
            Set oStructSystemParent = parent1
        ElseIf TypeOf parents(1) Is IJStructSystemParent Then
            Set oStructSystemParent = parents(1)
        End If
        oStructSystemParent.RemoveSplit pPropertyDescription.object
        
        SPSToDoErrorNotify SCToDoMsgCodelist, TDL_SPLITMACROS_NOVALID_INTERSECTION, ISOC, Nothing
        Err.Raise E_FAIL

    End If

    If TypeOf parent1 Is IJStructSystemParent Then
        InstallSplitPoint 2, 1, ISOC, parent1, pObject
    Else
        InstallSplitPoint 2, 1, ISOC, parents(1), pObject
    End If

    UpdateSplitPoint ISOC, pObject, x1, y1, z1

    Exit Sub

errorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, m_strSourceFile, MT).Number
    End If
End Sub

' The following methods are generic for all the Custom assembly
'
'
'*************************************************************************
'Function
'IJDUserSymbolServices_InstanciateDefinition
'
'Abstract
'Instantiates a persistent symbol definition object and initializes it for the first time,
'returning a pointer (ppSymbolDefDisp) to the IDispatch interface of the initialized symbol definition.
'
'Arguments:
'codeBase specifies the URL (or UNC) of the .cab file that can provides the dll associated to the symbol definition object (ActiveX® control packaging).
'definitionParameters  Definition parameters.
'pResourceMgr  resource manager to which the symbol definition will be connected.
'
'Return:
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions:
'
'***************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
   Const MT = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo errorHandler
 
  
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
errorHandler:  HandleError MODULE, MT
End Function
'*************************************************************************
'Function
'IJDUserSymbolServices_GetDefinitionName
'
'Abstract
'Used during the execution of IJDDefinitionCollection::GetDefinitionByProgId to get the definition name
'based upon the definitionParameters passed in. It returns the definition name (pDefName) if it already
'exists within the collection. The name of a definition is the identifier of the definition object
'in the definition collection and assures its uniqueness in the given resource manager.
'
'Arguments
'definitionParameters
'
'Return
'
'Exceptions
'
'***************************************************************************

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

Private Function IJStructILCHelper_GetParents(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection, pParent1 As Object, otherParents As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus
    Const MT = "ISPSILCInputHelper_GetParents"
    
    Dim IHStatus As StructSOCInputHelperStatus
    Dim obj1 As Object, obj2 As Object
    Dim parents As IJElements
    Dim count As Long

    Dim splitStat As splitStatus
    Dim iSplitResult As IJStructILCConnectionResult

    IHStatus = StructSOCInputHelper_UnexpectedError
    
    Set parents = GetRefCollObjects(ISOC)
    Set iSplitResult = ISOC
    splitStat = iSplitResult.SplitParentStatusResult

    count = parents.count

    If count > 0 Then
        Set obj1 = parents(1)
        If count > 1 Then
            Set obj2 = parents(2)
            If count > 2 Then
                IHStatus = StructSOCInputHelper_BadNumberOfObjects
            End If
        Else
            IHStatus = StructSOCInputHelper_BadNumberOfObjects
        End If
    Else
        IHStatus = StructSOCInputHelper_BadNumberOfObjects
    End If

    If IHStatus = StructSOCInputHelper_UnexpectedError Then ' ok so far.  obj1 and obj2 are set.
        If TypeOf obj1 Is ISPSMemberSystem Then
            'V6 limit surfaces to grid planes only.
            If ObjectIsOkSurface(obj2) And IsAllowedSurface(obj2) Then
                If splitStat <> ssSplitFirst Then
                    IHStatus = StructSOCInputHelper_InconsistentRelations
                End If
            Else
                IHStatus = StructSOCInputHelper_InvalidTypeOfObject
            End If
        ElseIf ObjectIsOkSurface(obj1) And IsAllowedSurface(obj1) Then
            If TypeOf obj2 Is ISPSMemberSystem Then
                If splitStat <> ssSplitSecond Then
                    IHStatus = StructSOCInputHelper_InconsistentRelations
                End If
            Else
                IHStatus = StructSOCInputHelper_InvalidTypeOfObject
            End If
        Else
            IHStatus = StructSOCInputHelper_InvalidTypeOfObject
        End If
    End If

    If Not obj1 Is Nothing Then
        Set pParent1 = obj1
    End If
    Set otherParents = New JObjectCollection
    If Not obj2 Is Nothing Then
        otherParents.Add obj2
    End If

    If IHStatus = StructSOCInputHelper_UnexpectedError Then
        IHStatus = StructSOCInputHelper_Ok
    End If

    Set obj1 = Nothing
    Set obj2 = Nothing
    Set parents = Nothing

    IJStructILCHelper_GetParents = IHStatus
    Exit Function

errorHandler:
    HandleError MODULE, MT
    IJStructILCHelper_GetParents = IHStatus
End Function

Private Property Get IJStructILCHelper_SelectionRule(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection) As SPSMembers.IJStructSOSelectRule
    Set IJStructILCHelper_SelectionRule = Me
End Property

Private Function IJStructILCHelper_SetParents(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection, _
    ByVal parent1 As Object, ByVal otherParents As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus

    Const MT = "ISPSILCHelper_SetParents"
    On Error GoTo errorHandler
    
    Dim IHStatus As StructSOCInputHelperStatus
    Dim splitStat As splitStatus
    Dim ILConnection As IJStructILCConnection
    
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim iSplitResult As IJStructILCConnectionResult
    Dim parent2 As Object

    IHStatus = StructSOCInputHelper_UnexpectedError

    Set ILConnection = ISOC
    splitStat = ILConnection.SplitParentStatus

    If parent1 Is Nothing Then
        IHStatus = StructSOCInputHelper_BadNumberOfObjects
        GoTo wrapup
    End If
    
    If otherParents Is Nothing Then
        IHStatus = StructSOCInputHelper_BadNumberOfObjects
        GoTo wrapup
    End If
    If Not TypeOf otherParents Is IJElements Then
        IHStatus = StructSOCInputHelper_InvalidTypeOfObject
        GoTo wrapup
    End If
    If otherParents.count <> 1 Then
        IHStatus = StructSOCInputHelper_BadNumberOfObjects
        GoTo wrapup
    End If
    Set parent2 = otherParents(1)

    If parent1 Is parent2 Then
        IHStatus = StructSOCInputHelper_DuplicateObject
        GoTo wrapup
    End If

    If TypeOf parent1 Is IJStructSystemParent Then
        If ObjectIsOkSurface(parent2) And IsAllowedSurface(parent2) Then
            If splitStat <> ssRule And splitStat <> ssSplitFirst Then
                IHStatus = StructSOCInputHelper_InvalidTypeOfObject
                GoTo wrapup
            End If
            splitStat = ssSplitFirst
        Else
            IHStatus = StructSOCInputHelper_InvalidTypeOfObject
            GoTo wrapup
        End If
    
    ElseIf ObjectIsOkSurface(parent1) And IsAllowedSurface(parent1) Then
        If TypeOf parent2 Is IJStructSystemParent Then
            If splitStat <> ssRule And splitStat <> ssSplitSecond Then
                IHStatus = StructSOCInputHelper_InvalidTypeOfObject
                GoTo wrapup
            End If
            splitStat = ssSplitSecond
        Else
            IHStatus = StructSOCInputHelper_InvalidTypeOfObject
            GoTo wrapup
        End If

    Else
        IHStatus = StructSOCInputHelper_InvalidTypeOfObject
        GoTo wrapup
    End If
    
    If splitStat = ssSplitFirst Then
        If CheckAccess(1, parent1, ISOC) <> StructSOCInputHelper_Ok Then
            IHStatus = StructSOCInputHelper_NoWriteAccess
            GoTo wrapup
        End If
    Else
        If CheckAccess(1, parent2, ISOC) <> StructSOCInputHelper_Ok Then
            IHStatus = StructSOCInputHelper_NoWriteAccess
            GoTo wrapup
        End If
    End If
    

    RemoveCommonFCs parent1, parent2

    'set interfaces for input member
    Set oRC = GetRefColl(ISOC)
    oRC.IJDEditJDArgument.RemoveAll

    If splitStat = ssSplitFirst Then
        oRC.IJDEditJDArgument.SetEntity 1, parent1, ConstISPSMemberSystemSuppingNotify3, "MembSysSuppingNotify3RC_DEST"
        oRC.IJDEditJDArgument.SetEntity 2, parent2, ConstIJGeometry, "SplitRefColl_DEST"
        AddSysChild 1, parent1, ISOC
    Else
        oRC.IJDEditJDArgument.SetEntity 1, parent1, ConstIJGeometry, "SplitRefColl_DEST"
        oRC.IJDEditJDArgument.SetEntity 2, parent2, ConstISPSMemberSystemSuppingNotify3, "MembSysSuppingNotify3RC_DEST"
        AddSysChild 1, parent2, ISOC
    End If

    Set iSplitResult = ISOC
    iSplitResult.SplitParentStatusResult = splitStat

    IHStatus = StructSOCInputHelper_Ok

wrapup:
    IJStructILCHelper_SetParents = IHStatus
    Exit Function

errorHandler:
    HandleError MODULE, MT
    IJStructILCHelper_SetParents = IHStatus
End Function

Private Function IJStructILCHelper_UpdateValuesFromPosition(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection) As SPSMembers.StructSOCInputHelperStatus
    IJStructILCHelper_UpdateValuesFromPosition = StructSOCInputHelper_Ok
End Function

Private Property Get IJStructILCHelper_UserAttributeMgmt() As SPSMembers.IJUserAttributeMgmt
    Set IJStructILCHelper_UserAttributeMgmt = Me
End Property

' insure that the given parents are valid types.
' make sure they intersect, or their minDist point is along each length.
' At least one of them must be a SPSMemberSystemLinear.
' It gets returned as pRelationParent1.

Private Function IJStructILCHelper_ValidateParents(ByVal ISOC As SPSMembers.IJStructSmartOccurrenceConnection, _
    ByVal inputParent1 As Object, ByVal inputOtherParents As SPSMembers.IJElements, _
    pRelationParent1 As Object, pRelationOtherParents As SPSMembers.IJElements) As SPSMembers.StructSOCInputHelperStatus

    Const MT = "IJStructILCHelper_ValidateParents"

    Dim splitStat As splitStatus
    Dim ILConnection As IJStructILCConnection
    Dim iStatus As StructSOCInputHelperStatus
    Dim parent2 As Object
    Dim parent1 As Object
    Dim ok As Boolean
    Dim x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double

    On Error GoTo errorHandler

    iStatus = StructSOCInputHelper_Ok

    Set ILConnection = ISOC
    splitStat = ILConnection.SplitParentStatus

    If inputParent1 Is Nothing Then
        iStatus = StructSOCInputHelper_BadNumberOfObjects
    End If
    Set parent1 = inputParent1
    If inputOtherParents Is Nothing Then
        iStatus = StructSOCInputHelper_BadNumberOfObjects
    ElseIf Not TypeOf inputOtherParents Is IJElements Then
        iStatus = StructSOCInputHelper_InvalidTypeOfObject
    ElseIf inputOtherParents.count <> 1 Then
        iStatus = StructSOCInputHelper_BadNumberOfObjects
    Else
        Set parent2 = inputOtherParents(1)
    End If

    If iStatus <> StructSOCInputHelper_Ok Then
        GoTo wrapup
    End If
    ' we know that both parent1 and parent2 are set.

    If TypeOf parent1 Is IJStructSystemParent Then
        If Not ObjectIsOkSurface(parent2) And Not IsAllowedSurface(parent2) Then
            iStatus = StructSOCInputHelper_InvalidTypeOfObject
        End If
        If splitStat <> ssSplitFirst And splitStat <> ssRule Then   ' object type inconsistent with split status.
            iStatus = StructSOCInputHelper_InvalidTypeOfObject
        End If
        splitStat = ssSplitFirst
    ElseIf ObjectIsOkSurface(parent1) And IsAllowedSurface(parent1) Then
        If Not TypeOf parent2 Is IJStructSystemParent Then
            iStatus = StructSOCInputHelper_InvalidTypeOfObject
        End If
        If splitStat <> ssSplitSecond And splitStat <> ssRule Then   ' object type inconsistent with split status.
            iStatus = StructSOCInputHelper_InvalidTypeOfObject
        End If
        splitStat = ssSplitSecond
    Else
        iStatus = StructSOCInputHelper_InvalidTypeOfObject
    End If
        
    If iStatus <> StructSOCInputHelper_Ok Then
        GoTo wrapup
    End If
    
    If CheckAccess(1, parent1, ISOC) <> StructSOCInputHelper_Ok Then
        iStatus = StructSOCInputHelper_NoWriteAccess
        GoTo wrapup
    End If

    ComputeIntersection parent1, parent2, splitStat, 0#, x1, y1, z1, x2, y2, z2, ok
    If Not ok Then
        iStatus = StructSOCInputHelper_BadGeometry
    ElseIf RedundantConnectionExists(ISOC, parent1, parent2) Then
        iStatus = StructSOCInputHelper_DuplicateObject
    ElseIf TypeOf parent1 Is IJStructSystemParent And IsNOTWritable(parent1) Then
        iStatus = StructSOCInputHelper_InvalidTypeOfObject
    ElseIf TypeOf parent2 Is IJStructSystemParent And IsNOTWritable(parent2) Then
        iStatus = StructSOCInputHelper_InvalidTypeOfObject
    Else
        Set pRelationParent1 = parent1
        Set pRelationOtherParents = New JObjectCollection
        
        ' for root plate system get the first port for the refColl relation
        If StrComp(GetObjectClassIDString(parent2), PlateSystemClassIDStr, vbTextCompare) = 0 Then
            Dim oConnectable As IJConnectable
            Dim oPorts As IJElements
            Set oConnectable = parent2
            oConnectable.enumConnectablePorts oPorts, PortFace
            pRelationOtherParents.Add oPorts.Item(1)
        Else
            pRelationOtherParents.Add parent2
        End If
    
        iStatus = StructSOCInputHelper_Ok
    End If

wrapup:
    IJStructILCHelper_ValidateParents = iStatus
    Exit Function

errorHandler:
    IJStructILCHelper_ValidateParents = StructSOCInputHelper_UnexpectedError
    Err.Raise E_FAIL
End Function


Private Function IJStructSOSelectRule_Evaluate(ByVal iSmartOccurrence As Object, strSelectedItem As String) As SPSMembers.StructSOCInputHelperStatus

    strSelectedItem = "SplitBySurface-1"
    IJStructSOSelectRule_Evaluate = StructSOCInputHelper_Ok

End Function

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnAttributeChange
'
'Abstract
'Gets called for each attribute change on the property page
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of attributes as currently displayed ( prior to Commit )
'pAttrToChange is which attribute is being edited
'varNewAttrValue is the value given by the user.
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo errorHandler
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_SPLITMACROS_ERROR, "ERROR")
    
    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long

    ' attribute value validation in reponse to changes on GOPC
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.interfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
'            IJUserAttributeMgmt_OnAttributeChange = "ERROR::Bad Value"
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
        End If
    End If
    
    ' any additional attribute state changes in response the change in the
    ' input attribute can be done below
    
    pAttrToChange.AttrValue = varNewAttrValue
    If (pAttrToChange.attrName = "Edge") Then
        If (varNewAttrValue <> 0) Then
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If (pAttrDescr.attrName = "OffsetDirection") Then
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
                If (pAttrDescr.attrName = "OffsetDirection") Then
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
errorHandler:  HandleError MODULE, METHOD
End Function

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnPreCommit
'
'Abstract
'Gets called before the attribute changes are committed to allow a check of validity.
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of attributes as currently displayed ( prior to Commit )
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

'*************************************************************************
'Function
'IJUserAttributeMgmt_OnPreLoad
'
'Abstract
'Gets called prior to display of attributes on the property page to set readOnly status
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of IJAttributeDescriptor's
'
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo errorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_SPLITMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String

    IJUserAttributeMgmt_OnPreLoad = ""
    
    Exit Function

errorHandler:
    HandleError MODULE, METHOD
End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String

End Function

Private Function IsAllowedSurface(oSurface As Object) As Boolean
Const METHOD = "IsAllowedSurface"
On Error GoTo errorHandler

    IsAllowedSurface = False

    Dim iControlFlags As IJControlFlags

    If Not oSurface Is Nothing Then
        If TypeOf oSurface Is ISPGGridPlane Then
            IsAllowedSurface = True
        ElseIf TypeOf oSurface Is ISPGElevationPlane Then
            IsAllowedSurface = True
        
        ElseIf TypeOf oSurface Is IJPort Then
            Dim oPort As IJPort
            Set oPort = oSurface
            
            IsAllowedSurface = IsAllowedSurface(oPort.Connectable)

        ' want to only locate Root Plate Systems, and not the Leaf Plate Systems.
        ElseIf StrComp(GetObjectClassIDString(oSurface), PlateSystemClassIDStr, vbTextCompare) = 0 Then
            
            Dim idesignChild As IJDesignChild
            Dim oParent As Object
            Set idesignChild = oSurface
            Set oParent = idesignChild.GetParent

            ' a Leaf Plate System will have a parent that is also a Plate System, and is not okay.
            If StrComp(GetObjectClassIDString(oParent), PlateSystemClassIDStr, vbTextCompare) = 0 Then
                IsAllowedSurface = 0

            ' a Root Plate System that has a DesignedMember as its parent is not okay.
            ElseIf TypeOf oParent Is ISPSDesignedMember Then
                IsAllowedSurface = 0
        
            ' it is a RootPlateSystem whose parent is not a DesignedMember.  okay.
            Else
                IsAllowedSurface = 1
            End If
        'DI-133354 Planning Surface is allowed now
        ElseIf StrComp(GetObjectClassIDString(oSurface), PlanningSurfaceClassIDStr, vbTextCompare) = 0 Then
            IsAllowedSurface = 1

        ElseIf TypeOf oSurface Is IJPlane Then     ' okay if CanRule output.  for now just check whether it is not displayed
            Set iControlFlags = oSurface
            If iControlFlags.ControlFlags(CTL_FLAG_NO_DISPLAY) = CTL_FLAG_NO_DISPLAY Then
                IsAllowedSurface = True
            End If
        
        End If
    End If

    Exit Function

errorHandler:
    HandleError MODULE, METHOD
End Function
Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

Private Function GetObjectClassIDString(oObj As Object) As String

    Dim oStructSymbolTools As IJStructSymbolTools
    Dim strOid As String

    Set oStructSymbolTools = New SP3DStructGenericTools.StructSymbolTools
    strOid = oStructSymbolTools.GetObjectOid(oObj)

    GetObjectClassIDString = Mid(strOid, 2, 8)
    
End Function

