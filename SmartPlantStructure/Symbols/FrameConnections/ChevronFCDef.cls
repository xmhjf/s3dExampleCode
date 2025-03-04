VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChevronFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : ChevronFCDEF.cls
'
'Author : A. Patil
'
'Description :
'    SmartPlant Structural chevron Frame Connection Definition file
'
'History:
'
' 05/21/03   J.Schwartz     Added header
'
' 08/14/03   R. Peringoth   Substituted interface name in place of guid
'                           in IJDPropertyDescriptions->AddProperty. The guid of
'                           a virtual interface may change during bulkload
' 09/27/06   AS             TR#105470 Remove Blank stubs for CMSetInput and CMRemoveInput

'********************************************************************


Private Const MODULE = "ChevronFCDef"
Private m_oLocalizer As IJLocalizer

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.ChevronFCDef"
Const m_ItemName As String = "SPSFCMacros.ChevronFCDef"

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements SPSMembers.ISPSFCInputHelper
Implements SPSMembers.ISPSFCInputHelperEx
Implements SP3DStructInterfaces.IJUserAttributeMgmt
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
  On Error GoTo ErrorHandler
  pIH.SetInput "SupportedRefColl"
  pIH.SetInput "SupportingRefColl"
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub



Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
 IJDUserSymbolServices_EditOccurence = False
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

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
  IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function

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
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler

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
  pAD.AggregatorClsid = "{5FEB4ADB-E5EC-45D6-8878-150ADDC04D0A}"     'CSPSFrameConnection
  Set pAD = Nothing
  
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  ' Add following interfaces as inputs so as to get notified when they are modified.
  'These interfaces carry properties that drive rotation and offset
  pAPDs.AddProperty "WPOProps", 1, "IJUASPSChevron"
  pAPDs.AddProperty "ChevronOffsetProps", 2, "IJUASPSFCOffset"
    
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 

  '*************************WPO*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("ChevronWPO", 1, "CMConstructChevronWPO", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalChevronWPO"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputChevronWPO"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseChevronWPO"
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisWPO
  pPropertyDescriptions.AddProperty "ComputeChevronWPO", 1, "{34DF6BED-41D5-4D19-B090-D58D93A1CF64}", "CMUpdateChevronWPO", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  '********************************************************************************************************
  
  Set pMemberDescriptions = Nothing
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConditionalChevronWPO
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

Public Sub CMConditionalChevronWPO(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalChevronWPO"
  On Error GoTo ErrorHandler
    bIsNeeded = True ' Always created
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMConstructChevronWPO
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

Public Sub CMConstructChevronWPO(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructChevronWPO"
  On Error GoTo ErrorHandler
  Dim oFrameConn As ISPSFrameConnection
  
  Set oFrameConn = pMemberDescription.CAO
  Set pObj = oFrameConn.WPO

  Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMSetInputChevronWPO
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
Public Sub CMSetInputChevronWPO(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputChevronWPO"
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMUpdateChevronWPO
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
Public Sub CMUpdateChevronWPO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMUpdateChevronWPO"
    On Error GoTo ErrorHandler

    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim obj As Object
    Dim oWPO As ISPSAxisWPO
        
    Dim oSys As ISPSMemberSystem
    Dim pIJAttribs As IJDAttributes
    Dim FC As ISPSFrameConnection

    Dim iSupportedEnd As SPSMemberAxisPortIndex
    Dim oSupportingVec As New DVector, oSupportedVec As New DVector
    Dim oRotation As ISPSAxisRotation
    Dim iJDT4x4 As iJDT4x4
    Dim oLocPos As New DPosition, oGloPos As DPosition

    Dim dOffsetAlong As Double       'distance along supporting member
    Dim iNearMidFar As Long     'near mid far
    Dim dOffsetX As Double, dOffsetY As Double, dOffsetZ As Double

    Set FC = pPropertyDescriptions.CAO
    Set pIJAttribs = FC

    IHStatus = FC.InputHelper.GetRelatedObjects(FC, oSys, obj)
   
    If IHStatus <> SPSFCInputHelper_Ok Then
        Set FC.definition = Nothing
        GoTo ErrorHandler
    End If
    
    dOffsetAlong = pIJAttribs.CollectionOfAttributes("IJUASPSFCOffset").Item("Offset").Value
    iNearMidFar = pIJAttribs.CollectionOfAttributes("IJUASPSFCChevron").Item("WorkPoint").Value

    FC.MemberSystem.Rotation.GetTransform iJDT4x4
    oSupportedVec.x = iJDT4x4.IndexValue(8)
    oSupportedVec.y = iJDT4x4.IndexValue(9)
    oSupportedVec.z = iJDT4x4.IndexValue(10)
    
    oSys.Rotation.GetTransform iJDT4x4
    iJDT4x4.IndexValue(12) = 0
    iJDT4x4.IndexValue(13) = 0
    iJDT4x4.IndexValue(14) = 0

    oSupportingVec.x = iJDT4x4.IndexValue(8)
    oSupportingVec.y = iJDT4x4.IndexValue(9)
    oSupportingVec.z = iJDT4x4.IndexValue(10)
    
    iSupportedEnd = FC.WPO.portIndex
    If iSupportedEnd = SPSMemberAxisStart Then
        If oSupportingVec.Dot(oSupportedVec) < 0# Then
            dOffsetAlong = -dOffsetAlong
        End If
    Else
        If oSupportingVec.Dot(oSupportedVec) > 0# Then
            dOffsetAlong = -dOffsetAlong
        End If
    End If

    'find dot product of supported / supporting axis to see whether to use negative value of offset
    oLocPos.Set 0#, 0#, dOffsetAlong
    
    ' for now, this code is ignoring the near/mid/far parameter
    Set oGloPos = iJDT4x4.TransformPosition(oLocPos)

'        FC.Services.ComputeChevronOffsets( FC.MemberSystem, FC.WPO.PortIndex, oSys, _
'                        iNearMidFar, dOffsetAlong, dOffsetX, dOffsetY, dOffsetZ )
        
    Set oWPO = pObject
        
    oWPO.SetWPO oGloPos.x, oGloPos.y, oGloPos.z, 0, 0, 0
    oWPO.WPOCardinalPoint = 0
    
    Dim currType As Long
    currType = oWPO.MemberSystem.MemberType.Type
    oWPO.MemberSystem.MemberType.Type = currType + 1
    oWPO.MemberSystem.MemberType.Type = currType
        
    Set oSys = Nothing
    Set pIJAttribs = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL    ' Force the FC to the ToDo List

End Sub

'*************************************************************************
'Function
'CMReleaseChevronWPO
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
Public Sub CMReleaseChevronWPO(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseChevronWPO"
On Error GoTo ErrorHandler
    
    Dim iFC As ISPSFrameConnection
    Dim iWPO As ISPSAxisWPO
    
    Set iFC = pMD.CAO
    Set iWPO = iFC.WPO
        
    ' Clear Work Point Offsets and Work Point CP
    iWPO.SetWPO 0#, 0#, 0#, 0#, 0#, 0#
    iWPO.WPOCardinalPoint = 0
    
  Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL    ' Force the FC to the ToDo List

End Sub
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
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
   Const MT = "IJDUserSymbolServices_InstanciateDefinition"
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
ErrorHandler:    HandleError MODULE, MT
 
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub


Private Property Get ISPSFCInputHelperEx_AMPOption() As SPSMembers.SPSFCAMPOptions
    ISPSFCInputHelperEx_AMPOption = SPSFCAMPOption_NONE
End Property

'*************************************************************************
'Function
'ISPSFCInputHelper_ExecuteSelectionRule
'
'Abstract
'Selects a particular type of FrameConnection to be used
'
'Arguments
'FC as FrameConnection
'name as String
'
'Return
'Returns the catalog smartitem name as string and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_ExecuteSelectionRule(ByVal FC As SPSMembers.ISPSFrameConnection, selection As String) As SPSMembers.SPSFCInputHelperStatus
    selection = "Chevron-1"
    ISPSFCInputHelper_ExecuteSelectionRule = SPSFCInputHelper_Ok
End Function


'*************************************************************************
'Function
'ISPSFCInputHelper_GetRelatedObjects
'
'Abstract
'Returns the objects in the ReferenceCollection used   as input by the FrameConnection.
'
'Arguments
'FC As ISPSFrameConnection
'RelatedObject1 As Object
'RelatedObject2 As Object
'
'Return
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_GetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, RelatedObject1 As Object, RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus

    Const MT = "ISPSFCInputHelper_GetRelatedObjects"
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim count As Long
    Dim oObj1 As Object, oObj2 As Object

    On Error GoTo ErrorHandler
    IHStatus = SPSFCInputHelper_UnexpectedError

    Set oRC = GetRefColl(FC)
  
    count = oRC.IJDEditJDArgument.GetCount
    If count < 2 Or count > 2 Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNEXPECTED_REFCOLL, "Unexpected refcoll count=") & count
    End If

    Set RelatedObject1 = oRC.IJDEditJDArgument.GetEntityByIndex(2)
    FC.Joint.GetPointOn oObj1, oObj2
    
    If RelatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_NOOBJ2_INREF, "No object 2 in Reference coll")
    End If
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNKNOWNOBJ1_INREF, "Unknown object 1 in Reference coll")
    End If

    If Not oObj1 Is RelatedObject1 Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNKNOWNOBJ1_INREF, "Unknown object 1 in Reference coll")
    End If

    If RelatedObject1 Is FC.MemberSystem Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_RELOBJ1_SHOULDBE_SUPPORTING_MEMB, "RelatedObject1 should be the supporting MemberSystem")
    End If

    If Not oObj2 Is Nothing Then
        If TypeOf oObj2 Is ISPSMemberSystem Then
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNKNOWNOBJ2_INREF, "Unknown object 2 in Reference coll")
        End If
        Set RelatedObject2 = oObj2
    End If

    IHStatus = SPSFCInputHelper_Ok
    ISPSFCInputHelper_GetRelatedObjects = IHStatus
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
    Err.Clear
End Function

'*************************************************************************
'Function
'ISPSFCInputHelper_SetRelatedObjects
'
'Abstract
'Sets the objects in the ReferenceCollection used as input by the FrameConnection.
'SetRelatedObjects is called by the command to enable this code to set connections to the FrameConnection's ReferenceCollection.
'
'Different connections will establish different relations according to what needs to be watched.
'This connection is the RootSelector connection, and serves as a hub for the "ByRule".
'
'See documentation and metadata for more details.
'
'Arguments
'FC As ISPSFrameConnection
'RelatedObject1 As Object
'RelatedObject2 As Object
'
'Return
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_SetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal RelatedObject1 As Object, ByVal RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus
    Const METHOD = "ISPSFCInputHelper_SetRelatedObjects"

    On Error GoTo ErrorHandler
    Dim IHStatus As SPSFCInputHelperStatus
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oJoint As ISPSAxisJoint
    Dim oMS As ISPSMemberSystem

    If Not IsFCCleared(FC) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_PREVIOUS_RELATIONS, "Previous Relations Not Cleared")
    End If
    
    If RelatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_NORELATED_OBJ1, "No RelatedObject1")
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_UNKNOWN_RELATEDOBJ1, "Unknown type of RelatedObject1")
    End If

    'The Supported MemberSystem cannot be the Supporting MemberSystem.
    If RelatedObject1 Is FC.MemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_RELOBJ1_SHOULDBE_SUPPORTING_MEMB, "RelatedObject1 should be the supporting MemberSystem")
    End If

    Set oRC = GetRefColl(FC)
    oRC.IJDEditJDArgument.RemoveAll

    Set oMS = FC.MemberSystem
    If FC.WPO.portIndex = SPSMemberAxisStart Then
        'Scoping includes logical axis of supported member, AxisRotationInput and XSectionNotify
        oRC.IJDEditJDArgument.SetEntity 1, oMS, ISPSMemberSystemStartEndNotify, "MemberSysStartNotifyRC_DEST"
    Else
        'Scoping includes logical axis of supported member, AxisRotation and XSectionNotify
        oRC.IJDEditJDArgument.SetEntity 1, oMS, ISPSMemberSystemEndEndNotify, "MemberSysEndNotifyRC_DEST"
    End If

    Set oMS = RelatedObject1
    'Scoping includes physical axis of supporting member, AxisRotation and XSectionNotify
    oRC.IJDEditJDArgument.SetEntity 2, oMS, ISPSMemberSystemSuppingNotify2, "MembSysSuppingNotify2RC_DEST"

    FindSameSideJoint oMS, FC, oJoint
    
    If Not oJoint Is Nothing Then
        oJoint.AddMember FC.MemberSystem, FC.WPO.portIndex
    Else
        FC.Joint.SetPointOn RelatedObject1, RelatedObject2
    End If
    
    ISPSFCInputHelper_SetRelatedObjects = SPSFCInputHelper_Ok
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    ISPSFCInputHelper_SetRelatedObjects = IHStatus
End Function

'*************************************************************************
'Function
'ISPSFCInputHelper_UserAttributeMgmt
'
'Abstract
'Function on ISPSFACInputHelper to return the UserAttributeMgmt interface
'
'Arguments
'
'Return
'The interface is returned.
'
'Exceptions
'
'***************************************************************************
Private Property Get ISPSFCInputHelper_UserAttributeMgmt() As SPSMembers.IJUserAttributeMgmt
    Set ISPSFCInputHelper_UserAttributeMgmt = Me
End Property

'*************************************************************************
'Function
'ISPSFCInputHelper_ValidateLocatedObjects
'
'Abstract
'Validates the objects to be set as input to the ReferenceCollection for the FrameConnection.
'ValidateLocatedObjects is called by the command during Mouse Move events of PlaceMemberSystem command
'Depending on the type of FrameConnection selected in the ribbon bar its repective ValidateLocatedObjects
'function will be called to determine if that object is valid for placing the selected type of FrameConnection
'the located objects are validated and returned as relatedobjects in the argument
'
'Arguments
'FC As ISPSFrameConnection
'LocatedObject1 As Object
'LocatedObject2 As Object
'RelatedObject1 As Object
'RelatedObject2 As Object
'RelatedPositions X,Y,Z as Doubles
'LocatedPositions X,Y,Z as Doubles
'The other arguments are not being used and is for future enhancements
'
'Returns the input objects as arguments and SPSFCInputHelperStatus representing the error occurred
'
'Exceptions
'
'***************************************************************************
Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal options As Long, ByVal snapDistance As Double, _
    ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, _
    RelatedObject1 As Object, RelatedObject2 As Object, RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus

Dim oFC As ISPSFrameConnection
Dim IHStatus As SPSFACInputHelperStatus
Dim oR1 As Object, oR2 As Object

IHStatus = SPSFACInputHelper_UnexpectedError
RelatedObjectX = LocateX
RelatedObjectY = LocateY
RelatedObjectZ = LocateZ
Set RelatedObject1 = Nothing
Set RelatedObject2 = Nothing

If LocatedObject1 Is Nothing And LocatedObject2 Is Nothing Then     ' no object located
    ISPSFCInputHelper_ValidateLocatedObjects = SPSFCInputHelper_BadNumberOfObjects
    GoTo wrapup
End If

If Not LocatedObject1 Is Nothing And Not LocatedObject2 Is Nothing Then     ' two objects located
    If TypeOf LocatedObject1 Is ISPSMemberSystem Then
        If TypeOf LocatedObject2 Is ISPSMemberSystem Then
            IHStatus = SPSFACInputHelper_InvalidTypeOfObject    'we COULD figure out which member is better and set the other to Nothing
        Else
            Set RelatedObject1 = LocatedObject1
            Set RelatedObject2 = LocatedObject2
            IHStatus = SPSFCInputHelper_Ok
        End If
    ElseIf TypeOf LocatedObject2 Is ISPSMemberSystem Then
        Set RelatedObject1 = LocatedObject2
        Set RelatedObject2 = LocatedObject1
        IHStatus = SPSFCInputHelper_Ok
    Else
        IHStatus = SPSFACInputHelper_InvalidTypeOfObject    'no go.  neither one is a MemberSystem
    End If
    GoTo wrapup
End If

'One object was located

If LocatedObject1 Is Nothing Then       ' make locatedObject1 be the located object
    Set LocatedObject1 = LocatedObject2
    Set LocatedObject2 = Nothing
End If

Dim oMS As ISPSMemberSystem
Dim oJoint As ISPSAxisJoint

If TypeOf LocatedObject1 Is ISPSMemberSystem Then

    Set oMS = LocatedObject1
    Set oJoint = oMS.JointAtEndPosition(LocateX, LocateY, LocateZ)
    If Not oJoint Is Nothing Then       'located at the end.  Is it a Chevron FC ?
        Set oFC = oMS.FrameConnectionAtEnd(oMS.ResolveEnd(oJoint))
        If oFC.DefinitionClassName = "Chevron" Then
            Set oMS = Nothing           'pretend he located the FC
        Else
            Set oFC = Nothing           'keep what was located.
        End If
    End If

ElseIf TypeOf LocatedObject1 Is ISPSFrameConnection Then
    Set oFC = LocatedObject1
    If oFC.DefinitionClassName <> "Chevron" Then
        Set oMS = oFC.MemberSystem
        Set oFC = Nothing
    End If
End If

If Not oFC Is Nothing Then      ' we know it is a Chevron
    IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oR1, oR2)
    If IHStatus <> SPSFACInputHelper_Ok Then
        ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
        GoTo wrapup
    End If

    Set RelatedObject1 = oR1
    Set RelatedObject2 = oR2
    IHStatus = SPSFACInputHelper_Ok

ElseIf Not oMS Is Nothing Then
    
    Set RelatedObject1 = oMS
    IHStatus = SPSFACInputHelper_Ok

Else        'return bad input
    IHStatus = SPSFACInputHelper_InvalidTypeOfObject

End If

wrapup:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
    Exit Function

ErrorHandler:
    Err.Clear
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

End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String

End Function

Public Sub FindSameSideJoint(oSingMS As ISPSMemberSystem, oFC As ISPSFrameConnection, ByRef oBestJoint As ISPSAxisJoint)

    Const METHOD = "FindSameSideJoint"
    On Error GoTo ErrorHandler

    Dim iEnd As SPSMemberAxisPortIndex
    Dim myQuadrant As Long, otherQuadrant As Long
    Dim ii As Long, jj As Long, iJointCount As Long, iMSCount As Long
    Dim elePOJoints As IJElements, eleSupportedMS As IJElements
    Dim oJoint As ISPSAxisJoint
    Dim oOtherMS As ISPSMemberSystem
    Dim dPosx As Double, dPosy As Double, dPosz As Double

    Set oBestJoint = Nothing

    oFC.Joint.Point.GetPoint dPosx, dPosy, dPosz
    Set elePOJoints = oSingMS.PointOnJointsAtPosition(dPosx, dPosy, dPosz)
    iJointCount = elePOJoints.count
    
    'For each joint, need to check whether the member is in the same quadrant of the supporting member.
    If iJointCount > 0 Then
        oFC.Services.ComputeIncidentMemberQuadrant oSingMS, oFC.MemberSystem, oFC.WPO.portIndex, myQuadrant
        For ii = 1 To iJointCount
            Set oJoint = elePOJoints.Item(ii)
            Set eleSupportedMS = oJoint.EndMemberSystems
            iMSCount = eleSupportedMS.count

            For jj = 1 To eleSupportedMS.count
                Set oOtherMS = eleSupportedMS.Item(jj)
                iEnd = oOtherMS.ResolveEnd(oJoint)
                If oOtherMS.FrameConnectionAtEnd(iEnd).DefinitionClassName = "Chevron" Then
                    
                    oFC.Services.ComputeIncidentMemberQuadrant oSingMS, oOtherMS, iEnd, otherQuadrant
                
                    'Accept the first joint we find that has a MS on the same quadrant
                    'as new supported member if that MS is connected using a Chevron FC
                
                    If otherQuadrant = myQuadrant Then
                        Set oBestJoint = oJoint
                        GoTo wrapup
                    End If
                End If
            Next jj
        Next ii
    End If

wrapup:
    Set oJoint = Nothing
    Set oOtherMS = Nothing
    Set eleSupportedMS = Nothing
    Set elePOJoints = Nothing
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub
Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
