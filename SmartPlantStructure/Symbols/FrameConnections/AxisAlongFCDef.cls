VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AxisAlongFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : AxisAlongFCDEF.cls
'
'Author : Structure Development
'
'Description :
'    SmartPlant Structural axis-along Frame Connection Definition file
'
'History:
'
' 05/21/03   Structure Development      Changed error handling in GetRelatedObjects
'                                       and SetRelatedObjects to raise the error;
'                                       also replaced the Msgbox with HandleError
'
' 08/14/03   Structure Development      Substituted interface name in place of guid
'                                       in IJDPropertyDescriptions->AddProperty. The guid of
'                                       a virtual interface may change during bulkload
' 04/18/06   Structure Development      Calls methods to propagate WPO to the other end and to clear
'                                       the offset on the othen end
' 05/05/06   Structure Development      97266  Select which supporting end offset to propagate
' 05/16/06   Structure Development      Curve Member impacts
' 08/23/06   RP                         Added  support for copy mirror(TR#72061)
' 09/27/06   AS                         TR#105470 Remove Blank stubs for CMSetInput and CMRemoveInput
' 09/26/06   Structure Development      TR 104591  ReadPathPointOnObjectsForFC
' 10/03/06   Struct Dev                 TR 103405  OnPreLoad AttributeDescriptor_ReadOnly for member with path
' 06/06/07   Struct Dev                 TR 120531  added CheckForUndefinedValueAndRaiseError
' 08/17/07   Struct Dev                 TR 126260  set default to "Distance" and do measure based on joint position
' 03/07/08   Struct Dev                 TR 88227  Make FCs work for Buil-up members
' 05/14/08   Struct Dev                 DM 142538 remove setting of ByDistance
' 05/11/09   RP                         CR#36434 modified UpdateWPO code to look for Can at conection point
' 07/1/09    RP                         CR165819 - Added code in SetRelatedObjects() to handle the FC's
'                                       association to CanRules
'******************************************************************************************
Option Explicit
Private Const MODULE = "AxisAlongFCDef"
Private m_oLocalizer As IJLocalizer

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.AxisAlongFCDef"
Const m_ItemName As String = "SPSFCMacros.AxisAlongFCDef"



' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements SPSMembers.ISPSFCInputHelper
Implements SPSMembers.ISPSFCInputHelperEx
Implements ISPSMirrorHelper
Implements ISPSTransformHelper
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
  On Error GoTo ErrorHandler

  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs -
  ' They are identical to the class inputs (i.e. penetrated and Penetrating)
  Dim pIH As IJDInputsHelper
  Set pIH = New inputHelper
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
  pAPDs.AddProperty "AxisProps", 1, "IJUASPSFCAxis"
  pAPDs.AddProperty "AxisOffsetProps", 2, "IJUASPSFCManualOffset"
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 

  '*************************WPO*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("AxisWPO", 1, "CMConstructAxisWPO", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalAxisWPO"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAxisWPO"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseAxisWPO"
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisWPO
  pPropertyDescriptions.AddProperty "ComputeAxisWPO", 1, "{34DF6BED-41D5-4D19-B090-D58D93A1CF64}", "CMUpdateAxisWPO", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  '********************************************************************************************************
  
  Set pMemberDescriptions = Nothing
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConditionalAxisWPO
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

Public Sub CMConditionalAxisWPO(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalAxisWPO"
  On Error GoTo ErrorHandler
    bIsNeeded = True ' Always created
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConstructAxisWPO
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

Public Sub CMConstructAxisWPO(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructAxisWPO"
  On Error GoTo ErrorHandler
    Dim oFrameConn As ISPSFrameConnection
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJAttribsConn As IJDAttributes
    Dim oAttrbs As IJDAttributes
    
    Set oFrameConn = pMemberDescription.CAO
    Set pObj = oFrameConn.WPO
    Set oSmartOcc = oFrameConn
    
    Set pIJAttribsConn = oSmartOcc '.ItemObject
    Set oAttrbs = oSmartOcc.ItemObject
    
    If Not IsSOOverridden(pIJAttribsConn.CollectionOfAttributes("IJUASPSFCAxis")) Then
        CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCAxis"), oAttrbs.CollectionOfAttributes("IJUASPSFCAxis")
    End If
    
    If Not IsSOOverridden(pIJAttribsConn.CollectionOfAttributes("IJUASPSFCManualOffset")) Then
        CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCManualOffset"), oAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset")
    End If
      
    SetRecatchOfWPO pObj
      
    Set oAttrbs = Nothing
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMSetInputAxisWPO
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
Public Sub CMSetInputAxisWPO(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputAxisWPO"
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMUpdateAxisWPO
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
Public Sub CMUpdateAxisWPO(pPropertyDescriptions As IJDPropertyDescription, pOutObject As Object)
    Const MT = "CMUpdateAxisWPO"
    On Error GoTo ErrorHandler

    UpdateAxisOffsetWPO pPropertyDescriptions.CAO, Me, pOutObject, lFCAxisAlong
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL    ' Force the FC to the ToDo List
End Sub

'*************************************************************************
'Function
'CMReleaseAxisWPO
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
Public Sub CMReleaseAxisWPO(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseAxisWPO"
On Error GoTo ErrorHandler
    
    Dim iFC As ISPSFrameConnection
    Dim iWPO As ISPSAxisWPO
    
    Set iFC = pMD.CAO
    If Not iFC Is Nothing Then
        ' Clear Work Point Offsets and Work Point CP
        ClearWPO iFC
    End If
    
    Exit Sub
    
ErrorHandler: HandleError MODULE, MT
End Sub

'
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
ErrorHandler:  HandleError MODULE, MT
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

Const METHOD = "IJUserAttributeMgmt_OnPreLoad"

    On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_FCMACROS_EROOR, "ERROR")

    Dim i As Integer, count As Integer
    Dim oFC As ISPSFrameConnection

    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor

    Set oFC = pIJDAttrs
    If Not oFC.MemberSystem.LogicalAxis.CurveObject Is Nothing Then
        Set pColl = CollAllDisplayedValues
        count = pColl.count
        For i = 1 To count
            Set pAttrDescr = pColl.Item(i)
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        Next
    End If

    IJUserAttributeMgmt_OnPreLoad = ""
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD

End Function


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

    selection = "Axis-Along"
    ISPSFCInputHelper_ExecuteSelectionRule = SPSFCInputHelper_Ok

End Function

'*************************************************************************
'Function
'ISPSFCInputHelper_GetRelatedObjects
'
'Abstract
'Returns the objects in the ReferenceCollection used  as input by the FrameConnection.
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
    On Error GoTo ErrorHandler
    
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    IHStatus = SPSFCInputHelper_UnexpectedError

    Dim oRel1 As Object, oRel2 As Object, oRelPath1 As Object, oRelPath2 As Object
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem, oMS3 As ISPSMemberSystem
    Dim oCR As ISPSCanRule
    
    Set oRC = GetRefColl(FC)
    If oRC.IJDEditJDArgument.GetCount <> 2 Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_MISSING_INPUT, "Axis-along frame connection is missing a required input. Change connection to unsupported and then re-select the supporting member.")
    End If

    If TypeOf oRC.IJDEditJDArgument.GetEntityByIndex(1) Is ISPSCanRule Then
        Set oCR = oRC.IJDEditJDArgument.GetEntityByIndex(1)
        oCR.GetPrimaryMemberSystem oMS1
        Set oCR = Nothing
    Else
        Set oMS1 = oRC.IJDEditJDArgument.GetEntityByIndex(1)
    End If
    Set oMS2 = FC.MemberSystem
    If Not oMS1 Is oMS2 Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_SUPPORTINGMEMB_NOTDEFINED, "Supporting member not defined. Change connection to unsupported and then re-select the supporting member.")
    End If

    Set oMS2 = Nothing
    If TypeOf oRC.IJDEditJDArgument.GetEntityByIndex(2) Is ISPSCanRule Then
        Set oCR = oRC.IJDEditJDArgument.GetEntityByIndex(2)
        oCR.GetPrimaryMemberSystem oMS2
        Set oCR = Nothing
    Else
        Set oMS2 = oRC.IJDEditJDArgument.GetEntityByIndex(2)
    End If
    If Not TypeOf oMS2 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_SUPPORTEDMEMB_NOTDEFINED, "Supported member not defined. Change connection to unsupported and then re-select the supporting member.")
    End If

    FC.Joint.GetPointOn oRel1, oRel2
    
    If Not oMS2 Is oRel1 Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_INVALIDPOINT, "Axis-along frame connection's supported member has invalid point-on relationship with another member. Change connection to unsupported and then re-select the supporting member.")
    End If

    If ReadPathPointOnObjectsForFC(FC, oRelPath1, oRelPath2) Then
        If oRelPath1 Is oRel1 And oRelPath2 Is oRel2 Then
            IHStatus = SPSFCInputHelper_Ok
        ElseIf oRelPath1 Is oRel2 And oRelPath2 Is oRel1 Then
            IHStatus = SPSFCInputHelper_Ok
        Else
            IHStatus = SPSFCInputHelper_InconsistentRelations
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_INVALIDPOINT, "Axis-along frame connection's supported member has invalid point-on relationship with another member. Change connection to unsupported and then re-select the supporting member.")
        End If
    Else
        IHStatus = SPSFCInputHelper_Ok
    End If

    'Return the PointOn object, and possibly the intersecting plane
    Set RelatedObject1 = oMS2
    Set RelatedObject2 = oRel2

    Set oRC = Nothing
    Set oMS1 = Nothing
    Set oMS2 = Nothing
    Set oMS3 = Nothing
    Set oRel1 = Nothing
    Set oRel2 = Nothing
    Set oRelPath1 = Nothing
    Set oRelPath2 = Nothing
    Set oCR = Nothing
    
    ISPSFCInputHelper_GetRelatedObjects = IHStatus

    Exit Function

ErrorHandler:
    If ObjectAssocFlags(FC, &H200, 0) Then      ' object is not RELATION_INSERTED_IN_TRANSACTION
         HandleError MODULE, MT                 ' log error
    End If
    Err.Clear
    ISPSFCInputHelper_GetRelatedObjects = IHStatus
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
    Const MT = "ISPSFCInputHelper_SetRelatedObjects"
    On Error GoTo ErrorHandler
    
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    IHStatus = SPSFCInputHelper_UnexpectedError
    
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oR1 As Object, oR2 As Object
    Dim oMeMS As ISPSMemberSystem
    Dim oSingMS As ISPSMemberSystem
    Dim oSingFC As ISPSFrameConnection
    Dim oSysCrSectionNotify As ISPSMemberSystemXSectionNotify
    Dim oRotation As ISPSAxisRotation
    Dim ePointOn As SPSMembers.StructPointOnCondition
    Dim oCanRulePrimary As ISPSCanRule
    Dim oCanRuleSupported As ISPSCanRule
    Dim oLegMembSys As ISPSMemberSystem
    
    
    If Not IsFCCleared(FC) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_PREVIOUSCONN_NOTREMOVED, "Axis-along frame connection is invalid because the previous connection was not removed. Change connection to unsupported and the re-select the supporting member.")
    End If
    If RelatedObject1 Is Nothing And Not RelatedObject2 Is Nothing Then
        Set RelatedObject1 = RelatedObject2
        Set RelatedObject2 = Nothing
    End If

    If RelatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_MISSING_SUPPORTINGMEMB_INFORM, "Axis-along frame connection is missing the supporting member information. Change connection to unsupported and then re-select the supporting member.")
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_CHANGE_CONN, "Axis-along frame connection is related to something other than a supporting member system. Change connection to unsupported and then re-select the supporting member.")
    End If
    
    If Not RelatedObject2 Is Nothing Then
        If TypeOf RelatedObject2 Is ISPSMemberSystem Or TypeOf RelatedObject2 Is IJPoint Then
            Set RelatedObject2 = Nothing
        End If
    End If

    Set oMeMS = FC.MemberSystem
    Set oSingMS = RelatedObject1
    'The Supported MemberSystem cannot be the Supporting MemberSystem.
    If oSingMS Is oMeMS Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_SAMEMEMB_FORBOTH, "Axis-along frame connection is using the same member system for both the supporting and supported member. Change connection to unsupported and then re-select the supporting member.")
    End If

    'save current state of Pointon
    ePointOn = FC.IsPointOn
    
    RemoveGAPRelations FC ' user may be switching from a GapFC, so clear GAP specific relations
    
    'Scoping includes logical axis of supported member, AxisRotationInput and XSectionNotify
    Set oRC = GetRefColl(FC)

    Set oCanRulePrimary = FC.GetCrossSectionObject(SPSFCPrimary)
    Set oCanRuleSupported = FC.GetCrossSectionObject(SPSFCSupported)
    
    oRC.IJDEditJDArgument.RemoveAll

    If oCanRuleSupported Is Nothing Then 'no can on the end, so establish relation with memb sys
        If FC.WPO.portIndex = SPSMemberAxisStart Then
            'Scoping includes logical axis of supported member, AxisRotationInput and XSectionNotify
            oRC.IJDEditJDArgument.SetEntity 1, oMeMS, ISPSMemberSystemStartEndNotify, "MemberSysStartNotifyRC_DEST"
        Else
            'Scoping includes logical axis of supported member, AxisRotation and XSectionNotify
            oRC.IJDEditJDArgument.SetEntity 1, oMeMS, ISPSMemberSystemEndEndNotify, "MemberSysEndNotifyRC_DEST"
        End If
    Else
        FC.SetCrossSectionObject SPSFCSupported, oCanRuleSupported
    End If
    
    If Not oCanRulePrimary Is Nothing Then
        oCanRulePrimary.GetPrimaryMemberSystem oLegMembSys
    End If
    
    If oLegMembSys Is oSingMS Then 'the new supporting member is same as the one with the can that I was related to before
        FC.SetCrossSectionObject SPSFCPrimary, oCanRulePrimary ' add the canrule again
    Else
        'Scoping includes physical axis of supporting member, AxisRotation and XSectionNotify
        oRC.IJDEditJDArgument.SetEntity 2, oSingMS, ISPSMemberSystemSuppingNotify2, "MembSysSuppingNotify2RC_DEST"
    End If


    'create my own joint, unless its ok to keep sharing the joint.
    ConditionallyTransferJoint FC

    'Axis connection does not share joints and is along the supporting member system
    FC.Joint.SetPointOn RelatedObject1, RelatedObject2

    SetPathPointOnObjectsForFC FC, RelatedObject1, RelatedObject2
    
    'For the supporting member, get it's supporting members via FC.InputHelper.GetRelatedObjects
    'If any of those are oFC.MemberSystem, then that FC gets set to Unsupported.
    Set oSingFC = oSingMS.FrameConnectionAtEnd(SPSMemberAxisStart)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If
    Set oSingFC = oSingMS.FrameConnectionAtEnd(SPSMemberAxisEnd)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If
    Set oSingFC = Nothing
    Set oR1 = Nothing
    Set oR2 = Nothing

    RemoveCommonSplitConnections oMeMS, oSingMS

    ' if prev def was NOT pointOn, then set my positionRule to ByDistance and measure based on joint's current position.
' DM 142538 remove setting of ByDistance
'    If ePointOn = SPSNotPointOn Then
'        Dim iJointPODimension As SPSMembers.ISPSPointOnDimension
'        Dim distAlong As Double
'
'        Set iJointPODimension = FC.Joint.PointOnDimension
'        iJointPODimension.Measure SPSPODimension_EndSelection_Auto, SPSPODimension_Position_Distance, distAlong
'        iJointPODimension.PointOnDistance = distAlong
'        iJointPODimension.PointOnEndRule = SPSPODimension_EndSelection_Auto
'        iJointPODimension.PointOnPositionRule = SPSPODimension_Position_Distance
'
'    End If

    AddOutputWPO FC, 1

    IHStatus = FC.inputHelper.GetRelatedObjects(FC, oR1, oR2)
    If IHStatus <> SPSFCInputHelper_Ok Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_AXISALONG_REVERIFIED_INVALID, "Axis-along frame connection was reverified and found to be invalid. Change connection to unsupported and then re-select the supporting member.")
    End If
    
    ISPSFCInputHelper_SetRelatedObjects = SPSFCInputHelper_Ok
    Exit Function

ErrorHandler:
    HandleError MODULE, MT
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
Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal options As Long, ByVal snapDistance As Double, ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, RelatedObject1 As Object, RelatedObject2 As Object, RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus
' FC is optional and can be Nothing
' what we want to do here is if snap is enabled, then
' based on LocatedObject1's type we can return the x,y,z of nearby points of interest such as
' grid-intersections or pointOn joints, endpoints
    
    On Error GoTo ErrorHandler
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim IHStatus2 As SPSMembers.SPSFCInputHelperStatus

    IHStatus = SPSFCInputHelper_UnexpectedError
    
    IHStatus = ExtractValidRelatedObjectFromLocatedObject(FC, False, LocatedObject1, RelatedObject1)
    IHStatus2 = ExtractValidRelatedObjectFromLocatedObject(FC, False, LocatedObject2, RelatedObject2)
    
    If IHStatus2 = SPSFCInputHelper_Ok And IHStatus <> SPSFCInputHelper_Ok Then
        Dim oTemp As Object
        Set oTemp = RelatedObject1
        Set RelatedObject1 = RelatedObject2
        Set RelatedObject2 = oTemp
        IHStatus = IHStatus2
    End If
    
    RelatedObjectX = LocateX
    RelatedObjectY = LocateY
    RelatedObjectZ = LocateZ
    
wrapup:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
    Exit Function

ErrorHandler:
    Err.Clear
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
End Function

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

Private Sub ISPSMirrorHelper_Mirror(ByVal pPlane As SP3DStructInterfaces.IJPlane, ByVal pMatrix As SP3DStructInterfaces.iJDT4x4, ByVal bCopy As Boolean, ByVal oOriginalFC As Object, ByVal oFC As Object)
    Const METHOD = "ISPSMirrorHelper_Mirror"
    On Error GoTo ErrorHandler
    
    Dim oSuppingSysNew As ISPSMemberSystem
    Dim oSuppingSysOld As ISPSMemberSystem
    Dim oSuppedSysOld As ISPSMemberSystem
    
    Dim oFrmConn As ISPSFrameConnection
    Dim oInputHelper As ISPSFCInputHelper
    Dim status As SPSFCInputHelperStatus
    
    Dim oRelObj1 As Object, oRelObj2 As Object
    Dim oPlaneNormal As New DVector, oVec As New DVector
    
    Dim oIJAttributes As IJDAttributes
    Dim oCollProxy As CollectionProxy
    Dim FCSuppingCP As Long
    Dim dotProd As Double
    
    Dim x As Double, y As Double, z As Double
    Dim gX As Double, gY As Double, gZ As Double, lX As Double, lY As Double, lZ As Double
    Dim nx As Double, ny As Double, nz As Double, uX As Double, uY As Double, uZ As Double
    Dim vX As Double, vY As Double, vZ As Double
    
    Dim oMat As New DT4x4
    Dim lCoordSys As Long
    
    'get frame connection
    Set oFrmConn = oFC
    'get input helper
    Set oInputHelper = oFrmConn.inputHelper
    'get the supporting member
    status = oInputHelper.GetRelatedObjects(oFrmConn, oRelObj1, oRelObj2)
    
    On Error Resume Next
    Set oSuppingSysNew = oRelObj1
    If oSuppingSysNew Is Nothing Then
        Set oSuppingSysNew = oRelObj2
    End If
    On Error GoTo ErrorHandler
    
    'get frame connection
    Set oFrmConn = oOriginalFC
    'get input helper
    Set oInputHelper = oFrmConn.inputHelper
    'get the supporting member
    status = oInputHelper.GetRelatedObjects(oFrmConn, oRelObj1, oRelObj2)
    
    On Error Resume Next
    Set oSuppingSysOld = oRelObj1
    If oSuppingSysOld Is Nothing Then
        Set oSuppingSysOld = oRelObj2
    End If
    On Error GoTo ErrorHandler
    
    'we currently handle only the case where the supporting is the same for the copy and
    'original FC.Other cases should work correctly assuming that the supporting is also
    'selected for mirror copy. Where it is not selected it is assumed that the new supporting
    'is a mirror copy of the original supporting.

    Set oIJAttributes = oFC
    
    If oSuppingSysOld Is oSuppingSysNew Then
        pPlane.GetNormal x, y, z
        oPlaneNormal.Set x, y, z
        
        oPlaneNormal.Length = 1
        'get the supporting local z and local y
        oSuppingSysNew.Rotation.GetOrientationVector x, y, z
        oVec.Set x, y, z
        oVec.Length = 1
        
        dotProd = Abs(oVec.Dot(oPlaneNormal))
        
        Set oCollProxy = oIJAttributes.CollectionOfAttributes("IJUASPSFCAxis")
        
        FCSuppingCP = oCollProxy.Item("SupportingCP").Value
        
        If Abs(dotProd - 1) < distTol Then        'check if mirror plane normal is parallel to local z
            If FCSuppingCP = 7 Or FCSuppingCP = 8 Or FCSuppingCP = 9 Then
                oCollProxy.Item("SupportingCP").Value = FCSuppingCP - 6
            ElseIf FCSuppingCP = 1 Or FCSuppingCP = 2 Or FCSuppingCP = 3 Then
                oCollProxy.Item("SupportingCP").Value = FCSuppingCP + 6
            End If
        Else
            oSuppingSysNew.Rotation.GetTransform oMat
            'get y vector
            oVec.Set oMat.IndexValue(4), oMat.IndexValue(5), oMat.IndexValue(6)
            oVec.Length = 1
            dotProd = Abs(oVec.Dot(oPlaneNormal))
            If Abs(dotProd - 1) < distTol Then 'check if mirror plane normal is parallel to local y
                If FCSuppingCP = 1 Or FCSuppingCP = 4 Or FCSuppingCP = 7 Then
                    oCollProxy.Item("SupportingCP").Value = FCSuppingCP + 2
                ElseIf FCSuppingCP = 3 Or FCSuppingCP = 6 Or FCSuppingCP = 9 Then
                    oCollProxy.Item("SupportingCP").Value = FCSuppingCP - 2
                End If
            End If
        End If
    End If
    
    'get offset vector of the old member
    Set oCollProxy = oIJAttributes.CollectionOfAttributes("IJUASPSFCManualOffset")
    lCoordSys = oCollProxy.Item("CoordinateSystem").Value

    ' if it is local coordinate system then the only time we need to transform it is
    ' when the local coordinate system has not been transformed. This happens when we
    ' have the same supporting system.

    If (lCoordSys = 2 And oSuppingSysOld Is oSuppingSysNew) Or lCoordSys = 1 Then
           
        'Need to mirror the offset vector now
         
        x = oCollProxy.Item("XOffset").Value
        y = oCollProxy.Item("YOffset").Value
        z = oCollProxy.Item("ZOffset").Value
        oVec.Set x, y, z

        If lCoordSys = 2 Then ' local to supporting
            oSuppingSysOld.Rotation.GetTransform oMat
            'transform to global
            Set oVec = oMat.TransformVector(oVec)
        End If
    
        'create a transformation matrix based on mirror plane
        pPlane.GetNormal nx, ny, nz
        pPlane.GetUDirection uX, uY, uZ
        pPlane.GetVDirection vX, vY, vZ
        
        oMat.LoadIdentity
         
        oMat.IndexValue(0) = uX
        oMat.IndexValue(1) = uY
        oMat.IndexValue(2) = uZ
         
        oMat.IndexValue(4) = vX
        oMat.IndexValue(5) = vY
        oMat.IndexValue(6) = vZ
         
        oMat.IndexValue(8) = nx
        oMat.IndexValue(9) = ny
        oMat.IndexValue(10) = nz
         
        oMat.Invert 'global to local
        Set oVec = oMat.TransformVector(oVec)
         
        'to mirror, reverse z
        oVec.z = -oVec.z
         
        oMat.Invert 'transform back to global
        Set oVec = oMat.TransformVector(oVec)
         
        'if the coordsystem is global then we are done
         
        'if the coordystem is local then conver the vector to local
        If lCoordSys = 2 Then ' local to supporting
            oSuppingSysNew.Rotation.GetTransform oMat
            oMat.Invert
            'transform to local
            Set oVec = oMat.TransformVector(oVec)
        End If
    
        oCollProxy.Item("XOffset").Value = oVec.x
        oCollProxy.Item("YOffset").Value = oVec.y
        oCollProxy.Item("ZOffset").Value = oVec.z
     
    End If
        
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub ISPSTransformHelper_Transform(ByVal Trans4x4 As SP3DStructInterfaces.iJDT4x4, ByVal oFC As Object)
    Const METHOD = "ISPSTransformHelper_Transform"
    On Error GoTo ErrorHandler
    
    Dim x As Double, y As Double, z As Double
    Dim oOffset As IJDVector
    Dim coordSys As Long
    
    Dim oFrameConn As ISPSFrameConnection
    
    Dim oAttributes As IJDAttributes
    Dim oOffsets As CollectionProxy
    Dim oRotationMatrix As iJDT4x4
    
    Set oFrameConn = oFC
    Set oAttributes = oFrameConn
    
    Set oOffsets = oAttributes.CollectionOfAttributes("IJUASPSFCManualOffset")
    coordSys = oOffsets.Item("CoordinateSystem").Value
    
    ' if coordinate system is global
    If coordSys = 1 Then
    
        x = oOffsets.Item("XOffset").Value
        y = oOffsets.Item("YOffset").Value
        z = oOffsets.Item("ZOffset").Value
        
        Set oOffset = New DVector
        
        oOffset.Set x, y, z
        
        Set oOffset = Trans4x4.TransformVector(oOffset)
        
        oOffsets.Item("XOffset").Value = oOffset.x
        oOffsets.Item("YOffset").Value = oOffset.y
        oOffsets.Item("ZOffset").Value = oOffset.z
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
