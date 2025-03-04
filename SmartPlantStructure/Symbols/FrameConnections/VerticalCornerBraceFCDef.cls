VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VerticalCornerBraceFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : VerticalCornerBraceFCDEF.cls
'
'Author : Structure Development
'
'Description :
'    SmartPlant Structural vertical corner barce Frame Connection
'       Definition file
'
'History:
'
' 05/21/03   Structure Development    Changed error handling in GetRelatedObjects
'                                     and SetRelatedObjects to raise the error;
'                                     also replaced the Msgbox with HandleError
' 08/13/03   Structure Development    Add brace as PointOn to existing Joint
'
' 08/14/03   Structure Development    Substituted interface name in place of guid
'                                     in IJDPropertyDescriptions->AddProperty. The guid of
'                                     a virtual interface may change during bulkload
'
' 01/21/04   Structure Development    Added additional checks in the GetRelatedObjects
'                                     function to validate that the brace end joint
'                                     is point on to another joint and the joint is one
'                                     of the supporting member joints (TR#54247)
' 03/03/04   Structure Development    Fixed a FooPah in GetRelatedObjects where
'                                     the end of the function was suppose to null
'                                     out the oPointOnJoint object but instead was
'                                     setting it to either a joint or a member system
'                                     which would cause a Type Mismatch error (TR#56022)
'
' 04/11/06   Structure Development    Added x,y,z offset and coord system to this FC
' 05/16/06   Structure Development    Curve Member impacts
' 15/09/06   RP                       Moved IsTwoMemberSystemsOkforVCB() to the common module from the selector.
'                                     Tha validate method calls this. Now checks if all three members are in the
'                                     same quadrant (TR#104603)
' 09/27/06   AS                       TR#105470 Remove Blank stubs for CMSetInput and CMRemoveInput
' 06/06/07   Struct Dev               TR 120531  added CheckForUndefinedValueAndRaiseError
' 07/1/09    RP                       CR165819 - Added code in SetRelatedObjects() to handle the FC's
'                                     association to CanRules
'*****************************************************************************
Option Explicit

Private Const MODULE = "VerticalCornerBraceFCDef"
Private m_oLocalizer As IJLocalizer

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.VerticalCornerBraceFCDef"
Const m_ItemName As String = "SPSFCMacros.VerticalCornerBraceFCDef"

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements ISPSTransformHelper
Implements ISPSMirrorHelper
Implements SP3DStructInterfaces.IJUserAttributeMgmt
Implements SPSMembers.ISPSFCInputHelper
Implements SPSMembers.ISPSFCInputHelperEx

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
  pIH.SetInput "Supporting1RefColl"
  pIH.SetInput "Supporting2RefColl"
  
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
  pAPDs.AddProperty "VerticalCornerBraceProps", 1, "IJUASPSFCVerticalCornerBrace"
  pAPDs.AddProperty "VerticalCornerBraceOffsetProps", 2, "IJUASPSFCManualOffset"
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 

  '*************************WPO*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("VerticalCornerBraceWPO", 1, "CMConstructVerticalCornerBraceWPO", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalVerticalCornerBraceWPO"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputVerticalCornerBraceWPO"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseVerticalCornerBraceWPO"
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisWPO {34DF6BED-41D5-4D19-B090-D58D93A1CF64}
  pPropertyDescriptions.AddProperty "ComputeVerticalCornerBraceWPO", 1, "{34DF6BED-41D5-4D19-B090-D58D93A1CF64}", "CMUpdateVerticalCornerBraceWPO", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  '********************************************************************************************************
  
  Set pMemberDescriptions = Nothing
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConditionalVerticalCornerBraceWPO
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

Public Sub CMConditionalVerticalCornerBraceWPO(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalVerticalCornerBraceWPO"
  On Error GoTo ErrorHandler
    bIsNeeded = True ' Always created
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConstructVerticalCornerBraceWPO
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

Public Sub CMConstructVerticalCornerBraceWPO(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructVerticalCornerBraceWPO"
  On Error GoTo ErrorHandler
    Dim oFrameConn As ISPSFrameConnection
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJAttribsConn As IJDAttributes
    Dim oAttrbs As IJDAttributes
    
    Set oFrameConn = pMemberDescription.CAO
    Set pObj = oFrameConn.WPO
    Set oSmartOcc = oFrameConn
    
    Set pIJAttribsConn = oSmartOcc
    Set oAttrbs = oSmartOcc.ItemObject
    If Not IsSOOverridden(pIJAttribsConn.CollectionOfAttributes("IJUASPSFCVerticalCornerBrace")) Then
        CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCVerticalCornerBrace"), oAttrbs.CollectionOfAttributes("IJUASPSFCVerticalCornerBrace")
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
'CMSetInputVerticalCornerBraceWPO
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
Public Sub CMSetInputVerticalCornerBraceWPO(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputVerticalCornerBraceWPO"
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMUpdateVerticalCornerBraceWPO
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
Public Sub CMUpdateVerticalCornerBraceWPO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMUpdateVerticalCornerBraceWPO"
On Error GoTo ErrorHandler

    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim iSO As IJSmartOccurrence
    Dim iWPO As ISPSAxisWPO
    Dim WPOx As Double, WPOy As Double, WPOz As Double
    Dim pIJAttrbs As IJDAttributes
    Dim iOriginalPhysicalPoint As IJPoint
    Dim oOriginalVec As New DVector
    Dim oSupportedMembSys As ISPSMemberSystem
    Dim oSupportingPrimaryMembSys As ISPSMemberSystem
    Dim oSupportingSecondaryMembSys As ISPSMemberSystem

    Dim oFC As ISPSFrameConnection
    Dim cornerPoint As Long
    Dim xOriginal As Double, yOriginal As Double, zOriginal As Double
    
    Set oFC = pPropertyDescriptions.CAO
    Set oSupportedMembSys = oFC.MemberSystem
    Set iOriginalPhysicalPoint = oFC
    If Not iOriginalPhysicalPoint Is Nothing Then
        iOriginalPhysicalPoint.GetPoint xOriginal, yOriginal, zOriginal
    End If
    IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oSupportingPrimaryMembSys, oSupportingSecondaryMembSys)
    
    ' if error occured at reading supporting objects, it may be fixable so call FixFC to see.
    ' if FixFC fails to fix it, it returns false, the FC has been set to Unsupported and we must exit.
    ' if FixFC returns okay, try again to get_RelatedObjects.  If it fails now, we set to unsupported here and exit.
    ' Otherwise, we have been healed and are okay to continue
    
    If IHStatus <> SPSFCInputHelper_Ok Then
        If FixFC(oFC) Then
            IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oSupportingPrimaryMembSys, oSupportingSecondaryMembSys)
            If IHStatus <> SPSFCInputHelper_Ok Then
                Set oFC.definition = Nothing
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    If IsRelatedSplitConnectionInBadState(oFC) Then
        GoTo ErrorHandler
    End If
    
    Set iSO = pPropertyDescriptions.CAO
    Set pIJAttrbs = iSO
    
    cornerPoint = pIJAttrbs.CollectionOfAttributes("IJUASPSFCVerticalCornerBrace").Item("WorkPoint").Value
    
    'Following attributes are added in v7. They may not exist in older catalog.
    Dim xOffset As Double, yOffset As Double, zOffset As Double
    
    Dim lCoordSys As Long
    Dim oLocalToGlobal As iJDT4x4
    
    xOffset = 0
    yOffset = 0
    zOffset = 0
    
    lCoordSys = 1
    
    On Error Resume Next
    lCoordSys = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("CoordinateSystem").Value
    On Error GoTo ErrorHandler
    If lCoordSys = 2 Then 'Local
        getLocalToGlobal oFC, oLocalToGlobal
    End If

    If Not lCoordSys = 0 Then 'VT_EMPTY (not set by user) which is not in the codelist
        CheckForUndefinedValueAndRaiseError oFC, lCoordSys, "StructCoordSysReference", 10
    End If
    CheckForUndefinedValueAndRaiseError oFC, cornerPoint, "StructuralWPOIntersectionIndex", 8

    ' It is possible that the following method will fail as
    '   a result of providing it invalid inputs hence this function
    '   must return E_FAIL to force the offending item to go to the ToDO list.
    'Dim bUsePhysicalPointAsInput As Boolean
    oFC.Services.ComputeWPOIntersection oSupportingPrimaryMembSys, oSupportingSecondaryMembSys, _
        oSupportedMembSys, oFC.WPO.portIndex, cornerPoint, WPOx, WPOy, WPOz
    Set iWPO = pObject
    
    Dim oVec As New DVector
      
    If oFC.IsImporting And Not iOriginalPhysicalPoint Is Nothing Then
        'Adjust the offset to keep the physical unchanged
        Dim xJointPosition As Double, yJointPosition As Double, zJointPosition As Double
        oFC.Joint.Point.GetPoint xJointPosition, yJointPosition, zJointPosition
        iWPO.SetWPO xOriginal - xJointPosition, yOriginal - yJointPosition, zOriginal - zJointPosition, 0, 0, 0
        
        xOffset = xOriginal - xJointPosition - WPOx
        yOffset = yOriginal - yJointPosition - WPOy
        zOffset = zOriginal - zJointPosition - WPOz
        If lCoordSys = 2 Then 'Local
            'convert global offset to local to save
            oLocalToGlobal.Invert
            oVec.Set xOffset, yOffset, zOffset
            Set oVec = oLocalToGlobal.TransformVector(oVec)
            oVec.Get xOffset, yOffset, zOffset
        End If
        pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("XOffset").Value = xOffset
        pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("YOffset").Value = yOffset
        pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("ZOffset").Value = zOffset
    Else
        On Error Resume Next
        xOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("XOffset").Value
        yOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("YOffset").Value
        zOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("ZOffset").Value
        On Error GoTo ErrorHandler
        If lCoordSys = 2 Then 'Local
            'convert local to global offset
            oVec.Set xOffset, yOffset, zOffset
            Set oVec = oLocalToGlobal.TransformVector(oVec)
            oVec.Get xOffset, yOffset, zOffset
        End If
        iWPO.SetWPO WPOx + xOffset, WPOy + yOffset, WPOz + zOffset, 0, 0, 0
    End If
    
    iWPO.WPOCardinalPoint = 0

    If oFC.IsImporting Then
        SetWPOAtOtherEndIfUnsupported oFC    ' if other end is unsupported, set its offsets to keep physical where it is.
        oFC.IsImporting = False
    Else
        PropagateWPOToOtherEnd oFC   'if the other end is unsupported, then propagate offset to the other end based on the align flag
    End If
    
    Set iOriginalPhysicalPoint = Nothing
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

'*************************************************************************
'Function
'CMReleaseVerticalCornerBraceWPO
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
Public Sub CMReleaseVerticalCornerBraceWPO(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseVerticalCornerBraceWPO"
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

End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String

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
    selection = "VerticalCornerBrace-WP1"
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
    Const METHOD = "ISPSFCInputHelper_GetRelatedObjects"
    On Error GoTo ErrorHandler
    
    Dim IHStatus As SPSFCInputHelperStatus
    Dim obj As Object, obj2 As Object
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem, oMS3 As ISPSMemberSystem
    Dim oPoint As IJPoint
    Dim lngRelated As Long
    Dim oCR As ISPSCanRule
       
    Set oRC = GetRefColl(FC)
    If oRC.IJDEditJDArgument.GetCount <> 3 Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_REQDINPUT, "Vertical corner brace frame connection is missing a required input. Change connection to unsupported and then re-select the supporting members.")
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
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_BRACEMEMBS, "Vertical corner brace frame connection is missing the brace member system. Change connection to unsupported and then re-select the supporting members.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_SUPPORTINGMEMB, "Vertical corner brace frame connection is missing a supporting member system. Change connection to unsupported and then re-select the supporting members.")
    End If
    
    If TypeOf oRC.IJDEditJDArgument.GetEntityByIndex(3) Is ISPSCanRule Then
        Set oCR = oRC.IJDEditJDArgument.GetEntityByIndex(3)
        oCR.GetPrimaryMemberSystem oMS3
        Set oCR = Nothing
    Else
        Set oMS3 = oRC.IJDEditJDArgument.GetEntityByIndex(3)
    End If
    If Not TypeOf oMS3 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_TWO_MEMBS, "Vertical corner brace frame connection is missing one of the two require supporting member systems. Change connection to unsupported and then re-select the supporting member.")
    End If

    ' Verify that the brace's FrameConnection's joint is point on to a joint or splitConn
    FC.Joint.GetPointOn obj, obj2
    If obj Is Nothing Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_EXPECTED_BRACEJOINT, "Vertical corner brace frame connection requires that the brace have a point on relationship to the intersection of the supporting members. Change connection to unsupported and then re-select the intersection.")
    ElseIf Not (TypeOf obj Is IJPoint) Then     ' joint or SplitConn
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTPOINTON_RELATED, "Brace is not point-on related to a split connection or an end-joint of a member. Change connection to unsupported and then re-select the intersection.")
    End If
    
    GetCommonConnection oMS2, oMS3, oPoint, lngRelated
    
    If (lngRelated = 0) Or (oPoint Is Nothing) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTPOINTON_RELATED, "Brace is not point-on related to a split connection or an end-joint of a member. Change connection to unsupported and then re-select the intersection.")
    End If

    If TypeOf oPoint Is ISPSFrameConnection Then
        Dim otmpFC As ISPSFrameConnection
        Set otmpFC = oPoint
        Set oPoint = otmpFC.Joint
    End If
    
    'is the PO obj same one as the common connection btw oMS2 and oMS3 ?
    If Not oPoint Is obj Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTPOINTON_RELATED, "Brace is not point-on related to a split connection or an end-joint of a member. Change connection to unsupported and then re-select the intersection.")
    End If
    
    'lngRelated = 2 should have been taken care of, and swapped, in SetRelatedObjects
    If Not (lngRelated = 1 Or lngRelated = 3) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_INVALID_RELATIONSHIP, "Vertical corner brace frame connection has invalid relationships between member systems. Change connection to unsupported and then re-select the intersection.")
    End If
    
    Set RelatedObject1 = oMS2
    Set RelatedObject2 = oMS3
    
    Set oRC = Nothing
    Set oMS1 = Nothing
    Set oMS2 = Nothing
    Set oMS3 = Nothing
    Set obj = Nothing
    Set obj2 = Nothing
    Set oCR = Nothing
    
    ISPSFCInputHelper_GetRelatedObjects = SPSFCInputHelper_Ok

    Exit Function

ErrorHandler:
    If ObjectAssocFlags(FC, &H200, 0) Then      ' object is not RELATION_INSERTED_IN_TRANSACTION
         HandleError MODULE, METHOD             ' log error
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
    Const METHOD = "ISPSFCInputHelper_SetRelatedObjects"
    On Error GoTo ErrorHandler
    
    Dim IHStatus As SPSFCInputHelperStatus
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    
    Dim oR1 As Object, oR2 As Object
    Dim portIndex As SPSMemberAxisPortIndex
    Dim oMeMS As ISPSMemberSystem
    Dim oSingMS1 As ISPSMemberSystem
    Dim oSingMS2 As ISPSMemberSystem
    Dim oSingFC As ISPSFrameConnection
    Dim oPointConn As IJPoint
    Dim otmpFC As ISPSFrameConnection
    Dim lngRelated As Long
    Dim oCanRulePrimary As ISPSCanRule
    Dim oCanRuleSecondary As ISPSCanRule
    Dim oCanRuleSupported As ISPSCanRule
    Dim oLegMembSys As ISPSMemberSystem

    IHStatus = SPSFCInputHelper_UnexpectedError

    If Not IsFCCleared(FC) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_PREVIOUSCONN_NOTREMOVED, "Vertical corner brace frame connection is invalid because the previous connection was not removed. Change connection to unsupported and the re-select the supporting members.")
    End If
    
    If RelatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_SUPPORTING_MEMBS, "Vertical corner brace frame connection is missing the first supporting member information. Change connection to unsupported and then re-select the supporting members.")
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTRELATED_SUPPORTINGMEMB, "Vertical corner brace frame connection is related to something other than the first supporting member system. Change connection to unsupported and then re-select the supporting member.")
    End If

    If RelatedObject2 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_SUPPRTINGMEMBS, "Vertical corner brace frame connection is missing the second supporting member information. Change connection to unsupported and then re-select the supporting members.")
    End If
    
    If Not TypeOf RelatedObject2 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTRELATED_SECOND_SUPPORTING, "Vertical corner brace frame connection is related to something other than the second supporting member system. Change connection to unsupported and then re-select the supporting members.")
    End If

    Set oMeMS = FC.MemberSystem
    Set oSingMS1 = RelatedObject1
    Set oSingMS2 = RelatedObject2

    If oSingMS1 Is oMeMS Or oSingMS2 Is oMeMS Then
        'The Supported MemberSystem cannot be the Supporting MemberSystem.
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_USING_BRACEMEMB_INCORRECTLY, "Vertical corner brace frame connection is using the brace member incorrectly as a supporting member. Change connection to unsupported and then re-select the supporting members.")
    End If

    If oSingMS1 Is oSingMS2 Then
        IHStatus = SPSFCInputHelper_DuplicateObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_SAMEMEMBS_FORBOTH, "Vertical corner brace frame connection is using the same member system for both supporting members. The two supporting members must be different members. Change connection to unsupported and then re-select the supporting members.")
    End If

'lngRelated     Meaning
' 0             not related
' 1             oMS2 is dependent on oMS1.  oPointConn is a FC at an end of oMS2
' 2             oMS1 is dependent on oMS2.  oPointConn is a FC at an end of oMS1
' 3             oPointConn is a split connection.  dependency not clear !

    GetCommonConnection oSingMS1, oSingMS2, oPointConn, lngRelated
    
    If (oPointConn Is Nothing) Or (lngRelated = 0) Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTRELATED_SECOND_SUPPORTING, "Vertical corner brace frame connection is related to something other than the second supporting member system. Change connection to unsupported and then re-select the supporting members.")
    End If
    
    If lngRelated = 2 Then
        Dim oTmp As Object
        Set oTmp = oSingMS1
        Set oSingMS1 = oSingMS2
        Set oSingMS2 = oTmp
    End If

    portIndex = oMeMS.ResolveEnd(FC.Joint)
    If portIndex <> SPSMemberAxisStart And portIndex <> SPSMemberAxisEnd Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_EXPECTED_BRACEJOINT, "Vertical corner brace frame cannot resolve FC joint with its MemberSystem.")
    End If
    
    'if it's a FC, get the joint.  Otherwise verify that it is a SplitConnection.
    'Brace is then PO to the split connection.

    If TypeOf oPointConn Is ISPSFrameConnection Then
        Set otmpFC = oPointConn
        Set oPointConn = otmpFC.Joint
        Set otmpFC = Nothing
    ElseIf Not TypeOf oPointConn Is ISPSSplitMemberConnection Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_UNKNOWNOBECT, "Unknown object type from port.get_ILC.")
    End If

    RemoveGAPRelations FC ' user may be switching from a GapFC, so clear GAP specific relations

    'set interfaces for supported member
    Set oRC = GetRefColl(FC)
    
    Set oCanRulePrimary = FC.GetCrossSectionObject(SPSFCPrimary)
    Set oCanRuleSecondary = FC.GetCrossSectionObject(SPSFCSecondary)
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
     
    If oLegMembSys Is oSingMS1 Then
        FC.SetCrossSectionObject SPSFCPrimary, oCanRulePrimary
    Else
        'Scoping includes physical axis of supporting member, AxisRotation and XSectionNotify
        oRC.IJDEditJDArgument.SetEntity 2, oSingMS1, ISPSMemberSystemSuppingNotify2, "MembSysSuppingNotify2RC_DEST"
    End If

    Set oLegMembSys = Nothing
    If Not oCanRuleSecondary Is Nothing Then
        oCanRuleSecondary.GetPrimaryMemberSystem oLegMembSys
    End If
    
    If oLegMembSys Is oSingMS2 Then
        FC.SetCrossSectionObject SPSFCSecondary, oCanRuleSecondary
    Else
        'Scoping includes physical axis of supporting member, AxisRotation and XSectionNotify
        oRC.IJDEditJDArgument.SetEntity 3, oSingMS2, ISPSMemberSystemSuppingNotify2, "MembSysSuppingNotify2RC_DEST"
    End If
    
    'create my own joint, unless its ok to keep sharing the joint.
    ConditionallyTransferJoint FC

    ' VerticalCornerBrace is PointOn to the end of the existing secondary member
    FC.Joint.SetPointOn oPointConn, Nothing

    'For the supporting member, get it's supporting members via FC.InputHelper.GetRelatedObjects
    'If any of those are oFC.MemberSystem, then that FC gets set to Unsupported.
    Set oSingFC = oSingMS1.FrameConnectionAtEnd(SPSMemberAxisStart)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If
    Set oSingFC = oSingMS1.FrameConnectionAtEnd(SPSMemberAxisEnd)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If
    Set oSingFC = oSingMS2.FrameConnectionAtEnd(SPSMemberAxisStart)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If
    Set oSingFC = oSingMS2.FrameConnectionAtEnd(SPSMemberAxisEnd)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If

    RemoveCommonSplitConnections oMeMS, oSingMS1
    RemoveCommonSplitConnections oMeMS, oSingMS2

    AddOutputWPO FC, 1

    IHStatus = FC.InputHelper.GetRelatedObjects(FC, oR1, oR2)
    If IHStatus <> SPSFCInputHelper_Ok Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_FAILURETO_ESTB_REL, "Failure to establish correct relations")
    End If

    ISPSFCInputHelper_SetRelatedObjects = SPSFCInputHelper_Ok

    Exit Function

ErrorHandler:
    If ObjectAssocFlags(FC, &H200, 0) Then      ' object is not RELATION_INSERTED_IN_TRANSACTION
         HandleError MODULE, METHOD             ' log error
    End If
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
Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, _
ByVal options As Long, ByVal snapDistance As Double, ByVal LocatedObject1 As Object, _
ByVal LocatedObject2 As Object, ByVal LocateX As Double, ByVal LocateY As Double, _
ByVal LocateZ As Double, RelatedObject1 As Object, RelatedObject2 As Object, _
RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus
' FC is optional and can be Nothing
' what we want to do here is if snap is enabled, then
' based on LocatedObject1's type we can return the x,y,z of nearby points of interest such as
' grid-intersections or pointOn joints, endpoints

    On Error GoTo ErrorHandler
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim oFC As SPSMembers.ISPSFrameConnection
    Dim oMembSys As SPSMembers.ISPSMemberSystem
    Dim O1 As Object, O2 As Object
    
    Dim oPoint As IJPoint
    Dim lngRelated As Long

    IHStatus = SPSFCInputHelper_Ok
    ISPSFCInputHelper_ValidateLocatedObjects = SPSFCInputHelper_UnexpectedError

    If LocatedObject1 Is Nothing And Not LocatedObject2 Is Nothing Then
        Set LocatedObject1 = LocatedObject2
        Set LocatedObject2 = Nothing
    End If

    If TypeOf LocatedObject1 Is ISPSFrameConnection Then
        Set oFC = LocatedObject1
        IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, O1, O2)
        If Not IHStatus = SPSFCInputHelper_Ok Then
            GoTo wrapup
        End If
        
        ' return what the FC is related to and the FC's owning MembSys...
        Set LocatedObject1 = O1                 'primary
        Set LocatedObject2 = oFC.MemberSystem   'secondary
        
        ' moh 04/21/2005
        ' ... unless it is also a VCB.  Then, return both supporting members.
        ' Same as if the user had located end of the beam...
        
        If Not O1 Is Nothing And Not O2 Is Nothing Then
            If TypeOf O1 Is ISPSMemberSystem And TypeOf O2 Is ISPSMemberSystem Then
                Set LocatedObject1 = O1
                Set LocatedObject2 = O2
            End If
        End If
    
    ElseIf TypeOf LocatedObject1 Is ISPSSplitMemberConnection Then
        Dim ii As Long
        Dim eles As IJElements
        Dim oSC As ISPSSplitMemberConnection
        Dim oSCResult As IJStructILCConnectionResult
        Dim iSplitStatus As SplitStatus
            
        Set oSC = LocatedObject1
        Set eles = oSC.InputObjects
        If eles.count = 2 Then
            Set oSCResult = oSC
            iSplitStatus = oSCResult.SplitParentStatusResult
            
            'if splitFirst, eles(1) is split, make the other primary.
            If iSplitStatus = ssSplitSecond Then
                Set LocatedObject1 = eles(1)
                Set LocatedObject2 = eles(2)
            Else
                Set LocatedObject1 = eles(2)
                Set LocatedObject2 = eles(1)
            End If
        End If
    End If

    If LocatedObject1 Is Nothing Or LocatedObject2 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        GoTo wrapup
    End If

    If TypeOf LocatedObject1 Is ISPSMemberPartPrismatic Then
        Dim iMemberPart As ISPSMemberPartPrismatic
        Set iMemberPart = LocatedObject1
        Set LocatedObject1 = iMemberPart.MemberSystem
    End If
    
    If TypeOf LocatedObject2 Is ISPSMemberPartPrismatic Then
        Set iMemberPart = LocatedObject2
        Set LocatedObject2 = iMemberPart.MemberSystem
    End If

    If Not TypeOf LocatedObject1 Is ISPSMemberSystemLinear Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        GoTo wrapup
    End If
    If Not TypeOf LocatedObject2 Is ISPSMemberSystemLinear Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        GoTo wrapup
    End If
    If Not FC Is Nothing Then
        If Not TypeOf FC.MemberSystem Is ISPSMemberSystemLinear Then
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject
            GoTo wrapup
        End If
    End If

    If IsMemberAxesColinear(LocatedObject1, LocatedObject2) Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject

    Else
        If Not FC Is Nothing Then
            If Not IsTwoMemberSystemsOkforVCB(FC, LocatedObject1, LocatedObject2) Then
                IHStatus = SPSFCInputHelper_InvalidTypeOfObject
                GoTo wrapup
            End If
        End If
        GetCommonConnection LocatedObject1, LocatedObject2, oPoint, lngRelated

        If oPoint Is Nothing Then   ' the two membersystems are not related
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject

        Else
            ' return the coordinates of the connection to the command
            ' so it can set the new member to that location.
            oPoint.GetPoint RelatedObjectX, RelatedObjectY, RelatedObjectZ
        
            If lngRelated = 1 Then         'MS2 is dependent on MS1
                Set RelatedObject1 = LocatedObject1
                Set RelatedObject2 = LocatedObject2
            
            ElseIf lngRelated = 2 Then      'MS1 is dependent on MS2
                Set RelatedObject1 = LocatedObject2
                Set RelatedObject2 = LocatedObject1
            
            ElseIf lngRelated = 3 Then      'its a split conn.  Which one is primary ?
                Set RelatedObject1 = LocatedObject1
                Set RelatedObject2 = LocatedObject2
            
            Else
                IHStatus = SPSFCInputHelper_InconsistentRelations   ' unknown problem
            
            End If
        End If
    End If
wrapup:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
    Exit Function

ErrorHandler:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
End Function

Public Function FindSplitConnection(oFC As IJPoint, oMS As ISPSMemberSystem) As IJPoint

    On Error GoTo ErrorHandler
    Const METHOD = "FindSplitConnection"

    Dim ii As Long, count As Long
    Dim oSplitConns As IJElements
    Dim oSCPoint As IJPoint
    
    Set FindSplitConnection = Nothing

    Set oSplitConns = oMS.SplitConnections
    count = oSplitConns.count
    For ii = 1 To count
        Set oSCPoint = oSplitConns(ii)
        If oSCPoint.IsEqual(oFC) Then
            Set FindSplitConnection = oSCPoint
            GoTo wrapup
        End If
    Next ii

wrapup:
    Set oSplitConns = Nothing
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
    Exit Function
End Function
Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

Private Function IsRelatedSplitConnectionInBadState(oFC As ISPSFrameConnection) As Boolean
    Const METHOD = "IsRelatedSplitConnectionInBadState"
    On Error GoTo ErrorHandler
        
    Dim lngRelated As Long
    Dim oSPSPersistFlagsAccess As ISPSPersistFlagsAccess
    Dim obj1 As Object, obj2 As Object
    
    IsRelatedSplitConnectionInBadState = False
    oFC.Joint.GetPointOn obj1, obj2
    
    If Not obj1 Is Nothing Then
        If TypeOf obj1 Is ISPSSplitMemberConnection Then
            Set oSPSPersistFlagsAccess = obj1
            If Not oSPSPersistFlagsAccess Is Nothing Then
                If (oSPSPersistFlagsAccess.PersistFlags And 1) Then
                    IsRelatedSplitConnectionInBadState = True
                End If
            End If
        Else
            ' Its not point on to a split connection
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    IsRelatedSplitConnectionInBadState = False
End Function

Private Sub ISPSMirrorHelper_Mirror(ByVal pPlane As SP3DStructInterfaces.IJPlane, ByVal pMatrix As SP3DStructInterfaces.iJDT4x4, ByVal bCopy As Boolean, ByVal oOriginalFC As Object, ByVal oFC As Object)
    Const METHOD = "ISPSMirrorHelper_Mirror"
    On Error GoTo ErrorHandler
     
    Dim oPlaneNormal As New DVector, oVec As New DVector
    Dim oFrmConn As ISPSFrameConnection
    Dim oVecX As New DVector, oVecY As New DVector, oVecZ As New DVector
   
    Dim oIJAttributes As IJDAttributes
    Dim oCollProxy As CollectionProxy
    
    Dim x As Double, y As Double, z As Double
    Dim gX As Double, gY As Double, gZ As Double, lX As Double, lY As Double, lZ As Double
    Dim nx As Double, ny As Double, nz As Double, uX As Double, uY As Double, uZ As Double
    Dim vX As Double, vY As Double, vZ As Double
    Dim oMat As New DT4x4, oLocMat As DT4x4
    Dim lCoordSys As Long

    'NOTE - It is assumed that primary and secondary are also being mirrored along with the supported member
    '(this is the most common case). Selecting a diffrent primary or secondary or both through the paste
    'dialog may not give the correct behavior as is not handled here. User shouldn't be using mirror for that case and instead use
    'copy paste.
    
    Set oIJAttributes = oFC
    
    'get offset vector of the old member
    Set oCollProxy = oIJAttributes.CollectionOfAttributes("IJUASPSFCManualOffset")
    lCoordSys = oCollProxy.Item("CoordinateSystem").Value
    
    'Need to mirror the offset vector now
         
    x = oCollProxy.Item("XOffset").Value
    y = oCollProxy.Item("YOffset").Value
    z = oCollProxy.Item("ZOffset").Value
    oVec.Set x, y, z
        
    
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
    If lCoordSys = 2 Then ' local
        'get original frame connection
        Set oFrmConn = oOriginalFC
        getLocalToGlobal oFrmConn, oLocMat
        Set oVec = oLocMat.TransformVector(oVec) ' convert to global
        
        'also update the local matrix for the frame connection based on the mirror plane
        ' we could have avoided this step by calculating the matrix based on the primary and secondary
        ' members of the clone frame connection. The problem with this is that there is no guarantee
        'that the secondary and primary are mirrored before this method on the FC is called
        
       
        oVecX.Set oLocMat.IndexValue(0), oLocMat.IndexValue(1), oLocMat.IndexValue(2)
        oVecY.Set oLocMat.IndexValue(4), oLocMat.IndexValue(5), oLocMat.IndexValue(6)
        'convert to mirror plane local
        Set oVecX = oMat.TransformVector(oVecX)
        Set oVecY = oMat.TransformVector(oVecY)
        
        oVecX.z = -oVecX.z
        oVecY.z = -oVecY.z
    
    End If
    
    Set oVec = oMat.TransformVector(oVec)
     
    'to mirror, reverse z
    oVec.z = -oVec.z
     
    oMat.Invert 'transform back to global
    Set oVec = oMat.TransformVector(oVec)
    If lCoordSys = 2 Then ' local
   
        Set oVecX = oMat.TransformVector(oVecX)
        oVecX.Length = 1 ' normalize
        Set oVecY = oMat.TransformVector(oVecY)
        oVecY.Length = 1 ' normalize
        Set oVecZ = oVecX.Cross(oVecY)
        oVecZ.Length = 1 ' normalize
        
        
        oLocMat.IndexValue(0) = oVecX.x
        oLocMat.IndexValue(1) = oVecX.y
        oLocMat.IndexValue(2) = oVecX.z
        
        oLocMat.IndexValue(4) = oVecY.x
        oLocMat.IndexValue(5) = oVecY.y
        oLocMat.IndexValue(6) = oVecY.z
        
        oLocMat.IndexValue(8) = oVecZ.x
        oLocMat.IndexValue(9) = oVecZ.y
        oLocMat.IndexValue(10) = oVecZ.z
        
        
        'transfor to FC's local coordinates
        oLocMat.Invert
        Set oVec = oLocMat.TransformVector(oVec)
    End If

    oCollProxy.Item("XOffset").Value = oVec.x
    oCollProxy.Item("YOffset").Value = oVec.y
    oCollProxy.Item("ZOffset").Value = oVec.z
     
       
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

Private Sub getLocalToGlobal(oFC As ISPSFrameConnection, ByRef oMatrix As iJDT4x4)
                            
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim xlocal As Double, ylocal As Double, zlocal As Double
    Dim oSupportedMembSys As ISPSMemberSystem
    Dim oSupportingPriMembSys As ISPSMemberSystem
    Dim oSupportingSecMembSys As ISPSMemberSystem
    Dim dotProd As Double
    Dim IHStatus As SPSFCInputHelperStatus
        
    Dim oVec As New DVector, oVec1 As New DVector, oVec2 As New DVector
    Dim oNvec As IJDVector
    
    Set oSupportedMembSys = oFC.MemberSystem
    IHStatus = oFC.InputHelper.GetRelatedObjects(oFC, oSupportingPriMembSys, oSupportingSecMembSys)
    If IHStatus <> SPSFCInputHelper_Ok Then
        Err.Raise E_FAIL
    End If
    
    oSupportedMembSys.PhysicalAxis.EndPoints X0, Y0, Z0, x1, y1, z1
    
    If oFC.WPO.portIndex = SPSMemberAxisStart Then
        oVec.Set x1 - X0, y1 - Y0, z1 - Z0
    Else
        oVec.Set X0 - x1, Y0 - y1, Z0 - z1
    End If
    
    oVec.Length = 1
    oSupportingSecMembSys.PhysicalAxis.EndPoints X0, Y0, Z0, x1, y1, z1
    oVec1.Set x1 - X0, y1 - Y0, z1 - Z0
    oVec1.Length = 1
    

    'We need to make sure that oVec1 points from the connected end to the other end of secondary
    dotProd = oVec.Dot(oVec1)
    If dotProd < 0 Then
        oVec1.Set X0 - x1, Y0 - y1, Z0 - z1
        oVec1.Length = 1
    End If
    oSupportingPriMembSys.PhysicalAxis.EndPoints X0, Y0, Z0, x1, y1, z1
    oVec2.Set x1 - X0, y1 - Y0, z1 - Z0
    oVec2.Length = 1
    'We need to make sure that oVec2 points from the connected end to the other end of primary
    dotProd = oVec.Dot(oVec2)
    If dotProd < 0 Then
        oVec2.Set X0 - x1, Y0 - y1, Z0 - z1
        oVec2.Length = 1
    End If
    
    Set oNvec = oVec1.Cross(oVec2)
    oNvec.Length = 1
    
    If oMatrix Is Nothing Then
        Set oMatrix = New DT4x4
    End If
    
    oMatrix.LoadIdentity
    
    ' load matrix with the vectors

    oMatrix.IndexValue(0) = oVec1.x
    oMatrix.IndexValue(1) = oVec1.y
    oMatrix.IndexValue(2) = oVec1.z
    
    oMatrix.IndexValue(4) = oVec2.x
    oMatrix.IndexValue(5) = oVec2.y
    oMatrix.IndexValue(6) = oVec2.z
    
    oMatrix.IndexValue(8) = oNvec.x
    oMatrix.IndexValue(9) = oNvec.y
    oMatrix.IndexValue(10) = oNvec.z
    
End Sub


