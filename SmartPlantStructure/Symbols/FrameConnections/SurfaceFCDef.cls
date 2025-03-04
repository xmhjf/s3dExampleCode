VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SurfaceFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : SurfaceFCDEF.cls
'
'Author : Structure Development
'
'Description :
'    SmartPlant Structural surface Frame Connection Definition file
'
'History:
'
' 05/21/03   Structure Development      Changed error handling in GetRelatedObjects
'                                       and SetRelatedObjects to raise the error;
'                                       also replaced the Msgbox with HandleError
' 08/07/03   Structure Development      Implemented offsets.  "Local" offset is ReadOnly.
'
' 08/14/03   Structure Development      Substituted interface name in place of guid
'                                       in IJDPropertyDescriptions->AddProperty. The guid of
'                                       a virtual interface may change during bulkload
' 04/15/05   Structure Development      GetRelatedObjects:  object 2 message, tolerate port
'
' 04/18/06   Structure Development      Calls methods to propagate WPO to the other end and to clear
'                                       the offset on the othen end
' 06/06/06   Structure Development      Added CMMigrate method stubs
' Jun062006  Structure Development      Fixed T99966:Added migration code for RC
' 09/27/06   RP                         SurfaceFC now works for curved members. The validate
'                                       method was returning notOK previously (TR#106105)
' 09/27/06   AS                         TR#105470 Remove Blank stubs for CMSetInput and CMRemoveInput
' 09/26/06   Struct Dev    TR 104591    ReadPathPointOnObjectsForFC
' 10/03/06   Struct Dev                 TR 103405  OnPreLoad AttributeDescriptor_ReadOnly for member with path
'
' 07/1/09    RP                         CR165819 - Added code in SetRelatedObjects() to remove
'                                       relations created for the GAP connection (the user may switch
'                                       from GAP FC to surface FC )
'*******************************************************************************************
Option Explicit
Private Const MODULE = "SurfaceFCDef"
Private m_oLocalizer As IJLocalizer

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.SurfaceFCDef"
Const m_ItemName As String = "SPSFCMacros.SurfaceFCDef"

Dim bOnPreLoad As Boolean


' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

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
  Set pIH = New InputHelper
  pIH.definition = pDefinition
  DefinitionInputs pIH

  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{5FEB4ADB-E5EC-45D6-8878-150ADDC04D0A}"     'CSPSFrameConnection
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  ' Add following interfaces as inputs so as to get notified when they are modified.
  'These interfaces carry properties that drive rotation and offset
  pAPDs.AddProperty "SurfaceProps", 1, "IJUASPSFCManualOffset"
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 

  '*************************WPO*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("SurfaceWPO", 1, "CMConstructSurfaceWPO", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSurfaceWPO"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSurfaceWPO"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseSurfaceWPO"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSurfaceWPO" 'added Migrate for split notification
  
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisWPO {34DF6BED-41D5-4D19-B090-D58D93A1CF64}
  pPropertyDescriptions.AddProperty "ComputeSurfaceWPO", 1, "{34DF6BED-41D5-4D19-B090-D58D93A1CF64}", "CMUpdateSurfaceWPO", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  '********************************************************************************************************
  
  Set pMemberDescriptions = Nothing
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConditionalSurfaceWPO
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

Public Sub CMConditionalSurfaceWPO(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSurfaceWPO"
  On Error GoTo ErrorHandler
    bIsNeeded = True ' Always created
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConstructSurfaceWPO
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

Public Sub CMConstructSurfaceWPO(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSurfaceWPO"
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
'CMSetInputSurfaceWPO
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
Public Sub CMSetInputSurfaceWPO(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSurfaceWPO"
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMUpdateSurfaceWPO
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
Public Sub CMUpdateSurfaceWPO(pPropertyDescription As IJDPropertyDescription, pObject As Object)
 Const MT = "CMUpdateSurfaceWPO"
    On Error GoTo ErrorHandler
    
    Dim oFrameConn As ISPSFrameConnection
    Dim oWPO As ISPSAxisWPO
    Dim xOffset As Double, yOffset As Double, zOffset As Double
    Dim oIJAttributes As IJDAttributes
    Dim oCollProxy As CollectionProxy
    
    Dim oObj1 As Object, oObj2 As Object
    Dim fcStatus As SPSFCInputHelperStatus
    
    Set oFrameConn = pPropertyDescription.CAO   'just a sanity check.
    Set oWPO = pObject

    fcStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oObj1, oObj2)
    
    ' if error occured at reading supporting objects, it may be fixable so call FixFC to see.
    ' if FixFC fails to fix it, it returns false, the FC has been set to Unsupported and we must exit.
    ' if FixFC returns okay, try again to get_RelatedObjects.  If it fails now, we set to unsupported here and exit.
    ' Otherwise, we have been healed and are okay to continue
    
    If fcStatus <> SPSFCInputHelper_Ok Then
        If FixFC(oFrameConn) Then
            fcStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oObj1, oObj2)
            If fcStatus <> SPSFCInputHelper_Ok Then
                Set oFrameConn.definition = Nothing
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    Set oIJAttributes = oFrameConn
    Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCManualOffset)

    ' if IsImporting, then set offset attributes based on current physical location of member axis
    If oFrameConn.IsImporting Then
        Dim iCurve As IJCurve
        Dim startX As Double, startY As Double, startZ As Double
        Dim endX As Double, endY As Double, endZ As Double
        Dim logicalX As Double, logicalY As Double, logicalZ As Double
        
        Set iCurve = oFrameConn.MemberSystem
        iCurve.EndPoints startX, startY, startZ, endX, endY, endZ
        
        If oFrameConn.WPO.portIndex = SPSMemberAxisStart Then
            oFrameConn.MemberSystem.LogicalAxis.GetLogicalStartPoint logicalX, logicalY, logicalZ
            xOffset = startX - logicalX
            yOffset = startY - logicalY
            zOffset = startZ - logicalZ
        Else
            oFrameConn.MemberSystem.LogicalAxis.GetLogicalEndPoint logicalX, logicalY, logicalZ
            xOffset = endX - logicalX
            yOffset = endY - logicalY
            zOffset = endZ - logicalZ
        End If
        oCollProxy.Item(UA_FCManualOffset_XOffset).Value = xOffset
        oCollProxy.Item(UA_FCManualOffset_YOffset).Value = yOffset
        oCollProxy.Item(UA_FCManualOffset_ZOffset).Value = zOffset
        oFrameConn.IsImporting = False
    Else
        xOffset = oCollProxy.Item(UA_FCManualOffset_XOffset).Value
        yOffset = oCollProxy.Item(UA_FCManualOffset_YOffset).Value
        zOffset = oCollProxy.Item(UA_FCManualOffset_ZOffset).Value
    End If
   
    oWPO.SetWPO xOffset, yOffset, zOffset, 0#, 0#, 0#
    oWPO.WPOCardinalPoint = 0
    
    If oFrameConn.IsImporting Then
        SetWPOAtOtherEndIfUnsupported oFrameConn    ' if other end is unsupported, set its offsets to keep physical where it is.
        oFrameConn.IsImporting = False
    Else
        PropagateWPOToOtherEnd oFrameConn   'if the other end is unsupported, then propagate offset to the other end based on the align flag
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL    ' Force the FC to the ToDo List
End Sub

'*************************************************************************
'Function
'CMReleaseSurfaceWPO
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
Public Sub CMReleaseSurfaceWPO(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseSurfaceWPO"
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
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FCMACROS_EROOR, "ERROR")
    
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
            IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FCMACROS_ERROR_BADVAL, "ERROR::Bad Value")
            Exit Function
        End If
    End If
    
    If (pAttrToChange.attrName <> "CoordinateSystem") Then
        pAttrToChange.AttrValue = varNewAttrValue
    End If
    
    IJUserAttributeMgmt_OnAttributeChange = ""
   
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
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

    IJUserAttributeMgmt_OnPreCommit = ""

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
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor

    Dim oFC As ISPSFrameConnection

    Set pColl = CollAllDisplayedValues
    count = pColl.count
    
    ' if supported member is driven by a path, no offsets are permitted.
    Set oFC = pIJDAttrs
    If Not oFC.MemberSystem.LogicalAxis.CurveObject Is Nothing Then
        For i = 1 To count
            Set pAttrDescr = pColl.Item(i)
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        Next
    
    Else

        For i = 1 To count
            Set pAttrDescr = pColl.Item(i)
            If (pAttrDescr.attrName = "CoordinateSystem") Then
                pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnlyByRule
            End If
        Next
    End If

    IJUserAttributeMgmt_OnPreLoad = ""
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String

    UserAttributeMgmt_Validate = ""

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

    selection = "Surface-Default"
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
    Dim obj As Object
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem

    IHStatus = SPSFCInputHelper_UnexpectedError

    Set oRC = GetRefColl(FC)
    If oRC.IJDEditJDArgument.GetCount <> 2 Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SURFACE_MISSSING_INPUT, "Surface frame connection is missing a required input. Change connection to unsupported and then re-select the supporting member.")
    End If

    Set oMS1 = oRC.IJDEditJDArgument.GetEntityByIndex(1)
    Set oMS2 = FC.MemberSystem
    If Not oMS1 Is oMS2 Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SURFACE_MISSING_MEMBS, "Surface frame connection is missing the member system. Change connection to unsupported and then re-select the supporting member.")
    End If

    Set oMS2 = Nothing
    
    Set obj = oRC.IJDEditJDArgument.GetEntityByIndex(2)

    If Not ObjectIsOkSurface(obj) Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SURFACE_MISSING_SURFACE, "Surface frame connection is missing the surface. Change connection to unsupported and then re-select the supporting member.")
    End If
            
    Dim oRel1 As Object, oRel2 As Object    ' objects related to possible sketch path
    If ReadPathPointOnObjectsForFC(FC, oRel1, oRel2) Then
        If Not oRel1 Is obj Then
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject
            Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SURFACE_MISSING_SURFACE, "Surface frame connection is missing the surface. Change connection to unsupported and then re-select the supporting member.")
        End If
    End If

    Set RelatedObject1 = obj
    Set RelatedObject2 = Nothing
    
    Set oRC = Nothing
    Set oMS1 = Nothing
    Set oMS2 = Nothing
    
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

    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim oRefColl As IMSSymbolEntities.IJDReferencesCollection
    Dim oR1 As Object, oR2 As Object

    'If only one is sent in, make it RelatedObject1
    If RelatedObject1 Is Nothing And Not RelatedObject2 Is Nothing Then
        Set RelatedObject1 = RelatedObject2
        Set RelatedObject2 = Nothing
    End If

    If Not ObjectIsOkSurface(RelatedObject1) Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SURFACE_NOTRELATED_VALIDSURFACE, "Surface frame connection is related to something other than a valid surface. Change connection to unsupported and then re-select the supporting member.")
    End If
    
    RemoveGAPRelations FC ' user may be switching from a GapFC, so clear GAP specific relations
    
    Set oRefColl = GetRefColl(FC)
    oRefColl.IJDEditJDArgument.RemoveAll

    If FC.WPO.portIndex = SPSMemberAxisStart Then
        oRefColl.IJDEditJDArgument.SetEntity 1, FC.MemberSystem, ISPSMemberSystemStartEndNotify, "MemberSysStartNotifyRC_DEST"
    Else
        oRefColl.IJDEditJDArgument.SetEntity 1, FC.MemberSystem, ISPSMemberSystemEndEndNotify, "MemberSysEndNotifyRC_DEST"
    End If
    oRefColl.IJDEditJDArgument.SetEntity 2, RelatedObject1, ConstIJSurface, "SurfaceToRC_DEST"

    FC.Joint.SetPointOn RelatedObject1, Nothing
    SetPathPointOnObjectsForFC FC, RelatedObject1, Nothing

    RemoveCommonSplitConnections FC.MemberSystem, RelatedObject1
    
    AddOutputWPO FC, 1

    IHStatus = FC.InputHelper.GetRelatedObjects(FC, oR1, oR2)
    If IHStatus <> SPSFCInputHelper_Ok Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SURFACE_REVERIFIED_INVALID, "Surface frame connection was reverified and found to be invalid. Change connection to unsupported and then re-select the supporting member.")
    End If

    ISPSFCInputHelper_SetRelatedObjects = SPSFCInputHelper_Ok
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    ISPSFCInputHelper_SetRelatedObjects = IHStatus
    Exit Function

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
'Criteria for validated object:
'1) must of type surface ( or port->Geometry is surface )

Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal options As Long, ByVal snapDistance As Double, ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, RelatedObject1 As Object, RelatedObject2 As Object, RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus
    Const METHOD = "ISPSFCInputHelper_ValidateLocatedObjects"
    On Error GoTo ErrorHandler
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus

    IHStatus = SPSFCInputHelper_UnexpectedError
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus

    If LocatedObject1 Is Nothing And Not LocatedObject2 Is Nothing Then
        Set LocatedObject1 = LocatedObject2
        Set LocatedObject2 = Nothing
    End If
        
    If ObjectIsOkSurface(LocatedObject1) Then
        IHStatus = SPSFCInputHelper_Ok
    Else
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
    End If
    

    Set RelatedObject1 = GetStablePort(LocatedObject1)
    Set RelatedObject2 = GetStablePort(LocatedObject2)
    
    RelatedObjectX = LocateX
    RelatedObjectY = LocateY
    RelatedObjectZ = LocateZ

    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
End Function

'*************************************************************************
'Function
'CMMigrateAggregator
'
'Abstract
'Migrates thr FC to the correct surface if it is split.
'
'Arguments
'IJDMemberDescription interface of the member
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CMMigrateAggregator(oAggregatorDesc As IJDAggregatorDescription, oMigrateHelper As IJMigrateHelper)

  Const MT = "CMMigrateAggregator"
  On Error GoTo ErrorHandler
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    
    Dim oObjectsReplacing() As Object
    Dim bIsInputMigrated As Boolean
    Dim oPoint As IJPoint
    Dim ii As Integer
    
    Set oSmartOccurrence = oAggregatorDesc.CAO
    Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOccurrence)
    
    On Error Resume Next
    Set oPoint = oSmartOccurrence
    Err.Clear
    On Error GoTo ErrorHandler
    'We can work with this point or let the default point which is 0,0,0 to be used.
    'Lets not error out here
    If Not oPoint Is Nothing Then
        MigrateRefColl oRefColl, oMigrateHelper, oObjectsReplacing, bIsInputMigrated, oPoint
    Else
        MigrateRefColl oRefColl, oMigrateHelper, oObjectsReplacing, bIsInputMigrated
    End If
    
    If bIsInputMigrated Then
        'If any of the inputs are indeed migrated, reset them on the ref coll
        oRefColl.IJDEditJDArgument.SetEntity 2, oObjectsReplacing(2), ConstIJSurface, "SurfaceToRC_DEST"
    End If
    
    For ii = LBound(oObjectsReplacing) To UBound(oObjectsReplacing)
        Set oObjectsReplacing(ii) = Nothing
    Next ii
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMMigrateSurfaceWPO
'
'Abstract
'Migrates thr FC to the correct surface if it is split.
'
'Arguments
'IJDMemberDescription interface of the member
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CMMigrateSurfaceWPO(ByVal pMemberDescription As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const MT = "CMMigrateSurfaceWPO"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
