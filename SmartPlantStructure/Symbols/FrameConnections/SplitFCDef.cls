VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : SplitFCDEF.cls
'
'Author : Structure Development
'
'Description :
'    SmartPlant Structural split Frame Connection Definition file
'
'History:
'
' 05/21/03   Structure Development     Changed error handling in GetRelatedObjects
'                                      and SetRelatedObjects to raise the error;
'                                      also replaced the Msgbox with HandleError
'
' 08/14/03   Structure Development     Substituted interface name in place of guid
'                                      in IJDPropertyDescriptions->AddProperty. The guid of
'                                      a virtual interface may change during bulkload
' 04/18/06   Structure Development     Calls methods to propagate WPO to the other end and to clear
'                                      the offset on the othen end
' 09/27/06   AS                        TR#105470 Remove Blank stubs for CMSetInput and CMRemoveInput
'**************************************************************************************
Option Explicit
Private Const MODULE = "SplitFCDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.SplitFCDef"
Const m_ItemName As String = "SPSFCMacros.SplitFCDef"

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements SPSMembers.ISPSFCInputHelper
Implements SPSMembers.ISPSFCInputHelperEx
Implements SP3DStructInterfaces.IJUserAttributeMgmt


'*************************************************************************
'Function
'
'<DefinitionInputs>
'
'Abstract
'
'<define the input on the inputhelpers object>
'
'Arguments
'
'<pIH as IJDInputsHelper>
'
'Return
'
'Exceptions
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
Public Sub CMUpdateAxisWPO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMUpdateAxisWPO"
    On Error GoTo ErrorHandler
        Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
        Dim obj As Object
        Dim oWPO As ISPSAxisWPO
'        Dim oSymbol                As IMSSymbolEntities.IJDSymbol
 '       Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
  '      Dim oRefColl1   As IMSSymbolEntities.IJDReferencesCollection, oRefColl2  As IMSSymbolEntities.IJDReferencesCollection
        Dim intPort As SPSMemberAxisPortIndex
        Dim oRotation  As ISPSAxisRotation
        Dim oMat1 As iJDT4x4, oMat2 As iJDT4x4
        Dim oMemberPart As ISPSMemberPartPrismatic
        Dim oSys As ISPSMemberSystem
        Dim suppingCP As Long, offsetCP As Long
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        Dim dx As Double, dy As Double
        Dim xOffset As Double, yOffset As Double, zOffset As Double
        Dim lCoordSys As Long
        Dim pIJAttrbs As IJDAttributes
        Dim dBearingPlateThk As Double, bWithBearingPlate As Boolean
        Dim oPos As New DPosition
        Dim oLocPos As New DPosition
        Dim FC As ISPSFrameConnection
        
        Set FC = pPropertyDescriptions.CAO
        IHStatus = FC.InputHelper.GetRelatedObjects(FC, oSys, obj)
    
        If IHStatus <> SPSFCInputHelper_Ok Then
            Exit Sub
        End If
    
        Set oRotation = oSys.Rotation
        Set oMemberPart = oSys.DesignPartAtEnd(FC.WPO.portIndex)
        suppingCP = oMemberPart.CrossSection.CardinalPoint ' the CP with which supporting is placed

        Set pIJAttrbs = pPropertyDescriptions.CAO
 ' the CP of the supporting where the supported needs be positioned
        offsetCP = pIJAttrbs.CollectionOfAttributes("IJUASPSFCAxis").Item("SupportingCP").Value
        
        xOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("XOffset").Value
        yOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("YOffset").Value
        zOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("ZOffset").Value
        lCoordSys = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("CoordinateSystem").Value
        
        If offsetCP = 0 Then 'if 0 then use the CP with which the supporting is placed
            offsetCP = suppingCP
        End If
        
        oMemberPart.CrossSection.GetCardinalPointOffset suppingCP, x1, y1
        oMemberPart.CrossSection.GetCardinalPointOffset offsetCP, x2, y2
        
        dx = x2 - x1 ' the x offset in cross section coordinates
        dy = y2 - y1 ' the y offset in cross section coordinates
        
 'calculate the matrix to apply
        oRotation.GetTransform oMat1
        oMat1.IndexValue(12) = 0
        oMat1.IndexValue(13) = 0
        oMat1.IndexValue(14) = 0
        Set oMat2 = CreateCSToMembTransform(oMat1)
                
                
        Set oWPO = pObject
        
        oPos.Set dx, dy, 0#
        Set oPos = oMat2.TransformPosition(oPos)
        
        If lCoordSys = 1 Then 'Global
            
            oPos.x = oPos.x + xOffset
            oPos.y = oPos.y + yOffset
            oPos.z = oPos.z + zOffset
            
        ElseIf lCoordSys = 2 Then 'Local
            oLocPos.Set xOffset, yOffset, zOffset
            Set oLocPos = oMat1.TransformPosition(oLocPos)
                     
            oPos.x = oPos.x + oLocPos.x
            oPos.y = oPos.y + oLocPos.y
            oPos.z = oPos.z + oLocPos.z
            
            
        End If
        
        oWPO.SetWPO oPos.x, oPos.y, oPos.z, 0, 0, 0
        
        'if the other end is unsupported, then propagate offset to the other end
        'based on the align flag
        PropagateWPOToOtherEnd FC
        
        Set oPos = Nothing
        Set oSys = Nothing
        Set oRotation = Nothing
        Set oMat1 = Nothing
        Set oMat2 = Nothing
        Set oMemberPart = Nothing
        Set pIJAttrbs = Nothing
    Exit Sub
ErrorHandler:  HandleError MODULE, MT
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

    selection = "Split-1"
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
    
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    IHStatus = SPSFCInputHelper_UnexpectedError

    Dim obj As Object
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem, oMS3 As ISPSMemberSystem
    
    Set oRC = GetRefColl(FC)
    If oRC.IJDEditJDArgument.GetCount <> 2 Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", "Unexpected refcoll count=" & oRC.IJDEditJDArgument.GetCount
    End If

    Set oMS1 = oRC.IJDEditJDArgument.GetEntityByIndex(1)
    Set oMS2 = FC.MemberSystem
    If Not oMS1 Is oMS2 Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", "Unknown object 1 in Reference coll"
    End If

    Set oMS2 = Nothing
    Set oMS2 = oRC.IJDEditJDArgument.GetEntityByIndex(2)
    If Not TypeOf oMS2 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", "Unknown object 2 in Reference coll"
    End If

    FC.Joint.GetPointOn oMS3, obj
    
    If Not oMS2 Is oMS3 Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", "PointOn MemberSystem is not same as RefColl MemberSystem"
    End If

    'Return the PointOn object, and possibly the intersecting plane
    Set RelatedObject1 = oMS2
    If obj Is Nothing Then
        Set RelatedObject2 = Nothing
    Else
        Set RelatedObject2 = obj
    End If
    
    Set oRC = Nothing
    Set oMS1 = Nothing
    Set oMS2 = Nothing
    Set oMS3 = Nothing
    Set obj = Nothing
    
    ISPSFCInputHelper_GetRelatedObjects = SPSFCInputHelper_Ok

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
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
    IHStatus = SPSFCInputHelper_UnexpectedError
    
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    
    Dim oMS As ISPSMemberSystem
    Dim oSysCrSectionNotify As ISPSMemberSystemXSectionNotify
    Dim oRotation As ISPSAxisRotation

    If Not IsFCCleared(FC) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", "Previous Relations Not Cleared"
    End If

    If RelatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", "No RelatedObject1"
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", "Unknown type of RelatedObject1"
    End If
    
    'Scoping includes logical axis of supported member, AxisRotationInput and XSectionNotify
    Set oRC = GetRefColl(FC)
    oRC.IJDEditJDArgument.SetEntity 1, FC.MemberSystem, ISPSMemberSystemStartEndNotify, "MemberSysStartNotifyRC_DEST"

    Set oMS = RelatedObject1
    'Scoping includes physical axis of supporting member, AxisRotation and XSectionNotify
    oRC.IJDEditJDArgument.SetEntity 2, oMS, ISPSMemberSystemSuppingNotify2, "MembSysSuppingNotify2RC_DEST"

    ' axis connection does not share joints and is along the supporting member system
    FC.Joint.SetPointOn RelatedObject1, RelatedObject2
    
    oMS.CreateSplit FC.Joint

    Dim dbgObj1 As Object, dbgObj2 As Object
    IHStatus = ISPSFCInputHelper_GetRelatedObjects(FC, dbgObj1, dbgObj2)
    If IHStatus = SPSFCInputHelper_Ok Then
        If Not dbgObj1 Is RelatedObject1 Then
            IHStatus = SPSFCInputHelper_InconsistentRelations
            Err.Raise IHStatus + SPSvbError, "", "Failure to establish correct relations"
        End If
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
Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal options As Long, ByVal snapDistance As Double, ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, RelatedObject1 As Object, RelatedObject2 As Object, RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus
' FC is optional and can be Nothing
' what we want to do here is if snap is enabled, then
' based on LocatedObject1's type we can return the x,y,z of nearby points of interest such as
' grid-intersections or pointOn joints, endpoints
    
    On Error GoTo ErrorHandler
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    IHStatus = SPSFCInputHelper_UnexpectedError
    ISPSFCInputHelper_ValidateLocatedObjects = SPSFCInputHelper_UnexpectedError
    
    If FC Is Nothing Then
    End If

    If LocatedObject1 Is Nothing And Not LocatedObject2 Is Nothing Then
        Set LocatedObject1 = LocatedObject2
        Set LocatedObject2 = Nothing
    End If
        
    If LocatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        GoTo wrapup
    ElseIf Not TypeOf LocatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        GoTo wrapup
    End If
    
    If Not FC Is Nothing Then
        If LocatedObject1 Is FC.MemberSystem Then
            IHStatus = SPSFCInputHelper_DuplicateObject
            GoTo wrapup
        End If
    End If

    If LocatedObject1 Is LocatedObject2 Then
        IHStatus = SPSFCInputHelper_DuplicateObject
        GoTo wrapup
    End If

'    If Not LocatedObject2 Is Nothing Then
'        If Not TypeOf LocatedObject2 Is IJSurface Then
'            IHstatus = SPSFCInputHelper_InvalidTypeOfObject
'            GoTo wrapup
'        End If
'    End If

    If Not LocatedObject1 Is Nothing Then
        Set RelatedObject1 = LocatedObject1
    Else
        Set RelatedObject1 = Nothing
    End If
    
    If Not LocatedObject2 Is Nothing Then
        Set RelatedObject2 = LocatedObject2
    Else
        Set RelatedObject2 = Nothing
    End If
    
    RelatedObjectX = LocateX
    RelatedObjectY = LocateY
    RelatedObjectZ = LocateZ

    IHStatus = SPSFCInputHelper_Ok

wrapup:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
    Exit Function

ErrorHandler:
    ISPSFCInputHelper_ValidateLocatedObjects = IHStatus
End Function


