VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeatedFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2003 Intergraph Corporation. All rights reserved.
'
'File : SeatedFCDEF.cls
'
'Author : Structure Development
'
'Description :
'    SmartPlant Structural Seated Frame Connection Definition file
'
'History:
'
' 03/31/03   Structure Development     Added the code to update the rotation and
'                                      WPO offsets
' 05/21/03   Structure Development     Changed error handling in GetRelatedObjects
'                                      and SetRelatedObjects to raise the error;
'                                      also replaced the Msgbox with HandleError
'
' 08/14/03   Structure Development     Substituted interface name in place of guid
'                                      in IJDPropertyDescriptions->AddProperty. The guid of
'                                      a virtual interface may change during bulkload
' 10/08/03   Structure Development     Implemented ISPSMirrorHelper to take care of mirroring
'                                      memberystem
'
' 04/18/06   Structure Development     Calls methods to propagate WPO to the other end and to clear
'                                      the offset on the othen end
' 05/16/06   Structure Development     Curve Member impacts
' 09/27/06   AS                         TR#105470 Remove Blank stubs for CMSetInput and CMRemoveInput
'
' 01/25/07   RP                        TR#113075 - SeatedFC works with curved supporting member now.
' 06/06/07   Struct Dev                TR 120531  added CheckForUndefinedValueAndRaiseError
' 08/17/07   Struct Dev                TR 126260  SetRelatedObjects, set default PointOnPositionRule to ByDistance
' 05/14/08   Struct Dev                DM 142538 remove setting of ByDistance
'
' 07/1/09    RP                        CR165819 - Added code in SetRelatedObjects() to handle the FC's
'                                      association to CanRules
'*******************************************************************************************************
Option Explicit

Private Const MODULE = "SeatedFCDef"
Private m_oLocalizer As IJLocalizer

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.SeatedFCDef"
Const m_ItemName As String = "SPSFCMacros.SeatedFCDef"
Dim bOnPreLoad As Boolean

' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Implements SP3DStructInterfaces.IJUserAttributeMgmt
Implements SP3DStructInterfaces.ISPSMirrorHelper
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
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  ' Add following interfaces as inputs so as to get notified when they are modified.
  'These interfaces carry properties that drive rotation and offset
  pAPDs.AddProperty "SeatedRotationProps", 1, "IJUASPSFCSeatedFlushRotation"
  pAPDs.AddProperty "SeatedOffsetProps", 2, "IJUASPSFCOffset"
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 
  '*************************Rotation*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("SeatedRotation", 1, "CMConstructSeatedRotation", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSeatedRotation"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSeatedRotation"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseSeatedRotation"
  pMemberDescription.IsDeletedWithAggregator = False
  
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisRotation {56CCD3A5-B756-4AB0-9C68-1F586D1E7C66}
  pPropertyDescriptions.AddProperty "ComputeSeatedRotation", 1, "{56CCD3A5-B756-4AB0-9C68-1F586D1E7C66}", "CMUpdateSeatedRotation", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing

  '*************************WPO*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("SeatedWPO", 2, "CMConstructSeatedWPO", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSeatedWPO"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSeatedWPO"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseSeatedWPO"
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisWPO {34DF6BED-41D5-4D19-B090-D58D93A1CF64}
  pPropertyDescriptions.AddProperty "ComputeSeatedWPO", 1, "{34DF6BED-41D5-4D19-B090-D58D93A1CF64}", "CMUpdateSeatedWPO", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  '********************************************************************************************************
  
  Set pMemberDescriptions = Nothing
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConditionalSeatedRotation
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

Public Sub CMConditionalSeatedRotation(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSeatedRotation"
  On Error GoTo ErrorHandler
    Dim oFrameConn As ISPSFrameConnection
    Set oFrameConn = pMemberDescription.CAO
    If oFrameConn.WPO.portIndex = SPSMemberAxisStart Then
            bIsNeeded = True
    Else
            bIsNeeded = False ' second end never drives the rotation
    End If

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConstructSeatedRotation
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

Public Sub CMConstructSeatedRotation(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSeatedRotation"
  On Error GoTo ErrorHandler
    Dim oFrameConn As ISPSFrameConnection
    Set oFrameConn = pMemberDescription.CAO
    Set pObj = oFrameConn.MemberSystem.Rotation
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMSetInputSeatedRotation
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
Public Sub CMSetInputSeatedRotation(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSeatedRotation"
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMUpdateSeatedRotation
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
Public Sub CMUpdateSeatedRotation(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMUpdateSeatedRotation"
    On Error GoTo ErrorHandler
    
    Dim oFrameConn As ISPSFrameConnection
    Dim oSupportedMem As ISPSMemberSystem
    Dim oSupportingMem As ISPSMemberSystem
    Dim oAxisRotation As ISPSAxisRotation
    Dim oObj As Object
    Dim fcStatus As SPSFCInputHelperStatus
    Dim whichSide As SPSCrossSectionSide
    Dim bIsTopRight As Boolean
    Dim bFlip As Boolean
    Dim bMirror As Boolean
    Dim lngOffsetSide As Long

    Dim orientX As Double, orientY As Double, orientZ As Double
    
    Dim oIJAttributes As IJDAttributes
    Dim oCollProxy As CollectionProxy
    
    Set oFrameConn = pPropertyDescriptions.CAO
    Set oSupportedMem = oFrameConn.MemberSystem
    
    fcStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oSupportingMem, oObj)
    
    ' if error occured at reading supporting objects, it may be fixable so call FixFC to see.
    ' if FixFC fails to fix it, it returns false, the FC has been set to Unsupported and we must exit.
    ' if FixFC returns okay, try again to get_RelatedObjects.  If it fails now, we set to unsupported here and exit.
    ' Otherwise, we have been healed and are okay to continue
    
    If fcStatus <> SPSFCInputHelper_Ok Then
        If FixFC(oFrameConn) Then
            fcStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oSupportingMem, oObj)
            If fcStatus <> SPSFCInputHelper_Ok Then
                Set oFrameConn.definition = Nothing
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    Set oIJAttributes = oFrameConn

    Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCSeatedFlushRotation)
    whichSide = oCollProxy.Item(UA_FCSeatedFlushRotation_Edge).Value
    bFlip = CBool(oCollProxy.Item(UA_FCSeatedFlushRotation_Reflect).Value)
    
    Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCOffset)
    ' Value of 1 is supported is on the TopRight of the supporting member (from
    '   the perspective of the supported the supporting is to the left or bottom
    '   of the supported member)
    lngOffsetSide = CLng(oCollProxy.Item(UA_FCOffset_Side).Value)
    If lngOffsetSide = 1 Then
        bIsTopRight = False
    Else
        bIsTopRight = True
    End If

    CheckForUndefinedValueAndRaiseError oFrameConn, whichSide, "StructFrmConnPositionRef", 5
    CheckForUndefinedValueAndRaiseError oFrameConn, lngOffsetSide, "StructFCSupportingSide", 6

    oFrameConn.Services.ComputeMemberOrientation SPSMemberPositionSeated, oSupportingMem, _
                                                oSupportedMem, whichSide, bIsTopRight, bFlip, bMirror, _
                                                orientX, orientY, orientZ
    Set oAxisRotation = pObject
    oAxisRotation.SetOrientationVector orientX, orientY, orientZ
    oAxisRotation.Mirror = bMirror
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL    ' Force the FC to the ToDo List
End Sub

'*************************************************************************
'Function
'CMReleaseSeatedRotation
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
Public Sub CMReleaseSeatedRotation(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseSeatedRotation"
On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConditionalSeatedWPO
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

Public Sub CMConditionalSeatedWPO(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSeatedWPO"
  On Error GoTo ErrorHandler
    bIsNeeded = True ' Always created
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConstructSeatedWPO
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

Public Sub CMConstructSeatedWPO(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSeatedWPO"
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
    
    If Not IsSOOverridden(pIJAttribsConn.CollectionOfAttributes("IJUASPSFCSeatedFlushRotation")) Then
        CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCSeatedFlushRotation"), oAttrbs.CollectionOfAttributes("IJUASPSFCSeatedFlushRotation")
    End If
    
    If Not IsSOOverridden(pIJAttribsConn.CollectionOfAttributes("IJUASPSFCOffset")) Then
        CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCOffset"), oAttrbs.CollectionOfAttributes("IJUASPSFCOffset")
    End If
    
    SetRecatchOfWPO pObj
    
    Set oAttrbs = Nothing
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMSetInputSeatedWPO
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
Public Sub CMSetInputSeatedWPO(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSeatedWPO"
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'*************************************************************************
'Function
'CMUpdateSeatedWPO
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
Public Sub CMUpdateSeatedWPO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMUpdateSeatedWPO"
    On Error GoTo ErrorHandler
  
    Dim oFrameConn As ISPSFrameConnection
    Dim oSupportedMem As ISPSMemberSystem
    Dim oSupportingMem As ISPSMemberSystem
    Dim oWPO As ISPSAxisWPO
    Dim oObj As Object
    Dim fcStatus As SPSFCInputHelperStatus
    Dim whichSide As SPSCrossSectionSide
    Dim WhichEnd As SPSMemberAxisPortIndex
    Dim bIsTopRight As Boolean
    Dim clearance As Double
    Dim offsetCP As Long
    Dim globalOffsetX As Double, globalOffsetY As Double, globalOffsetZ As Double
    Dim localOffsetX As Double, localOffsetY As Double, localOffsetZ As Double
    
    Dim oIJAttributes As IJDAttributes
    Dim oCollProxy As CollectionProxy
    Dim computedEdge As Long
    Dim lngOffsetSide As Long

    Set oFrameConn = pPropertyDescriptions.CAO
    Set oSupportedMem = oFrameConn.MemberSystem

    fcStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oSupportingMem, oObj)
    
    ' if error occured at reading supporting objects, it may be fixable so call FixFC to see.
    ' if FixFC fails to fix it, it returns false, the FC has been set to Unsupported and we must exit.
    ' if FixFC returns okay, try again to get_RelatedObjects.  If it fails now, we set to unsupported here and exit.
    ' Otherwise, we have been healed and are okay to continue
    
    If fcStatus <> SPSFCInputHelper_Ok Then
        If FixFC(oFrameConn) Then
            fcStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oSupportingMem, oObj)
            If fcStatus <> SPSFCInputHelper_Ok Then
                Set oFrameConn.definition = Nothing
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    Set oIJAttributes = oFrameConn

    WhichEnd = oFrameConn.WPO.portIndex
    Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCSeatedFlushRotation)
    whichSide = oCollProxy.Item(UA_FCSeatedFlushRotation_Edge).Value
    
    Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCOffset)
    ' Value of 1 is supported is on the TopRight of the supporting member (from
    '   the perspective of the supported the supporting is to the left or bottom
    '   of the supported member)
    lngOffsetSide = CLng(oCollProxy.Item(UA_FCOffset_Side).Value)
    If lngOffsetSide = 1 Then
        bIsTopRight = False
    Else
        bIsTopRight = True
    End If
    
    clearance = CDbl(oCollProxy.Item(UA_FCOffset_Offset).Value)
    
    CheckForUndefinedValueAndRaiseError oFrameConn, whichSide, "StructFrmConnPositionRef", 5
    CheckForUndefinedValueAndRaiseError oFrameConn, lngOffsetSide, "StructFCSupportingSide", 6
    
    oFrameConn.Services.ComputeWPOCrossing SPSMemberPositionSeated, oSupportingMem, oSupportedMem, WhichEnd, bIsTopRight, _
                                             whichSide, clearance, offsetCP, _
                                             globalOffsetX, globalOffsetY, globalOffsetZ, _
                                             localOffsetX, localOffsetY, localOffsetZ, computedEdge
    Set oWPO = pObject
    oWPO.SetWPO globalOffsetX, globalOffsetY, globalOffsetZ, localOffsetX, localOffsetY, localOffsetZ
    oWPO.WPOCardinalPoint = offsetCP
    
    'if the other end is unsupported, then propagate offset to the other end
    'based on the align flag
    PropagateWPOToOtherEnd oFrameConn
    
    
    ' If the edge was computed then set the edge property
    If computedEdge <> 0 Then
        Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCSeatedFlushRotation)
        oCollProxy.Item(UA_FCSeatedFlushRotation_Edge).Value = computedEdge
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL    ' Force the FC to the ToDo List
End Sub

'*************************************************************************
'Function
'CMReleaseSeatedWPO
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
Public Sub CMReleaseSeatedWPO(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseSeatedWPO"
On Error GoTo ErrorHandler
    
    Dim iFC As ISPSFrameConnection
    Dim iWPO As ISPSAxisWPO
    
    Set iFC = pMD.CAO
    If Not iFC Is Nothing Then
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
ErrorHandler:  HandleError MODULE, METHOD
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
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String

    Dim pFrameConn As ISPSFrameConnection
    Dim WhichEnd As SPSMemberAxisPortIndex  ' on which end of the member is the frame connection
    Set pAttrColl = CollAllDisplayedValues
    Set pFrameConn = pIJDAttrs
    WhichEnd = pFrameConn.WPO.portIndex
    
    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
        If ((pAttrDescr.attrName = "Edge") Or (pAttrDescr.attrName = "Reflect")) Then
            If (WhichEnd = SPSMemberAxisEnd) Then
                pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                pAttrDescr.AttrValue = Empty
            End If
        End If
    Next
    
    Set pFrameConn = Nothing
        
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

Private Function ISPSFCInputHelper_ExecuteSelectionRule(ByVal FC As SPSMembers.ISPSFrameConnection, ByRef selection As String) As SPSMembers.SPSFCInputHelperStatus

    selection = "Seated-Top"
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
    Const METHOD = ""
    On Error GoTo ErrorHandler

    Dim IHStatus As SPSFCInputHelperStatus
    Dim obj As Object
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem, oMS3 As ISPSMemberSystem
    Dim oCR As ISPSCanRule
    
    Set oRC = GetRefColl(FC)
    If oRC.IJDEditJDArgument.GetCount <> 2 Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_MISSING_INPUT, "Seated frame connection is missing a required input. Change connection to unsupported and then re-select the supporting member.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_MISSING_SUPPORTEDMEMBS, "Seated frame connection is missing the supported member system. Change connection to unsupported and then re-select the supporting member.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_MISSING_SUPPORTING_MEMBS, "Seated frame connection is missing the supporting member system. Change connection to unsupported and then re-select the supporting member.")
    End If

    FC.Joint.GetPointOn oMS3, obj
    
    If Not oMS2 Is oMS3 Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_INVALID_PORT_SUPPORTED, "Seated frame connection's supported member has invalid point-on relation to another member. Change connection to unsupported and then re-select the supporting member.")
    End If

    ' only return the PointOn object
    Set RelatedObject1 = oMS2
    Set RelatedObject2 = Nothing
    
    Set oRC = Nothing
    Set oMS1 = Nothing
    Set oMS2 = Nothing
    Set oMS3 = Nothing
    Set obj = Nothing
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
Private Function ISPSFCInputHelper_SetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, _
    ByVal RelatedObject1 As Object, ByVal RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus
    Const METHOD = "ISPSFCInputHelper_SetRelatedObjects"
    On Error GoTo ErrorHandler
    
    Dim IHStatus As SPSFCInputHelperStatus
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oR1 As Object, oR2 As Object
    Dim oMeMS As ISPSMemberSystem
    Dim oSingMS As ISPSMemberSystem
    Dim oSingFC As ISPSFrameConnection
    Dim ePointOn As SPSMembers.StructPointOnCondition
    Dim oCanRulePrimary As ISPSCanRule
    Dim oCanRuleSupported As ISPSCanRule
    Dim oLegMembSys As ISPSMemberSystem


    If Not IsFCCleared(FC) Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_PREVIOUSCONN_NOTREMOVED, "Seated frame connection is invalid because the previous connection was not removed. Change connection to unsupported and the re-select the supporting member.")
    End If
    If RelatedObject1 Is Nothing And Not RelatedObject2 Is Nothing Then
        Set RelatedObject1 = RelatedObject2
        Set RelatedObject2 = Nothing
    End If

    If RelatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_MISSING_SUPPORTED_INFORMATION, "Seated frame connection is missing the supporting member information. Change connection to unsupported and then re-select the supporting member.")
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_NOTRELATED_SUPPORTINGMEMBS, "Seated frame connection is related to something other than a supporting member system. Change connection to unsupported and then re-select the supporting member.")
    End If

    'The Supported MemberSystem cannot be the Supporting MemberSystem.
    If RelatedObject1 Is FC.MemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_SAMEMEMBS_FORBOTH, "Seated frame connection is using the same member system for both the supporting and supported member. Change connection to unsupported and then re-select the supporting member.")
    End If
    If Not RelatedObject2 Is Nothing Then
        If TypeOf RelatedObject2 Is ISPSMemberSystem Or TypeOf RelatedObject2 Is IJPoint Then
            Set RelatedObject2 = Nothing
        End If
    End If

    Set oMeMS = FC.MemberSystem
    Set oSingMS = RelatedObject1

    'save current state of Pointon
    ePointOn = FC.IsPointOn

    RemoveGAPRelations FC ' user may be switching from a GapFC, so clear GAP specific relations
    
    'set interfaces for supported member
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

    ' seated connection does not share joints and is along the supporting member system
    FC.Joint.SetPointOn RelatedObject1, RelatedObject2

    'For the supporting member, get it's supporting members via FC.InputHelper.GetRelatedObjects
    'If either of those are oFC.MemberSystem, then that FC gets set to Unsupported.
    Set oSingFC = oSingMS.FrameConnectionAtEnd(SPSMemberAxisStart)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If
    Set oSingFC = oSingMS.FrameConnectionAtEnd(SPSMemberAxisEnd)
    If HasMeAsSupportingMember(oMeMS, oSingFC) Then
        Set oSingFC.definition = Nothing
    End If
    Set oSingFC = Nothing

    RemoveCommonSplitConnections oMeMS, oSingMS

    ' if prev def was NOT pointOn, then set my positionRule to ByDistance and measure based on current position.
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

    AddOutputWPO FC, 2

    IHStatus = FC.InputHelper.GetRelatedObjects(FC, oR1, oR2)
    If IHStatus <> SPSFCInputHelper_Ok Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_SEATED_REVERIFIED_INVALID, "Failure to establish correct relations")
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
Private Function ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, _
    ByVal options As Long, ByVal snapDistance As Double, _
    ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, _
    ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, _
    RelatedObject1 As Object, RelatedObject2 As Object, _
    RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus

    On Error GoTo ErrorHandler
    Dim IHStatus As SPSMembers.SPSFCInputHelperStatus
    Dim IHStatus2 As SPSMembers.SPSFCInputHelperStatus
    
    IHStatus = SPSFCInputHelper_UnexpectedError
    ' Check to see if this is a curved member, if so they return invalid
    If Not FC Is Nothing Then
        If TypeOf FC.MemberSystem Is ISPSMemberSystemCurve Then ' is a curve member and we are aborting
            IHStatus = SPSFCInputHelper_InvalidTypeOfObject
            GoTo wrapup
        End If
    End If
    
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
 


Private Sub ISPSMirrorHelper_Mirror(ByVal pPlane As SP3DStructInterfaces.IJPlane, ByVal pMatrix As SP3DStructInterfaces.iJDT4x4, ByVal bCopy As Boolean, ByVal oOriginalFC As Object, ByVal oFC As Object)
    Const METHOD = "ISPSMirrorHelper_Mirror"
    On Error GoTo ErrorHandler
    Dim x#, y#, z#, x1#, y1#, z1#, FCx#, FCy#, FCz#
    Dim oFrmConn As ISPSFrameConnection
    Dim oPlaneNormal As New DVector
    Dim dotProd#
    Dim whichSide As SPSCrossSectionSide
    Dim oIJAttributes As IJDAttributes
    Dim oCollProxy As CollectionProxy
    Dim oSuppingSys As ISPSMemberSystem
    Dim oSuppedSys As ISPSMemberSystem
    Dim oInputHelper As ISPSFCInputHelper
    Dim oRelObj1 As Object, oRelObj2 As Object
    Dim oAxis As IJCurve
    Dim oSuppingAxis As New DVector
    Dim oSuppedAxis As New DVector
    Dim bSwitchSide As Boolean
    Dim status As SPSFCInputHelperStatus
    Dim zVec As New DVector
    Dim oCrossVec As IJDVector
    Dim bSuppedAndSuppingInVerticalPlane As Boolean
    Dim oMatAtPos As iJDT4x4
    
    'get frame connection
    Set oFrmConn = oFC
    'get input helper
    Set oInputHelper = oFrmConn.InputHelper
    'get the supporting member
    status = oInputHelper.GetRelatedObjects(oFrmConn, oRelObj1, oRelObj2)
    On Error Resume Next
    Set oSuppingSys = oRelObj1
    If oSuppingSys Is Nothing Then
        Set oSuppingSys = oRelObj2
    End If
    On Error GoTo ErrorHandler
    'get supported member
    Set oSuppedSys = oFrmConn.MemberSystem
    'Craete Z vector
    zVec.Set 0#, 0#, 1#
        
    
    'get connection point
    oFrmConn.Joint.Point.GetPoint FCx, FCy, FCz
    oSuppingSys.Rotation.GetTransformAtPosition FCx, FCy, FCz, oMatAtPos, Nothing
    
    'get old supping axis (tangent at connection point)
    
    'get tangent at connection point
    oSuppingAxis.Set oMatAtPos.IndexValue(0), oMatAtPos.IndexValue(1), oMatAtPos.IndexValue(2)
    
    oSuppingAxis.Length = 1 'normalize

    'get supped axis
    Set oAxis = oSuppedSys.PhysicalAxis
    oAxis.EndPoints x, y, z, x1, y1, z1
    oSuppedAxis.Set x1 - x, y1 - y, z1 - z
    oSuppedAxis.Length = 1 'normalize
    
    'check if supped and supping are in vertical plane
    Set oCrossVec = oSuppedAxis.Cross(oSuppingAxis)
    oCrossVec.Length = 1
    dotProd = zVec.Dot(oCrossVec)
    If Abs(dotProd) < distTol Then 'if supped and supping are in vertical plane, dotProd=0
        bSuppedAndSuppingInVerticalPlane = True
    End If
    
    pPlane.GetNormal x, y, z
    oPlaneNormal.Set x, y, z
    dotProd = oSuppingAxis.Dot(oPlaneNormal)
    'if dotprod=0 then supping axis is parallel to the plane
    'if dotprod=1 then supping axis is normal to the plane

    If bSuppedAndSuppingInVerticalPlane Then
        If Abs(dotProd) < 0.70711 Then ' supping axis and mirror plane close to parallel
            bSwitchSide = True
        End If
    Else
        dotProd = oCrossVec.Dot(oPlaneNormal)
        If Abs(dotProd) > 0.70711 Then ' supping axis and supped axis are close to parallel to mirror plane
            bSwitchSide = True
        End If
    End If
    Set oIJAttributes = oFrmConn
    Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCSeatedFlushRotation)
    
    'switch reflect property
    'Reflect is always switched during mirror
    If oCollProxy.Item(UA_FCSeatedFlushRotation_Reflect).Value = True Then
        oCollProxy.Item(UA_FCSeatedFlushRotation_Reflect).Value = False
    Else
        oCollProxy.Item(UA_FCSeatedFlushRotation_Reflect).Value = True
    End If
    If bSwitchSide Then
        'switch the side property
        Set oCollProxy = oIJAttributes.CollectionOfAttributes(UASPSINTERFACE_FCOffset)
        If oCollProxy.Item(UA_FCOffset_Side).Value = 1 Then 'Top or Right
            oCollProxy.Item(UA_FCOffset_Side).Value = 2 'bottom or left
        Else
            oCollProxy.Item(UA_FCOffset_Side).Value = 1  'Top or Right
        End If
    End If
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
