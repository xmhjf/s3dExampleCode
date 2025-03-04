VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gap1FCdef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'*******************************************************************
'
'Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'File : Gap1FCDef.cls
'
'Author : Structure Development
'
'Description :
'    SmartPlant Structural GAP1 Frame Connection
'       Definition file
'
'History:
'
' 06/06/07   Struct Dev              CR 70100  Implement Gap FrameConnection
' 12/04/08   Struct Dev              CR153113-  Added code to make GAP connections work for all CPs
' 03/06/09   RP                      CR153111-  Added support to offset the target member along
'                                               secondary member
' 03/31/09   RP                      CR140575-  Added support for overlap/centerline offsets
'
' 07/1/09    RP                      CR165819 - Added code in SetRelatedObjects() to handle the FC's
'                                    association to CanRules
'****************************************************************************************************
Option Explicit

Private Const MODULE = "Gap1FCDef"
Private m_oLocalizer As IJLocalizer

Const m_ItemProgId As String = "SPSFCMacros.Gap1FCDef"
Const m_ItemName As String = "SPSFCMacros.Gap1FCDef"

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
  
  pIH.SetInput "Target"
  pIH.SetInput "Supporting"
  pIH.SetInput "Secondary"
  
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
  ' Add following interface as input so as to get notified when they are modified.
  'These interfaces carry properties that drive the  offset
  pAPDs.AddProperty "GapFC1Props", 1, "ISPSFCGapProps"

  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 

  '*************************WPO*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("Gap1Offset", 1, "CMConstructGapOffset", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGapOffset"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGapOffset"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseGapOffset"
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSGapOffsetInput {DF8568C9-2B0F-4d39-BEF5-BAF4353D360F}
  pPropertyDescriptions.AddProperty "ComputeGapOffset", 1, "{DF8568C9-2B0F-4d39-BEF5-BAF4353D360F}", "CMUpdateGapOffset", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  '********************************************************************************************************
  
  Set pMemberDescriptions = Nothing
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConditionalGapOffset
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

Public Sub CMConditionalGapOffset(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalGapOffset"
  On Error GoTo ErrorHandler
    Dim pIJAttribsConn As IJDAttributes
    
    Set pIJAttribsConn = pMemberDescription.CAO
  
    'set the type to of GAP - offsetting only one member or moving both members
 
    pIJAttribsConn.CollectionOfAttributes("ISPSFCGapProps").Item("GapType").Value = 1 ' GAPSINGLE or GAP1
    
    bIsNeeded = True ' Always created
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

'*************************************************************************
'Function
'CMConstructGapOffset
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

Public Sub CMConstructGapOffset(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructGapOffset"
  On Error GoTo ErrorHandler
    Dim oFrameConn As ISPSFrameConnection
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJAttribsConn As IJDAttributes
    Dim oAttrbs As IJDAttributes
    
    Set oFrameConn = pMemberDescription.CAO
    Set pObj = oFrameConn.WPO
    
    Set pIJAttribsConn = pMemberDescription.CAO

    Set oSmartOcc = pIJAttribsConn
    
    Set oAttrbs = oSmartOcc.ItemObject
    
    CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("ISPSFCGapProps"), oAttrbs.CollectionOfAttributes("ISPSFCGapProps")
    
    SetRecatchOfWPO pObj
    
    Exit Sub
    
ErrorHandler:  HandleError MODULE, MT
Err.Raise E_FAIL
End Sub

'*************************************************************************
'Function
'CMSetInputGapOffset
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
Public Sub CMSetInputGapOffset(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputGapOffset"
  On Error GoTo ErrorHandler
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
Err.Raise E_FAIL
End Sub
'*************************************************************************
'Function
'CMUpdateGapOffset
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
Public Sub CMUpdateGapOffset(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMUpdateGapOffset"
On Error GoTo ErrorHandler
    Dim oFrameConn As ISPSFrameConnection
    Dim oSupportingMembSys As ISPSMemberSystem, oSecondaryMembSys As ISPSMemberSystem, oSupportedMemberSystem As ISPSMemberSystem
    Dim IHStatus As SPSFCInputHelperStatus
    Dim step As Integer
    Dim oAssocCompute As IJStructAssocCompute
    
    Dim oSuppingVec As New DVector, oSuppedVec As New DVector, oSecndaryVec As New DVector
    Dim normVec1 As IJDVector, normVec2 As IJDVector
    Dim pIJAttribsConn As IJDAttributes
    Dim lngClearanceMode As Long
    
    Dim dotProd As Double
    Dim sX As Double, sY As Double, sZ As Double, eX As Double, eY As Double, eZ As Double
    
    step = 0
   
    step = 1
    'check if the Cp5 lines of supported and secondary intersect with CP5 line of supporting
    Set oFrameConn = pPropertyDescriptions.CAO
    IHStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oSupportingMembSys, oSecondaryMembSys)
    
    ' if error occured at reading supporting objects, it may be fixable so call FixFC to see.
    ' if FixFC fails to fix it, it returns false, the FC has been set to Unsupported and we must exit.
    ' if FixFC returns okay, try again to get_RelatedObjects.  If it fails now, we set to unsupported here and exit.
    ' Otherwise, we have been healed and are okay to continue
    
    If IHStatus <> SPSFCInputHelper_Ok Then
        If FixFC(oFrameConn) Then
            IHStatus = oFrameConn.InputHelper.GetRelatedObjects(oFrameConn, oSupportingMembSys, oSecondaryMembSys)
            If IHStatus <> SPSFCInputHelper_Ok Then
                Set oFrameConn.definition = Nothing
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
      
    step = 2
    
    'check if any 2 members are colinear
    oSupportingMembSys.PhysicalAxis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    oSuppingVec.Set eX - sX, eY - sY, eZ - sZ
    oSuppingVec.Length = 1
    
    oSecondaryMembSys.PhysicalAxis.EndPoints sX, sY, sZ, eX, eY, eZ
    oSecndaryVec.Set eX - sX, eY - sY, eZ - sZ
    oSecndaryVec.Length = 1
    
    Set oSupportedMemberSystem = oFrameConn.MemberSystem
    
    oSupportedMemberSystem.LogicalAxis.GetLogicalStartPoint sX, sY, sZ
    oSupportedMemberSystem.LogicalAxis.GetLogicalEndPoint eX, eY, eZ
    oSuppedVec.Set eX - sX, eY - sY, eZ - sZ
    oSuppedVec.Length = 1
    
    dotProd = Abs(oSuppingVec.Dot(oSecndaryVec))
    
    If Abs(dotProd - 1) < angleTol Then
        GoTo ErrorHandler
    End If
    
    dotProd = Abs(oSuppingVec.Dot(oSuppedVec))
    
    If Abs(dotProd - 1) < angleTol Then
        GoTo ErrorHandler
    End If
    
    dotProd = Abs(oSecndaryVec.Dot(oSuppedVec))
    
    If Abs(dotProd - 1) < angleTol Then
        GoTo ErrorHandler
    End If
    
    
    step = 3
    Set normVec1 = oSuppingVec.Cross(oSuppedVec)
    normVec1.Length = 1
    Set normVec2 = oSuppingVec.Cross(oSecndaryVec)
    normVec2.Length = 1
    
    dotProd = Abs(normVec1.Dot(normVec2))
    
    Set pIJAttribsConn = oFrameConn
    lngClearanceMode = pIJAttribsConn.CollectionOfAttributes("ISPSFCGapProps").Item("OffsetDirection").Value
    'for radial clearance members cannot be coplanar
    If (Abs(dotProd - 1) < angleTol) And (lngClearanceMode = SPSSurfaceRadial) Then
        GoTo ErrorHandler
    End If
    
    
    step = 4

    'update ISPSGapOffsetInput
    Set oAssocCompute = New StructAssocTools
    oAssocCompute.UpdateObject pObject, "{DF8568C9-2B0F-4d39-BEF5-BAF4353D360F}"
    
    Exit Sub
ErrorHandler:
    If step = 1 Then
        SPSToDoErrorNotify FCToDoMsgCodelist, TDL_FCMACROS_GAP_MEMBERS_INVALID, oFrameConn, Nothing
    ElseIf step = 2 Then
        SPSToDoErrorNotify FCToDoMsgCodelist, TDL_FCMACROS_GAP_COLINEAR_MEMBERS, oFrameConn, Nothing
    ElseIf step = 3 Then
        SPSToDoErrorNotify FCToDoMsgCodelist, TDL_FCMACROS_GAP_RADIAL_COPLANAR_MEMBERS, oFrameConn, Nothing
    Else
        SPSToDoErrorNotify FCToDoMsgCodelist, TDL_FCMACROS_GAP_UNEXPECTED_ERROR, oFrameConn, Nothing
    End If
    
    Err.Raise SPS_MACRO_WARNING
End Sub

'*************************************************************************
'Function
'CMReleaseGapOffset
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
Public Sub CMReleaseGapOffset(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseGapOffset"
On Error GoTo ErrorHandler
    Dim oFC As ISPSFrameConnection
    Dim oAssocCompute As IJStructAssocCompute
    
    Set oFC = pMD.CAO
    
    If Not oFC Is Nothing Then

        oFC.WPO.SetWPO 0, 0, 0, 0, 0, 0
        oFC.WPO.WPOCardinalPoint = 0
        'update ISPSAxisWPO
        Set oAssocCompute = New StructAssocTools
        oAssocCompute.UpdateObject oFC.WPO, "{34DF6BED-41D5-4D19-B090-D58D93A1CF64}"
        
    End If
  Exit Sub
ErrorHandler: HandleError MODULE, MT
Err.Raise E_FAIL
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
  
    Dim j As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Set pAttrColl = CollAllDisplayedValues
    'when the offset direction is radial, the gap target is always the primary member. so set it to gap along primary and disable that field.
    If (pAttrToChange.attrName = "OffsetDirection") Then
        If varNewAttrValue = 3 Then 'Axial and Flush
            IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FCMACROS_VALUE_NOT_CURRENTLY_SUPPORTED, "Selected value is not currently supported.")
        Else
            For j = 1 To pAttrColl.count
                Set pAttrDescr = pAttrColl.Item(j)
                If (pAttrDescr.attrName = "OffsetAlong") Then
                    If varNewAttrValue = 2 Then 'SurfaceRadial
                        pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                        pAttrDescr.AttrValue = 1 ' gap along primary
                    Else
                        pAttrDescr.AttrState = 0 ' make it editable
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
   
    
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
    
  
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim pAttrDescr1 As IJAttributeDescriptor, pAttrDescr2 As IJAttributeDescriptor, pAttrDescr3 As IJAttributeDescriptor

    Set pAttrColl = CollAllDisplayedValues
    
    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
        If (pAttrDescr.attrName = "OffsetDirection") Then
            Set pAttrDescr1 = pAttrDescr
        ElseIf (pAttrDescr.attrName = "OffsetAlong") Then
            Set pAttrDescr2 = pAttrDescr
        ElseIf (pAttrDescr.attrName = "OffsetType") Then
            Set pAttrDescr3 = pAttrDescr
        ElseIf (pAttrDescr.attrName = "FlushDirection") Or _
         (pAttrDescr.attrName = "FlushOffset") Or _
         (pAttrDescr.attrName = "OverlapRatioMin") Or _
         (pAttrDescr.attrName = "OverlapRatioMax") Then 'set to read only as we added the properties but not the code to process them
         pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        End If
    Next
    
    If Not pAttrDescr2 Is Nothing Then
        If pAttrDescr2.AttrValue = Empty Then
            pAttrDescr2.AttrValue = 1 ' gap along primary
        End If
    End If
    
    If Not pAttrDescr3 Is Nothing Then
        If pAttrDescr3.AttrValue = Empty Then
            pAttrDescr3.AttrValue = 1 ' offset type = gap
        End If
    End If
    
    If Not pAttrDescr1 Is Nothing Then
        If pAttrDescr1.AttrValue = 2 Then 'SurfaceRadial
            pAttrDescr2.AttrState = pAttrDescr2.AttrState Or AttributeDescriptor_ReadOnly
        End If
    End If
    
  
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

Private Function ISPSFCInputHelper_ExecuteSelectionRule(ByVal FC As SPSMembers.ISPSFrameConnection, selection As String) As SPSMembers.SPSFCInputHelperStatus
    selection = "GAP1"
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_REQDINPUT, "Gap frame connection is missing a required input. Change connection to unsupported and then re-select the supporting members.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_BRACEMEMBS, "Gap frame connection is missing the brace member system. Change connection to unsupported and then re-select the supporting members.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_SUPPORTINGMEMB, "Gap frame connection is missing a supporting member system. Change connection to unsupported and then re-select the supporting members.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_TWO_MEMBS, "Gap frame connection is missing one of the two require supporting member systems. Change connection to unsupported and then re-select the supporting member.")
    End If

    ' Verify that the brace's FrameConnection's joint is point on to a joint or splitConn
    FC.Joint.GetPointOn obj, obj2
    If obj Is Nothing Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_EXPECTED_BRACEJOINT, "Gap frame connection requires that the brace have a point on relationship to the intersection of the supporting members. Change connection to unsupported and then re-select the intersection.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_INVALID_RELATIONSHIP, "Gap frame connection has invalid relationships between member systems. Change connection to unsupported and then re-select the intersection.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_PREVIOUSCONN_NOTREMOVED, "Gap frame connection is invalid because the previous connection was not removed. Change connection to unsupported and the re-select the supporting members.")
    End If
    
    If RelatedObject1 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_SUPPORTING_MEMBS, "Gap frame connection is missing the first supporting member information. Change connection to unsupported and then re-select the supporting members.")
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTRELATED_SUPPORTINGMEMB, "Gap frame connection is related to something other than the first supporting member system. Change connection to unsupported and then re-select the supporting member.")
    End If

    If RelatedObject2 Is Nothing Then
        IHStatus = SPSFCInputHelper_BadNumberOfObjects
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_MISSING_SUPPRTINGMEMBS, "Gap frame connection is missing the second supporting member information. Change connection to unsupported and then re-select the supporting members.")
    End If
    
    If Not TypeOf RelatedObject2 Is ISPSMemberSystem Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTRELATED_SECOND_SUPPORTING, "Gap frame connection is related to something other than the second supporting member system. Change connection to unsupported and then re-select the supporting members.")
    End If

    Set oMeMS = FC.MemberSystem
    Set oSingMS1 = RelatedObject1
    Set oSingMS2 = RelatedObject2

    If oSingMS1 Is oMeMS Or oSingMS2 Is oMeMS Then
        'The Supported MemberSystem cannot be the Supporting MemberSystem.
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_USING_BRACEMEMB_INCORRECTLY, "Gap frame connection is using the brace member incorrectly as a supporting member. Change connection to unsupported and then re-select the supporting members.")
    End If

    If oSingMS1 Is oSingMS2 Then
        IHStatus = SPSFCInputHelper_DuplicateObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_SAMEMEMBS_FORBOTH, "Gap frame connection is using the same member system for both supporting members. The two supporting members must be different members. Change connection to unsupported and then re-select the supporting members.")
    End If

'lngRelated     Meaning
' 0             not related
' 1             oMS2 is dependent on oMS1.  oPointConn is a FC at an end of oMS2
' 2             oMS1 is dependent on oMS2.  oPointConn is a FC at an end of oMS1
' 3             oPointConn is a split connection.  dependency not clear !

    GetCommonConnection oSingMS1, oSingMS2, oPointConn, lngRelated
    
    If (oPointConn Is Nothing) Or (lngRelated = 0) Then
        IHStatus = SPSFCInputHelper_InvalidTypeOfObject
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_NOTRELATED_SECOND_SUPPORTING, "Gap frame connection is related to something other than the second supporting member system. Change connection to unsupported and then re-select the supporting members.")
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
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_EXPECTED_BRACEJOINT, "Gap frame cannot resolve FC joint with its MemberSystem.")
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

    'set interfaces for supported member
    Set oRC = GetRefColl(FC)
    Set oCanRulePrimary = FC.GetCrossSectionObject(SPSFCPrimary)
    Set oCanRuleSecondary = FC.GetCrossSectionObject(SPSFCSecondary)
    Set oCanRuleSupported = FC.GetCrossSectionObject(SPSFCSupported)
    
    RemoveGAPRelations FC
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

    '  not an oversight.  AddOutputWPO FC, 1  we want to call CMConstructGapOffset during compute.

    IHStatus = FC.InputHelper.GetRelatedObjects(FC, oR1, oR2)
    If IHStatus <> SPSFCInputHelper_Ok Then
        IHStatus = SPSFCInputHelper_InconsistentRelations
        Err.Raise IHStatus + SPSvbError, "", m_oLocalizer.GetString(IDS_FCMACROS_VC_FAILURETO_ESTB_REL, "Failure to establish correct relations")
    End If

    SetGAPRelations FC, SPSGap1  'establishes relation to supporting
    'and secondary member and also the extend relation between FC and port to get the computed values
    'from the port and display on the FC proppage/filter dialog
    
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

    'initialize to nothing
    Set RelatedObject1 = Nothing
    Set RelatedObject2 = Nothing

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


