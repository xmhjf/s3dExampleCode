VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToEndFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE = "EndToEndFCDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.EndToEndFCDef"
Const m_ItemName As String = "SPSFCMacros.EndToEndFCDef"


' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Implements SP3DStructInterfaces.IJUserAttributeMgmt
Implements SPSMembers.ISPSFCInputHelper


Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo errorHandler
  
  pIH.SetInput "SupportedRefColl"
  pIH.SetInput "SupportingRefColl"
  
  Exit Sub
errorHandler:
  pIH.ReportError
End Sub

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo errorHandler

  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs -
  ' They are identical to the class inputs (i.e. penetrated and Penetrating)
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
  DefinitionInputs pIH

  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{5FEB4ADB-E5EC-45D6-8878-150ADDC04D0A}"     'CSPSFrameConnection
  pAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAggregator"
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  ' Add following interfaces as inputs so as to get notified when they are modified.
  'These interfaces carry properties that drive rotation and offset
  pAPDs.AddProperty "EndToEndProps", 1, "{57287753-2FEA-4220-AA05-989043DA7351}" 'IJUASPSFCEndToEnd

  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  pMemberDescriptions.RemoveAll  ' Remove all the previous member descriptions
 
  '*************************Rotation*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("EndToEndRotation", 1, "CMConstructEndToEndRotation", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputEndToEndRotation"
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndToEndRotation"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndToEndRotation"
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisRotation {56CCD3A5-B756-4AB0-9C68-1F586D1E7C66}
  pPropertyDescriptions.AddProperty "ComputeEndToEndRotation", 1, "{56CCD3A5-B756-4AB0-9C68-1F586D1E7C66}", "CMUpdateEndToEndRotation", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing

  '*************************WPO*******************************************************************************
  Set pMemberDescription = pMemberDescriptions.AddMember("EndToEndWPO", 2, "CMConstructEndToEndWPO", imsCOOKIE_ID_USS_LIB)
  'Declare all custom methods. Some of them would be just place holders. Adding them now avoids a bulkload.
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalEndToEndWPO"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputEndToEndWPO"
  pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseEndToEndWPO"
  pMemberDescription.IsDeletedWithAggregator = False
    
  Set pPropertyDescriptions = pMemberDescription
  'outputs ISPSAxisWPO {34DF6BED-41D5-4D19-B090-D58D93A1CF64}
  pPropertyDescriptions.AddProperty "ComputeEndToEndWPO", 1, "{34DF6BED-41D5-4D19-B090-D58D93A1CF64}", "CMUpdateEndToEndWPO", imsCOOKIE_ID_USS_LIB
  
  Set pPropertyDescriptions = Nothing
  Set pMemberDescription = Nothing
  '********************************************************************************************************
  
  Set pMemberDescriptions = Nothing
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMSetInputAggregator"
On Error GoTo errorHandler

  
Exit Sub
errorHandler:     HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalEndToEndRotation(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalEndToEndRotation"
  On Error GoTo errorHandler
    
    Dim oFrameConn As ISPSFrameConnection
    Set oFrameConn = pMemberDescription.CAO
    If oFrameConn.WPO.portIndex = SPSMemberAxisStart Then
            bIsNeeded = True
    Else
            bIsNeeded = False ' second end never drives the rotation
    End If

  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMConstructEndToEndRotation(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructEndToEndRotation"
  On Error GoTo errorHandler
    Dim oFrameConn As ISPSFrameConnection
    Set oFrameConn = pMemberDescription.CAO
    Set pObj = oFrameConn.MemberSystem.Rotation
  
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMSetInputEndToEndRotation(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputEndToEndRotation"
  On Error GoTo errorHandler
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMUpdateEndToEndRotation(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMUpdateEndToEndRotation"
    On Error GoTo errorHandler
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMReleaseEndToEndRotation(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseEndToEndRotation"
On Error GoTo errorHandler
  Exit Sub
errorHandler: HandleError MODULE, MT
End Sub

Public Sub CMConditionalEndToEndWPO(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalEndToEndWPO"
  On Error GoTo errorHandler
    bIsNeeded = True ' Always created
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMConstructEndToEndWPO(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructEndToEndWPO"
  On Error GoTo errorHandler
    Dim oFrameConn As ISPSFrameConnection
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJAttribsConn As IJDAttributes
    Dim oAttrbs As IJDAttributes
    
    Set oFrameConn = pMemberDescription.CAO
    Set pObj = oFrameConn.WPO
    Set oSmartOcc = oFrameConn
    
    Set pIJAttribsConn = oSmartOcc '.ItemObject
    Set oAttrbs = oSmartOcc.ItemObject
    CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCEndToEnd"), oAttrbs.CollectionOfAttributes("IJUASPSFCEndToEnd")
    Set oAttrbs = Nothing
  
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMSetInputEndToEndWPO(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputEndToEndWPO"
  On Error GoTo errorHandler
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMUpdateEndToEndWPO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMUpdateEndToEndWPO"
    On Error GoTo errorHandler
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMReleaseEndToEndWPO(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseEndToEndWPO"
On Error GoTo errorHandler
  Exit Sub
errorHandler: HandleError MODULE, MT
End Sub

'
' The following methods are generic for all the Custom assembly
'
'
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








Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String

End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function


Private Property Get ISPSFCInputHelper_ExecuteSelectionRule(ByVal FC As SPSMembers.ISPSFrameConnection, ByRef selection As String) As SPSMembers.SPSFCInputHelperStatus

    selection = "EndToEnd-Top"

End Property

Private Property Get ISPSFCInputHelper_GetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, RelatedObject1 As Object, RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus

    Dim errString As String
    Dim portIndex As SPSMemberAxisPortIndex
    Dim SingJoint As ISPSAxisJoint
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem
    
    Set oRC = GetRefColl(FC)
    If oRC.IJDEditJDArgument.GetCount <> 2 Then
        errString = "EndToEndFCSel: Unexpected refcoll count=" & oRC.IJDEditJDArgument.GetCount
        GoTo errorHandler
    End If

    Set oMS1 = oRC.IJDEditJDArgument.GetEntityByIndex(1)
    Set oMS2 = FC.MemberSystem
    If Not oMS1 Is oMS2 Then
        errString = "EndToEndFCSel: Unknown object 1 in Reference coll"
        GoTo errorHandler
    End If

    Set oMS2 = Nothing
    Set oMS2 = oRC.IJDEditJDArgument.GetEntityByIndex(2)
    If Not TypeOf oMS2 Is ISPSMemberSystem Then
        errString = "EndToEndFCSel: Unknown object 2 in Reference coll"
        GoTo errorHandler
    End If

    portIndex = oMS2.ResolveEnd(FC.Joint)
    If portIndex <> SPSMemberAxisStart And portIndex <> SPSMemberAxisEnd Then
        errString = "EndToEndFCSel: Members are not constrained by the same joint."
        GoTo errorHandler
    End If
    
    Set RelatedObject1 = oMS2
    Set RelatedObject2 = Nothing
    
    Exit Property

errorHandler:
    MsgBox errString
    Exit Property

End Property

Private Property Get ISPSFCInputHelper_SetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, _
    ByVal RelatedObject1 As Object, ByVal RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus

    Dim errString As String
    Dim IHstatus As SPSMembers.SPSFCInputHelperStatus
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    
    Dim SingJoint As ISPSAxisJoint
    Dim oMS As ISPSMemberSystem
    Dim EndPosX As Double, EndPosY As Double, EndPosZ As Double

    Set oRC = GetRefColl(FC)
    oRC.IJDEditJDArgument.RemoveAll
    
    If RelatedObject1 Is Nothing Then
        errString = "EndToEndFCSel: No RelatedObject1"
        GoTo errorHandler
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        errString = "EndToEndFCSel: Unknown type of RelatedObject1"
        GoTo errorHandler
    End If
    
    FC.Joint.Point.GetPoint EndPosX, EndPosY, EndPosZ
    Set oMS = RelatedObject1
    Set SingJoint = oMS.JointAtEndPosition(EndPosX, EndPosY, EndPosZ)
    If SingJoint Is Nothing Then
        errString = "EndToEndFCSel: Supported Member not at end of Supporting Member"
        GoTo errorHandler
    End If

    'We'll just ignore any RelatedObject2.  It might be a grid plane.
    'set interfaces for supported member
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

    'EndEnd connection shares the end joint with the supporting member
    SingJoint.AddMember FC.MemberSystem, FC.WPO.portIndex

    Dim dbgObj1 As Object, dbgObj2 As Object
    IHstatus = ISPSFCInputHelper_GetRelatedObjects(FC, dbgObj1, dbgObj2)
    If IHstatus <> SPSFCInputHelper_Ok Then
        GoTo errorHandler
    End If
    If Not dbgObj1 Is RelatedObject1 Then
        errString = "EndToEndFCSel: Failure to establish correct relations"
        GoTo errorHandler
    End If
    ISPSFCInputHelper_SetRelatedObjects = SPSFCInputHelper_Ok
    Exit Property

errorHandler:
    ISPSFCInputHelper_SetRelatedObjects = IHstatus
    MsgBox errString
    Exit Property
End Property


 

Private Property Get ISPSFCInputHelper_UserAttributeMgmt() As SPSMembers.IJUserAttributeMgmt
   Set ISPSFCInputHelper_UserAttributeMgmt = Me
End Property

Private Property Get ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, _
    ByVal options As Long, ByVal snapDistance As Double, _
    ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, _
    ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, _
    RelatedObject1 As Object, RelatedObject2 As Object, _
    RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus

' FC is optional and can be Nothing
' what we want to do here is if snap is enabled, then
' based on LocatedObject1's type we can return the x,y,z of nearby points of interest such as
' grid-intersections or pointOn joints, endpoints

    On Error GoTo errorHandler
    Dim IHstatus As SPSMembers.SPSFCInputHelperStatus
    IHstatus = SPSFCInputHelper_UnexpectedError
    ISPSFCInputHelper_ValidateLocatedObjects = SPSFCInputHelper_UnexpectedError

    If LocatedObject1 Is Nothing And Not LocatedObject2 Is Nothing Then
        Set LocatedObject1 = LocatedObject2
        Set LocatedObject2 = Nothing
    End If
        
    If LocatedObject1 Is Nothing Then
        IHstatus = SPSFCInputHelper_BadNumberOfObjects
        GoTo wrapup
    ElseIf Not TypeOf LocatedObject1 Is ISPSMemberSystem Then
        IHstatus = SPSFCInputHelper_InvalidTypeOfObject
        GoTo wrapup
    End If
    
    If Not FC Is Nothing Then
        If LocatedObject1 Is FC.MemberSystem Then
            IHstatus = SPSFCInputHelper_DuplicateObject
            GoTo wrapup
        End If
    End If

    If LocatedObject1 Is LocatedObject2 Then
        IHstatus = SPSFCInputHelper_DuplicateObject
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

    IHstatus = SPSFCInputHelper_Ok

wrapup:
    ISPSFCInputHelper_ValidateLocatedObjects = IHstatus
    Exit Property

errorHandler:
    ISPSFCInputHelper_ValidateLocatedObjects = IHstatus
End Property
 

