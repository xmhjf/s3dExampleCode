VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AxisFCDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE = "AxisFCDef"

' TODO : - Replace <defname> by your selected name
Const m_ItemProgId As String = "SPSFCMacros.AxisFCDef"
Const m_ItemName As String = "SPSFCMacros.AxisFCDef"



' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements SPSMembers.ISPSFCInputHelper
Implements SP3DStructInterfaces.IJUserAttributeMgmt



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
  pAPDs.AddProperty "AxisProps", 1, "{63A8CA52-E95B-41C6-90B7-AEB5C841CE92}" 'IJUASPSFCAxis
  pAPDs.AddProperty "AxisOffsetProps", 2, "{222B2D9C-5D91-44A0-9288-22404757B172}" 'IJUASPSFCManualOffset
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
errorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMSetInputAggregator"
On Error GoTo errorHandler

  
Exit Sub
errorHandler:     HandleError MODULE, METHOD
End Sub



Public Sub CMConditionalAxisWPO(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalAxisWPO"
  On Error GoTo errorHandler
    bIsNeeded = True ' Always created
  
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMConstructAxisWPO(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructAxisWPO"
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
    
    If Not IsSOOverridden(pIJAttribsConn.CollectionOfAttributes("IJUASPSFCAxis")) Then
        CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCAxis"), oAttrbs.CollectionOfAttributes("IJUASPSFCAxis")
    End If
    
    If Not IsSOOverridden(pIJAttribsConn.CollectionOfAttributes("IJUASPSFCManualOffset")) Then
        CopyValuesToSOFromItem pIJAttribsConn.CollectionOfAttributes("IJUASPSFCManualOffset"), oAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset")
    End If
      
    Set oAttrbs = Nothing
  
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMSetInputAxisWPO(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputAxisWPO"
  On Error GoTo errorHandler
  
  Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMUpdateAxisWPO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMUpdateAxisWPO"
    On Error GoTo errorHandler
        Dim IHstatus As SPSMembers.SPSFCInputHelperStatus
        Dim obj As Object
        Dim oWPO As ISPSAxisWPO
'        Dim oSymbol                As IMSSymbolEntities.IJDSymbol
 '       Dim oRefColl   As IMSSymbolEntities.IJDReferencesCollection
  '      Dim oRefColl1   As IMSSymbolEntities.IJDReferencesCollection, oRefColl2  As IMSSymbolEntities.IJDReferencesCollection
        Dim intPort As SPSMemberAxisPortIndex
        Dim oRotation  As ISPSAxisRotation
        Dim oMat1 As IJDT4x4, oMat2 As IJDT4x4
        Dim oMemberPart As ISPSMemberPartPrismatic
        Dim oSys As ISPSMemberSystem
        Dim suppingCP As Long, offsetCP As Long
        Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
        Dim dX As Double, dY As Double
        Dim XOffset As Double, yOffset As Double, zOffset As Double
        Dim lCoordSys As Long
        Dim pIJAttrbs As IJDAttributes
        Dim dBearingPlateThk As Double, bWithBearingPlate As Boolean
        Dim oPos As New DPosition
        Dim oLocPos As New DPosition
        Dim FC As ISPSFrameConnection
        
        Set FC = pPropertyDescriptions.CAO
        IHstatus = FC.InputHelper.GetRelatedObjects(FC, oSys, obj)
    
        If IHstatus <> SPSFCInputHelper_Ok Then
            Exit Sub
        End If
    
        Set oRotation = oSys.Rotation
        Set oMemberPart = oSys.DesignPartAtEnd(FC.WPO.portIndex)
        suppingCP = oMemberPart.CrossSection.CardinalPoint ' the CP with which supporting is placed

        Set pIJAttrbs = pPropertyDescriptions.CAO
 ' the CP of the supporting where the supported needs be positioned
        offsetCP = pIJAttrbs.CollectionOfAttributes("IJUASPSFCAxis").Item("SupportingCP").Value
        
        XOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("XOffset").Value
        yOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("YOffset").Value
        zOffset = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("ZOffset").Value
        lCoordSys = pIJAttrbs.CollectionOfAttributes("IJUASPSFCManualOffset").Item("CoordinateSystem").Value
        
        If offsetCP = 0 Then 'if 0 then use the CP with which the supporting is placed
            offsetCP = suppingCP
        End If
        
        oMemberPart.CrossSection.GetCardinalPointOffset suppingCP, X1, Y1
        oMemberPart.CrossSection.GetCardinalPointOffset offsetCP, X2, Y2
        
        dX = X2 - X1 ' the x offset in cross section coordinates
        dY = Y2 - Y1 ' the y offset in cross section coordinates
        
 'calculate the matrix to apply
        oRotation.GetTransform oMat1
        oMat1.IndexValue(12) = 0
        oMat1.IndexValue(13) = 0
        oMat1.IndexValue(14) = 0
        Set oMat2 = CreateCSToMembTransform(oMat1)
                
                
        Set oWPO = pObject
        
        oPos.Set dX, dY, 0#
        Set oPos = oMat2.TransformPosition(oPos)
        
        If lCoordSys = 1 Then 'Global
            
            oPos.x = oPos.x + XOffset
            oPos.y = oPos.y + yOffset
            oPos.z = oPos.z + zOffset
            
        ElseIf lCoordSys = 2 Then 'Local
            oLocPos.Set XOffset, yOffset, zOffset
            Set oLocPos = oMat1.TransformPosition(oLocPos)
                     
            oPos.x = oPos.x + oLocPos.x
            oPos.y = oPos.y + oLocPos.y
            oPos.z = oPos.z + oLocPos.z
            
            
        End If
        
        oWPO.SetWPO oPos.x, oPos.y, oPos.z, 0, 0, 0
        
        Set oPos = Nothing
        Set oSys = Nothing
        Set oRotation = Nothing
        Set oMat1 = Nothing
        Set oMat2 = Nothing
        Set oMemberPart = Nothing
        Set pIJAttrbs = Nothing
    Exit Sub
errorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMReleaseAxisWPO(ByVal pMD As IJDMemberDescription)
Const MT = "CMReleaseAxisWPO"
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

Private Property Get ISPSFCInputHelper_ExecuteSelectionRule(ByVal FC As SPSMembers.ISPSFrameConnection, selection As String) As SPSMembers.SPSFCInputHelperStatus

    selection = "Axis-Default"
    ISPSFCInputHelper_ExecuteSelectionRule = SPSFCInputHelper_Ok

End Property

Private Property Get ISPSFCInputHelper_GetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, RelatedObject1 As Object, RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus
    On Error GoTo errorHandler
    Dim IHstatus As SPSMembers.SPSFCInputHelperStatus
    IHstatus = SPSFCInputHelper_UnexpectedError

    Dim obj As Object
    Dim errString As String
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    Dim oMS1 As ISPSMemberSystem, oMS2 As ISPSMemberSystem, oMS3 As ISPSMemberSystem
    
    Set oRC = GetRefColl(FC)
    If oRC.IJDEditJDArgument.GetCount <> 2 Then
        IHstatus = SPSFCInputHelper_InconsistentRelations
        errString = "AxisFCSel: Unexpected refcoll count=" & oRC.IJDEditJDArgument.GetCount
        GoTo errorHandler
    End If

    Set oMS1 = oRC.IJDEditJDArgument.GetEntityByIndex(1)
    Set oMS2 = FC.MemberSystem
    If Not oMS1 Is oMS2 Then
        IHstatus = SPSFCInputHelper_InconsistentRelations
        errString = "AxisFCSel: Unknown object 1 in Reference coll"
        GoTo errorHandler
    End If

    Set oMS2 = Nothing
    Set oMS2 = oRC.IJDEditJDArgument.GetEntityByIndex(2)
    If Not TypeOf oMS2 Is ISPSMemberSystem Then
        IHstatus = SPSFCInputHelper_InvalidTypeOfObject
        errString = "AxisFCSel: Unknown object 2 in Reference coll"
        GoTo errorHandler
    End If

    FC.Joint.GetPointOn oMS3, obj
    
    If Not oMS2 Is oMS3 Then
        IHstatus = SPSFCInputHelper_InconsistentRelations
        errString = "AxisFCSel: PointOn MemberSystem is not same as RefColl MemberSystem"
        GoTo errorHandler
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

    Exit Property

errorHandler:
    MsgBox errString
    ISPSFCInputHelper_GetRelatedObjects = IHstatus
    Exit Property
End Property

Private Property Get ISPSFCInputHelper_SetRelatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal RelatedObject1 As Object, ByVal RelatedObject2 As Object) As SPSMembers.SPSFCInputHelperStatus

    On Error GoTo errorHandler
    Dim IHstatus As SPSMembers.SPSFCInputHelperStatus
    IHstatus = SPSFCInputHelper_UnexpectedError
    
    Dim errString As String
    Dim oRC As IMSSymbolEntities.IJDReferencesCollection
    
    Dim oMS As ISPSMemberSystem
    Dim oSysCrSectionNotify As ISPSMemberSystemXSectionNotify
    Dim oRotation As ISPSAxisRotation
    
    Set oRC = GetRefColl(FC)
    oRC.IJDEditJDArgument.RemoveAll
    
    If RelatedObject1 Is Nothing Then
        errString = "AxisFCSel: No RelatedObject1"
        GoTo errorHandler
    End If
    
    If Not TypeOf RelatedObject1 Is ISPSMemberSystem Then
        errString = "AxisFCSel: Unknown type of RelatedObject1"
        GoTo errorHandler
    End If
    
    'Scoping includes logical axis of supported member, AxisRotationInput and XSectionNotify
                                                         
    oRC.IJDEditJDArgument.SetEntity 1, FC.MemberSystem, ISPSMemberSystemStartEndNotify, "MemberSysStartNotifyRC_DEST"

    Set oMS = RelatedObject1
    'Scoping includes physical axis of supporting member, AxisRotation and XSectionNotify
    oRC.IJDEditJDArgument.SetEntity 2, oMS, ISPSMemberSystemSuppingNotify2, "MembSysSuppingNotify2RC_DEST"

    ' axis connection does not share joints and is along the supporting member system
    FC.Joint.SetPointOn RelatedObject1, RelatedObject2

    Dim dbgObj1 As Object, dbgObj2 As Object
    IHstatus = ISPSFCInputHelper_GetRelatedObjects(FC, dbgObj1, dbgObj2)
    If IHstatus <> SPSFCInputHelper_Ok Then
        GoTo errorHandler
    End If
    If Not dbgObj1 Is RelatedObject1 Then
        errString = "AxisFCSel: Failure to establish correct relations"
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

Private Property Get ISPSFCInputHelper_ValidateLocatedObjects(ByVal FC As SPSMembers.ISPSFrameConnection, ByVal options As Long, ByVal snapDistance As Double, ByVal LocatedObject1 As Object, ByVal LocatedObject2 As Object, ByVal LocateX As Double, ByVal LocateY As Double, ByVal LocateZ As Double, RelatedObject1 As Object, RelatedObject2 As Object, RelatedObjectX As Double, RelatedObjectY As Double, RelatedObjectZ As Double) As SPSMembers.SPSFCInputHelperStatus
' FC is optional and can be Nothing
' what we want to do here is if snap is enabled, then
' based on LocatedObject1's type we can return the x,y,z of nearby points of interest such as
' grid-intersections or pointOn joints, endpoints
    
    On Error GoTo errorHandler
    Dim IHstatus As SPSMembers.SPSFCInputHelperStatus
    IHstatus = SPSFCInputHelper_UnexpectedError
    ISPSFCInputHelper_ValidateLocatedObjects = SPSFCInputHelper_UnexpectedError
    
    If FC Is Nothing Then
    End If

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
