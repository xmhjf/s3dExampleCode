VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FtgPierDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    FtgPierDef.cls
'
'Author
'       28th January 2003        AP
'
'Description
'
'Notes
'
'History:
'   24-Jun-2006 JMS DI#60069 - Changes to allow editing of the weight and CG values
'                   changed call to SetWCG to call new interface to put weight and CG
'                   values since SetWCG is reservered for setting user defined values
'                   when the values here are the computed values
'       07-Jul-2006 JMS TR#101063 - Tolerate the nonexistence of IJWCGValueOrigin interface
'*******************************************************************
Option Explicit
Private Const MODULE = "FtgPierDef"

Private Const INTERFACE_WCGValueOrigin As String = "IJWCGValueOrigin"
Private Const PROPERTY_DryWCGOrigin As String = "DryWCGOrigin"
Private Enum enumWeightCGDerivation
    WEIGHTCG_Computed = 2
    WEIGHTCG_UserDefined = 4
End Enum

Private Const CONST_ItemProgId As String = "SPSFootingMacros.FtgPierDef"
Private m_oLocalizer As IJLocalizer

Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Dim bOnPreLoad As Boolean




Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
     IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
     IJDUserSymbolServices_GetDefinitionName = CONST_ItemProgId
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
     
     Dim oInput As IJDInput
     Dim oInputs As IJDInputs
     Set oInputs = pDefinition

     Set oInput = New DInput
     oInput.name = SUPPORTING_PLANE
     oInput.index = 1
     oInput.Properties = igDESCRIPTION_OPTIONAL
     oInputs.Add oInput

     Set oInput = Nothing
     Set oInputs = Nothing
     
     ' Aggregator Type
     Dim pAD As IJDAggregatorDescription
     Set pAD = pDefinition
     pAD.AggregatorClsid = "{A2A655C1-E2F5-11D4-9825-00104BD1CC25}" 'SmartClass
     pAD.UserTypeClsid = "{6A8BB3BE-D432-4900-BDB4-C65146CA34D2}" 'CPUAFootingPier
     pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructPier"
     pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructPier"
     pAD.SetCMSetInputs -1, -1
     pAD.SetCMRemoveInputs -1, -1
     Set pAD = Nothing
     
     
  ' Aggregator property
     Dim pAPDs As IJDPropertyDescriptions
     Set pAPDs = pDefinition
     pAPDs.RemoveAll ' Remove all the previous property descriptions
     
     pAPDs.AddProperty PIER_WCG, 1, IJWeightCG, "CMEvaluatePierWCG", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
  
     Set pAPDs = Nothing
   

End Sub


Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal pResourceMgr As Object) As Object
' This method is in charge of the creation of the symbol definition object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler
     
     Dim pDefinition As IJDSymbolDefinition
     Dim pFact As IJCAFactory
     Set pFact = New CAFactory
     Set pDefinition = pFact.CreateCAD(pResourceMgr)
     
     ' Set definition progId and codebase
     pDefinition.ProgId = CONST_ItemProgId
     pDefinition.CodeBase = CodeBase
     
     ' Initialize the definition
     IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
     pDefinition.name = IJDUserSymbolServices_GetDefinitionName(defParams)
     
     ' Persistence behavior
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
     
     'returned symbol definition
     Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
Exit Function
ErrorHandler:      HandleError MODULE, METHOD

End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub
Public Sub CMFinalConstructPier(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructPier"
On Error GoTo ErrorHandler
    

  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMConstructPier(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMConstructPier"
On Error GoTo ErrorHandler
    

  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluatePierWCG(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluatePierWCG"
On Error GoTo ErrHandler
 
    Dim oAttrs As IJDAttributes
    Dim oAttrCol As IJDAttributesCol
    Dim lWCGOrigin As Long

    Set oAttrs = pObject
    lWCGOrigin = WEIGHTCG_Computed
    On Error Resume Next
    Set oAttrCol = oAttrs.CollectionOfAttributes(INTERFACE_WCGValueOrigin)
    On Error GoTo ErrHandler
    If Not oAttrCol Is Nothing Then
        lWCGOrigin = oAttrCol.Item(PROPERTY_DryWCGOrigin).Value
    End If
    
    If lWCGOrigin <> WEIGHTCG_UserDefined Then
        Dim oSmartOcc As IJSmartOccurrence
        Dim iWCG As ISPSComputedWeightCG
        Set oSmartOcc = pObject
        Set iWCG = oSmartOcc
        
        Dim Grade As String, Material As String
        Dim density As Variant
        Dim iMaterial As IJDMaterial
        Set oAttrs = oSmartOcc
     
        Material = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_MATERIAL).Value
        Grade = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_GRADE).Value
        Set iMaterial = GetMaterialObject(Material, Grade)
        If Not iMaterial Is Nothing Then
            density = iMaterial.density
        Else
            density = 2400 'approx density of concrete
        End If
    
        'determine what shape it is to calculate the volume
        Dim length As Double, width As Double, Height As Double
        Dim Shape As Long, dSurfArea As Double, Volume As Double
        Shape = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_SHAPE).Value
        CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, Shape, PRISMATIC_FOOTING_SHAPES, 124
    
        length = oAttrs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value
        width = oAttrs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value
        Height = oAttrs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_HEIGHT).Value
    
    
        If Shape = 2 Then
     'If its rectangular volume is
            Volume = length * width * Height
            dSurfArea = (2 * length * width) + (2 * length * Height) + (2 * width * Height)
        ElseIf Shape = 3 Then
     'If its Circular volume is
            Volume = (3.14159) * width ^ 2 * Height / 4
            dSurfArea = (2 * 3.14159 * width ^ 2 / 4) + (2 * 3.14159 * width / 2 * Height)
        End If
     
        oAttrs.CollectionOfAttributes(IGENERIC_VOLUME).Item(ATTR_VOLUME).Value = Volume
        oAttrs.CollectionOfAttributes(ISURFACE_AREA).Item(SURFACE_AREA).Value = dSurfArea
        
        Dim pSymbol As IJDSymbol
        Dim pOcc As IJDOccurrence
        Dim Matrix As IJDT4x4
        Set Matrix = New DT4x4
        Set pSymbol = pObject
        On Error Resume Next
        Set pOcc = pSymbol
        Set Matrix = pOcc.Matrix
          
        ' The following put property values was originally SetWCG on the IJWeightCG interface
        '   (which is reserved for setting user defined properties), hence changed to put values
        '   on a new interface
        iWCG.Weight = Volume * density
        iWCG.CGx = Matrix.IndexValue(12)
        iWCG.CGy = Matrix.IndexValue(13)
        iWCG.CGz = Matrix.IndexValue(14) - (Height / 2)
    End If
    
Exit Sub
ErrHandler:
If Err.Description = "Undefined Value" Then
    Err.Raise E_FAIL
Else
    HandleError MODULE, METHOD
End If
End Sub
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler



    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")
    
    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long
    
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
        End If
    End If
    
    ' when we change an attribute, we set the AttributeDescriptor_Changed flag
    ' This flag is supposed to be cleared on the client side after updating GOPC
    ' with the changes
    '
    ' We also set the AttributeDescriptor_ChangeAtCommit flag; this flag remains
    ' once set, to give us an idea of the attribute set that changed in this transaction
    
    pAttrToChange.AttrValue = varNewAttrValue
    If (pAttrToChange.attrName = PIER_SIZING_RULE) Then
        If (varNewAttrValue <> 3) Then 'User defined option for sizing rule
            'gray out the pier length and width on the GOPC
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.Count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = PIER_LENGTH) Or (pAttrDescr.attrName = PIER_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                    Else
                        pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                    End If
                End If
            Next
        Else
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.Count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = PIER_LENGTH) Or (pAttrDescr.attrName = PIER_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                        NonStateRO = Not (AttributeDescriptor_ReadOnly)
                        pAttrDescr.AttrState = pAttrDescr.AttrState And NonStateRO
                      End If
                End If
            Next
        End If
    End If
    
    'for combnined slab & pier assembly ,  make pier height readonly
     'also for merged pier & combinedSlabAsm make pier height readonly if it is placed with bottom plane
     
    Dim PierHtReadOnly As Boolean
    Dim oSmartOcc As IJSmartOccurrence
    Dim sysChild As IJDesignChild
    Dim sFtgName As String
    PierHtReadOnly = False
    If (pAttrToChange.attrName = PIER_HEIGHT) Then

        Set sysChild = pIJDAttrs
    
        If Not sysChild Is Nothing Then
            Set oSmartOcc = sysChild.GetParent
            sFtgName = oSmartOcc.Item
            If sFtgName = RECT_PIER_SLAB_COMB_FTG_ASM Or sFtgName = "RectPierAndOctSlab_CombFootingAsm" Then
              PierHtReadOnly = True
            End If
            
            If sFtgName = MERGED_RECT_PIER_COMB_FOOTING_ASM Then
                Dim oRefColl As IJDReferencesCollection
                Dim oRefColl1 As IJDReferencesCollection
                Set oRefColl = GetRefCollection(oSmartOcc)
                Set oRefColl1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
                If Not oRefColl1 Is Nothing Then
                    PierHtReadOnly = True
                End If
                Set oRefColl = Nothing
                Set oRefColl1 = Nothing
            End If
    
            If PierHtReadOnly = True Then
                Set pColl = CollAllDisplayedValues
                For i = 1 To pColl.Count
                    Set pAttrDescr = pColl.Item(i)
                    If (pAttrDescr.attrName = PIER_HEIGHT) Then
                        If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                        Else
                            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                        End If
                    End If
                Next i
            End If
            
            Set pColl = Nothing
            Set sysChild = Nothing
            Set oSmartOcc = Nothing
        End If
    End If
    
    ' For merged pier make pier orientation as Global & readonly
    If (pAttrToChange.attrName = PIER_ORIENTATION) Then
        Set sysChild = pIJDAttrs
    
        If Not sysChild Is Nothing Then
            Set oSmartOcc = sysChild.GetParent
            sFtgName = oSmartOcc.Item
            If sFtgName = MERGED_RECT_PIER_COMB_FOOTING_ASM Then
                Set pColl = CollAllDisplayedValues
                For i = 1 To pColl.Count
                    Set pAttrDescr = pColl.Item(i)
                    If (pAttrDescr.attrName = PIER_ORIENTATION) Then
                        pAttrDescr.AttrValue = 1 ' Global
                        If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                        Else
                            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                        End If
                    End If
                Next i
            End If
            Set pColl = Nothing
            Set sysChild = Nothing
            Set oSmartOcc = Nothing
        End If
    End If
    
    IJUserAttributeMgmt_OnAttributeChange = vbNullString
   
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreCommit"
On Error GoTo ErrorHandler

    IJUserAttributeMgmt_OnPreCommit = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")

    IJUserAttributeMgmt_OnPreCommit = vbNullString
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String
    
    Set pAttrColl = CollAllDisplayedValues
    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
    Next
    
    bOnPreLoad = False

    IJUserAttributeMgmt_OnPreLoad = vbNullString
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function


Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler

' first of all check if the symbol definition has CMCheck methods defined - TBD
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")

    Dim dInputs As IJDInputs
    Dim CurrentInput As IJDInput
    Dim oAttribute As IJDAttribute
    Dim PC As DParameterContent
    Dim bvalid As Boolean
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = pIJDAttrs
    Dim oSymbolDef As IJDSymbolDefinition
    Dim ErrMessage As String
    Set oSymbolDef = oSymbolOcc.IJDSymbolDefinition(2)
    Set dInputs = oSymbolDef.IJDInputs
    Set PC = New DParameterContent
    
    Set oAttribute = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item(sAttributeName)

    If oAttribute.Value <> vbNullString Then
        If oAttribute.AttributeInfo.Type = igString Then    ' check for string type here
        Else
            PC.UomValue = oAttribute.Value
            Set CurrentInput = Nothing
            bvalid = True
            On Error Resume Next
            Set CurrentInput = dInputs.GetInputByName(oAttribute.AttributeInfo.name)
            If Not CurrentInput Is Nothing Then
                CurrentInput.IJDInputDuringGame.definition = oSymbolDef
                CurrentInput.IJDInputStdCustomMethod.InvokeCMCheck PC, bvalid, ErrMessage
                CurrentInput.IJDInputDuringGame.definition = Nothing
                Set oSymbolOcc = Nothing
                Set oSymbolDef = Nothing
                If bvalid = False Then
'                    UserAttributeMgmt_Validate = "Symbol CMCheck Failed"
                    UserAttributeMgmt_Validate = ErrMessage
                    Exit Function
                Else
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
' get the list of interfaces implemented by the schema from IJDAttributes
' make sure that you are not looking into a system interface
' from the input interfaceName and propertyName, get the property type from catalog info
' select case on the property types, and in there, mention the valid attribute values for each propertyName

    Dim InterfaceID As Variant
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
    Dim oAttrCol As IJDInfosCol
    Dim IsInterfaceFound As Boolean
    Dim AttrCount As Long
    Dim AttrType As Long
    
    Set oAttributeMetaData = pIJDAttrs
    IsInterfaceFound = False
    For Each InterfaceID In pIJDAttrs
        Set oInterfaceInfo = Nothing
        Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(InterfaceID)
        If (oInterfaceInfo.IsHardCoded = False) Then
            If (oInterfaceInfo.name = sInterfaceName) Then
                IsInterfaceFound = True
                Exit For
            End If
        End If
    Next
    
    Set oInterfaceInfo = Nothing
    
    If IsInterfaceFound = False Then
        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FTGMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.name = sAttributeName And oAttrObj.name <> PIER_ROTATION_ANGLE And oAttrObj.name <> PIER_SIZE_INC Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (varAttributeValue <= 0#) Then
                            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FTGMACROS_NEGATIVE_ATTRIBVAL, "Negative Attribute Value")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
            End Select
        End If
    Next
    
    UserAttributeMgmt_Validate = vbNullString
    Set oAttributeMetaData = Nothing
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructInterfaces.IJElements)

End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
