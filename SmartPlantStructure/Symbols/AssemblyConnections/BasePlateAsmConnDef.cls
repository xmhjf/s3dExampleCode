VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasePlateAsmConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : BasePlateAsmConnDef.cls
'
'Author : RP
'
'Description :
'    Macro for creating baseplate connection
'
'History:
'
' 08/14/03  RP   Substituted interface name in place of guid
'                in IJDPropertyDescriptions->AddProperty. The guid of
'                a virtual interface may change during bulkload
' 02/02/04  JS   Added code to verify the assembly connection
'                and check if an AC already exists between
'                the supported or supporting members TR#52040
' 02/05/04  JS   Changed all generic reference collection relations to
'                specific relations. This includes the relations
'                to the split axis end ports, split axis along ports
'                and the cutback planes (if a plane exists).
' 08/27/04  MH   CMMigrate
' 06/13/06  RP   Changes due to impact from curved members(DI#84001)
' 19/Sep/06 SS   TR#105292 - changes to base plate asm connection to add plate part
'                as an assembly child of its parent parent
' 27/Sep/06 AS   TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
' 07/Jun/07 MH   TR 121116 added CheckForUndefinedValueAndRaiseError
'
'********************************************************************

Option Explicit
Dim bOnPreLoad As Boolean
Private Const MODULE = "BasePlateAsmConnDef"

Const m_ItemProgId As String = "SPSACMacros.BasePlateAsmConnDef"
Const m_ItemName As String = "SPSACMacros.BasePlateAsmConnDef"
Private Const strSourceFile As String = "BasePlateAsmConnDef.cls"
Private m_oLocalizer As IJLocalizer
Implements ISPSFACInputHelper
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJUserAttributeMgmtParent
Implements IJStructCustomFoulCheck

'*************************************************************************
'Function
'DefinitionInputs
'
'Abstract
'Sets up the inputs for the base plate FrameConnection
'
'Arguments
'pIH As IJDInputsHelper - Input
'
'Exceptions
'
'***************************************************************************

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  Dim oInput As IJDInput
  Dim oInputs As IJDInputs
  
  Set oInputs = pIH.Definition 'optional Plane
  Set oInput = New DInput
  oInput.Name = "Plane"
  oInput.index = 1
  oInput.Properties = igDESCRIPTION_OPTIONAL
  oInputs.Add oInput
  
  Set oInput = Nothing
  Set oInputs = Nothing
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub


'*************************************************************************
'Function
'IJDUserSymbolServices_InitializeSymbolDefinition
'
'Abstract
'Bulk load called method that establishes the Base Plate assembly connection
'definition
'
'Arguments
'pDefinition - Symbol definition created for the base plate assembly connection
'
'Exceptions
'
'***************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler
  
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs for the base plate connection (the helper class is initialized to the
  ' symbol definition and is an aid to help set the inputs)
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{E43FD681-1B37-4CC1-BD94-F399F43F946F}"     'CStructAssemblyConnection
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  Set pAD = Nothing
  
  ' tr 74803
  Dim pCADefinition As IJCADefinition
  Set pCADefinition = pDefinition
  Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
  Set pCADefinition = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  ' Listens to IJUASPSBasePlateAsmConn
  pAPDs.AddProperty "PlateSize", 1, "IJUASPSBasePlateAsmConn", "CMComputePlateSize", imsCOOKIE_ID_USS_LIB
  ' Listens to IJUASPSPlateTypeCategory
  pAPDs.AddProperty "PlateTypeCategory", 2, "IJUASPSPlateTypeCategory"
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll

  '_________________________________________________________________________________________________________________________________________
  ' Plate Smart Occurrence
  Set pMemberDescription = pMemberDescriptions.AddMember("BasePlateAsmConnPlate", 1, "CMConstructBasePlateAsmConnPlate", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID

  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalBasePlateAsmConnPlate"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputBasePlateAsmConnPlate"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateBasePlateAsmConnPlate"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "BasePlateSize", 1, "IJUASPSPlatePartDim", "CMComputeBasePlateSize", imsCOOKIE_ID_USS_LIB
  ' Outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "BasePlateThickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeBasePlateThickness", imsCOOKIE_ID_USS_LIB
  ' Outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "BasePlatePosition", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeBasePlatePosition", imsCOOKIE_ID_USS_LIB
  
  '_________________________________________________________________________________________________________________________________________
  ' Cutback Plane
  Set pMemberDescription = pMemberDescriptions.AddMember("BasePlateAsmConnCutbackPlane", 2, "CMConstructBasePlateAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalBasePlateAsmConnCutbackPlane"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputBasePlateAsmConnCutbackPlane"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateBasePlateAsmConnCutbackPlane"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJPlane.
  pPropertyDescriptions.AddProperty "ComputeBasePlateAsmConnCutbackPlane", 1, "{4317C6B3-D265-11D1-9558-0060973D4824}", "CMComputeBasePlateAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB
  
  '_________________________________________________________________________________________________________________________________________
  ' Cutback Smart Occurrence
  Set pMemberDescription = pMemberDescriptions.AddMember("BasePlateAsmConnCutback", 3, "CMConstructBasePlateAsmConnCutback", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalBasePlateAsmConnCutback"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputBasePlateAsmConnCutback"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateBasePlateAsmConnCutback"
  Set pPropertyDescriptions = pMemberDescription
  ' Outputs IJSmartOccurrence (this is just to make sure that the parent computes before the child)
  pPropertyDescriptions.AddProperty "ComputeBasePlateAsmConnCutback", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeBasePlateAsmConnCutback", imsCOOKIE_ID_USS_LIB
  
  Set pMemberDescriptions = Nothing
  Set pMemberDescription = Nothing
  Set pPropertyDescriptions = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMFinalConstructAggregator
'
'Abstract
'Permits one-time initialization of the user-defined aggregator class.
'
'Arguments
'IJDAggregatorDescription interface of the CustomAssemblyDefinition wrapper
'
'Exceptions
'
'***************************************************************************
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructAggregator"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    
    MigrateTheAggregator pAggregatorDescription, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub



'*************************************************************************
'Function
'CMComputePlateSize
'
'Abstract
'Callback Method to notify the base plate connection that the "PlateSize"
'property changed
'
'Arguments
'pPropertyDescriptions -    collection of the property descriptions for the custom
'                           assembly defined at bulk load time
'                           (see IJDUserSymbolServices_InitializeSymbolDefinition).
'                           Contains the metadata descriptions of such
'                           properties as BasePlateSize.
'pObject - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMComputePlateSize(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMComputePlateSize"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'*************************************************************************
'Function
'CMConditionalBasePlateAsmConnPlate
'
'Abstract
'Callback Method to determine whether the current member is needed as an output.
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'bIsNeeded - boolean which is always returned as true to indicate that a
'            base plate is always needed regardless of the input part or port.
'
'Exceptions
'
'Raises an "Unknown section encountered" error when the section type is not
'one of the standard section types. Raising the error at this time will force
'the base plate custom assembly into the ToDo list.
'
'***************************************************************************
Public Sub CMConditionalBasePlateAsmConnPlate(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalBasePlateAsmConnPlate"
    On Error GoTo ErrorHandler
    
    Dim oSuppedPort As ISPSSplitAxisEndPort
    Dim oMembPort As ISPSSplitAxisPort
    Dim oSuppingPort As IJPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements

    Dim strError As String
    ' Retrieve the inputs of the custom assembly occurrence to access the ports
    Set oStructConn = pMemberDescription.CAO
    oStructConn.enumPorts colPorts
    
    If colPorts.count = 1 Then
        Set oSuppedPort = colPorts.Item(1)
        If oSuppedPort Is Nothing Then
            SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_NOOF_INPUTS_BASEPLATE, oStructConn, Nothing
            Err.Raise E_FAIL
        End If
    ElseIf colPorts.count = 2 Then ' a plane is also input
        Set oSuppedPort = GetEndPort(colPorts)
        If oSuppedPort Is colPorts.Item(1) Then
            Set oSuppingPort = colPorts.Item(2)
        Else
            Set oSuppingPort = colPorts.Item(1)
        End If
        If oSuppedPort Is Nothing Or oSuppingPort Is Nothing Then
            SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_NOOF_INPUTS_BASEPLATE, oStructConn, Nothing
            Err.Raise E_FAIL
        End If
    Else
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_NOOF_INPUTS_BASEPLATE, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
    
    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed and the assembly connection added to the ToDo list
    If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
 
    Set oMembPort = oSuppedPort
    Set oSuppedPart = oMembPort.Part
    If Not ValidSectionType(oSuppedPart) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_UNKNOWNSEC_ENCOUNT, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
    bIsNeeded = True
    Exit Sub
    
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If
End Sub

'*************************************************************************
'Function
'CMConstructBasePlateAsmConnPlate
'
'Abstract
'Callback Method to construct the base plate object for the Custom assembly
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'pResourceManager - resource manager for the plant model database
'pObj - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMConstructBasePlateAsmConnPlate(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructBasePlateAsmConnPlate"
    On Error GoTo ErrorHandler
    Dim bCreateAtEnd As Boolean
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl   As IJDReferencesCollection
    Dim oSmartOccCAO      As IJSmartOccurrence
    Dim oSymbolOfCAO                As IJDSymbol
    Dim oMembObjects As IJDMemberObjects
    Dim oDesignParent As IJDesignParent
    Dim PlateThickness As Double
    Dim oAttrbs As IJDAttributes
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    
    ' Retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oMembObjects = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    oStructConn.enumPorts colPorts

    Set oSuppedPort = GetEndPort(colPorts)
    Set oSuppedPart = oSuppedPort.Part
    
    Set oFactory = New StructCustomPlatePartFactory
    ' Create the custom Plate that is the Base Plate.
    Set oPlate = oFactory.CreateCustomPlatePart(pResourceManager)
    
    Set oAttrbs = oSmartOccCAO.ItemObject
        
    On Error Resume Next
        Dim materialProxy As CollectionProxy
        'the interface below is added in V6. It may not exist in older catalog
        'New code and old catalog is unusual. The resume next above takes care of this though
        Set materialProxy = oAttrbs.CollectionOfAttributes("IJUASPSPlateMaterial")
    On Error GoTo ErrorHandler
    
    If Not materialProxy Is Nothing Then
        Dim strMaterial As String, strGrade As String
        strMaterial = materialProxy.Item("SPSMaterial").Value
        strGrade = materialProxy.Item("SPSGrade").Value
        SetPlateMaterial oPlate, strMaterial, strGrade
    Else
        'use hard coded material, as done prior to V6
        SetPlateMaterial oPlate, "Steel - Carbon", "A"
    End If
    
    
    ' Get default from the parent
    PlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value
    ' Set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, PlateThickness
    
    Set oSmartOcc = oPlate
    ' Initialize the SmartOccurrence
    oSmartOcc.RootSelectorClass = "GenericRectPlatePart"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)

    ' Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    Call AddAsmConnToAsmParent(oSmartOccCAO, oPlate)
    
    ' Create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefColl
    Set pObj = oPlate ' Return the created plate object to the custom assembly
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMSetInputBasePlateAsmConnPlate
'
'Abstract
'Callback Method used to set inputs on the parent SmartOccurrence prior to each CM evaluate
'
'Arguments
'pMemberDesc - Metadata describing the components of the custom
'              assembly which include the plate, cutback and
'              cutback plane. This metdata information was defined
'              at bulkload time via the
'              IJDUserSymbolServices_InitializeSymbolDefinition function.
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputBasePlateAsmConnPlate(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputBasePlateAsmConnPlate"
  On Error GoTo ErrorHandler
  
  Dim pIJAttribsCAO As IJDAttributes
  Dim oSmartOccCAO As IJSmartOccurrence
  Dim oCustomPlate As StructCustomPlatePart
  
  Set oSmartOccCAO = pMemberDesc.CAO
  Set pIJAttribsCAO = oSmartOccCAO
  On Error Resume Next ' Suppress the error that VB throws when QI fails
  Set oCustomPlate = pMemberDesc.object
  On Error GoTo ErrorHandler
  
  If Not oCustomPlate Is Nothing Then
      ' Checks if the occurrence attributes are Empty
      If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory")) Then
          ' Copies defintion values to occurrence values
          Dim oItemAttrbs As IJDAttributes
          Set oItemAttrbs = oSmartOccCAO.ItemObject
          CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory"), oItemAttrbs.CollectionOfAttributes("IJUASPSPlateTypeCategory")
      End If

      ' Copy plate category and type from parent to child
      oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Category").Value
      oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Type").Value
  End If
  
  Exit Sub
  
ErrorHandler:
  HandleError MODULE, MT
End Sub
Public Sub CMMigrateBasePlateAsmConnPlate(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateBasePlateAsmConnPlate"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub


'*************************************************************************
'Function
'CMComputeBasePlateSize
'
'Abstract
'Callback Method to re-evaluate the baseplate due to a change in the base
'plate's property size
'
'Arguments
'pPropertyDescriptions - collection of the property descriptions for the custom
'                        assembly defined at bulk load time
'                        (see IJDUserSymbolServices_InitializeSymbolDefinition).
'                        Contains the metadata descriptions of such
'                        properties as BasePlateSize.
'pObject - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMComputeBasePlateSize(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeBasePlateSize"
    On Error GoTo ErrorHandler

    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim DepthClr As Double, WidthClr As Double
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim sizeRule As Long
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim secDepth As Double, SecWidth As Double
    Dim depth As Double, Width As Double
    Dim oPlate As IJStructPlate
    Dim PlateThickness As Double
    Dim varThickness As Variant
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppingPlane As IJPlane
    Dim oMembDir As New DVector
    Dim oSuppingPlaneNorm As New DVector
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim nx#, ny#, nz#
    Dim oZDir As New DVector
    Dim cosA#
    Dim oRefColl As IJDReferencesCollection
 
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    oStructConn.enumPorts colPorts
    If colPorts.count = 1 Then
        Set oSuppedPort = colPorts.Item(1)
        Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOccCAO)
        On Error Resume Next
        ' Check if there is a plane in the ref coll
        ' if there is on we use that for positioning the plate
        Set oSuppingPlane = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        On Error GoTo ErrorHandler
    ElseIf colPorts.count = 2 Then ' a plane is also input
        Set oSuppedPort = GetEndPort(colPorts)
        If oSuppedPort Is colPorts.Item(1) Then
            Set oSuppingPlane = colPorts.Item(2)
        Else
            Set oSuppingPlane = colPorts.Item(1)
        End If
    End If
    Set oSuppedPart = oSuppedPort.Part
    
    Set oSectionAttrbs = oSuppedPart.CrossSection.Definition
    secDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    SecWidth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    
    If Not oSuppingPlane Is Nothing Then
        oSuppedPart.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
        If oSuppedPort.portIndex = SPSMemberAxisStart Then
            oMembDir.Set eX - sX, eY - sY, eZ - sZ
        Else
            oMembDir.Set sX - eX, sY - eY, sZ - eZ
        End If
        oMembDir.length = 1 ' Normalize the vector
        oSuppingPlane.GetNormal nx, ny, nz
        oZDir.Set nx, ny, nz
        cosA = oZDir.Dot(oMembDir) 'A is the angle between plane's normal and member axis
        If Abs(cosA) < 0.1 Then     ' member is within 6 degrees of parallel to surface
            GoTo ErrorHandler
        End If

        If cosA < 0 Then ' Make sure the plate's Z is towards the other end of the member. The plate is thickned along positive Z
            oZDir.x = -oZDir.x
            oZDir.y = -oZDir.y
            oZDir.z = -oZDir.z
        End If
        secDepth = secDepth / Abs(cosA) ' Depth projected on the plane
        SecWidth = SecWidth / Abs(cosA) ' Width projected on the plane
    End If
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSBasePlateAsmConn")
    End If
      

    DepthClr = pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn").Item("DepthClearance").Value
    WidthClr = pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn").Item("WidthClearance").Value
    sizeRule = pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn").Item("SizingRule").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, sizeRule, "StructACSizingRule", 20

    Set oSmartOccChild = pPropertyDescriptions.object
    Set pIJAttribsChild = oSmartOccChild
    
    If Not pIJAttribsChild Is Nothing Then
        If sizeRule <> 2 Then ' Depth and Width are calculated from clearances
            pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value = secDepth + 2 * DepthClr
            pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = SecWidth + 2 * WidthClr
        Else ' Clearance is calculated from depth and width
            
            If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim")) Then
            ' This probably called when the parent and child are just constructed and the child's def attributes are
            ' not copied to occ
                ' Copy defaults from parent
                depth = oAttrbs.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value
                Width = oAttrbs.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value
                'set defaults from AC on the plate
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value = depth
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
            Else
                depth = pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value
                Width = pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value
            End If
            
            DepthClr = (depth - secDepth) / 2#
            WidthClr = (Width - SecWidth) / 2#
            pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn").Item("DepthClearance").Value = DepthClr
            pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn").Item("WidthClearance").Value = WidthClr
        End If

    End If
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub


Public Sub CMComputeBasePlateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeBasePlateThickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMComputeBasePlatePosition
'
'Abstract
'Callback Method to re-evaluate the baseplate due to a change in the base
'plate's position property
'
'Arguments
'pPropertyDescriptions - collection of the property descriptions for the custom
'                        assembly defined at bulk load time
'                        (see IJDUserSymbolServices_InitializeSymbolDefinition).
'                        Contains the metadata descriptions of such
'                        properties as BasePlateSize.
'pObject - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMComputeBasePlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeBasePlatePosition"
    On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oSuppedPort As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim sX#, sY#, sZ#, eX#, eY#, eZ#
    Dim nx#, ny#, nz#
    Dim dX1 As Double, dY1 As Double, dX2 As Double, dY2 As Double
    Dim oShapeOcc As Object
    Dim oCrossService As New CrossSectionServices
    Dim oMat As IJDT4x4
    Dim oDeltaPos As New DPosition
    Dim oXDir As New DVector
    Dim oZDir As New DVector
    Dim oYDir As New DVector
    Dim oOriginPos As New DPosition
    Dim oOcc As IJDOccurrence
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppingPlane As IJPlane
    Dim oMembDir As New DVector
    Dim oSuppingPlaneNorm As New DVector
    Dim oPlane As IJPlane
    Dim oSurface As IJSurface
    Dim oGeom3DFactory As New GeometryFactory
    Dim oCP5Line As IJLine
    
    Dim code As Geom3dIntersectConstants
    Dim colElms As IJElements
    Dim oPoint As IJPoint
    Dim oStructPlate As IJStructPlate
    Dim thickness As Double
    Dim offsetPos As New DPosition
    Dim oRefColl As IJDReferencesCollection
    Dim bStartPoint As Boolean
   
    Set oOcc = pObject
    Set oStructPlate = pObject
    Set oSmartOccChild = pObject
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    oStructConn.enumPorts colPorts

    'Get input ports.
    If colPorts.count = 1 Then
        Set oSuppedPort = colPorts.Item(1)
        Set oRefColl = GetRefCollFromSmartOccurrence(oSmartOccCAO)
        On Error Resume Next
        ' Check if there is a plane in the ref coll;
        '  if there is one we use that for positioning the plate
        Set oSuppingPlane = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        On Error GoTo ErrorHandler
    
    ElseIf colPorts.count = 2 Then ' a plane is also input
        Set oSuppedPort = GetEndPort(colPorts)
        If oSuppedPort Is colPorts.Item(1) Then
            Set oSuppingPlane = colPorts.Item(2)
        Else
            Set oSuppingPlane = colPorts.Item(1)
        End If
    End If


    Set oSuppedPart = oSuppedPort.Part
    
    oSuppedPart.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
    
    'get the transformation matrix for the connected end
    If oSuppedPort.portIndex = SPSMemberAxisStart Then
        oSuppedPart.Rotation.GetTransformAtPosition sX, sY, sZ, oMat, Nothing
        ' Z dir of the plate is the tangent vector to the axis at the current end to the other end
        oMembDir.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oOriginPos.Set sX, sY, sZ
        bStartPoint = True
    
    Else
        oSuppedPart.Rotation.GetTransformAtPosition eX, eY, eZ, oMat, Nothing
        ' Z dir of the plate is the tangent vector to the axis at the current end to the other end
        oMembDir.Set -oMat.IndexValue(0), -oMat.IndexValue(1), -oMat.IndexValue(2)
        oOriginPos.Set eX, eY, eZ
        bStartPoint = False
    End If
    
    ' X direction of baseplate is the Z-axis of the member part
    oXDir.Set oMat.IndexValue(8), oMat.IndexValue(9), oMat.IndexValue(10)
    oXDir.length = 1 ' Normalize the vector
    oMembDir.length = 1 ' Normalize the vector
    
    
    ' Get rid of the translation component of the transformation matrix. It represents the start point of the member
    oMat.IndexValue(12) = 0
    oMat.IndexValue(13) = 0
    oMat.IndexValue(14) = 0
    
    Set oShapeOcc = oSuppedPart.CrossSection.Symbol
    oCrossService.GetCardinalPoint oShapeOcc, 5, dX1, dY1
    oCrossService.GetCardinalPoint oShapeOcc, oSuppedPart.CrossSection.CardinalPoint, dX2, dY2
    oDeltaPos.x = dX1 - dX2
    oDeltaPos.y = dY1 - dY2
    oDeltaPos.z = 0#
    ' Transform from cross section coordinate system to member coordinate system
    Set oMat = CreateCSToMembTransform(oMat)
    Set oDeltaPos = oMat.TransformPosition(oDeltaPos)
    
    
    If Not oSuppingPlane Is Nothing Then ' If a plane is input then the plate will be along that plane
        oSuppingPlane.GetNormal nx, ny, nz
        oSuppingPlaneNorm.Set nx, ny, nz
        If oSuppingPlaneNorm.Dot(oMembDir) < 0 Then ' Make sure the plate's Z is towards the other end of the member. The plate is thickned along positive Z
            oSuppingPlaneNorm.x = -oSuppingPlaneNorm.x
            oSuppingPlaneNorm.y = -oSuppingPlaneNorm.y
            oSuppingPlaneNorm.z = -oSuppingPlaneNorm.z
        End If
        Set oZDir = oSuppingPlaneNorm
        oZDir.length = 1
        oSuppingPlane.GetRootPoint sX, sY, sZ
        oOriginPos.Set sX, sY, sZ
    Else ' Otherwise the plate is normal to the member axis
        Set oZDir = oMembDir
        oZDir.length = 1
    End If
    ' Now create another plane by projecting the input plane by the thickness
    thickness = oStructPlate.thickness
    oZDir.length = thickness
    Set oOriginPos = oOriginPos.Offset(oZDir)
    oZDir.length = 1
    Set oSurface = oGeom3DFactory.Planes3d.CreateByPointNormal(Nothing, oOriginPos.x, oOriginPos.y, oOriginPos.z, oZDir.x, oZDir.y, oZDir.z)
    ' We are trying to get the top center of the plate
    Set oCP5Line = GetTangentLineAtCP(oSuppedPart, 5, bStartPoint)
    oCP5Line.Infinite = True
    oSurface.Intersect oCP5Line, colElms, code
    If Not colElms Is Nothing Then
        If colElms.count > 0 Then
            Set oPoint = colElms.Item(1)
            oPoint.GetPoint sX, sY, sZ
            ' Top center of the plate
            oOriginPos.Set sX, sY, sZ
        End If
    End If
    ' The y dir of the plate may not be in the positive Y direction of the member part
    Set oYDir = oZDir.Cross(oXDir)
    oYDir.length = 1
    Set oXDir = oYDir.Cross(oZDir)
    oXDir.length = 1
    
    
    oMat.LoadIdentity
    'X-axis of plate
    oMat.IndexValue(0) = oXDir.x
    oMat.IndexValue(1) = oXDir.y
    oMat.IndexValue(2) = oXDir.z
    'Y-axis of plate
    oMat.IndexValue(4) = oYDir.x
    oMat.IndexValue(5) = oYDir.y
    oMat.IndexValue(6) = oYDir.z
    'Z-axis of plate
    oMat.IndexValue(8) = oZDir.x
    oMat.IndexValue(9) = oZDir.y
    oMat.IndexValue(10) = oZDir.z
    offsetPos.Set 0, 0, -thickness
    
    ' So far we assumed the origin of plate at the top center
    ' now we offset the plate by its thickness in global coordinates
    Set offsetPos = oMat.TransformPosition(offsetPos)
    
    'Origin of plate
    oMat.IndexValue(12) = oOriginPos.x + offsetPos.x
    oMat.IndexValue(13) = oOriginPos.y + offsetPos.y
    oMat.IndexValue(14) = oOriginPos.z + offsetPos.z
    ' Set the matrix on the Baseplate occurrence
    oOcc.Matrix = oMat
    oSmartOccChild.Update ' Mark IJSmartOccurrence as modified so that the baseplate generation semantic fires
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConditionalBasePlateAsmConnCutbackPlane
'
'Abstract
'Callback Method to determine whether the current member is needed as an output.
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'bIsNeeded - boolean which is always returned as true to indicate that a
'            base plate plane is always needed.
'
'Exceptions
'
'***************************************************************************
Public Sub CMConditionalBasePlateAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalBasePlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    bIsNeeded = True
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConstructBasePlateAsmConnCutbackPlane
'
'Abstract
'Callback method for cnstructing the Base Plate Plane for the custom assmbly
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'pResourceManager - resource manager for the plant model database
'pObj - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMConstructBasePlateAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructBasePlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    Dim oGeomFactory As New GeometryFactory
    Dim oPlane As IJPlane
    Dim oBasePlate As IJStructPlate
    Dim oAssMembs As IJDMemberObjects
    Dim dblPlateThickness As Double
    Dim oBasePlateOcc As IJDOccurrence
    Dim oMat As IJDT4x4


    Dim oX#, oY#, oZ#, nx#, ny#, nz#, dx#, dy#, dz#
    Dim oControlFlags As IJControlFlags

    Set oAssMembs = pMemberDescription.CAO
    Set oBasePlate = oAssMembs.Item(1)
    
    Set oBasePlateOcc = oAssMembs.Item(1)
    Set oMat = oBasePlateOcc.Matrix
    
    dblPlateThickness = oBasePlate.thickness
    nx = oMat.IndexValue(8)
    ny = oMat.IndexValue(9)
    nz = oMat.IndexValue(10)
    
    oX = oMat.IndexValue(12)
    oY = oMat.IndexValue(13)
    oZ = oMat.IndexValue(14)
    
    
    dx = nx * dblPlateThickness
    dy = ny * dblPlateThickness
    dz = nz * dblPlateThickness

    Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(pResourceManager, oX + dx, oY + dy, oZ + dz, nx, ny, nz)
    ' Hide the plane
    Set oControlFlags = oPlane
    oControlFlags.ControlFlags(&H4) = &H4
    Set oControlFlags = Nothing
    
    Set pObj = oPlane ' Return the created object to the custom assembly
    
    Set oAssMembs = Nothing
    Set oBasePlateOcc = Nothing
    Set oBasePlate = Nothing
    Set oPlane = Nothing
    Set oGeomFactory = Nothing

   
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMSetInputBasePlateAsmConnPlate
'
'Abstract
'Callback Method used to set inputs on the parent SmartOccurrence prior to each CM evaluate
'
'Arguments
'pMemberDesc - Metadata describing the components of the custom
'              assembly which include the plate, cutback and
'              cutback plane. This metdata information was defined
'              at bulkload time via the
'              IJDUserSymbolServices_InitializeSymbolDefinition function.
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputBasePlateAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputBasePlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMMigrateBasePlateAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateBasePlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    
    'MsgBox METHOD
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'*************************************************************************
'Function
'CMComputeBasePlateAsmConnCutbackPlane
'
'Abstract
'Callback Method to notify the base plate connection that the cutback
'plane needs to be recomputed
'
'Arguments
'pPropertyDescriptions - collection of the property descriptions for the custom
'                        assembly defined at bulk load time
'                        (see IJDUserSymbolServices_InitializeSymbolDefinition).
'                        Contains the metadata descriptions of such
'                        properties as BasePlateSize.
'pObject - Base plate custom assembly's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMComputeBasePlateAsmConnCutbackPlane(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeBasePlateAsmConnCutbackPlane"
    On Error GoTo ErrorHandler

    Dim oPlane As IJPlane
    Dim oBasePlate As IJStructPlate
    Dim oAssMembs As IJDMemberObjects
    Dim dblPlateThickness As Double
    Dim oVec As New DVector
    Dim oX#, oY#, oZ#, nx#, ny#, nz#, dx#, dy#, dz#
    Dim oBasePlateOcc As IJDOccurrence
    Dim oMat As IJDT4x4


    Set oAssMembs = pPropertyDescriptions.CAO
    Set oBasePlate = oAssMembs.Item(1) ' the baseplate
    Set oBasePlateOcc = oAssMembs.Item(1)
    
    Set oMat = oBasePlateOcc.Matrix
    
    dblPlateThickness = oBasePlate.thickness
    nx = oMat.IndexValue(8)
    ny = oMat.IndexValue(9)
    nz = oMat.IndexValue(10)
    
    
    oX = oMat.IndexValue(12)
    oY = oMat.IndexValue(13)
    oZ = oMat.IndexValue(14)
    
    

    dx = nx * dblPlateThickness
    dy = ny * dblPlateThickness
    dz = nz * dblPlateThickness

    Set oPlane = pObject
    oPlane.SetRootPoint oX + dx, oY + dy, oZ + dz
    oPlane.SetNormal nx, ny, nz
    
   
    Set oAssMembs = Nothing
    Set oBasePlate = Nothing
    Set oBasePlateOcc = Nothing
    Set oMat = Nothing
    Set oPlane = Nothing
    Set oVec = Nothing
    
   
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

'*************************************************************************
'Function
'CMConditionalBasePlateAsmConnCutback
'
'Abstract
'Callback method to determine whether the current member is needed as an output.
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'bIsNeeded - boolean which is always returned as true to indicate that a
'            cutback is always needed.
'
'Exceptions
'
'***************************************************************************
Public Sub CMConditionalBasePlateAsmConnCutback(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalBasePlateAsmConnCutback"
    On Error GoTo ErrorHandler
    bIsNeeded = True
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMConstructBasePlateAsmConnCutback
'
'Abstract
'Callback method for constructing the Base Plate Plane for the custom assembly
'
'Arguments
'pMemberDescription - Metadata describing the components of the custom
'                     assembly which include the plate, cutback and
'                     cutback plane. This metdata information was defined
'                     at bulkload time via the
'                     IJDUserSymbolServices_InitializeSymbolDefinition function.
'pResourceManager - resource manager for the plant model database
'pObj - Base plate assembly connection's Smart Occurrence instance
'
'Exceptions
'
'***************************************************************************
Public Sub CMConstructBasePlateAsmConnCutback(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructBasePlateAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim oFeatureFactory As New StructFeatureFactory
    
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oSmartOccFac As IJSmartOccurrenceFactory
    Dim oDesignParent As IJDesignParent
    Dim oCutbackPlane As IJPlane
    Dim oPlate As IJStructCustomPlatePart
    Dim oAssemblyMembs As IJDMemberObjects
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSuppedPort As ISPSSplitAxisEndPort
    
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    oStructConn.enumPorts colPorts
    Set oSuppedPort = GetEndPort(colPorts)

    Set oSmartOccFac = New SmartOccurrenceFactory
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)

    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    ' Generate name
    GenerateNameForFeature oSOChild

    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTPlanarCutbackFeatureOcc"
    oSOChild.RootSelection = "PlanarCutbackFeature"
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)
    
    Set oAssemblyMembs = oSOParent
    
    Set oCutbackPlane = oAssemblyMembs.Item(2)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSuppedPort, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oCutbackPlane, IJPlane, "SPSSuppPlaneToRC_DEST"
    ' Connect the Smart Occurrence to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    
    Set pObj = oSOChild ' Return the newly created object
    
    Set oAssemblyMembs = Nothing
    Set oPlate = Nothing
    Set oRefCollChild = Nothing
    Set oSymFactory = Nothing
    Set oSOParent = Nothing
    Set oSOChild = Nothing
    Set iUserType = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMSetInputBasePlateAsmConnCutback
'
'Abstract
'Callback method used to set inputs on the parent SmartOccurrence prior to each CM evaluate
'
'Arguments
'pMemberDesc - Metadata describing the components of the custom
'              assembly which include the plate, cutback and
'              cutback plane. This metdata information was defined
'              at bulkload time via the
'              IJDUserSymbolServices_InitializeSymbolDefinition function.
'
'Exceptions
'
'***************************************************************************
Public Sub CMSetInputBasePlateAsmConnCutback(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputBasePlateAsmConnCutback"
    On Error GoTo ErrorHandler
    Dim pIJAttribsChild    As IJDAttributes

  
    Set pIJAttribsChild = pMemberDesc.object
    If Not IsSOOverridden(pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback")) Then
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("Clearance").Value = 0#
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("SquaredEnd").Value = False
    End If
    Set pIJAttribsChild = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMMigrateBasePlateAsmConnCutback(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateBasePlateAsmConnCutback"
    On Error GoTo ErrorHandler
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Dim pReferencesCollection As IJDReferencesCollection
    Dim bIsInputMigrated As Boolean
    Dim oPoint As IJPoint
    Dim ii As Integer, eleCount As Integer
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    
    Dim oOld As Object
    Dim oReplacing() As Object

    'MsgBox METHOD

    Set pSmartOccurrence = pMemberDesc.object
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)
    
    GetPositionFromRefColl pReferencesCollection, oPoint

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    ReDim oReplacing(1 To eleCount)

    For ii = 1 To eleCount

        Set oOld = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        
        Call pMigrateHelper.ObjectsReplacing(oOld, pObjectCollectionReplacing, bIsDeleted)
    
        If Not pObjectCollectionReplacing Is Nothing Then
            bIsInputMigrated = True
            SelectReplacingObject pObjectCollectionReplacing, oPoint, oReplacing(ii)
        Else
            Set oReplacing(ii) = oOld
        End If
        
        Set oOld = Nothing
        Set pObjectCollectionReplacing = Nothing
        
    Next ii
     
    If bIsInputMigrated Then
        Call pReferencesCollection.IJDEditJDArgument.RemoveAll
        pReferencesCollection.IJDEditJDArgument.SetEntity 1, oReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), IJPlane, "SPSSuppPlaneToRC_DEST"
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'*************************************************************************
'Function
'CMComputeBasePlateAsmConnCutback
'
'Abstract
'Callback Method to notify the base plate connection that the cutback
'needs to be recomputed
'
'Arguments
'pPropertyDescriptions - collection of the property descriptions for the custom
'                        assembly defined at bulk load time
'                        (see IJDUserSymbolServices_InitializeSymbolDefinition).
'                        Contains the metadata descriptions of such
'                        properties as BasePlateSize.
'pObject - Base plate custom assembly's Smart Occurrence instance
'
'Exceptions
' E_FAIL -raised when any of the input ports are missing which
'           forces the baseplate to the ToDo list
'
'***************************************************************************
Public Sub CMComputeBasePlateAsmConnCutback(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeBasePlateAsmConnCutback"
    On Error GoTo ErrorHandler
    
    'Nothing to compute
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

'*************************************************************************
'Function
'IJDUserSymbolServices
'   InstanciateDefinition
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
    
  ' Returned symbol definition
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function
ErrorHandler:
  HandleError MODULE, MT
End Function


'*************************************************************************
'Function
'IJDUserSymbolServices
'   GetDefinitionName
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


'*************************************************************************
'Function
'
'IJDUserSymbolServices
'   InvokeRepresentation
'
'Abstract
'Obsolete method
'
'***************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub


'*************************************************************************
'Function
'IJDUserSymbolServices_EditOccurence
'
'Abstract
'Obsolete method --instead you can record your custom command within the
'definition (see IJDCommandDescription interface)
'
'***************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function


'*************************************************************************
'Function
'IJStructCustomFoulCheck
'   GetConnectedParts
'
'Abstract
'Adds objects that are related with the SmartOccurrence ( input objects ) to the list of connected objects
'
'Arguments
'pPartObject is the SmartOccurrence object
'pIJMonUnks is the list of monikers
'
'Return
'
'Exceptions
'
'***************************************************************************
Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
End Sub


'*************************************************************************
'Function
'IJStructCustomFoulCheck
'   GetFoulInterfaceType
'
'Abstract
'Returns type of interference participant that this object is.
'
'Arguments
'pFoulInterfaceType is the InterferenceChecking type
'
'Return
'
'Exceptions
'
'***************************************************************************
Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnAttributeChange
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
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
End Function


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnPreCommit
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
Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
End Function


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnPreLoad
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
Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const MT = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler


Exit Function
ErrorHandler:  HandleError MODULE, MT
End Function


'*************************************************************************
'Function
'ISPSFACInputHelper
'   ValidateObjects
'
'Abstract
'Validates the given objects and returns the objects to be used for assemblyConnection inputs
'
'Arguments
'inputObjs is a collection of objects
'relationObjs is the collection to be used to build relationships
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
End Function




Private Function IJUserAttributeMgmtParent_OnAttributeChange(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String

End Function

Private Function IJUserAttributeMgmtParent_OnPreCommit(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmtParent_OnPreLoad(ByVal pParent As Object, ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

    'get the parent
    Dim oParent As IJDMemberObjects
    Dim oAttribMgmntChild As IJUserAttributeMgmtParent
    Dim oSmartItem As IJSmartItem
    Dim ii As Long
    Dim oChild As Object
    
    Set oParent = pParent
    
    If Not oParent Is Nothing Then
       
        Set oChild = oParent.ItemByDispid(1)
        If oChild Is pIJDAttrs Then
            Dim pIJAttribsCAO As IJDAttributes
            Dim lngSizeRule As Long
            

            Set pIJAttribsCAO = oParent
            lngSizeRule = pIJAttribsCAO.CollectionOfAttributes("IJUASPSBasePlateAsmConn").Item("SizingRule").Value
            Dim pAttrColl As Collection
            Set pAttrColl = CollAllDisplayedValues
            Dim pAttrDescr As IJAttributeDescriptor
            For ii = 1 To pAttrColl.count
                Set pAttrDescr = pAttrColl.Item(ii)
                If pAttrDescr.InterfaceName = "IJUASPSPlatePartDim" Then
                    If (lngSizeRule = 1) Then 'By Rule
                        'controlling  child's property
                        'make property read only in plate's GUI
                        pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                    End If
                End If
            Next

        End If
    End If
End Function

'*************************************************************************
'Function
'ISPSFACInputHelper
'   UserAttributeMgmt
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
Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
End Property


Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal inputObjs As SP3DStructInterfaces.IJElements, relationObjs As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const MT = "ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler

    Dim oInputObj1 As Object
    Dim oInputObj2 As Object
    Dim count As Integer
    Dim portCol As IJElements
    
    ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_Ok
    Set portCol = New JObjectCollection
    'filter out ports to portCol
    For count = 1 To inputObjs.count
        If TypeOf inputObjs.Item(count) Is IJPort Then
            portCol.Add inputObjs.Item(count)
        End If
    Next count
    
    
    If portCol.count = 1 Then
        Set oInputObj1 = portCol.Item(1)
        If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            If relationObjs Is Nothing Then
                Set relationObjs = New JObjectCollection
            End If
            relationObjs.Clear
            relationObjs.Add oInputObj1
            'check if there is a plane  in the inputs
            'the baseplate AC takes the plane as an optional input
            For count = 1 To inputObjs.count
                If TypeOf inputObjs.Item(count) Is IJPlane Then
                    relationObjs.Add inputObjs.Item(count)
                    Exit For
                End If
            Next count
        Else
            ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
        End If
    ElseIf portCol.count = 2 Then
        Set oInputObj1 = portCol.Item(1)
        Set oInputObj2 = portCol.Item(2)
        If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            If TypeOf oInputObj2 Is IJPlane Then
                If relationObjs Is Nothing Then
                    Set relationObjs = New JObjectCollection
                End If
                relationObjs.Clear
                relationObjs.Add oInputObj1
                relationObjs.Add oInputObj2
            Else
                ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
            End If
        Else
            ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
        End If
    Else
        ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
    End If
    Exit Property
    
ErrorHandler:
    HandleError MODULE, MT
End Property

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
