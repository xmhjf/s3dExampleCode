VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitreAsmConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : MitreAsmConnDef.cls
'
'Author : RP
'
'Description :
'    Macro for creating Mitre   connection
'
'History:
'
' 08/14/03   RP   Substituted interface name in place of guid in IJDPropertyDescriptions->AddProperty.
'                 The guid of a virtual interface may change during bulkload
' 02/02/04  JS    Added code to verify the assembly connection and check if an AC already exists
'                 between the supported or supporting members TR#52040
' 02/05/04  JS    Changed all generic reference collection relations to specific relations.
'                 This includes the relations to the split axis end ports, split axis along ports
'                 and the cutback planes (if a plane exists).
'02/16/2004 RP    Added code to position the plate based on top, bottom
'                 left,right distances. TR#54134
' 08/27/04  MH    CMMigrate
' 06/13/06  RP   Changes due to impact from curved members(DI#84001)
' 27/Sep/06 AS   TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
' 08/07/07  MH   TR#121116 added CheckForUndefinedValueAndRaiseError
'12/May/14 NK   TR-CP-253696  ATP - 26 minidumps at 'SPSACMacros' module : Adding checks to avoid access of Null objects
' 18/06/14  knukala CR-CP-223345  Rules should not replace root selector inputs
'*************************************************************************************

Option Explicit


Private Const MODULE = "MitreAsmConnDef"

Const m_ItemProgId As String = "SPSACMacros.MitreAsmConnDef"
Const m_ItemName As String = "SPSACMacros.MitreAsmConnDef"
Private Const strSourceFile As String = "MitreAsmConnDef.cls"
Private m_oLocalizer As IJLocalizer
Implements ISPSFACInputHelper
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Public Sub DefinitionInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    pIH.ReportError
    End Sub


Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
    Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    
    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' Define the inputs -
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
    'listens to IJUASPSMitreAsmConn
    pAPDs.AddProperty "MitreAsmConnProperties", 1, "IJUASPSMitreAsmConn", "CMComputeMitreAsmConnProperties", imsCOOKIE_ID_USS_LIB
    'listens to IJUASPSPlateTypeCategory
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
    'Plate1 SO
    Set pMemberDescription = pMemberDescriptions.AddMember("MitreAsmConnPlate1", 1, "CMConstructMitreAsmConnPlate1", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
    pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalMitreAsmConnPlate1"
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputMitreAsmConnPlate1"
    pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateMitreAsmConnPlate1"
    Set pPropertyDescriptions = pMemberDescription
    'outputs  IJUASPSPlatePartDim.
    pPropertyDescriptions.AddProperty "MitrePlate1Size", 1, "IJUASPSPlatePartDim", "CMComputeMitrePlate1Size", imsCOOKIE_ID_USS_LIB
    'outputs IJStructPlate.
    pPropertyDescriptions.AddProperty "MitrePlate1Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeMitrePlate1Thickness", imsCOOKIE_ID_USS_LIB
    'outputs IJSmartOccurrence.
    pPropertyDescriptions.AddProperty "MitrePlate1Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeMitrePlate1Position", imsCOOKIE_ID_USS_LIB
    '_________________________________________________________________________________________________________________________________________
    'Plate2 SO
    Set pMemberDescription = pMemberDescriptions.AddMember("MitreAsmConnPlate2", 2, "CMConstructMitreAsmConnPlate2", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
    pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalMitreAsmConnPlate2"
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputMitreAsmConnPlate2"
    pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateMitreAsmConnPlate2"
    Set pPropertyDescriptions = pMemberDescription
    'outputs  IJUASPSPlatePartDim.
    pPropertyDescriptions.AddProperty "MitrePlate2Size", 1, "IJUASPSPlatePartDim", "CMComputeMitrePlate2Size", imsCOOKIE_ID_USS_LIB
    'outputs IJStructPlate.
    pPropertyDescriptions.AddProperty "MitrePlate2Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeMitrePlate2Thickness", imsCOOKIE_ID_USS_LIB
    'outputs IJSmartOccurrence.
    pPropertyDescriptions.AddProperty "MitrePlate2Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeMitrePlate2Position", imsCOOKIE_ID_USS_LIB
    
    '_________________________________________________________________________________________________________________________________________
    'Cutback Plane
    Set pMemberDescription = pMemberDescriptions.AddMember("MitreAsmConnCutbackPlane", 3, "CMConstructMitreAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
    pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalMitreAsmConnCutbackPlane"
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputMitreAsmConnCutbackPlane"
    pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateMitreAsmConnCutbackPlane"
    Set pPropertyDescriptions = pMemberDescription
    'outputs IJPlane.
    pPropertyDescriptions.AddProperty "ComputeMitreAsmConnCutbackPlane", 1, "{4317C6B3-D265-11D1-9558-0060973D4824}", "CMComputeMitreAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB
    
    '_______________________________________________________________________________________________________________________________________
    'Cutback1 SO
    Set pMemberDescription = pMemberDescriptions.AddMember("MitreAsmConnCutback1", 4, "CMConstructMitreAsmConnCutback1", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
    pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalMitreAsmConnCutback1"
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputMitreAsmConnCutback1"
    pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateMitreAsmConnCutback1"
    Set pPropertyDescriptions = pMemberDescription
    'outputs IJSmartOccurrence. This is just to make sure that the parent computes before the child.
    pPropertyDescriptions.AddProperty "ComputeMitreAsmConnCutback1", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeMitreAsmConnCutback1", imsCOOKIE_ID_USS_LIB
    '_______________________________________________________________________________________________________________________________________
    'Cutback2 SO
    Set pMemberDescription = pMemberDescriptions.AddMember("MitreAsmConnCutback2", 5, "CMConstructMitreAsmConnCutback2", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
    pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalMitreAsmConnCutback2"
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputMitreAsmConnCutback2"
    pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateMitreAsmConnCutback2"
    Set pPropertyDescriptions = pMemberDescription
    'outputs IJSmartOccurrence. This is just to make sure that the parent computes before the child.
    pPropertyDescriptions.AddProperty "ComputeMitreAsmConnCutback2", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeMitreAsmConnCutback2", imsCOOKIE_ID_USS_LIB
    '___________________________________________________________________________________________________________________________________
    Set pMemberDescriptions = Nothing
    Set pMemberDescription = Nothing
    Set pPropertyDescriptions = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
'Aggregator custom methods.......................................................
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructAggregator"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMComputeMitreAsmConnProperties(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMComputeMitreAsmConnProperties"
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

'Plate1 SO custom methods............................................................................................
Public Sub CMConditionalMitreAsmConnPlate1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalMitreAsmConnPlate1"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim bWithPlate As Boolean
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim strError As String
      
    Set oSmartOccCAO = pMemberDescription.CAO
    
    'retrieve the inputs of the custom assembly occurrence
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)

    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped1Port Is Nothing Or oSupped2Port Is Nothing Then
        Dim iJDObject As iJDObject
        Set iJDObject = oStructConn
        iJDObject.Remove
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_INPUTS_MITRED, oStructConn, Nothing
        Err.Raise E_FAIL
    End If
    
    Set oSupped1Part = oSupped1Port.Part
    Set oSupped2Part = oSupped2Port.Part
      
    If Not (ValidSectionType(oSupped1Part) And ValidSectionType(oSupped2Part)) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_UNKNOWNSEC_ENCOUNT, oStructConn, Nothing
        Err.Raise E_FAIL
        
    End If
    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed and the assembly connection added to the ToDo list
    If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
        Err.Raise E_FAIL
        
    End If
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSMitreAsmConn")) Then
      'This is executed when the parent and child are just constructed
        Set oAttrbs = oSmartOccCAO.ItemObject
          'copy defaults from parent
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSMitreAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSMitreAsmConn")
    End If
    
    bWithPlate = pIJAttribsCAO.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("WithEndPlates").Value
      
    If bWithPlate = True Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    
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

Public Sub CMConstructMitreAsmConnPlate1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructMitreAsmConnPlate1"
    On Error GoTo ErrorHandler
    
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl As IJDReferencesCollection
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSymbolOfCAO As IJDSymbol
    Dim oMembObjects As IJDMemberObjects
    Dim oDesignParent As IJDesignParent
    Dim PlateThickness As Double
    Dim oAttrbs As IJDAttributes
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oMembObjects = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    
    
    Set oFactory = New StructCustomPlatePartFactory
    'Create the custom Plate that is the Base Plate.
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
    

    'get default from the parent
    PlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, PlateThickness
    'set the found thickness on attrbs
    '    oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value = PlateThickness
    
    Set oSmartOcc = oPlate
    ' Initialize the SmartOccurrence
    oSmartOcc.RootSelectorClass = "GenericRectPlatePart"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)
    
    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefColl
    
    Set pObj = oPlate 'return the created plate object to the custom assembly

    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMSetInputMitreAsmConnPlate1(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputMitreAsmConnPlate1"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oCustomPlate As StructCustomPlatePart
    
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    On Error Resume Next 'suppress the error that VB throws when QI fails
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
        
        'copy plate category and type from parent to child
        oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Category").Value
        oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Type").Value
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitrePlate1Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitrePlate1Size"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitrePlate1Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitrePlate1Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitrePlate1Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitrePlate1Position"
    On Error GoTo ErrorHandler
    
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oSurf As IJSurface
    Dim oGeom3DFactory As New GeometryFactory
    Dim iPort As SPSMemberAxisPortIndex
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim oNorm As IJDVector
    Dim oLine(1 To 4) As IJLine
    Dim idx As Integer
    Dim code As Geom3dIntersectConstants
    Dim colElms As IJElements
    Dim oPoint(1 To 4) As IJPoint
    Dim length As Double, Width As Double
    Dim topClr#, botClr#, leftClr#, rightClr#
    Dim oAttrbParent As IJDAttributes
    Dim oAttrbChild As IJDAttributes
    Dim xVec As IJDVector, yVec As IJDVector
    Dim zVec As New DVector
    Dim oPlate As IJStructPlate
    Dim thickness As Double
    Dim cosA As Double
    Dim oPos As New DPosition, oPlateOrg As New DPosition, oEndPos As New DPosition
    Dim dist As Double
    Dim oMat As New DT4x4
    Dim oOcc As IJDOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oSmartOcc As IJSmartOccurrence
    Dim align As Long
    Dim clearance As Double
    Dim idPort As SPSMemberAxisPortIndex
    Dim sizeRule As Long
    Dim bStartPos As Boolean

    
    Set oOcc = pObject
    Set oSmartOccChild = pObject
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set oAttrbParent = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
     ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If (oSupped1Port Is Nothing) Or (oSupped2Port Is Nothing) Then
        Err.Raise E_FAIL
    End If
    Set oSupped1Part = oSupped1Port.Part
    idPort = oSupped1Port.portIndex
    
    'Get normal as the bisector of member axes
    Set oNorm = GetMiterPlaneNormal(oSupped1Port, oSupped2Port)
    
    oSupped1Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    'it is assumed that both ports are at the same point
    If idPort = SPSMemberAxisStart Then
        oSupped1Part.Rotation.GetTransformAtPosition X0, Y0, Z0, oMat, Nothing
        zVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    ElseIf idPort = SPSMemberAxisEnd Then
        oSupped1Part.Rotation.GetTransformAtPosition x1, y1, z1, oMat, Nothing
        zVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        zVec.[Scale] -1
    End If
    zVec.length = 1
    
    'Find position at most "exterior" intersection of member notional boxes
    Set oPos = GetMiterPlanePosition(oSupped1Port, oSupped2Port)
    If Not oPos Is Nothing Then
        oPos.Get X0, Y0, Z0
    Else
        'it is assumed that both ports are at the same point
        oSupped1Part.PointAtEnd(idPort).GetPoint X0, Y0, Z0
    End If

    topClr = oAttrbParent.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("TopDist").Value
    botClr = oAttrbParent.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("BottDist").Value
    leftClr = oAttrbParent.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("LeftDist").Value
    rightClr = oAttrbParent.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("RightDist").Value
    align = oAttrbParent.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("Symmetry").Value
    clearance = oAttrbParent.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("clearance").Value
    sizeRule = oAttrbParent.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("SizingRule").Value
    
    CheckForUndefinedValueAndRaiseError oSmartOcc, align, "StructAlignment", 21
    CheckForUndefinedValueAndRaiseError oSmartOcc, sizeRule, "StructACSizingRule", 20

    Set oAttrbChild = pObject 'get child's property
    If IsSOOverridden(oAttrbChild.CollectionOfAttributes("IJUASPSPlatePartDim")) Then
        length = oAttrbChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value
        Width = oAttrbChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value
    Else
        Set oAttrbParent = oSmartOcc.ItemObject
        length = oAttrbParent.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value
        Width = oAttrbParent.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value
    End If
    
    Set oPlate = pObject
    thickness = oPlate.thickness
    If align = 1 Then 'center
        dist = thickness / 2
    ElseIf align = 2 Then 'left, only supporte1 is cutback
        dist = thickness + clearance
    ElseIf align = 3 Then 'right , only supported2 is cutback
        dist = -clearance
    End If
    cosA = oNorm.Dot(zVec)
    
    If cosA > 0 Then
        oPos.x = oNorm.x * dist
        oPos.y = oNorm.y * dist
        oPos.z = oNorm.z * dist
        oNorm.x = -oNorm.x
        oNorm.y = -oNorm.y
        oNorm.z = -oNorm.z
    Else
        oPos.x = -oNorm.x * dist
        oPos.y = -oNorm.y * dist
        oPos.z = -oNorm.z * dist
    End If
    
    oPlateOrg.Set X0 + oPos.x, Y0 + oPos.y, Z0 + oPos.z
    
    Set oSurf = oGeom3DFactory.Planes3d.CreateByPointNormal(Nothing, oPlateOrg.x, oPlateOrg.y, oPlateOrg.z, oNorm.x, oNorm.y, oNorm.z)
    
    oSupped1Part.PointAtEnd(idPort).GetPoint x, y, z
    
    oEndPos.Set x, y, z
    
    If idPort = SPSMemberAxisStart Then
        bStartPos = True
    End If
    
    Set oLine(1) = GetTangentLineAtCP(oSupped1Part, 5, bStartPos) ' center of the cross section range box
    oLine(1).Infinite = True
    Set oLine(2) = GetTangentLineAtCP(oSupped1Part, 1, bStartPos)
    oLine(2).Infinite = True
    Set oLine(3) = GetTangentLineAtCP(oSupped1Part, 3, bStartPos)
    oLine(3).Infinite = True
    Set oLine(4) = GetTangentLineAtCP(oSupped1Part, 7, bStartPos)
    oLine(4).Infinite = True
    
    For idx = 1 To 4
        oSurf.Intersect oLine(idx), colElms, code
        If colElms.count > 0 Then
            Set oPoint(idx) = colElms.Item(1)
        End If
    Next idx
 
    
    
    Set yVec = GetYVector(oSupped1Part, oEndPos)
    Set xVec = yVec.Cross(oNorm)
    xVec.length = 1
    Set yVec = oNorm.Cross(xVec)
    yVec.length = 1
    oMat.LoadIdentity
    'X-axis of plate
    oMat.IndexValue(0) = xVec.x
    oMat.IndexValue(1) = xVec.y
    oMat.IndexValue(2) = xVec.z
    'Y-axis of plate
    oMat.IndexValue(4) = yVec.x
    oMat.IndexValue(5) = yVec.y
    oMat.IndexValue(6) = yVec.z
    'Z-axis of plate
    oMat.IndexValue(8) = oNorm.x
    oMat.IndexValue(9) = oNorm.y
    oMat.IndexValue(10) = oNorm.z
    
    oPlateOrg.Set 0, 0, 0
    If sizeRule = 1 Then ' calculate plate size from clearnce
        length = oPoint(4).DistFromPt(oPoint(2))
        Width = oPoint(3).DistFromPt(oPoint(2))

        length = length + topClr + botClr
        Width = Width + leftClr + rightClr
        'plate needs to be offset based on the clearances
        'the offset is in plate's local coord system
        oPlateOrg.x = (topClr - botClr) / 2
        oPlateOrg.y = (leftClr - rightClr) / 2
        oPlateOrg.z = 0
        'transform offset to global
        Set oPlateOrg = oMat.TransformPosition(oPlateOrg)
        oAttrbChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value = length
        oAttrbChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
    End If
    
    'now apply this offset to plate's origin,which is at CP5 of supped1
    oPoint(1).GetPoint X0, Y0, Z0
    oPlateOrg.x = oPlateOrg.x + X0
    oPlateOrg.y = oPlateOrg.y + Y0
    oPlateOrg.z = oPlateOrg.z + Z0
    
    'set Origin of plate on the matrix
    oMat.IndexValue(12) = oPlateOrg.x
    oMat.IndexValue(13) = oPlateOrg.y
    oMat.IndexValue(14) = oPlateOrg.z
    
    'set the matrix on the Baseplate occurrence
    oOcc.Matrix = oMat
    oSmartOccChild.Update ' mark IJSmartOccurrence as modified so that the baseplate generation semantic fires
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMMigrateMitreAsmConnPlate1(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateMitreAsmConnPlate1"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub
'Plate2 SO custom methods............................................................................................
Public Sub CMConditionalMitreAsmConnPlate2(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalMitreAsmConnPlate2"
    On Error GoTo ErrorHandler
    
    bIsNeeded = False
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMConstructMitreAsmConnPlate2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructMitreAsmConnPlate2"
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMSetInputMitreAsmConnPlate2(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputMitreAsmConnPlate2"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oCustomPlate As StructCustomPlatePart
    
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    On Error Resume Next 'suppress the error that VB throws when QI fails
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
        
        'copy plate category and type from parent to child
        oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Category").Value
        oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateTypeCategory").Item("Type").Value
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitrePlate2Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitrePlate2Size"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitrePlate2Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitrePlate2Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitrePlate2Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitrePlate2Position"
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMMigrateMitreAsmConnPlate2(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateMitreAsmConnPlate2"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

'Cutback Plane custom methods......................................................................................................
Public Sub CMConditionalMitreAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalMitreAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    bIsNeeded = True
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMConstructMitreAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructMitreAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oPlane As IJPlane
    Dim oGeom3DFactory As New GeometryFactory
    Dim iPort As SPSMemberAxisPortIndex
    Dim x As Double, y As Double, z As Double
    Dim oNorm As IJDVector
    Dim oPos As IJDPosition
    Dim idPort As SPSMemberAxisPortIndex
    Dim oControlFlags As IJControlFlags
    
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
     ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If (oSupped1Port Is Nothing) Or (oSupped2Port Is Nothing) Then
        Err.Raise E_FAIL
    End If
    
    idPort = oSupped1Port.portIndex
    Set oSupped1Part = oSupped1Port.Part
    Set oNorm = GetMiterPlaneNormal(oSupped1Port, oSupped2Port)
    Set oPos = GetMiterPlanePosition(oSupped1Port, oSupped2Port)
    If Not oPos Is Nothing Then
        oPos.Get x, y, z
    Else
        'it is assumed that both ports are at the same point
        oSupped1Part.PointAtEnd(idPort).GetPoint x, y, z
    End If
    Set oPlane = oGeom3DFactory.Planes3d.CreateByPointNormal(pResourceManager, x, y, z, oNorm.x, oNorm.y, oNorm.z)
     'Hide the plane
    Set oControlFlags = oPlane
    oControlFlags.ControlFlags(&H4) = &H4
    
    Set pObj = oPlane ' return the plane to custom assembly
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMSetInputMitreAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputMitreAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitreAsmConnCutbackPlane(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitreAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oPlane As IJPlane
    Dim oGeom3DFactory As New GeometryFactory
    Dim iPort As SPSMemberAxisPortIndex
    Dim x As Double, y As Double, z As Double
    Dim oNorm As IJDVector
    Dim oPos As IJDPosition
    Dim idPort As SPSMemberAxisPortIndex
    
    Set oPlane = pObject ' get the plane
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If (oSupped1Port Is Nothing) Or (oSupped2Port Is Nothing) Then
        Err.Raise E_FAIL
    End If

    Set oSupped1Part = oSupped1Port.Part
    idPort = oSupped1Port.portIndex
    Set oNorm = GetMiterPlaneNormal(oSupped1Port, oSupped2Port)
    oPlane.SetNormal oNorm.x, oNorm.y, oNorm.z
    Set oPos = GetMiterPlanePosition(oSupped1Port, oSupped2Port)
    If Not oPos Is Nothing Then
        oPos.Get x, y, z
    Else
        'it is assumed that both ports are at the same point
        oSupped1Part.PointAtEnd(idPort).GetPoint x, y, z
    End If
    oPlane.SetRootPoint x, y, z
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMMigrateMitreAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateMitreAsmConnCutbackPlane"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

'Cutback1 SO custom methods
Public Sub CMConditionalMitreAsmConnCutback1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalMitreAsmConnCutback1"
    On Error GoTo ErrorHandler
    bIsNeeded = True
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMConstructMitreAsmConnCutback1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructMitreAsmConnCutback1"
    On Error GoTo ErrorHandler
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oDesignParent As IJDesignParent
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As IJPort
    Dim oSupped2Port As IJPort
    Dim oMembObjs As IJDMemberObjects
    Dim oFeatureFactory As New StructFeatureFactory
    
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Set oMembObjs = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
     ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped1Port Is Nothing Then
        Err.Raise E_FAIL
    End If

 
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    'create name
    GenerateNameForFeature oSOChild

    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTPlanarCutbackFeatureOcc"
    oSOChild.RootSelection = "PlanarCutbackFeature"
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSupped1Port, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oMembObjs.Item(3), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane

    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    Set pObj = oSOChild ' return the newly created object
  
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMSetInputMitreAsmConnCutback1(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputMitreAsmConnCutback1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Sub CMComputeMitreAsmConnCutback1(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitreAsmConnCutback1"
    On Error GoTo ErrorHandler
    
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim pltThickness As Double
    Dim oPlate As IJStructPlate
    Dim oAssmblyChildren As IJDMemberObjects
    Dim align As Long
    Dim clearance As Double, clearance1 As Double, clearance2 As Double
      
    Set oAssmblyChildren = pPropertyDescriptions.CAO
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    Set oPlate = oAssmblyChildren.Item(1)
    
    Set pIJAttribsChild = pPropertyDescriptions.object
    If Not pIJAttribsChild Is Nothing Then
        align = pIJAttribsCAO.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("Symmetry").Value
        clearance = pIJAttribsCAO.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("Clearance").Value
        clearance1 = clearance
        clearance2 = clearance
    
        CheckForUndefinedValueAndRaiseError oSmartOccCAO, align, "StructAlignment", 21

        If Not oPlate Is Nothing Then ' plate is a conditional member so may not exist
            pltThickness = oPlate.thickness 'oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value
            clearance2 = 2 * clearance2
        Else
            pltThickness = 0
            clearance1 = clearance1 / 2
        End If

        If align = 1 Then 'center
            clearance = pltThickness / 2 + clearance1
        ElseIf align = 2 Then 'left, only supporte1 is cutback
            clearance = pltThickness + clearance2
        ElseIf align = 3 Then 'right , only supported2 is cutback
            clearance = 0
        End If
        
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("Clearance").Value = clearance
    End If
    
    Set oAttrbs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIJAttribsCAO = Nothing
    Set pIJAttribsChild = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMMigrateMitreAsmConnCutback1(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateMitreAsmConnCutback1"
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
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMConditionalMitreAsmConnCutback2(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    Const MT = "CMConditionalMitreAsmConnCutback2"
    On Error GoTo ErrorHandler
    bIsNeeded = True
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMConstructMitreAsmConnCutback2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const MT = "CMConstructMitreAsmConnCutback2"
    On Error GoTo ErrorHandler
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oDesignParent As IJDesignParent
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As IJPort
    Dim oSupped2Port As IJPort
    Dim oMembObjs As IJDMemberObjects
    Dim oFeatureFactory As New StructFeatureFactory
  
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Set oMembObjs = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
     ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped2Port Is Nothing Then
       Err.Raise E_FAIL
    End If

 
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    'create name
    GenerateNameForFeature oSOChild

    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTPlanarCutbackFeatureOcc"
    oSOChild.RootSelection = "PlanarCutbackFeature"
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSupped2Port, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oMembObjs.Item(3), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane

    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    Set pObj = oSOChild ' return the newly created object
  
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMSetInputMitreAsmConnCutback2(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputMitreAsmConnCutback2"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

Public Sub CMComputeMitreAsmConnCutback2(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeMitreAsmConnCutback2"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim pltThickness As Double
    Dim oPlate As IJStructPlate
    Dim oAssmblyChildren As IJDMemberObjects
    Dim align As Long
    Dim clearance As Double, clearance1 As Double, clearance2 As Double
    
    Set oAssmblyChildren = pPropertyDescriptions.CAO
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    Set oPlate = oAssmblyChildren.Item(1)
    
    Set pIJAttribsChild = pPropertyDescriptions.object
    If Not pIJAttribsChild Is Nothing Then
        align = pIJAttribsCAO.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("Symmetry").Value
        clearance = pIJAttribsCAO.CollectionOfAttributes("IJUASPSMitreAsmConn").Item("Clearance").Value
        clearance1 = clearance
        clearance2 = clearance

        CheckForUndefinedValueAndRaiseError oSmartOccCAO, align, "StructAlignment", 21

        If Not oPlate Is Nothing Then ' plate is a conditional member so may not exist
            pltThickness = oPlate.thickness 'oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value
            clearance2 = 2 * clearance2
        Else
            pltThickness = 0
            clearance1 = clearance1 / 2
        End If

        If align = 1 Then 'center
            clearance = pltThickness / 2 + clearance1
        ElseIf align = 2 Then 'left, only supporte1 is cutback
            clearance = 0
        ElseIf align = 3 Then 'right , only supported2 is cutback
            clearance = pltThickness + clearance2
        End If
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("Clearance").Value = clearance
    End If
    
    Set oAttrbs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIJAttribsCAO = Nothing
    Set pIJAttribsChild = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMMigrateMitreAsmConnCutback2(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateMitreAsmConnCutback2"
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
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

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
    
ErrorHandler:
      HandleError MODULE, MT
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


Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)

End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String

End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String

End Function

Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
    Set ISPSFACInputHelper_UserAttributeMgmt = Me
End Property

Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal inputObjs As SP3DStructInterfaces.IJElements, relationObjs As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const MT = "ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler
    
    Dim oInputObj1 As Object, oInputObj2 As Object
    Dim portIdx As SPSMemberAxisPortIndex
    Dim oSupped1Part As ISPSMemberPartPrismatic, oSupped2Part As ISPSMemberPartPrismatic
    Dim oSupped1Port As ISPSSplitAxisPort, oSupped2Port As ISPSSplitAxisPort
    Dim oPoint1 As IJPoint, oPoint2 As IJPoint, oPoint3 As IJPoint
    Dim dist1 As Double, dist2 As Double
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
    
    '2 ports are expected
    If portCol.count = 2 Then
        Set oInputObj1 = portCol.Item(1)
        Set oInputObj2 = portCol.Item(2)
        If (TypeOf oInputObj1 Is ISPSSplitAxisPort) And (TypeOf oInputObj2 Is ISPSSplitAxisPort) Then
            If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
                Set oSupped1Port = oInputObj1
                Set oSupped2Port = oInputObj2
            ElseIf TypeOf oInputObj2 Is ISPSSplitAxisEndPort Then
                Set oSupped1Port = oInputObj2
                Set oSupped2Port = oInputObj1
            Else
                ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
                Exit Property
            End If
            If TypeOf oSupped2Port Is ISPSSplitAxisAlongPort Then
                'get the end port
                Set oSupped2Part = oSupped2Port.Part
                portIdx = oSupped1Port.portIndex
                Set oSupped1Part = oSupped1Port.Part
                Set oPoint1 = oSupped1Part.PointAtEnd(portIdx)
                Set oPoint2 = oSupped2Part.PointAtEnd(SPSMemberAxisStart)
                Set oPoint3 = oSupped2Part.PointAtEnd(SPSMemberAxisEnd)
                dist1 = oPoint1.DistFromPt(oPoint2)
                dist2 = oPoint1.DistFromPt(oPoint3)
                If dist1 < dist2 Then
                    Set oSupped2Port = oSupped2Part.AxisPort(SPSMemberAxisStart)
                Else
                    Set oSupped2Port = oSupped2Part.AxisPort(SPSMemberAxisEnd)
                End If
            End If
            If relationObjs Is Nothing Then
                Set relationObjs = New JObjectCollection
            End If
            relationObjs.Clear
            relationObjs.Add oSupped1Port
            relationObjs.Add oSupped2Port
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
