VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameFndAsmDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************
' Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File
'    FrameFndAsmDef.cls
'
'Author
'       1-May-03        SS
'
'Description
'
'Notes
'
'History:
'
'   09/10/03    SS   tr# 48263 - occ attrs were used in determining the size of the plate
'                       instead of defn attrs. Also clearance was getting added twice.
'
'  13-Sep-04    SS   React to CmnApp SmartOcc impact for
'                           TR 59875 - CAODisplayCES semantic adds more to the graph
'                           adding new relation interface for children
' 06/06/06   Structure Development      Added CMMigrate method stubs
'   24-Jun-2006 JMS DI#60069 - Changes to allow editing of the weight and CG values
'                   changed call to SetWCG to call new interface to put weight and CG
'                   values since SetWCG is reservered for setting user defined values
'                   when the values here are the computed values
' 04-Aug-2006 AS TR#99968 Added functionality to handle Split Migration
'
'  19-Sep-06  SS    TR#105292 - add each of the member outputs to IJAssembly for it to have
'                   proper hierarchy in Assembly tab.
'
'  24-Sep-06  SS    TR#104973 - asserts during the sync. The reason for this assertion is that the symbol
'                   cannot access its inputs by reference because the ReferencesCollectionToSymbolRelation
'                   relationship does not exist.The custom method CMSetInputs defined on the
'                   IJDAggregatorDescription interface of a CustomAssemblyDefinition has the responsability
'                   to establish this ReferencesCollectionToSymbolRelation relationship.If this custom method
'                   is not defined, then the SmartOccurrence semantic establishes this relationship with the
'                   ReferencesCollection already connected to the SmartOccurrence through the SOtoArgs_R relationship.
'                   If this custom method is defined and does nothing, then the ReferencesCollectionToSymbolRelation
'                   relationship is never established. To address this:
'                   1. provide a migration script in SqlServer and Oracle for existing databases, to add the missing relationship (by CmnApp)
'                   2. removed the dummy implementations of this custom method in the project,
'                   incremented the .dll version number and a new synchronize)
'
'  02-Nov-07  SS    TR#128057 - A delta of .127m and .25m was being added to fndn port
'                   calculations based on standard pump foundation port. This will lead eqp fndns with diff dimensions
'                   when fndn ports are being added manually in eqp environment. Removed them.
'
'  14-Mar-08  SS    DI#134831  Changed the code from CreateObject() to SP3DCreateObject()
'                   as the symbol is no longer registered.
'
'  25-Aug-08   RP   CR#115597 - If the equipment port is missing then a warning is
'                               returned instead of error. This is necessary to
'                               get the children transformed to the pasted location
'  15-Sept-08  RP   TR#149287 - Added code to set the occurrence matrice on the founfdation
'*******************************************************************

Private Const MODULE = "FrameFndAsmDef"
Private Const strSourceFile = "FrameFndAsmDef.def"
Private Const CONST_ItemProgId As String = "SPSEqpFndMacros.FrameFndAsmDef"

Private Const FRAMEFNDASM_IFACE = "IJUASPSFrameFndnAsm"
Private Const PLATESIZEBYRULE_ATTRNAME = "IsTopPlateSizeDrivenByRule"
Private Const PLATEEDGECLEAR_ATTRNAME = "TopPlateEdgeClearance"
Private Const TOPPLATE_LENGTH_ATTRNAME = "TopPlateLength"
Private Const TOPPLATE_WIDTH_ATTRNAME = "TopPlateWidth"
Private Const TOPPLATE_THICKNESS_ATTRNAME = "TopPlateThickness"
Private Const TOPPLATE_MATERIAL_ATTRNAME = "TopPlateSPSMaterial"
Private Const TOPPLATE_GRADE_ATTRNAME = "TopPlateSPSGrade"

Private Const SIDEPLATE_LENGTH_ATTRNAME = "SidePlateLength"
Private Const SIDEPLATE_WIDTH_ATTRNAME = "SidePlateWidth"
Private Const SIDEPLATE_THICKNESS_ATTRNAME = "SidePlateThickness"
Private Const SIDEPLATE_MATERIAL_ATTRNAME = "SidePlateSPSMaterial"
Private Const SIDEPLATE_GRADE_ATTRNAME = "SidePlateSPSGrade"

Private Const LEGANGLE_MATERIAL_ATTRNAME = "LegAngleSPSMaterial"
Private Const LEGANGLE_GRADE_ATTRNAME = "LegAngleSPSGrade"
Private Const LEGANGLE_CS_STD_ATTRNAME = "LegAngleSPSSectionRefStandard"
Private Const LEGANGLE_CS_NAME_ATTRNAME = "LegAngleSPSSectionName"
Private Const LEGANGLE_HT_ATTRNAME = "LegAngleHeight"

' these for generic plate part, defined in assembly connections schema
Private Const PLATEPARTDIM_IFACE = "IJUASPSPlatePartDim"
Private Const PLATETHICK_IFACE = "IJUASPSPlateThickness"
Private Const PLATEPARTDIM_DEPTH_ATTRNAME = "Length"
Private Const PLATEPARTDIM_WIDTH_ATTRNAME = "Width"
Private Const PLATEPARTDIM_THICKNESS_ATTRNAME = "Thickness"
Private Const INTERFACE_WCGValueOrigin As String = "IJWCGValueOrigin"
Private Const PROPERTY_DryWCGOrigin As String = "DryWCGOrigin"

Private Enum enumWeightCGDerivation
    WEIGHTCG_Computed = 2
    WEIGHTCG_UserDefined = 4
End Enum

Private m_bOnPreLoad As Boolean

Const FNDNCOMPPROGID = "SPSEquipFoundations.SPSEquipFoundation"

Private Const ROOTSELECTOR_CLASS = "GenericRectPlatePart"
Private Const ROOTSELECTION = "GenericPlatePart_1"

Private Const DOUBLE_VALUE = 8
Private Const CHAR = 1

Private Const PI = 3.14159265358979

Private m_dAngleRotated As Double
Private m_oLocalizer As IJLocalizer

Implements IJDUserSymbolServices
Implements ISPSEquipFndnDefServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSFoundationInputHelper 'TR#71850



Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    'Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = CONST_ItemProgId
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler
     
    ' This method is in charge of the creation of the symbol definition object
    ' Can keep the current design unchanged for basic VB symbol definition.
    Dim pDefinition As IJDSymbolDefinition
    Dim pFact As IJCAFactory
    Set pFact = New CAFactory
    Set pDefinition = pFact.CreateCAD(pResourceMgr)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = CONST_ItemProgId
    pDefinition.CodeBase = CodeBase
    
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    pDefinition.name = IJDUserSymbolServices_GetDefinitionName(definitionParameters)
    
    ' Persistence behavior
    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
    'returned symbol definition
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
    Exit Function
     
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
 
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
    Dim oInput As IJDInput
    Dim oInputs As IJDInputs
    Set oInputs = pDefinition
    
    Set oInput = New DInput
    oInput.name = "EquipmentPorts"
    oInput.index = 1
    oInputs.Add oInput
    oInput.Reset
    
    oInput.name = "SupportedPlane"
    oInput.index = 2
    oInput.Properties = igDESCRIPTION_OPTIONAL
    oInputs.Add oInput
    oInput.Reset
    
    Set oInput = Nothing
    Set oInputs = Nothing
    
     ' Aggregator Type
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    pAD.AggregatorClsid = "{94F2C800-ECF1-47C0-88D5-01460B80ECD5}" 'CSPSEquipFoundation
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAsm"
    pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructAsm"
    pAD.SetCMSetInputs -1, -1
    pAD.SetCMRemoveInputs -1, -1
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
        
    Set pAD = Nothing
    
    ' tr 74802
    Dim pCADefinition As IJCADefinition
    Set pCADefinition = pDefinition
    Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
    Set pCADefinition = Nothing
    
     ' Aggregator property
    Dim pAPDs As IJDPropertyDescriptions
    Set pAPDs = pDefinition
    pAPDs.RemoveAll ' Remove all the previous property descriptions
    
    'listens to IJUASPSFrameFndnAsm
    pAPDs.AddProperty FRAMEFNDASM_IFACE, 1, FRAMEFNDASM_IFACE, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
    Set pAPDs = Nothing

     ' Define the members
    Dim pMemberDescriptions As IJDMemberDescriptions
    Dim pMemberDescription As IJDMemberDescription
    Dim pPropertyDescriptions As IJDPropertyDescriptions
    
    Set pMemberDescriptions = pDefinition
    
    ' Remove all the previous member descriptions
    pMemberDescriptions.RemoveAll
    
    'Top Plate
    Set pMemberDescription = pMemberDescriptions.AddMember("TopPlate", 1, "CMConstructTopPlate", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputTopPlate"
    pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTopPlate"
    pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
    
    Set pPropertyDescriptions = pMemberDescription
    'output watching IJDAttributes
    pPropertyDescriptions.AddProperty "TopPlateProperties", 1, IJDAttributes, "CMEvaluateTopPlate", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "TopPlateThickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeTopPlateThickness", imsCOOKIE_ID_USS_LIB

    'Leg Angle 1
    Set pMemberDescription = pMemberDescriptions.AddMember("LegAngle1", 2, "CMConstructLegAngle1", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputLegAngle1"
    pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructLegAngle1"
    pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
    
    Set pPropertyDescriptions = pMemberDescription
    'output
    pPropertyDescriptions.AddProperty "LegAngle1Properties", 1, IJDAttributes, "CMEvaluateLegAngle1", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "LegAngle1ModifiedbyMemberPart", 2, IJGeometry, "CMEvaluateLegAngle1Modified", imsCOOKIE_ID_USS_LIB
    
    'Leg Angle 2
    Set pMemberDescription = pMemberDescriptions.AddMember("LegAngle2", 3, "CMConstructLegAngle2", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputLegAngle2"
    pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructLegAngle2"
    pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID

    Set pPropertyDescriptions = pMemberDescription
    'output
    pPropertyDescriptions.AddProperty "LegAngle2Properties", 1, IJDAttributes, "CMEvaluateLegAngle2", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "LegAngle2ModifiedbyMemberPart", 2, IJGeometry, "CMEvaluateLegAngle2Modified", imsCOOKIE_ID_USS_LIB

    'Leg Angle 3
    Set pMemberDescription = pMemberDescriptions.AddMember("LegAngle3", 4, "CMConstructLegAngle3", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputLegAngle3"
    pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructLegAngle3"
    pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID

    Set pPropertyDescriptions = pMemberDescription
    'output
    pPropertyDescriptions.AddProperty "LegAngle3Properties", 1, IJDAttributes, "CMEvaluateLegAngle3", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "LegAngle3ModifiedbyMemberPart", 2, IJGeometry, "CMEvaluateLegAngle3Modified", imsCOOKIE_ID_USS_LIB

    'Leg Angle 4
    Set pMemberDescription = pMemberDescriptions.AddMember("LegAngle4", 5, "CMConstructLegAngle4", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputLegAngle4"
    pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructLegAngle4"
'    pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseLegAngle4"
    pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID

    Set pPropertyDescriptions = pMemberDescription
    'output
    pPropertyDescriptions.AddProperty "LegAngle4Properties", 1, IJDAttributes, "CMEvaluateLegAngle4", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "LegAngle4ModifiedbyMemberPart", 2, IJGeometry, "CMEvaluateLegAngle4Modified", imsCOOKIE_ID_USS_LIB

    Set pMemberDescriptions = Nothing
    Set pMemberDescription = Nothing
    Set pPropertyDescriptions = Nothing
   
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMFinalConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAsm"
On Error GoTo ErrorHandler
      
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMConstructAsm"
On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub SetPlateOrigin(pPropertyDescriptions As IJDPropertyDescription, ByRef eqPos As Automath.DPosition)
Const METHOD = "SetPlateOrigin"
On Error GoTo ErrorHandler

    Dim oFndn As IJLocalCoordinateSystem
    Set oFndn = pPropertyDescriptions.CAO
    
    Dim oPosition As DPosition
    Set oPosition = oFndn.Position
    
    Dim dx, dy, dz As Double
    dx = oPosition.x
    dy = oPosition.y
    dz = oPosition.z
    eqPos.Set dx, dy, dz
    
    Set oFndn = Nothing
    Set oPosition = Nothing
    
Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub SetMemberOrigin(pPropertyDescriptions As IJDMemberDescription, ByRef eqPos As Automath.DPosition)
Const METHOD = "SetMemberOrigin"
On Error GoTo ErrorHandler

    Dim oLeg As IJLocalCoordinateSystem
    Set oLeg = pPropertyDescriptions.CAO
    
    Dim oPosition As DPosition
    Set oPosition = oLeg.Position
    
    Dim dx, dy, dz As Double
    dx = oPosition.x
    dy = oPosition.y
    dz = oPosition.z
    eqPos.Set dx, dy, dz
    
    Set oLeg = Nothing
    Set oPosition = Nothing
    
Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub



Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO"
On Error GoTo ErrorHandler
    Dim pIRCAsm As IJDReferencesCollection
    Dim pIRCAsm1 As IJDReferencesCollection
    Dim pIRCAsm2 As IJDReferencesCollection
    Dim FoundationPorts As IJElements
    Dim oSmartOcc As IJSmartOccurrence
    Dim oOccAttribs As IJDAttributes
    Dim Trans As IJDT4x4
    Dim HoleLocations() As Automath.DPosition
    Dim NumberOfHoles As Integer
    Dim newLength As Double, newWidth As Double, newHeight As Double
    Dim xVec As New DVector, yVec As New DVector, zVec As New DVector
    Dim newTmx As New DT4x4
    Dim eqPos As New DPosition
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim iloop As Long
    Dim bPlateSizeByRule As Boolean
    Dim dblClear As Double
    Dim dblOffX As Double, dblOffY As Double
    Dim pOcc As IJDOccurrence
    Dim vProp As Variant
    
    'this methos sets the occurrence matrics for the equipment foundation
    
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIRCAsm = GetRefCollection(oSmartOcc)

    Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
    On Error Resume Next
    Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
    Set FoundationPorts = New JObjectCollection 'Elements
    
    Dim pEnumJDArgument As IEnumJDArgument
    Set pEnumJDArgument = pIRCAsm1
    
    Dim bPtOption As Boolean
    Call GetPortAndHoleLocations(pEnumJDArgument, FoundationPorts, HoleLocations, bPtOption)
    
    Set oOccAttribs = oSmartOcc
    
    Set newTmx = New Automath.DT4x4
    Set eqPos = New Automath.DPosition
    Dim xX As Double, yy As Double, zZ  As Double
    
    If bPtOption Or UBound(HoleLocations) > 1 Then
        Dim oFndnMatrix As IJDMatrixAccess
        Set oFndnMatrix = pPropertyDescriptions.CAO
        Set newTmx = oFndnMatrix.Matrix
        Dim oFndn As IJLocalCoordinateSystem
        Set oFndn = pPropertyDescriptions.CAO
        Dim oPosition As DPosition
        Set oPosition = oFndn.Position
        xX = oPosition.x
        yy = oPosition.y
        zZ = oPosition.z
        eqPos.Set xX, yy, zZ
        
        Set oFndnMatrix = Nothing
        Set oFndn = Nothing
    Else
        ' calculate holes from foundation ports positions
        Call GetTransformAndHoles(HoleLocations(), NumberOfHoles, FoundationPorts, Trans)
        
        'calculate the centroid of all the holes of all the foundation ports
        Dim centroidX As Double, centroidY As Double, centroidZ As Double
        GetCentroidOfPositions HoleLocations(), NumberOfHoles, centroidX, centroidY, centroidZ
        eqPos.Set centroidX, centroidY, centroidZ
        
        ' Calculate the length vector and width vector based on the hole locations.
        On Error GoTo PORTS_TOO_CLOSE
        SetRectangleOrientationByClearance HoleLocations(), 0, 0, xVec, yVec
        On Error GoTo ErrorHandler
    
'        newWidth = xVec.Length
'        newLength = yVec.Length
        
        'Transform the centre and the directions to global
        Set eqPos = Trans.TransformPosition(eqPos)
        Set xVec = Trans.TransformVector(xVec)
        Set yVec = Trans.TransformVector(yVec)
        Set zVec = yVec.Cross(xVec)

        xVec.Length = 1
        yVec.Length = 1
        zVec.Length = 1
      
        ' In case the supported objects(equipments) are not in same heights then adjust the EF also
        ' so that the EF should start from the bottom most point in the normal direction.
        ' therefore the origin of the EF should be shifted to the lower most point along the normal.
        
        ' Identify the maximum distance between the origin and all the holes along normal
        Dim dShift As Double
        dShift = GetOriginShiftForUnevenSupportedObjects(HoleLocations, eqPos, zVec, Trans)
        dShift = dShift + oOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
        
        ' Shift the origin along the vector by the magnitude of maximum difference.
        Dim dTol As Double
        dTol = 0.0001
        If dShift - 0# > dTol Then
        zVec.Length = 1
        zVec.Set zVec.x * dShift, zVec.y * dShift, zVec.z * dShift
        eqPos.Set eqPos.x + zVec.x, eqPos.y + zVec.y, eqPos.z + zVec.z
        End If
        
        'Set the transform
        SetTransform newTmx, eqPos, xVec, yVec
    End If
    
    On Error Resume Next
    Set pOcc = pObject
    If pOcc Is Nothing Then
        Exit Sub
    End If
    pOcc.Matrix = newTmx
        
    If Not FoundationPorts Is Nothing Then
        FoundationPorts.Clear
        Set FoundationPorts = Nothing
    End If
   
    Set xVec = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set eqPos = Nothing
    Set newTmx = Nothing
    Set pIRCAsm = Nothing
    Set pIRCAsm1 = Nothing
    Set pIRCAsm2 = Nothing
    Set oSmartOcc = Nothing
    Set oOccAttribs = Nothing
    Set pOcc = Nothing
'    Set SupportPlane = Nothing
    Set oFndn = Nothing
    Set zVec = Nothing
    Set Trans = Nothing
    
    Erase HoleLocations
   
    Exit Sub
PORTS_TOO_CLOSE:
    SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFND_CLOSEFOR_SELECTED, pObject, Nothing
    Err.Raise E_FAIL
    Exit Sub

ErrorHandler:
    ' For errors logged with SPS_MACRO_WARNING, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If


End Sub

' Top Plate methods
Public Sub CMConstructTopPlate(ByVal pMemberDescription As IJDMemberDescription, _
                                ByVal pResourceManager As IUnknown, _
                                ByRef pObj As Object)
Const METHOD = "CMConstructTopPlate"
On Error GoTo ErrorHandler
    
    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim strMaterial As String, strGrade As String
    Dim oPlateSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oPlate As IJStructCustomPlatePart
    Dim oCustomPlatePartFactory As IJStructCustomPlatePartFactory
    Dim TopPlateThickness As Double
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    
    Set oCustomPlatePartFactory = New StructCustomPlatePartFactory
    
    'Create the custom Plate, that is the Top Plate.
    Set oPlate = oCustomPlatePartFactory.CreateCustomPlatePart(pResourceManager)

    Set oAttrbs = oSmartOccCAO  ' occ attrs
    
    ' get material and grade
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_GRADE_ATTRNAME).Value
   
    SetPlateMaterial oPlate, strMaterial, strGrade

    'get default from the parent
    TopPlateThickness = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
    
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, TopPlateThickness
    'set the found (from the available in catalog) thickness on attrbs
    oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value = TopPlateThickness

    Set oPlateSmartOcc = oPlate
    
'     Initialize the SmartOccurrence
    oPlateSmartOcc.RootSelectorClass = ROOTSELECTOR_CLASS
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oPlate
    Set oAssemblyparent = Nothing
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oPlateSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
    
    Set oPlate = Nothing
    Set oCustomPlatePartFactory = Nothing
    Set oPlateSmartOcc = Nothing
    Set oRefColl = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputTopPlate(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputTopPlate"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTopPlate(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructTopPlate"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTopPlate(pPropertyDescriptions As IJDPropertyDescription, _
                              pObject As Object)
Const METHOD = "CMEvaluateTopPlate"
On Error GoTo ErrorHandler

    Dim pIRCAsm As IJDReferencesCollection
    Dim pIRCAsm1 As IJDReferencesCollection
    Dim pIRCAsm2 As IJDReferencesCollection
    Dim FoundationPorts As IJElements
    Dim SupportPlane As IJPlane
    Dim pIJOccAttribs As IJDAttributes
'    Dim pIJDefAttrs As IJDAttributes
    Dim oPort As Object
    Dim pOcc As IJDOccurrence
    Dim rc1cnt As Integer, rc2cnt As Integer
    Dim rccnt As Integer, i As Integer
    Dim bFoundPorts As Boolean, bFoundSuppSurface As Boolean
    Dim Trans As IngrGeom3D.IJDT4x4
    Dim HoleLocations() As Automath.DPosition
    Dim NumberOfHoles As Integer
    Dim newLength As Double, newWidth As Double, newHeight As Double
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim newTmx As Automath.DT4x4
    Dim eqPos As Automath.DPosition
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oAngleCalcVec1 As IJDVector
    Dim oAngleCalcVec2 As IJDVector
    
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set oSmartOccChild = pObject
    
    oSmartOccChild.Update ' mark IJSmartOccurrence as modified so that the baseplate generation semantic fires
    'this is needed to get the plate transformed during a move
    'move this to here from  down below so that the update is done even when the equipment is missing
    
    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)   ' pPropertyDescriptions.CAO

    rccnt = pIRCAsm.IJDEditJDArgument.GetCount
    Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
    On Error Resume Next
    Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
    
    Set FoundationPorts = New JObjectCollection 'Elements
    
    Dim pEnumJDArgument As IEnumJDArgument
    Set pEnumJDArgument = pIRCAsm1
    
    Dim bPtOption As Boolean
    Call GetPortAndHoleLocations(pEnumJDArgument, FoundationPorts, HoleLocations, bPtOption)
    
    Set pIJOccAttribs = oSmartOccCAO    ' occ attrs
    
    ' get the SizingByRule attribute
    Dim bPlateSizeByRule As Boolean
    bPlateSizeByRule = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(PLATESIZEBYRULE_ATTRNAME).Value

    newWidth = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value
    newLength = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value
    Dim plateThick As Double
    plateThick = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
    
    Set newTmx = New Automath.DT4x4
    Set eqPos = New Automath.DPosition
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set zVec = New Automath.DVector
    
    Dim xX As Double, yy As Double, zZ As Double
    
    If bPtOption Or UBound(HoleLocations) > 1 Then
        Dim oFndnMatrix As IJDMatrixAccess
        Set oFndnMatrix = pPropertyDescriptions.CAO
        Set newTmx = oFndnMatrix.Matrix
        Dim oFndn As IJLocalCoordinateSystem
        Set oFndn = pPropertyDescriptions.CAO
        Dim oPosition As DPosition
        Set oPosition = oFndn.Position
        xX = oPosition.x
        yy = oPosition.y
        zZ = oPosition.z
        eqPos.Set xX, yy, zZ
        
        'Set the Z axis vector
        Dim zVecX As Double, zVecY As Double, zVecZ As Double
        zVecX = newTmx.IndexValue(8)
        zVecY = newTmx.IndexValue(9)
        zVecZ = newTmx.IndexValue(10)
        zVec.Set zVecX, zVecY, zVecZ
        zVec.Length = 1

        If UBound(HoleLocations) > 1 Then ' multiple point case commented for now

        End If
          
        Set oFndnMatrix = Nothing
        Set oFndn = Nothing
    Else
        ' calculate holes from foundation ports positions
        Call GetTransformAndHoles(HoleLocations(), NumberOfHoles, FoundationPorts, Trans)
        
        'calculate the centroid of all the holes of all the foundation ports
        Dim centroidX As Double, centroidY As Double, centroidZ As Double
        GetCentroidOfPositions HoleLocations(), NumberOfHoles, centroidX, centroidY, centroidZ
        eqPos.Set centroidX, centroidY, centroidZ
        

        ' Calculate the length vector and width vector based on the hole locations.
        On Error GoTo PORTS_TOO_CLOSE
        SetRectangleOrientationByClearance HoleLocations(), 0, 0, xVec, yVec
        On Error GoTo ErrorHandler
    
        newWidth = xVec.Length
        newLength = yVec.Length
        
        'Set the first vector before rotating the object -- snkota (TR CP 42798)
        Set oAngleCalcVec1 = xVec
        
        'Transform the centre and the directions to global
        Set eqPos = Trans.TransformPosition(eqPos)
        Set xVec = Trans.TransformVector(xVec)
        Set yVec = Trans.TransformVector(yVec)
        Set zVec = yVec.Cross(xVec)
        
        xVec.Length = 1
        yVec.Length = 1
        zVec.Length = 1
      
        ' In case the supported objects(equipments) are not in same heights then adjust the EF also
        ' so that the EF should start from the bottom most point in the normal direction.
        ' therefore the origin of the EF should be shifted to the lower most point along the normal.
        
        ' Identify the maximum distance between the origin and all the holes along normal
        Dim dShift As Double
        dShift = GetOriginShiftForUnevenSupportedObjects(HoleLocations, eqPos, zVec, Trans)
        dShift = dShift + pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
    
        ' Shift the origin along the vector by the magnitude of maximum difference.
        Dim dTol As Double
        dTol = 0.0001
        If dShift - 0# > dTol Then
        zVec.Length = 1
        zVec.Set zVec.x * dShift, zVec.y * dShift, zVec.z * dShift
        eqPos.Set eqPos.x + zVec.x, eqPos.y + zVec.y, eqPos.z + zVec.z
        End If
        
        'Set the transform
        SetTransform newTmx, eqPos, xVec, yVec
        Set zVec = yVec.Cross(xVec)
        
        'Get the Second vector for calculation of the rotaion angle after transform
        '       -- snkota (TR CP 42798)
        Set oAngleCalcVec2 = xVec
        
        Dim otempnormVec As IJDVector
        Set otempnormVec = oAngleCalcVec1.Cross(oAngleCalcVec2)
        'Calculate the angle of rotation of the object from start point
        m_dAngleRotated = oAngleCalcVec2.angle(oAngleCalcVec1, otempnormVec)
    End If
    
    ' get the clearance attribute value
    Dim dClear As Double
    dClear = 0#
    If bPlateSizeByRule = False And pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(PLATEEDGECLEAR_ATTRNAME).Value <> "" Then
        dClear = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(PLATEEDGECLEAR_ATTRNAME).Value
    End If
    
    Dim dblOffX As Double, dblOffY As Double
    dblOffX = 0#
    dblOffY = 0#
    Dim vProp As Variant
    If bPlateSizeByRule = False Then
        vProp = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value
        If vProp <> vbEmpty Then
            If vProp <> newWidth Then
                dblOffX = vProp - newWidth
            End If
        End If
        vProp = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value
        If vProp <> vbEmpty Then
            If vProp <> newLength Then
                dblOffY = vProp - newLength
            End If
        End If
        newWidth = newWidth + dblOffX
        newLength = newLength + dblOffY
    End If
    
    ' assign new values to block
    If newWidth <= 0# Or newLength <= 0# Then
        On Error GoTo 0
        SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_FNDWIDTH_LENGTH_ZERO, oSmartOccCAO, Nothing
        Err.Raise E_FAIL
    End If
    pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value = newLength
    pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value = newWidth
    
    rc2cnt = 0
    If pIRCAsm2 Is Nothing Then Else rc2cnt = pIRCAsm2.IJDEditJDArgument.GetCount
    If rc2cnt >= 1 Then
        bFoundSuppSurface = True
        Set SupportPlane = pIRCAsm2.IJDEditJDArgument.GetEntityByIndex(1)
    End If
    If bFoundSuppSurface = True Then
        zVec.Length = 1     ' normalize
        ' calculate the block ht from the support plane
        CalculateHeightFromOriginToSupporingPlane SupportPlane, eqPos, zVec, newHeight
        If newHeight <= 0# Then
            On Error GoTo 0
            SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_ASSOCEQP_MODIFIED_NOHEIGHT, oSmartOccCAO, Nothing
            Err.Raise E_FAIL
            
        End If
        pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value = newHeight - plateThick
    End If
                
    On Error Resume Next
    Set pOcc = pObject
    If pOcc Is Nothing Then
        Exit Sub
    End If
    
    pOcc.Matrix = newTmx
    oSmartOccChild.Update ' mark IJSmartOccurrence as modified so that the baseplate generation semantic fires
    
    ' assign them to the plate object (child)
    Dim pIJAttribsChild As IJDAttributes
    
    Set oSmartOccChild = pPropertyDescriptions.object
    Set pIJAttribsChild = oSmartOccChild

    ' get the clearance attribute value
    dClear = 0#
    If bPlateSizeByRule = False Then
        dClear = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(PLATEEDGECLEAR_ATTRNAME).Value
    End If
    
    pIJAttribsChild.CollectionOfAttributes(PLATEPARTDIM_IFACE).Item(PLATEPARTDIM_DEPTH_ATTRNAME).Value = newWidth + dClear ' tr 48639, fix plate part length
    pIJAttribsChild.CollectionOfAttributes(PLATEPARTDIM_IFACE).Item(PLATEPARTDIM_WIDTH_ATTRNAME).Value = newLength + dClear ' tr 48639, fix plate part width
    
    Dim oPlate As IJStructCustomPlatePart
    Set oPlate = oSmartOccChild
    
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, plateThick

    pIJAttribsChild.CollectionOfAttributes(PLATETHICK_IFACE).Item(PLATEPARTDIM_THICKNESS_ATTRNAME).Value = plateThick
    ' for optional plane type following line is creating fndn with -ve ht, corrected.
    If bFoundSuppSurface = True Then pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value = newHeight - plateThick
    pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value = plateThick
    
    Set oPlate = Nothing
    Set pIJAttribsChild = Nothing
    Set oSmartOccChild = Nothing
    
    If Not FoundationPorts Is Nothing Then
        FoundationPorts.Clear
        Set FoundationPorts = Nothing
    End If
    
    Set xVec = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set eqPos = Nothing
    Set newTmx = Nothing
    Set pIRCAsm = Nothing
    Set pIRCAsm1 = Nothing
    Set pIRCAsm2 = Nothing
    Set oSmartOccCAO = Nothing
    Set pIJOccAttribs = Nothing
    Set pOcc = Nothing
    Set SupportPlane = Nothing
    Set oSmartOccChild = Nothing
    
    Erase HoleLocations
        
    Exit Sub
PORTS_TOO_CLOSE:
    SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFND_CLOSEFOR_SELECTED, pObject, Nothing
    Err.Raise E_FAIL
    Exit Sub

ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub

Public Sub CMComputeTopPlateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const MT = "CMComputeTopPlateThickness"
    On Error GoTo ErrorHandler
    
    'This is a notification function and is called when the plate thickness is changed.
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub

' Leg Angle1 methods
Public Sub CMConstructLegAngle1(ByVal pMemberDescription As IJDMemberDescription, _
                                ByVal pResourceManager As IUnknown, _
                                ByRef pObj As Object)
Const METHOD = "CMConstructLegAngle1"
On Error GoTo ErrorHandler
    
    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oMemberFactory  As SPSMembers.SPSMemberFactory
    Dim oMemberPart     As SPSMembers.ISPSMemberPartPrismatic
    Dim oSupportSurface As IJPlane

    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO

    Set oAttrbs = oSmartOccCAO  ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value

    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)

    Dim legHt As Double
    legHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value

    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pMemberDescription.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Call SetOrigin(pMemberDescription, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, oAttrbs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
'    ' get top plate thickness
    Dim dTopPHt As Double
    dTopPHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    Set oMemberFactory = New SPSMembers.SPSMemberFactory
    Set oMemberPart = oMemberFactory.CreateMemberPartPrismatic(3, pResourceManager)

    Set oMemberPart.CrossSection.definition = oCrossSection

    Dim cp As Integer
    Dim angle As Double

    cp = 5
    angle = 1.5 * PI        ' 270 degrees

    oMemberPart.CrossSection.CardinalPoint = cp
'    oMemberPart.Rotation.BetaAngle = angle
    oMemberPart.MemberType.TypeCategory = 1
    oMemberPart.MemberType.Type = 100

    Set oMemberPart.MaterialDefinition = oMaterialDefinition

    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
            
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = oMemberPart.Axis

    eqPos.x = eqPos.x + dLegSecWidth
    eqPos.y = eqPos.y
    eqPos.z = eqPos.z + (zaxis.z * dTopPHt)

    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)

    
'Dim topPlateWidth As Double, topPlateLength As Double
'topPlateWidth = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value
'topPlateLength = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value
'Dim widthOffset As Double, lengthOffset As Double
'widthOffset = topPlateWidth / 2 - dLegSecWidth / 2
'lengthOffset = topPlateLength / 2 - dLegSecWidth / 2
'Dim oStartPos As New DPosition
'oStartPos.Set eqPos.x - widthOffset, eqPos.y - lengthOffset, eqPos.z
'oLine3dPart.SetStartPoint oStartPos.x, oStartPos.y, oStartPos.z
'oLine3dPart.SetEndPoint oStartPos.x + (zaxis.x * legHt), oStartPos.y + (zaxis.y * legHt), oStartPos.z + (zaxis.z * legHt)
'Set oStartPos = Nothing
    
    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oMemberPart

    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oMemberPart
    Set oAssemblyparent = Nothing
    
    'create name for the member part
    GenerateNameForMember oMemberPart

    Set pObj = oMemberPart 'return the created member part object to the custom assembly

    Set oMemberPart = Nothing
    Set oMemberFactory = Nothing
    Set oRefColl = Nothing
    Set oSmartOcc = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing

    Set xVec = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set eqPos = Nothing
    Set pIRCAsm = Nothing
    Set oSmartOcc = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputLegAngle1(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputLegAngle1"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructLegAngle1(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructLegAngle1"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMEvaluateLegAngle1(pPropertyDescriptions As IJDPropertyDescription, _
                               pObject As Object)
Const METHOD = "CMEvaluateLegAngle1"
On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes
    Dim oSupportSurface As IJPlane
    
    
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO  ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value
    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

   'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition
    
    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis
    
'    oLine3dPart.length = legHt
  'Based on the center position found, find the position where the leg1 has to
  'be placed -- snkota (TR CP 42798)
    '   .1__________
    '   |           |
    '   |           |
    '   |           |
    '    ----------
    Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z
    yVec.Set -tmpX, -tmpY, -tmpZ
    Set newPos = eqPos.Offset(xVec)
    Set eqPos = newPos.Offset(yVec)
    
    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If

    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector tmpX, tmpY, tmpZ
    pMemberPart.Rotation.Mirror = True
    
    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateLegAngle1Modified(pPropertyDescriptions As IJDPropertyDescription, _
                                       pObject As Object)
Const METHOD = "CMEvaluateLegAngle1Modified"
On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes
    Dim oSupportSurface As IJPlane
    
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO  ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value
    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

   'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis

'Based on the center position found, find the position where the leg1 has to
  'be placed -- snkota (TR CP 42798)
    '   .1__________
    '   |           |
    '   |           |
    '   |           |
    '    ----------
   Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z

    yVec.Set -tmpX, -tmpY, -tmpZ

    Set newPos = eqPos.Offset(xVec)
    Set eqPos = newPos.Offset(yVec)

    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)

    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If
    
    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector tmpX, tmpY, tmpZ
    pMemberPart.Rotation.Mirror = True

    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' Leg Angle2 methods
Public Sub CMConstructLegAngle2(ByVal pMemberDescription As IJDMemberDescription, _
                                ByVal pResourceManager As IUnknown, _
                                ByRef pObj As Object)
Const METHOD = "CMConstructLegAngle2"
On Error GoTo ErrorHandler
    
    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oMemberFactory  As SPSMembers.SPSMemberFactory
    Dim oMemberPart     As SPSMembers.ISPSMemberPartPrismatic
    Dim oSupportSurface As IJPlane
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    
    Set oAttrbs = oSmartOccCAO  ' occ attrs
            
    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value
    
    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)
    
    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth
    
    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value
    
    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)

    Dim legHt As Double
    legHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value

    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition
    
    Set pIRCAsm = GetRefCollection(pMemberDescription.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Call SetOrigin(pMemberDescription, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, oAttrbs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    ' get top plate thickness
    Dim dTopPHt As Double
    dTopPHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
        
    Set oMemberFactory = New SPSMembers.SPSMemberFactory
    Set oMemberPart = oMemberFactory.CreateMemberPartPrismatic(3, pResourceManager)

    Set oMemberPart.CrossSection.definition = oCrossSection
    
    Dim cp As Integer
    Dim angle As Double
    
    cp = 5
    angle = 0#         ' 0 degrees
    
    oMemberPart.CrossSection.CardinalPoint = cp
'    oMemberPart.Rotation.BetaAngle = angle
    oMemberPart.MemberType.TypeCategory = 1
    oMemberPart.MemberType.Type = 100
    
    Set oMemberPart.MaterialDefinition = oMaterialDefinition

    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = oMemberPart.Axis

    ' get top plate width
    Dim dTopPlateWidth As Double
    dTopPlateWidth = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value
   
    eqPos.x = eqPos.x + dLegSecWidth
    eqPos.y = eqPos.y - dTopPlateWidth + dLegSecWidth
    eqPos.z = eqPos.z + (zaxis.z * dTopPHt)

    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oMemberPart
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oMemberPart
    Set oAssemblyparent = Nothing
    
    'create name for the member part
    GenerateNameForMember oMemberPart
    
    Set pObj = oMemberPart 'return the created member part object to the custom assembly
    
    Set oMemberPart = Nothing
    Set oMemberFactory = Nothing
    Set oRefColl = Nothing
    Set oSmartOcc = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Set xVec = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set eqPos = Nothing
    Set pIRCAsm = Nothing
    Set oSmartOcc = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputLegAngle2(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputLegAngle2"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructLegAngle2(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructLegAngle2"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateLegAngle2(pPropertyDescriptions As IJDPropertyDescription, _
                               pObject As Object)
Const METHOD = "CMEvaluateLegAngle2"
On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes

    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO    ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value

    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    ' get top plate ht
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    ' get top plate width
    Dim dTopPlateWidth As Double
    dTopPlateWidth = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis
  'Based on the center position found, find the position where the leg2 has to
  'be placed -- snkota (TR CP 42798)
    '   .___________2
    '   |           |
    '   |           |
    '   |           |
    '    ----------
    Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    
    tmpX = xVec.x
    tmpY = xVec.y
    tmpZ = xVec.z
    xVec.Set -tmpX, -tmpY, -tmpZ
       
    Set newPos = eqPos.Offset(xVec)
    
    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z
    yVec.Set -tmpX, -tmpY, -tmpZ
    
    Set eqPos = newPos.Offset(yVec)
    
    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If
    
    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector tmpX, tmpY, tmpZ
    pMemberPart.Rotation.Mirror = False
    
    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateLegAngle2Modified(pPropertyDescriptions As IJDPropertyDescription, _
                                        pObject As Object)
Const METHOD = "CMEvaluateLegAngle2Modified"
On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes

    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO    ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value

    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    ' get top plate ht
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
    
    ' get top plate width
    Dim dTopPlateWidth As Double
    dTopPlateWidth = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis
  'Based on the center position found, find the position where the leg2 has to
  'be placed -- snkota (TR CP 42798)
    '   .___________2
    '   |           |
    '   |           |
    '   |           |
    '    ----------
    Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    
    tmpX = xVec.x
    tmpY = xVec.y
    tmpZ = xVec.z
    xVec.Set -tmpX, -tmpY, -tmpZ
       
    Set newPos = eqPos.Offset(xVec)
    
    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z
    yVec.Set -tmpX, -tmpY, -tmpZ
    
    Set eqPos = newPos.Offset(yVec)

    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If
    
    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector tmpX, tmpY, tmpZ
    pMemberPart.Rotation.Mirror = False
    
    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' Leg Angle 3 methods
Public Sub CMConstructLegAngle3(ByVal pMemberDescription As IJDMemberDescription, _
                                ByVal pResourceManager As IUnknown, _
                                ByRef pObj As Object)
Const METHOD = "CMConstructLegAngle3"
On Error GoTo ErrorHandler
    
    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oMemberFactory  As SPSMembers.SPSMemberFactory
    Dim oMemberPart     As SPSMembers.ISPSMemberPartPrismatic
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    
    Set oAttrbs = oSmartOccCAO      ' occ attrs
            
    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value
    
    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)
    
    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth
    
    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value
    
    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)

    Dim legHt As Double
    legHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value

    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition
    
    Set pIRCAsm = GetRefCollection(pMemberDescription.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pMemberDescription, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, oAttrbs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)

    Set pIJDefAttrs = Nothing
    
'    ' get top plate thickness
    Dim dTopPHt As Double
    dTopPHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
        
    Set oMemberFactory = New SPSMembers.SPSMemberFactory
    Set oMemberPart = oMemberFactory.CreateMemberPartPrismatic(3, pResourceManager)

    Set oMemberPart.CrossSection.definition = oCrossSection
    
    Dim cp As Integer
    Dim angle As Double
    
    cp = 5
    angle = PI        ' 180 degrees
    
    oMemberPart.CrossSection.CardinalPoint = cp
'    oMemberPart.Rotation.BetaAngle = angle
    oMemberPart.MemberType.TypeCategory = 1
    oMemberPart.MemberType.Type = 100
    
    Set oMemberPart.MaterialDefinition = oMaterialDefinition

    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = oMemberPart.Axis

    ' get top plate width
    Dim dTopPlateLength As Double
    dTopPlateLength = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value

    eqPos.x = eqPos.x + dTopPlateLength
    eqPos.y = eqPos.y
    eqPos.z = eqPos.z + (zaxis.z * dTopPHt)

    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oMemberPart
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oMemberPart
    Set oAssemblyparent = Nothing
    
    'create name for the member part
    GenerateNameForMember oMemberPart
    
    Set pObj = oMemberPart 'return the created member part object to the custom assembly
    
    Set oMemberPart = Nothing
    Set oMemberFactory = Nothing
    Set oRefColl = Nothing
    Set oSmartOcc = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Set xVec = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set eqPos = Nothing
    Set pIRCAsm = Nothing
    Set oSmartOcc = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputLegAngle3(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputLegAngle3"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructLegAngle3(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructLegAngle3"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateLegAngle3(pPropertyDescriptions As IJDPropertyDescription, _
                               pObject As Object)
Const METHOD = "CMEvaluateLegAngle3"
On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes

    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO    ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value

    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis
  
     ' get top plate width
    Dim dTopPlateLength As Double
    dTopPlateLength = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value

'Based on the center position found, find the position where the leg3 has to
  'be placed -- snkota (TR CP 42798)
    '   .___________
    '   |           |
    '   |           |
    '   |           |
    '   3-----------
    Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2
    
    Set newPos = eqPos.Offset(xVec)
    
    Set eqPos = newPos.Offset(yVec)
    
    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If
    
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z
    
    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = 2 * PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector -tmpX, -tmpY, -tmpZ
    pMemberPart.Rotation.Mirror = False
    
    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateLegAngle3Modified(pPropertyDescriptions As IJDPropertyDescription, _
                                       pObject As Object)
Const METHOD = "CMEvaluateLegAngle3Modified"
On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes

    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO    ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value

    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis
  
     ' get top plate width
    Dim dTopPlateLength As Double
    dTopPlateLength = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value
'Based on the center position found, find the position where the leg3 has to
  'be placed -- snkota (TR CP 42798)
    '   .___________
    '   |           |
    '   |           |
    '   |           |
    '   3-----------
      Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2

    Set newPos = eqPos.Offset(xVec)
    
    Set eqPos = newPos.Offset(yVec)
    
    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If
    
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z
    
    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = 2 * PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector -tmpX, -tmpY, -tmpZ
    pMemberPart.Rotation.Mirror = False
    
    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' Leg Angle 4 methods
Public Sub CMConstructLegAngle4(ByVal pMemberDescription As IJDMemberDescription, _
                                ByVal pResourceManager As IUnknown, _
                                ByRef pObj As Object)
Const METHOD = "CMConstructLegAngle4"
On Error GoTo ErrorHandler
    
    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oMemberFactory  As SPSMembers.SPSMemberFactory
    Dim oMemberPart     As SPSMembers.ISPSMemberPartPrismatic
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    
    Set oAttrbs = oSmartOccCAO  ' occ attrs
            
    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value
    
    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)
    
    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth
    
    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value
    
    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)

    Dim legHt As Double
    legHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value

    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition
    
    Set pIRCAsm = GetRefCollection(pMemberDescription.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pMemberDescription, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, oAttrbs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)

    Set pIJDefAttrs = Nothing
    
'    ' get top plate thickness
    Dim dTopPHt As Double
    dTopPHt = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
        
    Set oMemberFactory = New SPSMembers.SPSMemberFactory
    Set oMemberPart = oMemberFactory.CreateMemberPartPrismatic(3, pResourceManager)

    Set oMemberPart.CrossSection.definition = oCrossSection
    
    Dim cp As Integer
    Dim angle As Double
    
    cp = 5
    angle = PI / 2
    
    oMemberPart.CrossSection.CardinalPoint = cp
'    oMemberPart.Rotation.BetaAngle = angle
    oMemberPart.MemberType.TypeCategory = 1
    oMemberPart.MemberType.Type = 100
    
    Set oMemberPart.MaterialDefinition = oMaterialDefinition

    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = oMemberPart.Axis

    ' get top plate width
    Dim dTopPlateLength As Double, dTopPlateWidth As Double
    dTopPlateLength = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value
    dTopPlateWidth = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value

    eqPos.x = eqPos.x + dTopPlateLength
    eqPos.y = eqPos.y + dLegSecWidth - dTopPlateWidth
    eqPos.z = eqPos.z + (zaxis.z * dTopPHt) - dTopPHt

    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oMemberPart
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oMemberPart
    Set oAssemblyparent = Nothing
    
    'create name for the member part
    GenerateNameForMember oMemberPart
    
    Set pObj = oMemberPart 'return the created member part object to the custom assembly
    
    Set oMemberPart = Nothing
    Set oMemberFactory = Nothing
    Set oRefColl = Nothing
    Set oSmartOcc = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Set xVec = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set eqPos = Nothing
    Set pIRCAsm = Nothing
    Set oSmartOcc = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputLegAngle4(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputLegAngle4"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructLegAngle4(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructLegAngle4"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateLegAngle4(pPropertyDescriptions As IJDPropertyDescription, _
                                pObject As Object)
Const METHOD = "CMEvaluateLegAngle4"
On Error GoTo ErrorHandler

    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes

    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO    ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value

    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis
  
    ' get top plate width
    Dim dTopPlateLength As Double, dTopPlateWidth As Double
    dTopPlateLength = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value
    dTopPlateWidth = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value
'Based on the center position found, find the position where the leg4 has to
  'be placed -- snkota (TR CP 42798)
    '   .___________
    '   |           |
    '   |           |
    '   |           |
    '    -----------4
   Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    tmpX = xVec.x
    tmpY = xVec.y
    tmpZ = xVec.z
      
    xVec.Set -tmpX, -tmpY, -tmpZ
    
    Set newPos = eqPos.Offset(xVec)
    Set eqPos = newPos.Offset(yVec)

    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If
    
    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z
    
    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = 2 * PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector -tmpX, -tmpY, -tmpZ
    pMemberPart.Rotation.Mirror = True
    
    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateLegAngle4Modified(pPropertyDescriptions As IJDPropertyDescription, _
                                      pObject As Object)
Const METHOD = "CMEvaluateLegAngle4Modified"
On Error GoTo ErrorHandler
    
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes

    Set oSmartOccCAO = pPropertyDescriptions.CAO
    
    Set pIJOccAttribs = oSmartOccCAO    ' occ attrs

    'get cross section and its ref standard
    Dim LegSecStandard As String, LegSectionName As String
    LegSecStandard = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_STD_ATTRNAME).Value
    LegSectionName = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_CS_NAME_ATTRNAME).Value

    'get cross section object
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = GetCrossSection(LegSecStandard, LegSectionName)

    ' from cross section info get its properties
    Dim dLegSecWidth As Double, dLegSecDepth As Double
    GetCrossSecData LegSectionName, LegSecStandard, dLegSecWidth, dLegSecDepth

    ' get material and grade
    Dim strMaterial As String, strGrade As String
    strMaterial = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_MATERIAL_ATTRNAME).Value
    strGrade = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_GRADE_ATTRNAME).Value

    'get material defn
    Dim oMaterialDefinition As Object
    Set oMaterialDefinition = GetMaterialObject(strMaterial, strGrade)
    
    ' get the equipment port information and its positional coordinates
    Dim pIRCAsm As IJDReferencesCollection
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim eqPos As Automath.DPosition

    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set eqPos = New Automath.DPosition

    Dim pIJDefAttrs As IJDAttributes
    Set pIJDefAttrs = oSmartOccCAO.ItemObject
    
    Dim oSupportSurface As IJPlane
    Call SetOrigin(pPropertyDescriptions, eqPos)
    Call CalculateEqpPortandPlatePos(pIRCAsm, pIJOccAttribs, pIJDefAttrs, eqPos, xVec, yVec, oSupportSurface)
    Set pIJDefAttrs = Nothing
    
    Dim dTopPHt As Double
    dTopPHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value

    Dim legHt As Double
    legHt = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(LEGANGLE_HT_ATTRNAME).Value
    
    Dim pMemberPart As ISPSMemberPartPrismatic
    Set pMemberPart = pObject
    
    Set pMemberPart.CrossSection.definition = oCrossSection
    
    Set pMemberPart.MaterialDefinition = oMaterialDefinition
    
    Dim zaxis As New Automath.DVector
    Set zaxis = yVec.Cross(xVec)
    zaxis.Length = 1     ' normalize
    
    Dim oLine3dPart     As Line3d
    Set oLine3dPart = pMemberPart.Axis
  
    ' get top plate width
    Dim dTopPlateLength As Double, dTopPlateWidth As Double
    dTopPlateLength = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value
    dTopPlateWidth = pIJOccAttribs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value
'Based on the center position found, find the position where the leg4 has to
  'be placed -- snkota (TR CP 42798)
    '   .___________
    '   |           |
    '   |           |
    '   |           |
    '    -----------4
     Dim newPos As IJDPosition
    xVec.Length = xVec.Length / 2 - dLegSecWidth / 2
    yVec.Length = yVec.Length / 2 - dLegSecDepth / 2
    Dim tmpX As Double, tmpY As Double, tmpZ As Double
    tmpX = xVec.x
    tmpY = xVec.y
    tmpZ = xVec.z
      
    xVec.Set -tmpX, -tmpY, -tmpZ
    
    Set newPos = eqPos.Offset(xVec)
    Set eqPos = newPos.Offset(yVec)
    oLine3dPart.SetStartPoint eqPos.x, eqPos.y, eqPos.z
    oLine3dPart.SetEndPoint eqPos.x + (zaxis.x * legHt), eqPos.y + (zaxis.y * legHt), eqPos.z + (zaxis.z * legHt)
    
    ' tr 69627 - optimized code, if old and new cutback surface are same
    If oSupportSurface Is Nothing Then
        If pMemberPart.Cutbacks(SPSMemberAxisEnd).count > 0 Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
        End If
    Else
        If Not pMemberPart.Cutbacks(SPSMemberAxisEnd).Contains(oSupportSurface) Then
            RemoveAllCutbacksOnMemPart pMemberPart  ' tr 50065  ' tr 52153
            pMemberPart.AddCutbackSurface SPSMemberAxisEnd, oSupportSurface
        End If
    End If

    tmpX = yVec.x
    tmpY = yVec.y
    tmpZ = yVec.z
    
    'Set the Rotation angle of the member based on the rotation of the equipment. -- snkota (TR CP 42798)
'    pMemberPart.Rotation.BetaAngle = 2 * PI - m_dAngleRotated
    pMemberPart.Rotation.SetOrientationVector -tmpX, -tmpY, -tmpZ
    pMemberPart.Rotation.Mirror = True
    
    Set oLine3dPart = Nothing
    Set pMemberPart = Nothing
    Set pIJOccAttribs = Nothing
    Set oSmartOccCAO = Nothing
    Set pIRCAsm = Nothing
    Set eqPos = Nothing
    Set xVec = Nothing
    Set yVec = Nothing
    Set oMaterialDefinition = Nothing
    Set oCrossSection = Nothing
    Set oSupportSurface = Nothing

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' tr 50065 - remove all existing cutbacks if any at member part AxisEnd
Private Sub RemoveAllCutbacksOnMemPart(pMemberPart As SPSMembers.ISPSMemberPartPrismatic)
Const METHOD = "RemoveAllCutbacksOnMemPart"
On Error GoTo ErrorHandler

    Dim oPartGenerator As Object
    Set oPartGenerator = pMemberPart.Generator
    
    Dim pRelationHelper As IMSRelation.DRelationHelper
    Dim pRelColl As IJDRelationshipCol
        
    Set pRelationHelper = oPartGenerator
        
    Set pRelColl = pRelationHelper.CollectionRelations("{5205F933-09DD-4F6E-8454-9BE22C31E960}", "EndTrimSurface_ORIG")

    If pRelColl.count > 0 Then
        Dim pRevision As IJRevision
        Set pRevision = New JRevision
        
        Dim i As Integer
        Dim pRelatioShip As Object
        
        For i = 1 To pRelColl.count
            Set pRelatioShip = pRelColl.Item(i)
            pRevision.RemoveRelationship pRelatioShip
            Set pRelatioShip = Nothing
        Next i
    End If
        
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

 
        
' Side Plate 1 methods
Public Sub CMConstructSidePlate1(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObj As Object)
Const METHOD = "CMConstructSidePlate1"
On Error GoTo ErrorHandler

    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim strMaterial As String, strGrade As String
    Dim oPlateSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oPlate As IJStructCustomPlatePart
    Dim oCustomPlatePartFactory As IJStructCustomPlatePartFactory
    Dim SidePlateThickness As Double
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    
    Set oCustomPlatePartFactory = New StructCustomPlatePartFactory
    
    'Create the custom Plate, that is the Top Plate.
    Set oPlate = oCustomPlatePartFactory.CreateCustomPlatePart(pResourceManager)

    Set oAttrbs = oSmartOccCAO.ItemObject
    
    ' get material and grade
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_GRADE_ATTRNAME).Value
   
    SetPlateMaterial oPlate, strMaterial, strGrade
   
    'get default from the parent
    SidePlateThickness = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_THICKNESS_ATTRNAME).Value
    
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, SidePlateThickness

    'set the found (from the available in catalog) thickness on attrbs
    oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value = SidePlateThickness

    Set oPlateSmartOcc = oPlate
    
    ' Initialize the SmartOccurrence
    oPlateSmartOcc.RootSelectorClass = ROOTSELECTOR_CLASS
    
    oPlateSmartOcc.ROOTSELECTION = ROOTSELECTION
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oPlate
    Set oAssemblyparent = Nothing
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oPlateSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
    
    Set oPlate = Nothing
    Set oCustomPlatePartFactory = Nothing
    Set oPlateSmartOcc = Nothing
    Set oRefColl = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputSidePlate1(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputSidePlate1"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructSidePlate1(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructSidePlate1"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseSidePlate1(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseSidePlate1"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateSidePlate1(pPropertyDescriptions As IJDPropertyDescription, _
                                pObject As Object)
Const METHOD = "CMEvaluateSidePlate1"
On Error GoTo ErrorHandler
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' Side Plate 2 methods
Public Sub CMConstructSidePlate2(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObj As Object)
Const METHOD = "CMConstructSidePlate2"
On Error GoTo ErrorHandler

    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim strMaterial As String, strGrade As String
    Dim oPlateSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oPlate As IJStructCustomPlatePart
    Dim oCustomPlatePartFactory As IJStructCustomPlatePartFactory
    Dim SidePlateThickness As Double
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    
    Set oCustomPlatePartFactory = New StructCustomPlatePartFactory
    
    'Create the custom Plate, that is the Top Plate.
    Set oPlate = oCustomPlatePartFactory.CreateCustomPlatePart(pResourceManager)

    Set oAttrbs = oSmartOccCAO.ItemObject
    
    ' get material and grade
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_GRADE_ATTRNAME).Value
   
    SetPlateMaterial oPlate, strMaterial, strGrade
   
    'get default from the parent
    SidePlateThickness = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_THICKNESS_ATTRNAME).Value
    
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, SidePlateThickness

    'set the found (from the available in catalog) thickness on attrbs
    oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value = SidePlateThickness

    Set oPlateSmartOcc = oPlate
    
    ' Initialize the SmartOccurrence
    oPlateSmartOcc.RootSelectorClass = ROOTSELECTOR_CLASS
    
    oPlateSmartOcc.ROOTSELECTION = ROOTSELECTION '"GenericPlatePart1"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oPlate
    Set oAssemblyparent = Nothing
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oPlateSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
    
    Set oPlate = Nothing
    Set oCustomPlatePartFactory = Nothing
    Set oPlateSmartOcc = Nothing
    Set oRefColl = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputSidePlate2(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputSidePlate2"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructSidePlate2(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructSidePlate2"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseSidePlate2(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseSidePlate2"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateSidePlate2(pPropertyDescriptions As IJDPropertyDescription, _
                                pObject As Object)
Const METHOD = "CMEvaluateSidePlate2"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' Side Plate 3 methods
Public Sub CMConstructSidePlate3(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObj As Object)
Const METHOD = "CMConstructSidePlate3"
On Error GoTo ErrorHandler

    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim strMaterial As String, strGrade As String
    Dim oPlateSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oPlate As IJStructCustomPlatePart
    Dim oCustomPlatePartFactory As IJStructCustomPlatePartFactory
    Dim SidePlateThickness As Double
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    
    Set oCustomPlatePartFactory = New StructCustomPlatePartFactory
    
    'Create the custom Plate, that is the Top Plate.
    Set oPlate = oCustomPlatePartFactory.CreateCustomPlatePart(pResourceManager)

    Set oAttrbs = oSmartOccCAO.ItemObject
    
    ' get material and grade
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_GRADE_ATTRNAME).Value
   
    SetPlateMaterial oPlate, strMaterial, strGrade
   
    'get default from the parent
    SidePlateThickness = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_THICKNESS_ATTRNAME).Value
    
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, SidePlateThickness

    'set the found (from the available in catalog) thickness on attrbs
    oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value = SidePlateThickness

    Set oPlateSmartOcc = oPlate
    
    ' Initialize the SmartOccurrence
    oPlateSmartOcc.RootSelectorClass = ROOTSELECTOR_CLASS
    
    oPlateSmartOcc.ROOTSELECTION = ROOTSELECTION '"GenericPlatePart1"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oPlate
    Set oAssemblyparent = Nothing
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oPlateSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
    
    Set oPlate = Nothing
    Set oCustomPlatePartFactory = Nothing
    Set oPlateSmartOcc = Nothing
    Set oRefColl = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputSidePlate3(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputSidePlate3"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructSidePlate3(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructSidePlate3"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseSidePlate3(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseSidePlate3"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateSidePlate3(pPropertyDescriptions As IJDPropertyDescription, _
                                pObject As Object)
Const METHOD = "CMEvaluateSidePlate3"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' Side Plate 4 methods
Public Sub CMConstructSidePlate4(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObj As Object)
Const METHOD = "CMConstructSidePlate4"
On Error GoTo ErrorHandler

    Dim oAttrbs As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim strMaterial As String, strGrade As String
    Dim oPlateSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oDesignParent As IJDesignParent
    Dim oPlate As IJStructCustomPlatePart
    Dim oCustomPlatePartFactory As IJStructCustomPlatePartFactory
    Dim SidePlateThickness As Double
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    
    Set oCustomPlatePartFactory = New StructCustomPlatePartFactory
    
    'Create the custom Plate, that is the Top Plate.
    Set oPlate = oCustomPlatePartFactory.CreateCustomPlatePart(pResourceManager)

    Set oAttrbs = oSmartOccCAO.ItemObject
    
    ' get material and grade
    strMaterial = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value
    strGrade = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_GRADE_ATTRNAME).Value
   
    SetPlateMaterial oPlate, strMaterial, strGrade
   
    'get default from the parent
    SidePlateThickness = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_THICKNESS_ATTRNAME).Value
    
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, SidePlateThickness

    'set the found (from the available in catalog) thickness on attrbs
    oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(SIDEPLATE_MATERIAL_ATTRNAME).Value = SidePlateThickness

    Set oPlateSmartOcc = oPlate
    
    ' Initialize the SmartOccurrence
    oPlateSmartOcc.RootSelectorClass = ROOTSELECTOR_CLASS
    
    oPlateSmartOcc.ROOTSELECTION = ROOTSELECTION '"GenericPlatePart1"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild oPlate
    Set oAssemblyparent = Nothing

    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oPlateSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
    
    Set oPlate = Nothing
    Set oCustomPlatePartFactory = Nothing
    Set oPlateSmartOcc = Nothing
    Set oRefColl = Nothing
    Set oSmartOccCAO = Nothing
    Set oDesignParent = Nothing
    Set oSymbolEntitiesFactory = Nothing
    
    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputSidePlate4(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputSidePlate4"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructSidePlate4(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructSidePlate4"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseSidePlate4(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseSidePlate4"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateSidePlate4(pPropertyDescriptions As IJDPropertyDescription, _
                                pObject As Object)
Const METHOD = "CMEvaluateSidePlate4"
On Error GoTo ErrorHandler

    Exit Sub
   
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub CalculateEqpPortandPlatePos(pIRCAsm As IJDReferencesCollection, _
                                        oAttrbs As IJDAttributes, _
                                        oDefAttrbs As IJDAttributes, _
                                        ByRef eqPos As DPosition, _
                                        ByRef xVec As DVector, _
                                        ByRef yVec As DVector, _
                                        pSupportSurface As IJPlane)
Const METHOD = "CalculateEqpPortandPlatePos"
On Error GoTo ErrorHandler

    Dim pIRCAsm1 As IJDReferencesCollection
    Dim pIRCAsm2 As IJDReferencesCollection
    Dim FoundationPorts As IJElements
    Dim SupportPlane As IJPlane
    Dim oPort As Object
    Dim rc1cnt As Integer, rc2cnt As Integer
    Dim rccnt As Integer, i As Integer
    Dim bFoundPorts As Boolean, bFoundSuppSurface As Boolean
    Dim Trans As IngrGeom3D.IJDT4x4
    Dim HoleLocations() As Automath.DPosition
    Dim NumberOfHoles As Integer
    Dim newTmx As Automath.DT4x4
    
    rccnt = pIRCAsm.IJDEditJDArgument.GetCount
    Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
    On Error Resume Next
    Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
    
    Set FoundationPorts = New JObjectCollection 'Elements
    
    Dim pEnumJDArgument As IEnumJDArgument
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim iloop As Long
     
    Set pEnumJDArgument = pIRCAsm1
    If Not pEnumJDArgument Is Nothing Then
        pEnumJDArgument.Reset
        iloop = 1
        Do
           pEnumJDArgument.Next 1, arg1, found
           If found <> 0 Then
                Set oPort = arg1.Entity
                FoundationPorts.Add oPort
                Set oPort = Nothing
                Set arg1 = Nothing
                iloop = iloop + 1
           Else: Exit Do
           End If
        Loop
    End If
    
    ' tr 77250 raise an error when all fnd port inputs are deleted
    If iloop = 0 Then
        On Error GoTo 0
        SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFND_NOTASSOC_WITHEQP, oAttrbs, Nothing
        Err.Raise SPS_MACRO_WARNING 'E_FAIL
    End If
    
    Dim bSinglePt As Boolean
    bSinglePt = False
    If FoundationPorts.count = 0 Then
        bSinglePt = True
    End If
    
    If iloop >= 1 Then bFoundPorts = True
    
    rc2cnt = 0
    If pIRCAsm2 Is Nothing Then Else rc2cnt = pIRCAsm2.IJDEditJDArgument.GetCount
    
    If rc2cnt >= 1 Then
        bFoundSuppSurface = True
        Set SupportPlane = pIRCAsm2.IJDEditJDArgument.GetEntityByIndex(1)
        Set pSupportSurface = SupportPlane
    End If
    
    ' get the SizingByRule attribute
    Dim bPlateSizeByRule As Boolean
    bPlateSizeByRule = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(PLATESIZEBYRULE_ATTRNAME).Value
    Dim plateThick As Double
    plateThick = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
    
    If bSinglePt Then
        eqPos.z = eqPos.z + plateThick
        Set xVec = New Automath.DVector
        Set yVec = New Automath.DVector
        yVec.y = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item("TopPlateLength").Value
        xVec.x = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item("TopPlateWidth").Value
    Else
        ' calculate holes from foundation ports positions
        Call GetTransformAndHoles(HoleLocations(), NumberOfHoles, FoundationPorts, Trans)
        
        Set newTmx = New Automath.DT4x4
        eqPos.Set 0, 0, 0
        
        On Error GoTo PORTS_TOO_CLOSE
        Call CalculateRectangleOrientation(HoleLocations(), xVec, yVec)
        On Error GoTo ErrorHandler
    
    
        Set eqPos = Trans.TransformPosition(eqPos)
        Set xVec = Trans.TransformVector(xVec)
        Set yVec = Trans.TransformVector(yVec)
        
        Dim dOffX As Double, dOffY As Double
        If bPlateSizeByRule = False And oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value <> "" And oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value <> "" Then
'            tr 70323 - this is not needed and giving wrong results as we were checking with def attributes
'            top plate length and width will always be same as these vectors.
'            This logic need to be changed when attrs can be modified by rule/user defined like in footings
'            dOffX = oDefAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_WIDTH_ATTRNAME).Value - xVec.length
'            dOffY = oDefAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_LENGTH_ATTRNAME).Value - yVec.length
        End If
        
        eqPos.Set 0, 0, 0
        Set xVec = Nothing
        Set yVec = Nothing
        Set xVec = New Automath.DVector
        Set yVec = New Automath.DVector

        ' get the clearance attribute value
        Dim dClear As Double
        dClear = 0#
        If bPlateSizeByRule = False And oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(PLATEEDGECLEAR_ATTRNAME).Value <> "" Then
            dClear = oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(PLATEEDGECLEAR_ATTRNAME).Value
        End If
        
        On Error GoTo PORTS_TOO_CLOSE
        SetRectangleOrientationByClearance HoleLocations(), dOffX + dClear, dOffY + dClear, xVec, yVec
        On Error GoTo ErrorHandler
        
        Set eqPos = Trans.TransformPosition(eqPos)
        Set xVec = Trans.TransformVector(xVec)
        Set yVec = Trans.TransformVector(yVec)
        
        
        'Get the center of the plate using which we will place the frames. -- snkota (TR CP 42798)
        Dim cenX As Double, cenY As Double, cenZ As Double
        GetCentroidOfPositions HoleLocations(), NumberOfHoles, cenX, cenY, cenZ
        eqPos.Set 0, 0, 0
        eqPos.x = cenX
        eqPos.y = cenY
        eqPos.z = cenZ + plateThick
        Set eqPos = Trans.TransformPosition(eqPos)
        
        Dim zVec As DVector
        Set zVec = yVec.Cross(xVec)
        zVec.Length = 1
        ' In case the supported objects(equipments) are not in same heights then adjust the EF also
        ' so that the EF should start from the bottom most point in the normal direction.
        ' therefore the origin of the EF should be shifted to the lower most point along the normal.
        
        ' Identify the maximum distance between the origin and all the holes along normal
        Dim dShift As Double
        dShift = GetOriginShiftForUnevenSupportedObjects(HoleLocations, eqPos, zVec, Trans)
        dShift = dShift - oAttrbs.CollectionOfAttributes(FRAMEFNDASM_IFACE).Item(TOPPLATE_THICKNESS_ATTRNAME).Value
        ' Shift the origin along the vector by the magnitude of maximum difference.
        Dim dTol As Double
        dTol = 0.0001
        If dShift - 0# > dTol Then
        zVec.Length = 1
        zVec.Set zVec.x * dShift, zVec.y * dShift, zVec.z * dShift
        eqPos.Set eqPos.x + zVec.x, eqPos.y + zVec.y, eqPos.z + zVec.z
        End If
        
    End If
    

    
    If Not FoundationPorts Is Nothing Then
        FoundationPorts.Clear
        Set FoundationPorts = Nothing
    End If
    
    Set pIRCAsm1 = Nothing
    Set pIRCAsm2 = Nothing
    
    Erase HoleLocations
    
    Exit Sub
PORTS_TOO_CLOSE:
    SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFND_CLOSEFOR_SELECTED, oAttrbs, Nothing
    Err.Raise E_FAIL
    Exit Sub

ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructInterfaces.IJElements)
    Dim i As Integer
    Dim cnt As Integer
    Dim oItem As IJSmartItem
    Dim oSmartOcc As IJSmartOccurrence
    Dim FndDefServices As ISPSEquipFndnDefServices
    Dim MemberObj As IJDMemberObjects
    Dim pObjectCollection As IJDObjectCollection
    Dim pDesignParent As IJDesignParent
    Dim object As Object
    
    Set oSmartOcc = pPartObject
    Set oItem = oSmartOcc.ItemObject
    Set MemberObj = oSmartOcc
    Set FndDefServices = SP3DCreateObject(oItem.definition)

    Dim Supported As IJElements 'IMSElements.IJElements
    Set Supported = New JObjectCollection 'IMSElements.Elements
    Dim Supporting As IJElements ' IMSElements.Elements
    Set Supporting = New JObjectCollection ' IMSElements.Elements

    FndDefServices.GetInputs pPartObject, Supported, Supporting
    
    Dim oFndPort As IJPort
    Dim oEquipment As IJEquipment
    
    For i = 1 To Supported.count
        Set oFndPort = Supported.Item(i)
        If Not oFndPort Is Nothing Then
            Set oEquipment = oFndPort.Connectable
            If Not oEquipment Is Nothing Then
                pIJMonUnks.Add oEquipment
            End If
        End If
        Set oFndPort = Nothing
        Set oEquipment = Nothing
    Next i
    For i = 1 To MemberObj.count
        On Error Resume Next
        Set pDesignParent = MemberObj.ItemByDispid(i)
        If Not pDesignParent Is Nothing Then
            Set pObjectCollection = New JObjectCollection
            pDesignParent.GetChildren pObjectCollection, vbNullString
            ' tr 66565 - should check for list count otherwise add each component
            If pObjectCollection.count > 0 Then
                For Each object In pObjectCollection
                    pIJMonUnks.Add object
                Next
            Else
                pIJMonUnks.Add MemberObj.ItemByDispid(i)
            End If
            pObjectCollection.Clear
            Set pObjectCollection = Nothing
            Set pDesignParent = Nothing
        Else
            pIJMonUnks.Add MemberObj.ItemByDispid(i)
        End If
    Next i
    
    Dim pUnk As Object
    Dim oSupportingSurface As Object
    Dim oRefProxy As IJDReferenceProxy
    
    If Supporting.count > 0 Then
        Set oSupportingSurface = Supporting.Item(1)
        If Not oSupportingSurface Is Nothing Then
            On Error Resume Next
            Set oRefProxy = oSupportingSurface

            If Not oRefProxy Is Nothing Then
                Set pUnk = oRefProxy.Reference
                pIJMonUnks.Add pUnk
                Set pUnk = Nothing
            Else ' for slab surfaces
                Dim oStructPort As IJPort
                On Error Resume Next
                Set oStructPort = oSupportingSurface
                Set pUnk = oStructPort.Connectable
                If Not pUnk Is Nothing Then
                    pIJMonUnks.Add pUnk
                End If
                Set pUnk = Nothing
            End If
            Set oSupportingSurface = Nothing
            Set oRefProxy = Nothing
        End If
                
    End If
End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = StandardGraphicEntity
End Sub

Private Sub ISPSEquipFndnDefServices_GetInputs(ByVal FndObject As Object, ByVal Supported As SP3DSPSEquipFoundations.IJElements, ByVal Supporting As SP3DSPSEquipFoundations.IJElements)
Const METHOD = "ISPSEquipFndnDefServices_GetInputs"
On Error GoTo ErrorHandler

    Call GetInputs_Supported_Supporting(FndObject, Supported, Supporting)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub ISPSEquipFndnDefServices_SetInputs(ByVal FndObject As Object, ByVal FndDefnObject As Object, ByVal Supported As SP3DSPSEquipFoundations.IJElements, ByVal Supporting As SP3DSPSEquipFoundations.IJElements)
Const METHOD = "ISPSEquipFndnDefServices_SetInputs"
On Error GoTo ErrorHandler

    Call SetInputs_Supported_Supporting(FndObject, FndDefnObject, Supported, Supporting)
    
    Exit Sub
    
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_EQPFNDMACROS_ERROR, "ERROR")
    
    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long
    
    If m_bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
'            IJUserAttributeMgmt_OnAttributeChange = "ERROR::Bad Value"
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
        
    If (pAttrToChange.attrName = "IsTopPlateSizeDrivenByRule") Then
        If (varNewAttrValue = True) Then 'User defined option for sizing rule
            'gray out the top plate length, width and thickness on the GOPC
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If pAttrDescr.attrName = "TopPlateEdgeClearance" Or pAttrDescr.attrName = "TopPlateLength" Or pAttrDescr.attrName = "TopPlateWidth" Then
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
                If pAttrDescr.attrName = "TopPlateEdgeClearance" Or pAttrDescr.attrName = "TopPlateLength" Or pAttrDescr.attrName = "TopPlateWidth" Then
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

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_EQPFNDMACROS_ERROR, "ERROR")
    m_bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim ErrStr As String
    Dim pSmartItem As IJSmartItem
    Dim pSO As IJSmartOccurrence
    Dim pAttrs As IJDAttributes
    Dim attrName As String
    
    Set pAttrColl = CollAllDisplayedValues
       
    Dim NumSupportingPlanes As String
    Set pSO = pIJDAttrs
    Set pSmartItem = pSO.ItemObject
    Set pAttrs = pSmartItem
    NumSupportingPlanes = pAttrs.CollectionOfAttributes("ISPSEquipFndInputCriteria").Item("NumberSupporting").Value
    
    ' check if any support plane defined
    Dim Supported As IJElements
    Set Supported = New JObjectCollection
    Dim Supporting As IJElements
    Set Supporting = New JObjectCollection
    Call GetInputs_Supported_Supporting(pSO, Supported, Supporting)
                    
    ' need to gray out the block height if it has a constraining plane
    If NumSupportingPlanes <> "0" Then
        For i = 1 To pAttrColl.count
            Set pAttrDescr = pAttrColl.Item(i)
            If (pAttrDescr.attrName = "LegAngleHeight") Then
                If Supporting.count >= 1 Then
                    pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                Else
                    pAttrDescr.AttrState = Not (AttributeDescriptor_ReadOnly)
                End If
                Exit For
            End If
        Next
    End If
    
    If Supported.count > 0 Then
        If TypeOf Supported.Item(1) Is IJPoint Then 'Multiple point case
            For i = 1 To pAttrColl.count
                Set pAttrDescr = pAttrColl.Item(i)
                If pAttrDescr.attrName = "IsTopPlateSizeDrivenByRule" Then
                    pAttrDescr.AttrValue = False
                    pAttrDescr.AttrState = Not (AttributeDescriptor_ReadOnly)
                End If
            Next
'        Else 'Equipment case
'            For i = 1 To pAttrColl.count
'                Set pAttrDescr = pAttrColl.Item(i)
'                If pAttrDescr.attrName = "IsTopPlateSizeDrivenByRule" Then
'                    pAttrDescr.AttrValue = True
'                    pAttrDescr.AttrState = Not (AttributeDescriptor_ReadOnly)
'                End If
'            Next
        End If
    Else 'single point case
        For i = 1 To pAttrColl.count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.attrName = "IsTopPlateSizeDrivenByRule" Then
                pAttrDescr.AttrValue = False
                pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
            End If
        Next
    End If
    
    Set pSO = Nothing
    Set Supported = Nothing
    Set Supporting = Nothing
    
    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
            If (pAttrDescr.attrName = "SidePlateSPSMaterial") Or (pAttrDescr.attrName = "SidePlateSPSGrade") Then
                    pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
            End If
            If (pAttrDescr.attrName = "SidePlateLength") Or (pAttrDescr.attrName = "SidePlateWidth") Or (pAttrDescr.attrName = "SidePlateThickness") Then
                    pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
            End If
            If (pAttrDescr.attrName = "TopPlateLength") Or (pAttrDescr.attrName = "TopPlateWidth") Then
                    pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
            End If
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                m_bOnPreLoad = False
                Exit Function
            End If
    Next
    
    m_bOnPreLoad = False

    IJUserAttributeMgmt_OnPreLoad = ""
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler

' first of all check if the symbol definition has CMCheck methods defined - TBD
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_EQPFNDMACROS_ERROR, "ERROR")

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

    If oAttribute.Value <> "" Then
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
'    Dim oAttrObj As IJDAttribute
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
'    Dim oAttrCol As IJDAttributesCol
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
    
'    Set oAttributeMetaData = Nothing
    Set oInterfaceInfo = Nothing
    
    If IsInterfaceFound = False Then
        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_EQPFNDMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
'    Set oAttrCol = pIJDAttrs.CollectionOfAttributes(InterfaceID)
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.name = sAttributeName Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (sAttributeName = "TopPlateEdgeClearance") And (varAttributeValue < 0#) Then
                            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_EQPFNDMACROS_INVALID_ATTRIB_VALUE, "Invalid Negative Attribute Value ")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        ElseIf (sAttributeName <> "TopPlateEdgeClearance") And (varAttributeValue <= 0#) Then
'                            UserAttributeMgmt_Validate = sAttributeName
                            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_EQPFNDMACROS_INVALID_ATTRIB_VALUE, "Invalid Negative Attribute Value ")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
            End Select
        End If
    Next
    
    UserAttributeMgmt_Validate = ""
    Set oAttributeMetaData = Nothing
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

'TR#71850
Private Property Get ISPSFoundationInputHelper_ValidateObjects(ByVal inputSupported As Object, ByVal inputSupporting As Object, ByVal SupportedObjList As SP3DStructInterfaces.IJElements, ByVal ObjsinSelectSet As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFoundationInputHelperStatus
Const MT = "ISPSFoundationInputHelper_ValidateObjects"
On Error GoTo ErrorHandler
    
If inputSupporting Is Nothing Then
    ISPSFoundationInputHelper_ValidateObjects = HasFoundationPort(inputSupported)
Else
    ISPSFoundationInputHelper_ValidateObjects = IsValidPlane(inputSupporting, SupportedObjList, ObjsinSelectSet)
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
'*************************************************************************
'Function
'CMMigrateAggregator
'
'Abstract
'Migrates the foundation to the correct surface if it is split.
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
  MigrateEqpFnd oAggregatorDesc, oMigrateHelper

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
