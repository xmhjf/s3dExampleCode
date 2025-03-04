VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlockSlabFndAsmDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************
' Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File
'    BlockSlabFndAsmDef.cls
'
'Author
'       1-Mar-03        SS
'
'Description
'
'Notes
'
'History:
'   SS     13-Sep-04   React to CmnApp SmartOcc impact for
'                           TR 59875 - CAODisplayCES semantic adds more to the graph
'                           adding new relation interface for children
' 06/06/06   AS      Added CMMigrate method stubs
' 24-Jun-2006 JMS DI#60069 - Changes to allow editing of the weight and CG values
'                   changed call to SetWCG to call new interface to put weight and CG
'                   values since SetWCG is reservered for setting user defined values
'                   when the values here are the computed values
'       07-Jul-2006 JMS TR#101063 - Tolerate the nonexistence of IJWCGValueOrigin interface
' 04-Aug-2006 AS    TR#99968 Added new function to handle EqpFnd Migration
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
'  01.May.2014  RUK TR-CP-253694  ATP - 81 minidumps at 'SPSEqpFndMacros' module. => Stack for BlockFndDef.cls @ 431
'*******************************************************************

Private Const MODULE = "BlockSlabFndAsmDef"
Private Const strSourceFile = "BlockSlabFndAsmDef.def"
Private Const CONST_ItemProgId As String = "SPSEqpFndMacros.BlockSlabFndAsmDef"
Private Const MODELDATABASE = "Model"
Private Const NUMKNOWNINPUTS = 3

Private Const BLOCKANDSLABFNDASM_IFACE = "IJUASPSBlockAndSlabFndnAsm"
Private Const BLOCKFND_IFACE = "IJUASPSBlockFndn"
Private Const BLOCK_COMP_NAME = "BlockComponent"
Private Const SLAB_COMP_NAME = "SlabComponent"
Private Const MATERIAL_ATTRNAME = "BlockSPSMaterial"
Private Const GRADE_ATTRNAME = "BlockSPSGrade"
Private Const LENGTH_ATTRNAME = "BlockLength"
Private Const WIDTH_ATTRNAME = "BlockWidth"
Private Const HT_ATTRNAME = "BlockHeight"
Private Const BLOCKSIZEBYRULE_ATTRNAME = "IsBlockSizeDrivenByRule"
Private Const BLOCKEDGECLEAR_ATTRNAME = "BlockEdgeClearance"

Private Const WITH_SLAB_ATTRNAME = "WithSlab"

Private Const SLABFND_IFACE = "IJUASPSSlabFndn"
Private Const SLABLENGTH_ATTRNAME = "SlabLength"
Private Const SLABWIDTH_ATTRNAME = "SlabWidth"
Private Const SLABHT_ATTRNAME = "SlabHeight"
Private Const SLABMAT_ATTRNAME = "SlabSPSMaterial"
Private Const SLABGR_ATTRNAME = "SlabSPSGrade"
Private Const SLABSIZEBYRULE_ATTRNAME = "IsSlabSizeDrivenByRule"
Private Const SLABEDGECLEAR_ATTRNAME = "SlabEdgeClearance"
Private Const INTERFACE_WCGValueOrigin As String = "IJWCGValueOrigin"
Private Const PROPERTY_DryWCGOrigin As String = "DryWCGOrigin"

Private Enum enumWeightCGDerivation
    WEIGHTCG_Computed = 2
    WEIGHTCG_UserDefined = 4
End Enum

Const FNDNCOMPPROGID = "SPSEquipFoundations.SPSEquipFoundation"

Private Const DOUBLE_VALUE = 8
Private Const CHAR = 1
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

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                             ByVal defParams As Variant, _
                                                             ByVal pResourceMgr As Object) As Object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler
     
     ' This method is in charge of the creation of the symbol definition object
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
    pAPDs.RemoveAll         ' Remove all the previous property descriptions

    'BLOCKATTR_IFACE
    pAPDs.AddProperty BLOCKANDSLABFNDASM_IFACE, 1, BLOCKANDSLABFNDASM_IFACE, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
    pAPDs.AddProperty "IJWeightCG", 2, IJWeightCG, "CMEvaluateCAOWCG", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
                                        
    Set pAPDs = Nothing
    
    Dim pMemberDescriptions As IJDMemberDescriptions
    Dim pMemberDescription As IJDMemberDescription
    Dim pPropertyDescriptions As IJDPropertyDescriptions
    Set pMemberDescriptions = pDefinition
    
    ' Remove all the previous member descriptions
    pMemberDescriptions.RemoveAll
    
    Set pMemberDescription = pMemberDescriptions.AddMember("Block", 1, "CMConstructBlock", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputBlock"
    pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructBlock"
    pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseBlock"
    pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
    Set pPropertyDescriptions = pMemberDescription
    pPropertyDescriptions.AddProperty "BlockProperties", 1, IJDAttributes, "CMEvaluateBlock", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "BlockMaterialProp", 2, IJStructMaterial, "CMEvaluateBlockMaterial", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "BlockGeometry", 3, IJGeometry, "CMEvaluateBlockGeometry", imsCOOKIE_ID_USS_LIB       ' tr 50305

    Set pMemberDescription = pMemberDescriptions.AddMember("Slab", 2, "CMConstructSlab", imsCOOKIE_ID_USS_LIB)
    pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSlab"
    pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSlab"
    pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructSlab"
    pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseSlab"
    pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
    
    Set pPropertyDescriptions = pMemberDescription
    pPropertyDescriptions.AddProperty "SlabProperties", 1, IJDAttributes, "CMEvaluateSlab", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "SlabMaterialProp", 2, IJStructMaterial, "CMEvaluateSlabMaterial", imsCOOKIE_ID_USS_LIB
    pPropertyDescriptions.AddProperty "SlabGeometry", 3, IJGeometry, "CMEvaluateSlabGeometry", imsCOOKIE_ID_USS_LIB     ' tr 50305

    Set pMemberDescriptions = Nothing
    Set pMemberDescription = Nothing
    Set pPropertyDescriptions = Nothing
     
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

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

Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO"
On Error GoTo ErrorHandler
    Dim pIRCAsm As IJDReferencesCollection
    Dim pIRCAsm1 As IJDReferencesCollection
    Dim FoundationPorts As IJElements
    Dim oSmartOcc As IJSmartOccurrence
    Dim oDefAttribs As IJDAttributes
    Dim Trans As IJDT4x4
    Dim HoleLocations() As Automath.DPosition
    Dim NumberOfHoles As Integer
    Dim newLength As Double, newWidth As Double
    Dim xVec As New DVector, yVec As New DVector, zVec As New DVector
    Dim newTmx As New DT4x4
    Dim eqPos As New DPosition
    Dim pEnumJDArgument As IEnumJDArgument
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim iloop As Long
    Dim bBlockSizeByRule As Boolean
    Dim strBlockComp As String
    Dim dClear As Double
    Dim dblOffX As Double, dblOffY As Double
    Dim pOcc As IJDOccurrence
    Dim vProp As Variant
    Dim bPtOption As Boolean
    
    'this methos sets the occurrence matrics for the equipment foundation
    
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIRCAsm = GetRefCollection(oSmartOcc)

    Set FoundationPorts = New JObjectCollection 'Elements
    
    'TR-CP-253694 -> Added the if condition to check the object is not nothing
    If Not pIRCAsm Is Nothing Then
        If Not pIRCAsm.IJDEditJDArgument Is Nothing Then
            If pIRCAsm.IJDEditJDArgument.GetCount > 0 Then
                Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
                Set pEnumJDArgument = pIRCAsm1
        
                Call GetPortAndHoleLocations(pEnumJDArgument, FoundationPorts, HoleLocations, bPtOption)
            End If
        End If
    End If
    
    Set oDefAttribs = oSmartOcc.ItemObject
    'get name of the catalog part for the block
    strBlockComp = oDefAttribs.CollectionOfAttributes(BLOCKANDSLABFNDASM_IFACE).Item(BLOCK_COMP_NAME).Value
    Set oDefAttribs = GetDefinition(strBlockComp)
    
    ' get the SizingByRule attribute of the block
    bBlockSizeByRule = oDefAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(BLOCKSIZEBYRULE_ATTRNAME).Value
    
    dClear = 0#
    vProp = oDefAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(BLOCKEDGECLEAR_ATTRNAME).Value
    If bBlockSizeByRule = False And vProp <> vbEmpty Then
        dClear = vProp
    End If
    
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    
    Dim dx, dy, dz As Double
    If UBound(HoleLocations) >= 1 Or bPtOption Then
        Dim oFndnMatrix As IJDMatrixAccess
        Set oFndnMatrix = pPropertyDescriptions.CAO
        Set newTmx = oFndnMatrix.Matrix
        Dim oFndn As IJLocalCoordinateSystem
        Set oFndn = pPropertyDescriptions.CAO
        Dim oPosition As DPosition
        Set oPosition = oFndn.Position
        Dim oPoint As IJPoint
        Dim xX As Double, yy As Double, zZ As Double
        xX = oPosition.x
        yy = oPosition.y
        zZ = oPosition.z
        
        If UBound(HoleLocations) >= 1 Then ' multiple point case commented for now
'            On Error GoTo PORTS_TOO_CLOSE
'            Call CalculateRectangleOrientation(HoleLocations(),  xVec, yVec)
'            On Error GoTo ErrorHandler
'            GetCentroidOfPositions  HoleLocations(), UBound(HoleLocations), xX, yy, zZ
'            newTmx.IndexValue(12) = xX
'            newTmx.IndexValue(13) = yy
'            newTmx.IndexValue(14) = zZ
        End If
                        
   Else
        ' calculate holes and get the transformation matrix from foundation ports
        Call GetTransformAndHoles(HoleLocations(), NumberOfHoles, FoundationPorts, Trans)
        
        'calculate the centroid of all the holes of all the foundation ports
        Dim centroidX As Double, centroidY As Double, centroidZ As Double
        GetCentroidOfPositions HoleLocations(), NumberOfHoles, centroidX, centroidY, centroidZ
        eqPos.Set centroidX, centroidY, centroidZ
        
        ' Calculate the length vector and width vector based on the hole locations.
        On Error GoTo PORTS_TOO_CLOSE
        SetRectangleOrientationByClearance HoleLocations(), 0, 0, xVec, yVec
        On Error GoTo ErrorHandler
        
        'Transform the centre and the directions to global
        Set eqPos = Trans.TransformPosition(eqPos)
        Set xVec = Trans.TransformVector(xVec)
        Set yVec = Trans.TransformVector(yVec)
        Set zVec = yVec.Cross(xVec)
        newWidth = xVec.Length
        newLength = yVec.Length
        
        xVec.Length = 1
        yVec.Length = 1
        zVec.Length = 1
        
        ' In case the supported objects(equipments) are not in same heights then adjust the EF also
        ' so that the EF should start from the bottom most point in the normal direction.
        ' therefore the origin of the EF should be shifted to the lower most point along the normal.
        
        ' Identify the maximum distance between the origin and all the holes along normal
        Dim dShift As Double
        dShift = GetOriginShiftForUnevenSupportedObjects(HoleLocations, eqPos, zVec, Trans)
        
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
    Set oSmartOcc = Nothing
    Set pOcc = Nothing
    Set oFndnMatrix = Nothing
    Set oFndn = Nothing
    
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

Public Sub CMEvaluateCAOWCG(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAOWCG"
On Error GoTo ErrHandler
    
    Exit Sub

ErrHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMConstructBlock(ByVal pMemberDescription As IJDMemberDescription, _
                            ByVal pResourceManager As IUnknown, _
                            ByRef pObj As Object)
Const METHOD = "CMConstructBlock"
On Error GoTo ErrorHandler
 
     Dim oBlockComp As ISPSFoundationComponent
     Dim strBlockComp As String
     
     strBlockComp = GetCAODefAttribute(pMemberDescription, BLOCKANDSLABFNDASM_IFACE, BLOCK_COMP_NAME)
     Call CreateComponent(strBlockComp, pResourceManager, oBlockComp, pMemberDescription, True)
    
     Set pObj = oBlockComp
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputBlock(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputBlock"
On Error GoTo ErrorHandler

    Dim strBlockComp As String
    strBlockComp = GetCAODefAttribute(pMemberDesc, BLOCKANDSLABFNDASM_IFACE, BLOCK_COMP_NAME)
       
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.object
    oSmartOcc.ROOTSELECTION = strBlockComp
    
    ' tr 59726  Modifying a block fnd to a block fnd w/ support plane, support plane wrong
    ' problem cause: new plane input is not getting copied on to the child SO.
    ' hence copying it during the setinput of the child SO
           
    Dim pIRCAsm As IJDReferencesCollection
    Set pIRCAsm = GetRefCollection(pMemberDesc.CAO)

    If Not pIRCAsm Is Nothing Then
                        
        Dim pIRCAsm2 As IJDReferencesCollection
        On Error Resume Next
        Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
        
        Dim oReferencesCollection As IJDReferencesCollection
        On Error Resume Next
        Set oReferencesCollection = GetRefCollection(oSmartOcc)
        
        On Error GoTo ErrorHandler
        If Not oReferencesCollection Is Nothing Then
            If Not pIRCAsm2 Is Nothing Then
                oReferencesCollection.IJDEditJDArgument.SetEntity 2, pIRCAsm2
            End If
        End If
                
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructBlock(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructBlock"
On Error GoTo ErrorHandler

     Call AddSystemAndNameRule(pMemberDesc)
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseBlock(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseBlock"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateBlock(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateBlock"
On Error GoTo ErrorHandler
         
    Dim pIRCAsm As IJDReferencesCollection
    Dim pIRCAsm1 As IJDReferencesCollection
    Dim pIRCAsm2 As IJDReferencesCollection
    Dim FoundationPorts As IJElements
    Dim SupportPlane As IJPlane
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes
    Dim pOcc As IJDOccurrence
    Dim rc1cnt As Integer, rc2cnt As Integer
    Dim rccnt As Integer, i As Integer
    Dim bFoundPorts As Boolean, bFoundSuppSurface As Boolean
    Dim Trans As IngrGeom3D.IJDT4x4
    Dim HoleLocations() As Automath.DPosition
    Dim NumberOfHoles As Integer
    Dim newLength As Double, newWidth As Double
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim newTmx As Automath.DT4x4
    Dim eqPos As Automath.DPosition
    Dim newHeight As Double
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)   ' oSmartOcc ' pPropertyDescriptions.CAO

    rccnt = pIRCAsm.IJDEditJDArgument.GetCount
    Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
    On Error Resume Next
    Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)

    Dim bPtOption As Boolean
    Dim pEnumJDArgument As IEnumJDArgument
    Set pEnumJDArgument = pIRCAsm1
    Set FoundationPorts = New JObjectCollection 'Elements
    Call GetPortAndHoleLocations(pEnumJDArgument, FoundationPorts, HoleLocations, bPtOption)
    
    
    'set the occ attributes
    Set pIJOccAttribs = pObject

    rc2cnt = 0
    If pIRCAsm2 Is Nothing Then Else rc2cnt = pIRCAsm2.IJDEditJDArgument.GetCount

    If rc2cnt >= 1 Then
        bFoundSuppSurface = True
        Set SupportPlane = pIRCAsm2.IJDEditJDArgument.GetEntityByIndex(1)
    End If

    'get WithSlab attr
    Dim oMemberObj As IJDMemberObjects
    Set oMemberObj = pPropertyDescriptions.CAO

    Dim oAttr As IJDAttributes
    Set oAttr = oMemberObj

    Dim bWithSlab As Boolean
    bWithSlab = oAttr.CollectionOfAttributes(BLOCKANDSLABFNDASM_IFACE).Item(WITH_SLAB_ATTRNAME).Value

    Set oMemberObj = Nothing
    Set oAttr = Nothing


    ' get the SizingByRule attribute
    Dim bBlockSizeByRule As Boolean

    bBlockSizeByRule = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(BLOCKSIZEBYRULE_ATTRNAME).Value

    ' get the clearance attribute value
    Dim dClear As Double
    dClear = 0#             ' tr 75005
    If bBlockSizeByRule = False And pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(BLOCKEDGECLEAR_ATTRNAME).Value <> "" Then
        dClear = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(BLOCKEDGECLEAR_ATTRNAME).Value
    End If
    Set newTmx = New Automath.DT4x4
    Set eqPos = New Automath.DPosition
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set zVec = New Automath.DVector
    
    Dim dOffX As Double, dOffY As Double
    dOffX = dOffY = 0#
    
    newLength = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value
    newWidth = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value
    newHeight = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(HT_ATTRNAME).Value
    
    Dim dx, dy, dz As Double
    If UBound(HoleLocations) >= 1 Or bPtOption Then
        Dim oFndnMatrix As IJDMatrixAccess
        Set oFndnMatrix = pPropertyDescriptions.CAO
        Set newTmx = oFndnMatrix.Matrix
        Dim oFndn As IJLocalCoordinateSystem
        Set oFndn = pPropertyDescriptions.CAO
        Dim oPosition As DPosition
        Set oPosition = oFndn.Position
        Dim oPoint As IJPoint
        Dim xX As Double, yy As Double, zZ As Double
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
        
        If UBound(HoleLocations) >= 1 Then ' multiple point case commented for now.
'            On Error GoTo PORTS_TOO_CLOSE
'            Call CalculateRectangleOrientation(HoleLocations(), xVec, yVec)
'            On Error GoTo ErrorHandler
'            ' assign new values to block
'            newLength = xVec.Length
'            newWidth = yVec.Length
'            If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value <> "" Then
'                If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value > newWidth Then
'                    dOffX = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value - newWidth
'                    newWidth = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value
'                Else
'                    dOffX = 0#
'                End If
'            End If
'            If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value <> "" Then
'                If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value > newLength Then
'                    dOffY = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value - newLength
'                    newLength = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value
'                Else
'                    dOffY = 0#
'                End If
'            End If
'            GetCentroidOfPositions  HoleLocations(), UBound(HoleLocations), xX, yY, zZ
'            newTmx.IndexValue(12) = xX
'            newTmx.IndexValue(13) = yY
'            newTmx.IndexValue(14) = zZ
        End If
    Else
        ' calculate holes and get the transformation matrix from foundation ports
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
    
    'Check the existing values for the length and width ,then modify.
    dOffX = dOffY = 0#
    If bBlockSizeByRule = False Then
        If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value <> "" Then
            If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value <> newWidth Then
                dOffX = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value - newWidth
            Else
                dOffX = 0#
            End If
        End If
        If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value <> "" Then
            If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value <> newLength Then
                dOffY = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value - newLength
            Else
                dOffY = 0#
            End If
        End If
        newWidth = newWidth + dOffX
        newLength = newLength + dOffY
    End If
    
    'set the new length and width.
    If newWidth <= 0# Or newLength <= 0# Then
        On Error GoTo 0
        SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_FNDWIDTH_LENGTH_ZERO, oSmartOcc, Nothing
        Err.Raise E_FAIL
    End If
    'set the neew length and width
    pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value = newLength
    pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value = newWidth
    
    If bFoundSuppSurface = True Then
        zVec.Length = 1     ' normalize
        ' calculate the block ht from the support plane
        CalculateHeightFromOriginToSupporingPlane SupportPlane, eqPos, zVec, newHeight
        If newHeight <= 0# Then
             On Error GoTo 0
             SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_ASSOCEQP_MODIFIED_NOHEIGHT, oSmartOcc, Nothing
             Err.Raise E_FAIL
        Else
            newHeight = newHeight + 0.00001 ' tr 74038
        End If
        'set the new height
        pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(HT_ATTRNAME).Value = newHeight
    End If
    

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
    Set pIJOccAttribs = Nothing
    Set pOcc = Nothing
    Set SupportPlane = Nothing
    Set zVec = Nothing
    Set oFndnMatrix = Nothing
    Set oFndn = Nothing

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

Public Sub CMEvaluateBlockMaterial(pPropertyDescriptions As IJDPropertyDescription, _
                                    pObject As Object)
Const METHOD = "CMEvaluateBlockMaterial"
On Error GoTo ErrHandler

    Dim MemberObj As IJDMemberObjects
    Dim oAttrs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Set MemberObj = pPropertyDescriptions.CAO
    Set oSmartOcc = MemberObj.ItemByDispid(1)
    On Error Resume Next
    Set oAttrs = oSmartOcc
    
    Dim Material As String
    Dim oStructMatl As IJStructMaterial
    Dim Grade As String

    Material = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(MATERIAL_ATTRNAME).Value
    Grade = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(GRADE_ATTRNAME).Value
    
    If Not Material = "" And Not Grade = "" Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        Set oStructMatl = oSmartOcc
        oStructMatl.StructMaterial = iMaterial
    End If

    Exit Sub
    
ErrHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateBlockGeometry(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateBlockGeometry"
On Error GoTo ErrHandler
    
    Call CMEvaluateBlock(pPropertyDescriptions, pObject)
    Exit Sub

ErrHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructSlab(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructSlab"
On Error GoTo ErrHandler
 
     Dim oSlabComp As ISPSFoundationComponent
     Dim strSlabComp As String
     strSlabComp = GetCAODefAttribute(pMemberDescription, BLOCKANDSLABFNDASM_IFACE, SLAB_COMP_NAME)
     Call CreateComponent(strSlabComp, pResourceManager, oSlabComp, pMemberDescription, True)

     Set pObj = oSlabComp
     
    Exit Sub
    
ErrHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputSlab(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputSlab"
On Error GoTo ErrorHandler

    Dim strSlabComp As String
    strSlabComp = GetCAODefAttribute(pMemberDesc, BLOCKANDSLABFNDASM_IFACE, SLAB_COMP_NAME)
       
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.object
    oSmartOcc.ROOTSELECTION = strSlabComp
    
    ' tr 59726  Modifying a block fnd to a block fnd w/ support plane, support plane wrong
    ' problem cause: new plane input is not getting copied on to the child SO.
    ' hence copying it during the setinput of the child SO
           
    Dim pIRCAsm As IJDReferencesCollection
    Set pIRCAsm = GetRefCollection(pMemberDesc.CAO)

    If Not pIRCAsm Is Nothing Then
                
        Dim pIRCAsm2 As IJDReferencesCollection
        On Error Resume Next
        Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
        
        Dim oReferencesCollection As IJDReferencesCollection
        On Error Resume Next
        Set oReferencesCollection = GetRefCollection(oSmartOcc)
        
        On Error GoTo ErrorHandler
        If Not oReferencesCollection Is Nothing Then
            If Not pIRCAsm2 Is Nothing Then
                oReferencesCollection.IJDEditJDArgument.SetEntity 2, pIRCAsm2
            End If
        End If
                
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
 
Public Sub CMFinalConstructSlab(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructSlab"
On Error GoTo ErrorHandler

    Call AddSystemAndNameRule(pMemberDesc)
        
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseSlab(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseSlab"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalSlab(ByVal pMemberDescription As IJDMemberDescription, _
                             ByRef bIsNeeded As Boolean)
Const METHOD = "CMConditionalSlab"
On Error GoTo ErrorHandler

    Dim OccAttribs As IJDAttributes, oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDescription.CAO
    Set OccAttribs = oSmartOcc
    Set oAttribs = oSmartOcc.ItemObject
    
    If Not IsSOOverridden(OccAttribs.CollectionOfAttributes(BLOCKANDSLABFNDASM_IFACE)) Then
        CopyValuesToSOFromItem OccAttribs.CollectionOfAttributes(BLOCKANDSLABFNDASM_IFACE), oAttribs.CollectionOfAttributes(BLOCKANDSLABFNDASM_IFACE)
    End If

    Dim oAttr As IJDAttributes
    Set oAttr = oSmartOcc
    Dim Needed As Boolean
    If IsEmpty(oAttr.CollectionOfAttributes(BLOCKANDSLABFNDASM_IFACE).Item(WITH_SLAB_ATTRNAME).Value) Then
        bIsNeeded = True
    Else
        Needed = oAttr.CollectionOfAttributes(BLOCKANDSLABFNDASM_IFACE).Item(WITH_SLAB_ATTRNAME).Value
        bIsNeeded = Needed
    End If
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateSlab(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateSlab"
On Error GoTo ErrorHandler

    Dim MemberObj As IJDMemberObjects
    Dim pIRCAsm As IJDReferencesCollection
    Dim pIRCAsm1 As IJDReferencesCollection
    Dim pIRCAsm2 As IJDReferencesCollection
    Dim FoundationPorts As IJElements
    Dim SupportPlane As IJPlane
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJOccAttribs As IJDAttributes
    Dim pIJBlockOccAttribs As IJDAttributes
    Dim oPort As Object
    Dim pOcc As IJDOccurrence
    Dim rc1cnt As Integer, rc2cnt As Integer
    Dim rccnt As Integer, i As Integer
    Dim bFoundPorts As Boolean, bFoundSuppSurface As Boolean
    Dim Trans As IngrGeom3D.IJDT4x4
    Dim HoleLocations() As Automath.DPosition
    Dim NumberOfHoles As Integer
    Dim newLength As Double, newWidth As Double
    Dim xVec As Automath.DVector, yVec As Automath.DVector, zVec As Automath.DVector
    Dim newTmx As Automath.DT4x4
    Dim eqPos As Automath.DPosition
    Dim bPtOption As Boolean
    Dim newHeight As Double
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set MemberObj = pPropertyDescriptions.CAO
    Set pIRCAsm = GetRefCollection(pPropertyDescriptions.CAO)
    rccnt = pIRCAsm.IJDEditJDArgument.GetCount
    
    Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
    On Error Resume Next
    Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
    
    Dim pEnumJDArgument As IEnumJDArgument
    Set pEnumJDArgument = pIRCAsm1
    Set FoundationPorts = New JObjectCollection 'Elements
    Call GetPortAndHoleLocations(pEnumJDArgument, FoundationPorts, HoleLocations, bPtOption)
     
    Set pIJOccAttribs = pObject
    rc2cnt = 0
    If pIRCAsm2 Is Nothing Then Else rc2cnt = pIRCAsm2.IJDEditJDArgument.GetCount

    If rc2cnt >= 1 Then
        bFoundSuppSurface = True
        Set SupportPlane = pIRCAsm2.IJDEditJDArgument.GetEntityByIndex(1)
    End If



    Dim bSlabSizeByRule As Boolean
    Dim dBlockHeight As Double
    Dim dOffX As Double, dOffY As Double
    dOffX = dOffY = 0#
    bSlabSizeByRule = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABSIZEBYRULE_ATTRNAME).Value
    Set newTmx = New Automath.DT4x4
    Set eqPos = New Automath.DPosition
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set zVec = New Automath.DVector
    ' get the clearance attribute value
    Dim dClear As Double
    dClear = 0#     ' tr 75005
    If bSlabSizeByRule = False And pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABEDGECLEAR_ATTRNAME).Value <> "" Then
        dClear = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABEDGECLEAR_ATTRNAME).Value
    End If
    Set pIJBlockOccAttribs = MemberObj.ItemByDispid(1)
    dBlockHeight = pIJBlockOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(HT_ATTRNAME).Value
    Dim dx, dy, dz As Double
    
    newLength = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value
    newWidth = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value
    newHeight = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABHT_ATTRNAME).Value
    
    If UBound(HoleLocations) >= 1 Or bPtOption Then
        Dim oFndnMatrix As IJDMatrixAccess
        Set oFndnMatrix = pPropertyDescriptions.CAO
        Set newTmx = oFndnMatrix.Matrix
        Dim oFndn As IJLocalCoordinateSystem
        Set oFndn = pPropertyDescriptions.CAO
        Dim oPosition As DPosition
        Set oPosition = oFndn.Position
        Dim oPoint As IJPoint
        Dim xX As Double, yy As Double, zZ As Double
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
        
        If UBound(HoleLocations) >= 1 Then ' multiple point case commenting for now
'            On Error GoTo PORTS_TOO_CLOSE
'            Call CalculateRectangleOrientation(HoleLocations(), xVec, yVec)
'            On Error GoTo ErrorHandler
'            ' assign new values to block
'            newLength = xVec.Length + dClear
'            newWidth = yVec.Length + dClear
'            If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value <> "" Then
'                If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value > newWidth * 1.5 Then
'                    dOffX = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value - newWidth * 1.5
'                    newWidth = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value 'newWidth + dOffX * 2
'                Else
'                    dOffX = 0#
'                End If
'            End If
'            If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value <> "" Then
'                If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value > newLength * 1.5 Then
'                    dOffY = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value - newLength * 1.5
'                    newLength = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value 'newLength + dOffY * 2
'                Else
'                    dOffY = 0#
'                End If
'            End If
'            dOffX = (newWidth - yVec.Length) / 2
'            dOffY = (newLength - xVec.Length) / 2
'            GetCentroidOfPositions  HoleLocations(), UBound(HoleLocations), xX, yY, ZZ
'            newTmx.IndexValue(12) = xX
'            newTmx.IndexValue(13) = yY
'            newTmx.IndexValue(14) = ZZ
        End If
    Else
        ' calculate holes and get the transformation matrix from foundation ports
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
    
    'Check the existing values for the length and width ,then modify.
    dOffX = dOffY = 0#
    If bSlabSizeByRule = False Then
        If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value <> "" Then
            If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value <> newWidth Then
                dOffX = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value - newWidth
            Else
                dOffX = 0#
            End If
        End If
        If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value <> "" Then
            If pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value <> newLength Then
                dOffY = pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value - newLength
            Else
                dOffY = 0#
            End If
        End If
        newWidth = newWidth + dOffX
        newLength = newLength + dOffY
    End If
    
    If newWidth <= 0# Or newLength <= 0# Then
        On Error GoTo 0
        SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_FNDWIDTH_LENGTH_ZERO, oSmartOcc, Nothing
        Err.Raise E_FAIL
    End If
    
    ' assign new values to block
    pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABLENGTH_ATTRNAME).Value = newLength
    pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABWIDTH_ATTRNAME).Value = newWidth
    
    If bFoundSuppSurface = True Then
        zVec.Length = 1     ' normalize
        ' calculate the block ht from the support plane
        CalculateHeightFromOriginToSupporingPlane SupportPlane, eqPos, zVec, newHeight
        If newHeight <= 0# Then
             On Error GoTo 0
             SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_ASSOCEQP_MODIFIED_NOHEIGHT, oSmartOcc, Nothing
             Err.Raise E_FAIL
        Else
            newHeight = newHeight + 0.00001 ' tr 74038
        End If
        ' tr 53900 - set block and slab ht same
        pIJOccAttribs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABHT_ATTRNAME).Value = newHeight / 2
        pIJBlockOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(HT_ATTRNAME).Value = newHeight / 2

    End If
    
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
    Set pIJOccAttribs = Nothing
    Set pOcc = Nothing
    Set SupportPlane = Nothing
    Set pIJBlockOccAttribs = Nothing
    Set oFndnMatrix = Nothing
    Set oFndn = Nothing
    
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

Public Sub CMEvaluateSlabMaterial(pPropertyDescriptions As IJDPropertyDescription, _
                                  pObject As Object)
Const METHOD = "CMEvaluateSlabMaterial"
On Error GoTo ErrorHandler
    
    Dim MemberObj As IJDMemberObjects
    Dim oAttrs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Set MemberObj = pPropertyDescriptions.CAO
    Set oSmartOcc = MemberObj.ItemByDispid(2)
    On Error Resume Next
    Set oAttrs = oSmartOcc
    Dim Material As String
    Dim oStructMatl As IJStructMaterial
    Dim Grade As String

    Material = oAttrs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABMAT_ATTRNAME).Value
    Grade = oAttrs.CollectionOfAttributes(SLABFND_IFACE).Item(SLABGR_ATTRNAME).Value

    If Not Material = "" And Not Grade = "" Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        Set oStructMatl = oSmartOcc
        oStructMatl.StructMaterial = iMaterial
    End If

    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateSlabGeometry(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateSlabGeometry"
On Error GoTo ErrHandler
    
    Call CMEvaluateSlab(pPropertyDescriptions, pObject)

    Exit Sub
ErrHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub CreateComponent(Component As String, _
                            ByVal pResourceManager As IUnknown, _
                            ByRef oFndnComp As ISPSFoundationComponent, _
                            ByVal pMembDescr As IJDMemberDescription, _
                            bSetInput As Boolean)
Const METHOD = "CreateComponent"
On Error GoTo ErrorHandler
         
     Dim oEqpFndnFactory As SPSEquipFoundationFactory
     Set oEqpFndnFactory = New SPSEquipFoundationFactory
     Set oFndnComp = oEqpFndnFactory.CreateFoundationComponent(pResourceManager)
     
     Dim oReferencesCollection As IMSSymbolEntities.IJDReferencesCollection
     Dim oSymbolEntitiesFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
     Set oReferencesCollection = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)
     
     If bSetInput Then
        Dim oInputObj As Object
        Dim pIRCAsm As IJDReferencesCollection
        Set pIRCAsm = GetRefCollection(pMembDescr.CAO)
        
        Dim pIRCAsm1 As IJDReferencesCollection
        Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
        
        Dim pIRCAsm2 As IJDReferencesCollection
        On Error Resume Next
        Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
        ' tr 76833
'        If pIRCAsm2 Is Nothing Then
'            Set pIRCAsm2 = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)
'        End If

        oReferencesCollection.IJDEditJDArgument.SetEntity 1, pIRCAsm1
        If Not pIRCAsm2 Is Nothing Then             ' tr 76833
            oReferencesCollection.IJDEditJDArgument.SetEntity 2, pIRCAsm2
        End If
     End If
     
     Dim oSmartOcc As IJSmartOccurrence
     Set oSmartOcc = oFndnComp
     Dim oItem As Object
     Dim strUserType As String
     Dim strSCName As String
     Dim oSmartItem As IJSmartItem
     Dim oSmartClass As IJSmartClass
     Dim oUserType As IJDUserType
     
     Set oItem = GetDefinition(Component)
     Set oSmartItem = oItem
     Set oSmartClass = oSmartItem.Parent
     strUserType = oSmartClass.SOUserType
     Set oUserType = oSmartOcc
     oUserType.UserType = strUserType
     strSCName = oSmartClass.SCName
     oSmartOcc.RootSelectorClass = strSCName
     oSmartOcc.ROOTSELECTION = oSmartItem.name
     ConnectSmartOccurrence oSmartOcc, oReferencesCollection
     Exit Sub
     
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub GenerateNameForFoundation(Obj As Object)
Const METHOD = "GenerateNameForFoundation"
On Error GoTo ErrorHandler

    Dim NameRule As String
    Dim found As Boolean
    found = False
    On Error Resume Next
      
    Dim NamingRules As IJElements
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oActiveNRHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Call oNameRuleHlpr.GetEntityNamingRulesGivenProgID(FNDNCOMPPROGID, NamingRules)
    Dim ncount As Integer
    Dim oNameRuleAE As GSCADGenNameRuleAE.IJNameRuleAE
      
    For ncount = 1 To NamingRules.count
        Set oNameRuleHolder = NamingRules.Item(1)
    Next ncount

    Call oNameRuleHlpr.AddNamingRelations(Obj, oNameRuleHolder, oNameRuleAE)
    Set oNameRuleHolder = Nothing
    
    Set oActiveNRHolder = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
 
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub AddSystemAndNameRule(pMemberDescription As IJDMemberDescription)
Const METHOD = "AddSystemAndNameRule"
On Error GoTo ErrorHandler

    Dim oDesignParent As IJDesignParent
    Set oDesignParent = pMemberDescription.CAO
    oDesignParent.AddChild pMemberDescription.object
         
    Dim oAssemblyparent As IJAssembly
    Set oAssemblyparent = oDesignParent
    oAssemblyparent.AddChild pMemberDescription.object
    
    Call GenerateNameForFoundation(pMemberDescription.object)
        
    Set oAssemblyparent = Nothing
    Set oDesignParent = Nothing
     
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructInterfaces.IJElements)
    Dim i As Integer
    Dim cnt As Integer
    Dim oItem As IJSmartItem
    Dim oSmartOcc As IJSmartOccurrence
    Dim FndDefServices As ISPSEquipFndnDefServices
    
    Set oSmartOcc = pPartObject
    Set oItem = oSmartOcc.ItemObject
    Set FndDefServices = SP3DCreateObject(oItem.definition)

    Dim Supported As IJElements 'IMSElements.IJElements
    Set Supported = New JObjectCollection 'IMSElements.Elements
    Dim Supporting As IJElements ' IMSElements.Elements
    Set Supporting = New JObjectCollection 'IMSElements.Elements

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
                If Not oStructPort Is Nothing Then
                    Set pUnk = oStructPort.Connectable
                End If
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

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String

End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

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

Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String

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
'Migrates thr foundation to the correct surface if it is split.
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
