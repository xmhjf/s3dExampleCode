VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlockFndDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************
' Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File
'    BlockFndDef.cls
'
'Author
'       1-Mar-03        SS
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
' 08-Aug-2006 AS    TR#99968 Added CMMigrate to handle EqpFnd Migration
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
'  01.May.2014  RUK TR-CP-253694  ATP - 81 minidumps at 'SPSEqpFndMacros' module. => Stack for BlockFndDef.cls @ 431
'''*******************************************************************

Private Const MODULE = "BlockFndDef"
Private Const strSourceFile = "BlockFndDef.def"
Private Const CONST_ItemProgId As String = "SPSEqpFndMacros.BlockFndDef"
Private Const MODELDATABASE = "Model"
Private Const NUMKNOWNINPUTS = 7
Private Const BLOCKFND_IFACE = "IJUASPSBlockFndn"
Private Const LENGTH_ATTRNAME = "BlockLength"
Private Const WIDTH_ATTRNAME = "BlockWidth"
Private Const HT_ATTRNAME = "BlockHeight"
Private Const MATERIAL_ATTRNAME = "BlockSPSMaterial"
Private Const GRADE_ATTRNAME = "BlockSPSGrade"
Private Const BLOCKSIZEBYRULE_ATTRNAME = "IsBlockSizeDrivenByRule"
Private Const BLOCKEDGECLEAR_ATTRNAME = "BlockEdgeClearance"
Private Const INTERFACE_WCGValueOrigin As String = "IJWCGValueOrigin"
Private Const PROPERTY_DryWCGOrigin As String = "DryWCGOrigin"

Private Enum enumWeightCGDerivation
    WEIGHTCG_Computed = 2
    WEIGHTCG_UserDefined = 4
End Enum

Private m_bOnPreLoad As Boolean

Private Const DOUBLE_VALUE = 8
Private Const BOOL = -7
Private Const CHAR = 1
Private m_oLocalizer As IJLocalizer

Implements IJDUserSymbolServices
Implements ISPSEquipFndnDefServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSFoundationInputHelper 'TR#71850
Implements ICustomOutputHelper

Private Function ICustomOutputHelper_GetComponents(ByVal assemblyObject As Object) As SP3DStructInterfaces.IJElements
Const METHOD = "ICustomOutputHelper_GetComponents"
On Error GoTo ErrHandler
    
    Dim colComponents As IJElements
    Dim colBlockAttribs As IJElements
    
    Dim colBlockFaces As IJElements

    Dim oEqpFndnComp As IAssemblyComponent
    Dim oOutput As Object
    Dim oAttrs As IJDAttributes
    Dim strGrade As String, strMaterial As String
    Dim iSymbolOccMisc As iSymbolOccMisc
    Dim oOutputProxy As IJDProxy
    Dim oRelation As IJDAssocRelation
    Dim oOutputs As IJDTargetObjectCol
    Dim oRelationship As IJDRelationship
    Dim lngCount As Long
    Dim lngIndex As Long
    Dim colAttribs As CollectionProxy
    Dim oAttrib As AttributeProxy
    Dim oSymbol As IJDSymbol
    
    
    Set oSymbol = assemblyObject
    
    Set iSymbolOccMisc = assemblyObject
    Set oAttrs = assemblyObject
    
    Set colComponents = New JObjectCollection
    Set colBlockFaces = New JObjectCollection
    
    Set oOutputProxy = iSymbolOccMisc.BindToOCIfExists(SYMBOL_REP_NAME)
    Set oRelation = oOutputProxy.Source

    Set oOutputs = oRelation.CollectionRelations(IJOUTPUTCOLLECTION, TO_OUTPUTS)
    lngCount = oOutputs.count
    For lngIndex = 1 To lngCount
        'the output below is not transformed by the symbol's transformation matrix
        Set oOutput = oOutputs.Item(lngIndex)
        Set oRelationship = oOutputs.GetRelationshipToTarget(oOutput)
        Dim oOutPutRefProxy As Object
        ' Get the output that is transformed by the symbol's transformation matrix
        Set oOutPutRefProxy = oSymbol.BindToOutput(SYMBOL_REP_NAME, oRelationship.name)
        
        colBlockFaces.Add oOutPutRefProxy
        Set oOutPutRefProxy = Nothing
        Set oOutput = Nothing
    Next

    'create Block
    
    Set oEqpFndnComp = New SPSAssemblyComponent
    'put name
    oEqpFndnComp.name = "Block"
    
    'put material and grade
    strMaterial = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(MATERIAL_ATTRNAME).Value
    strGrade = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(GRADE_ATTRNAME).Value
    
    oEqpFndnComp.SetMaterialTypeAndGrade strMaterial, strGrade
    
    'put faces
    oEqpFndnComp.Geometry = colBlockFaces
    
    'put attributes
    
    Set colBlockAttribs = New JObjectCollection
    
    For Each oAttrib In oAttrs.CollectionOfAttributes(BLOCKFND_IFACE)
        colBlockAttribs.Add oAttrib
    Next
    
    oEqpFndnComp.Attributes = colBlockAttribs
    colComponents.Add oEqpFndnComp
    
    Set ICustomOutputHelper_GetComponents = colComponents
    Exit Function
ErrHandler:
    Err.Raise ReportError(Err, strSourceFile, METHOD).Number

End Function

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

    'Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    IJDUserSymbolServices_EditOccurence = False

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = CONST_ItemProgId
    
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler

    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
    ' Aggregator Type
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    pAD.AggregatorClsid = "{94F2C800-ECF1-47C0-88D5-01460B80ECD5}" 'CSPSEquipFoundation
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructCA"
    pAD.SetCMSetInputs -1, -1
    pAD.SetCMRemoveInputs -1, -1
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
    Set pAD = Nothing
    
    ' Aggregator property
    Dim pAPDs As IJDPropertyDescriptions
    Set pAPDs = pDefinition
    pAPDs.RemoveAll         ' Remove all the previous property descriptions

    'BLOCKATTR_IFACE
    pAPDs.AddProperty BLOCKFND_IFACE, 1, BLOCKFND_IFACE, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
    pAPDs.AddProperty "IJStructMaterial", 2, IJStructMaterial, "CMEvaluateCAOMaterial", imsCOOKIE_ID_USS_LIB
    pAPDs.AddProperty "IJWeightCG", 3, IJWeightCG, "CMEvaluateCAOWCG", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
                                        
    Set pAPDs = Nothing
    
    Exit Sub
    
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAOMaterial(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAOMaterial"
On Error GoTo ErrHandler
    
    Dim oAttrs As IJDAttributes
    Dim oDefAttr As IJDAttributes
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Dim Material As String, defMaterial As String
    Dim Grade As String, defGrade As String

    Set oSmartOcc = pObject
    Set oAttrs = oSmartOcc
    
    Set oDefAttr = oSmartOcc.ItemObject
    
    Material = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(MATERIAL_ATTRNAME).Value
    Grade = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(GRADE_ATTRNAME).Value
    
    If Material = "" Then '        copy from the default
        defMaterial = oDefAttr.CollectionOfAttributes(BLOCKFND_IFACE).Item(MATERIAL_ATTRNAME).Value
        oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(MATERIAL_ATTRNAME).Value = defMaterial
    End If
    
    If Grade = "" Then  '        copy from the default
        defGrade = oDefAttr.CollectionOfAttributes(BLOCKFND_IFACE).Item(GRADE_ATTRNAME).Value
        oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(GRADE_ATTRNAME).Value = defGrade
    End If
    
    Material = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(MATERIAL_ATTRNAME).Value
    Grade = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(GRADE_ATTRNAME).Value
    
    If Not Material = "" And Not Grade = "" Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        If Not iMaterial Is Nothing Then Call AddMaterialRelationShip(oSmartOcc, iMaterial, "BlockFndn")
    End If

Exit Sub
ErrHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAOWCG(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAOWCG"
On Error GoTo ErrHandler

    Dim oSmartOcc As IJSmartOccurrence
    Dim oAttrs As IJDAttributes
    Dim Grade As String, Material As String
    Dim density As Variant
    Dim Volume As Double
    Dim BlockWt As Double
    Dim dSurfArea As Double
    Dim iMaterial As IJDMaterial
    Dim Length As Double, Width As Double, BlockHeight As Double
    Dim iWCG As ISPSComputedWeightCG
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
        Set oSmartOcc = pObject
    
        Material = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(MATERIAL_ATTRNAME).Value
        Grade = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(GRADE_ATTRNAME).Value
        
        Set iMaterial = GetMaterialObject(Material, Grade)
        
        If Not iMaterial Is Nothing Then
            density = iMaterial.density
        Else
            density = 2400 'approx density of concrete
        End If
        
        Length = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value
        Width = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value
        BlockHeight = oAttrs.CollectionOfAttributes(BLOCKFND_IFACE).Item(HT_ATTRNAME).Value
          
        Dim cgX As Double, cgY As Double, cgZ As Double
        Dim pOcc As IJDOccurrence
        Dim oMatrix As IJDT4x4
        Dim xOrigin As Double
        Dim yOrigin As Double
        Dim zOrigin As Double
        Dim pos As New DPosition
        Dim posOffset As IJDPosition
        
        Set pOcc = pObject
        Set oMatrix = pOcc.Matrix
        xOrigin = oMatrix.IndexValue(12)
        yOrigin = oMatrix.IndexValue(13)
        zOrigin = oMatrix.IndexValue(14)
        
        oMatrix.IndexValue(12) = 0#
        oMatrix.IndexValue(13) = 0#
        oMatrix.IndexValue(14) = 0#
        pos.Set Width / 2#, Length / 2#, BlockHeight / 2#
        Set posOffset = oMatrix.TransformPosition(pos)
        
        'volume computed from attr as we can't get it right if they are trimmed
        Volume = Length * Width * BlockHeight
      
        BlockWt = (Volume * density)
    
        oAttrs.CollectionOfAttributes("IJGenericVolume").Item("Volume").Value = Volume
        oAttrs.CollectionOfAttributes("IJSurfaceArea").Item("SurfaceArea").Value = dSurfArea
    
        Set iWCG = oSmartOcc
        ' The following put property values was originally SetWCG on the IJWeightCG interface
        '   (which is reserved for setting user defined properties), hence changed to put values
        '   on a new interface
        iWCG.Weight = BlockWt
        iWCG.cgX = xOrigin + posOffset.x
        iWCG.cgY = yOrigin + posOffset.y
        iWCG.cgZ = zOrigin - posOffset.z
    End If
    
    Exit Sub

ErrHandler:  HandleError MODULE, METHOD
End Sub


Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal pResourceMgr As Object) As Object
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

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Public Sub CMFinalConstructCA(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructCA"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO"
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
    Set oSmartOcc = pObject
    Set pIRCAsm = GetRefCollection(oSmartOcc)   ' pPropertyDescriptions.CAO

    'TR-CP-253694 -> Added the if condition to check the object is not nothing
    If Not pIRCAsm Is Nothing Then
        If Not pIRCAsm.IJDEditJDArgument Is Nothing Then
            rccnt = pIRCAsm.IJDEditJDArgument.GetCount
            If rccnt > 0 Then
                Set pIRCAsm1 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
            End If
            If rccnt > 1 Then
                Set pIRCAsm2 = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
            End If
        End If
    End If
    
    Set FoundationPorts = New JObjectCollection 'Elements
    Dim pEnumJDArgument As IEnumJDArgument
    
    If Not pIRCAsm1 Is Nothing Then
    Set pEnumJDArgument = pIRCAsm1
    
    Dim bPtOption As Boolean
    Call GetPortAndHoleLocations(pEnumJDArgument, FoundationPorts, HoleLocations, bPtOption)
    
    End If
    
    'set the occ attributes
    Set pIJOccAttribs = pObject
    
    ' get the SizingByRule attribute
    Dim bBlockSizeByRule As Boolean

    bBlockSizeByRule = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(BLOCKSIZEBYRULE_ATTRNAME).Value
    
    'Get the existing values for length, width and height
    newLength = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value
    newWidth = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value
    newHeight = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(HT_ATTRNAME).Value
    
    Set newTmx = New Automath.DT4x4
    Set eqPos = New Automath.DPosition
    Set xVec = New Automath.DVector
    Set yVec = New Automath.DVector
    Set zVec = New Automath.DVector
    
    Dim xX As Double, yy As Double, zZ As Double

    If UBound(HoleLocations) >= 1 Or bPtOption Then ' both single and multiple point cases
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
        
        If UBound(HoleLocations) >= 1 Then 'Multiple points case
'            On Error GoTo PORTS_TOO_CLOSE
'            Call CalculateRectangleOrientation(HoleLocations(),  xVec, yVec)
'            On Error GoTo ErrorHandler
'            ' assign new values to block
'            newLength = xVec.Length
'            newWidth = yVec.Length
'            If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value <> "" Then
'                If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value > newWidth Then
'                    newWidth = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value
'                End If
'            End If
'            If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value <> "" Then
'                If pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value > newLength Then
'                    newLength = pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value
'                End If
'            End If
'            'Need to set the centre at the centroid of the given points
'            GetCentroidOfPositions  HoleLocations(), UBound(HoleLocations), xX, yy, zZ
'            newTmx.IndexValue(12) = xX
'            newTmx.IndexValue(13) = yy
'            newTmx.IndexValue(14) = zZ
        End If
    
        Set oFndnMatrix = Nothing
        Set oFndn = Nothing
    Else ' equipment foundation case
        
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
       
    'Check whether the existing values for the length and width are different,then accordingly modify.
    Dim dOffX As Double, dOffY As Double
    dOffX = 0#
    dOffY = 0#
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
        
    'set the new length and width of the component
    If newWidth <= 0# Or newLength <= 0# Then
        On Error GoTo 0
        SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_FNDWIDTH_LENGTH_ZERO, oSmartOcc, Nothing
        Err.Raise E_FAIL
    End If
    pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(LENGTH_ATTRNAME).Value = newLength
    pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(WIDTH_ATTRNAME).Value = newWidth
    
    rc2cnt = 0
    'TR-CP-253694 -> Added/corrected the if condition to check the object is not nothing
    If Not pIRCAsm2 Is Nothing Then
        If Not pIRCAsm2.IJDEditJDArgument Is Nothing Then
            rc2cnt = pIRCAsm2.IJDEditJDArgument.GetCount
    
            'Calculate and set the height of the component(s) based on the support plane
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
                     SPSToDoErrorNotify EqpToDoMsgCodelist, TDL_EQPFNDMACROS_ASSOCEQP_MODIFIED_NOHEIGHT, oSmartOcc, Nothing
                     Err.Raise E_FAIL
                Else
                    newHeight = newHeight + 0.00001 ' tr 74038
                End If
            End If
            pIJOccAttribs.CollectionOfAttributes(BLOCKFND_IFACE).Item(HT_ATTRNAME).Value = newHeight
        End If
    End If
    
    'Set the occurence matrix
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
    If (pAttrToChange.attrName = "IsBlockSizeDrivenByRule") Then
        If (varNewAttrValue = True) Then 'User defined option for sizing rule
            'gray out the block length and width on the GOPC
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If pAttrDescr.attrName = "BlockEdgeClearance" Or pAttrDescr.attrName = "BlockLength" Or pAttrDescr.attrName = "BlockWidth" Then 'TR#75176
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
                If pAttrDescr.attrName = "BlockEdgeClearance" Or pAttrDescr.attrName = "BlockLength" Or pAttrDescr.attrName = "BlockWidth" Then  'TR#75176
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
    Dim pSmartItem As IJSmartItem
    Dim pSO As IJSmartOccurrence
    Dim pAttrs As IJDAttributes
    Dim attrName As String
    Dim ErrStr As String
    
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
            If (pAttrDescr.attrName = "BlockHeight") Then
                If Supporting.count >= 1 Then
                    pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                Else
                    pAttrDescr.AttrState = Not (AttributeDescriptor_ReadOnly)
                End If
                Exit For
            End If
        Next
    End If
    
    'Dim oProps As CollectionProxy
    If Supported.count > 0 Then
        If TypeOf Supported.Item(1) Is IJPoint Then 'Multiple point case
            For i = 1 To pAttrColl.count
                Set pAttrDescr = pAttrColl.Item(i)
                If pAttrDescr.attrName = "IsBlockSizeDrivenByRule" Then
                    pAttrDescr.AttrValue = False
                    pAttrDescr.AttrState = Not (AttributeDescriptor_ReadOnly)
                End If
            Next
'        Else 'Equipment case
'            For i = 1 To pAttrColl.count
'                Set pAttrDescr = pAttrColl.Item(i)
'                If pAttrDescr.attrName = "IsBlockSizeDrivenByRule" Then
'                    pAttrDescr.AttrValue = True
'                    pAttrDescr.AttrState = Not (AttributeDescriptor_ReadOnly)
'                End If
'            Next
        End If
    Else 'single point case
        For i = 1 To pAttrColl.count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.attrName = "IsBlockSizeDrivenByRule" Then
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

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
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
    Dim Supporting As IJElements 'IMSElements.Elements
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
                ' just to verify that we can do whatever we like to for specific attributes
'                    If (sAttributeName = "BlockLength") Then
'                        If (varAttributeValue < 1.29) Then
'                            UserAttributeMgmt_Validate = sAttributeName
'                            Set oAttributeMetaData = Nothing
'                            Exit Function
'                        End If
'                    Else
                        If (varAttributeValue <= 0#) Then
'                            UserAttributeMgmt_Validate = sAttributeName
                            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_EQPFNDMACROS_INVALID_ATTRIB_VALUE, "Invalid Negative Attribute Value ")
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
'                    End If
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
'Migrates thr Foundation to the correct surface if it is split.
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
