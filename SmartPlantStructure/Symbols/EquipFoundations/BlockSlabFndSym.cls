VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlockSlabFndSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    BlockSlabFndSym.cls
'
'Author
'       1-Mar-03        Sudha Srikakolapu
'
'Description
'
'Notes
'
'History:
'
'   Jun 24, 2009        WR      TR-145118 - Updated the Physical method to call
'                               the new CreateSolidAsPlanes method
'                               in EqpFndCommon.
'*******************************************************************

Private Const MODULE = "BlockSlabFndSym"
Const m_ItemProgId As String = "SPSEqpFndMacros.BlockSlabFndSym"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"

Implements IJDUserSymbolServices

Private Enum InputIndex
    EQPPORT_INDEX = 1
    SUPPPLANE_INDEX
    BLOCKLENGTH_INDEX
    BLOCKWIDTH_INDEX
    BLOCKHT_INDEX
    BLOCKSIZEBYRULE_INDEX
    BLOCKEDGECLEARANCE_INDEX
    BLOCKMATERIAL_INDEX
    BLOCKGRADE_INDEX
    SLABLENGTH_INDEX
    SLABWIDTH_INDEX
    SLABHT_INDEX
    SLABSIZEBYRULE_INDEX
    SLABEDGECLEARANCE_INDEX
    SLABMATERIAL_INDEX
    SLABGRADE_INDEX
End Enum

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler

    'Create a Symbol Definition Object.
    Dim pSymbolFactory As New DSymbolEntitiesFactory
    Dim pSymbolDefinition As IJDSymbolDefinition
    
    Set pSymbolDefinition = pSymbolFactory.CreateEntity(definition, pResourceMgr)
    pSymbolDefinition.ProgId = m_ItemProgId
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.name = pSymbolDefinition.ProgId

    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
       
    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition
    Set pSymbolFactory = Nothing
    Set pSymbolDefinition = Nothing
    
    Exit Function
 
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
     
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    On Error GoTo ErrorHandler
    
    ' Set the input to the definition
    Dim iInputs As IMSSymbolEntities.IJDInputs
    Set iInputs = pSymbolDefinition
    
    Dim iUM As IMSSymbolEntities.IJDUserMethods
    Set iUM = pSymbolDefinition
    
    Dim libDesc As New DLibraryDescription
    Dim mCookie As Long
    Dim libCookie As Long
    Dim methodCookie As Long
         
    libDesc.name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = m_ItemProgId
    
    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc
    
    Dim ChecklibDesc As New DLibraryDescription
    Dim ChecklibCookie As Long
    Dim GTZeroCheck As Long

    ChecklibDesc.name = "CMCheckLib"
    ChecklibDesc.Type = imsLIBRARY_IS_ACTIVEX
    ChecklibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    ChecklibDesc.Source = CheckProgId
    pSymbolDefinition.IJDUserMethods.SetLibrary ChecklibDesc
    ChecklibCookie = ChecklibDesc.Cookie

    GTZeroCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("GTZero", ChecklibCookie)
    
'    Get the lib/method cookie
    libCookie = libDesc.Cookie
    
    ' set to variable number of inputs for suppored equipments and supporting plane
    Dim inputsProp As IMSDescriptionProperties
    inputsProp = pSymbolDefinition.IJDInputs.Property
    pSymbolDefinition.IJDInputs.Property = inputsProp Or igCOLLECTION_VARIABLE
        
    Dim pIJDInput As IMSSymbolEntities.IJDInput
    Set pIJDInput = New IMSSymbolEntities.DInput
    
    Dim oInput As IMSSymbolEntities.IJDInput
    Set oInput = New IMSSymbolEntities.DInput
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    PC.Type = igValue
    
    oInput.name = "EquipmentPorts"
    oInput.Description = "Supported Equipment Port(s)"
    iInputs.SetInput oInput, EQPPORT_INDEX
    oInput.Reset

    oInput.name = "SupportPlane"
    oInput.Description = "Supporting Surface/Plane"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    iInputs.SetInput oInput, SUPPPLANE_INDEX
    oInput.Reset
    
    oInput.name = "BlockLength"
    oInput.Description = "Length of the equipment containing bolt-holes position"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 1.05
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, BLOCKLENGTH_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "BlockWidth"
    oInput.Description = "Mounting width which comes from the equipment"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.55
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, BLOCKWIDTH_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "BlockHeight"
    oInput.Description = "Height of the Block foundation"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.5
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, BLOCKHT_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "IsBlockSizeDrivenByRule"
    oInput.Description = "Is Block Size Driven By Rule (bolt hole locations)"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, BLOCKSIZEBYRULE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "BlockEdgeClearance"
    oInput.Description = "BlockEdgeClearance"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.0001
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, BLOCKEDGECLEARANCE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "BlockSPSMaterial"
    oInput.Description = "Block SPSMaterial"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Concrete"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, BLOCKMATERIAL_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "BlockSPSGrade"
    oInput.Description = "Block SPSGrade"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Fc 4000"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, BLOCKGRADE_INDEX
    oInput.Reset
    PC.Reset
    
     oInput.name = "SlabLength"
    oInput.Description = "Slab Length"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 1.05
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SLABLENGTH_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SlabWidth"
    oInput.Description = "Slab Width"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.55
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SLABWIDTH_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SlabHeight"
    oInput.Description = "Height of the Slab Part"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.5
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SLABHT_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "IsSlabSizeDrivenByRule"
    oInput.Description = "Slab Size Driven By Rule (bolt hole locations)"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SLABSIZEBYRULE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SlabEdgeClearance"
    oInput.Description = "SlabEdgeClearance"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.0001
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SLABEDGECLEARANCE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SlabSPSMaterial"
    oInput.Description = "Slab SPSMaterial"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Concrete"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SLABMATERIAL_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SlabSPSGrade"
    oInput.Description = "Slab SPSGrade"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Fc 4000"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SLABGRADE_INDEX
    oInput.Reset
    PC.Reset
    
    Dim pIReps As IMSSymbolEntities.IJDRepresentations
    Set pIReps = pSymbolDefinition
    Dim pIRep As IMSSymbolEntities.IJDRepresentation
    Set pIRep = New IMSSymbolEntities.DRepresentation
    
    pIRep.name = "Physical"
    pIRep.Description = "Physical representation"
    pIRep.RepresentationId = SimplePhysical
    mCookie = iUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    
    Dim pOutputs As IMSSymbolEntities.IJDOutputs
    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    
    Dim output As IMSSymbolEntities.IJDOutput
    Set output = New IMSSymbolEntities.DOutput
    
    output.name = "Slab"
    output.Description = "Slab"
    pOutputs.SetOutput output
    output.Reset
    
    pIReps.SetRepresentation pIRep 'Add representation to definition
    
    pIRep.name = "DetailPhysical"
    pIRep.Description = "DetailPhysical representation"
    pIRep.RepresentationId = DetailPhysical
    mCookie = iUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    Set pOutputs = pIRep
    pIReps.SetRepresentation pIRep 'Add representation to definition

    ' Set definition cache properties
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
     IJDUserSymbolServices_GetDefinitionName = m_ItemProgId
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)
Const METHOD = "Physical"
On Error GoTo ErrorHandler
    
    Dim BlockLength As Double, BlockWidth As Double, BlockHeight As Double, BlockClearance As Double
    Dim SlabLength As Double, SlabWidth As Double, SlabHeight As Double, SlabClearance As Double
    Dim Blockz As Double, deltalength As Double, deltawidth As Double
    
    Dim pRepDG As IJDRepresentationDuringGame
    Dim pOC As IJDOutputCollection
    
    Set pRepDG = pIRepSCM
    Set pOC = pRepDG.outputCollection
    
    Dim oVec As IJDVector
    Dim elems As IJElements
    Dim pObj As Object
    Dim OutStr As String
    Dim ii As Integer
    
    'remove all outputs
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep
    oOutputs.RemoveAllOutput
    
    Dim pInputs As IJDInputs
    Set pInputs = pRepDG.definition.IJDInputs
   
    Dim tmpLineString As IngrGeom3D.LineString3d

    'assign to meaningful variables from the input array
    BlockClearance = pInputs.GetInputByIndex(BLOCKEDGECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    BlockLength = pInputs.GetInputByIndex(BLOCKLENGTH_INDEX).IJDInputDuringGame.Result.UomValue + BlockClearance
    BlockWidth = pInputs.GetInputByIndex(BLOCKWIDTH_INDEX).IJDInputDuringGame.Result.UomValue + BlockClearance

    BlockHeight = pInputs.GetInputByIndex(BLOCKHT_INDEX).IJDInputDuringGame.Result.UomValue

    SlabClearance = pInputs.GetInputByIndex(SLABEDGECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    SlabLength = pInputs.GetInputByIndex(SLABLENGTH_INDEX).IJDInputDuringGame.Result.UomValue + SlabClearance
    SlabWidth = pInputs.GetInputByIndex(SLABWIDTH_INDEX).IJDInputDuringGame.Result.UomValue + SlabClearance
    SlabHeight = pInputs.GetInputByIndex(SLABHT_INDEX).IJDInputDuringGame.Result.UomValue

    Set pInputs = Nothing
    Set pRepDG = Nothing

'   fix for tr 53056 - BlockSlabEqpFndn Block and Slab are mixed up on properties page
    Dim slabZ As Double
    
    slabZ = -BlockHeight
'    deltawidth = (SlabWidth - BlockWidth) / 2
'    deltalength = (SlabLength - BlockLength) / 2
        
    '=================
    'BUILD block
    '=================
    Dim Block1(0 To 14) As Double
    
    Dim initialX As Double, initialY As Double, initialZ As Double
    initialX = 0 - BlockWidth / 2
    initialY = 0 - BlockLength / 2
    initialZ = 0#
    
    Block1(0) = initialX
    Block1(1) = initialY
    Block1(2) = initialZ

    Block1(3) = initialX + BlockWidth
    Block1(4) = initialY
    Block1(5) = initialZ

    Block1(6) = Block1(3)
    Block1(7) = initialY + BlockLength
    Block1(8) = initialZ

    Block1(9) = initialX
    Block1(10) = Block1(7)
    Block1(11) = initialZ

    Block1(12) = initialX
    Block1(13) = initialY
    Block1(14) = initialZ
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    
    Set tmpLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, Block1)

    OutStr = "Block"
    ii = 0
    Set oVec = New DVector
    oVec.Set 0, 0, -1
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, BlockHeight)
    For Each pObj In elems
        Call InitNewOutput(pOC, OutStr)
        pOC.AddOutput OutStr, pObj
        ii = ii + 1
        OutStr = OutStr & Trim$(Str$(ii))
    Next
    Set elems = Nothing
    
    '=================
    'BUILD slab
    '=================
    Dim dSlab(0 To 14) As Double
    initialX = 0 - SlabWidth / 2
    initialY = 0 - SlabLength / 2
    initialZ = slabZ
    
    dSlab(0) = initialX
    dSlab(1) = initialY
    dSlab(2) = initialZ

    dSlab(3) = initialX + SlabWidth
    dSlab(4) = initialY
    dSlab(5) = initialZ

    dSlab(6) = dSlab(3)
    dSlab(7) = initialY + SlabLength
    dSlab(8) = initialZ

    dSlab(9) = initialX
    dSlab(10) = dSlab(7)
    dSlab(11) = initialZ

    dSlab(12) = initialX
    dSlab(13) = initialY
    dSlab(14) = initialZ
    
    Set tmpLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dSlab)
    OutStr = "slab"
    ii = 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, BlockHeight)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
        Call InitNewOutput(pOC, OutStr)
    Next
    Set elems = Nothing
    Set oVec = Nothing
        
    Set tmpLineString = Nothing
    Set oGeomFactory = Nothing
    Set pOC = Nothing
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


