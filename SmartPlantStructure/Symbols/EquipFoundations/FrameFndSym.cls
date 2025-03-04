VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameFndSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    FrameFndSym.cls
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
'   Jun 24, 2009        WR      TR-145118 - Updated the AddPad method to call
'                               the new CreateBySingleSweepWCapsOpts method.
'*******************************************************************

Private Const MODULE = "FrameFndSym"
Const m_ItemProgId As String = "SPSEqpFndMacros.FrameFndSym"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"

Implements IJDUserSymbolServices

Private Enum InputIndex
    EQPPORT_INDEX = 1
    SUPPPLANE_INDEX
    PLATELENGTH_INDEX
    PLATEWIDTH_INDEX
    PLATETHICKNESS_INDEX
    SIDEPLATEWIDTH_INDEX
    SIDEPLATETHICKNESS_INDEX
    LEGANGLEWIDTH_INDEX
    LEGANGLETHICKNESS_INDEX
    WITHPAD_INDEX
    PADTHICKNESS_INDEX
    PADRADIUS_INDEX
    HT_INDEX
    PLATESIZEBYRULE_INDEX
    PLATEEDGECLEARANCE_INDEX
    MATERIAL_INDEX
    GRADE_INDEX
End Enum

Private m_oComplex As IngrGeom3D.ComplexString3d
Private m_Transform As IngrGeom3D.IJDT4x4

Private Sub Class_Initialize()
    Set m_oComplex = Nothing
    Set m_Transform = Nothing
End Sub

Private Sub Class_Terminate()
    Set m_oComplex = Nothing
    Set m_Transform = Nothing
End Sub

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
    oInput.Description = "Supporting Equipment Port(s)"
    iInputs.SetInput oInput, EQPPORT_INDEX
    oInput.Reset

    oInput.name = "SupportingPlane"
    oInput.Description = "Supporting Surface/Plane"
    oInput.Properties = igDESCRIPTION_OPTIONAL
    iInputs.SetInput oInput, SUPPPLANE_INDEX
    oInput.Reset

    oInput.name = "PlateLength"
    oInput.Description = "Length of the equipment containing bolt-holes position"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.55
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, PLATELENGTH_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "PlateWidth"
    oInput.Description = "Mounting width which comes from the equipment"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.55
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, PLATEWIDTH_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "PlateThickness"
    oInput.Description = "Thickness of the top plate"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.012
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, PLATETHICKNESS_INDEX
    oInput.Reset
    PC.Reset

    oInput.name = "SidePlateWidth"
    oInput.Description = "Thickness of the top plate"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.13
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SIDEPLATEWIDTH_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SidePlateThickness"
    oInput.Description = "Thickness of the top plate"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.13
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, SIDEPLATETHICKNESS_INDEX
    oInput.Reset
    PC.Reset

    oInput.name = "LegAngleWidth"
    oInput.Description = "LegAngle SPS Section Name from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.13
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, LEGANGLEWIDTH_INDEX
    oInput.Reset
    PC.Reset

    oInput.name = "LegAngleThickness"
    oInput.Description = "LegAngle Cross Section Reference Standard "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.012
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, LEGANGLETHICKNESS_INDEX
    oInput.Reset
    PC.Reset

    oInput.name = "WithPads"
    oInput.Description = "Selects whether to place pads"
    PC.UomValue = 0#
    oInput.Properties = igINPUT_IS_A_PARAMETER
'    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, IsBoolean
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, WITHPAD_INDEX
    oInput.Reset
    PC.Reset

    oInput.name = "PadThickness"
    oInput.Description = "Thickness of the pad under the leg"
    PC.UomValue = 0.025
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, PADTHICKNESS_INDEX
    oInput.Reset
    PC.Reset

    oInput.name = "PadRadius"
    oInput.Description = "Radius of pad corners"
    PC.UomValue = 0.012
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, PADRADIUS_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "Height"
    oInput.Description = "Height(min) of frame foundation"
    PC.UomValue = 0.5
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, HT_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "IsPlateSizeDrivenByRule"
    oInput.Description = "Plate Size Driven By Rule (bolt hole locations)"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, PLATESIZEBYRULE_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "PlateEdgeClearance"
    oInput.Description = "PlateEdgeClearance"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    PC.UomValue = 0.0001
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, PLATEEDGECLEARANCE_INDEX
    oInput.Reset
    PC.Reset
    
    
    oInput.name = "SPSMaterial"
    oInput.Description = "SPSMaterial"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Concrete"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, MATERIAL_INDEX
    oInput.Reset
    PC.Reset
    
    oInput.name = "SPSGrade"
    oInput.Description = "SPSGrade"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    PC.Type = igString
    PC.String = "Fc 4000"
    oInput.DefaultParameterValue = PC
    iInputs.SetInput oInput, GRADE_INDEX
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
    
    output.name = "TopPlate"
    output.Description = "TopPlate"
    pOutputs.SetOutput output
    output.Reset

    output.name = "Angle1"
    output.Description = "FirstAngle"
    pOutputs.SetOutput output
    output.Reset
  
    output.name = "Angle2"
    output.Description = "SecondAngle"
    pOutputs.SetOutput output
    output.Reset
  
    output.name = "Angle3"
    output.Description = "ThirdAngle"
    pOutputs.SetOutput output
    output.Reset
  
    output.name = "Angle4"
    output.Description = "FourthAngle"
    pOutputs.SetOutput output
    output.Reset

    output.name = "SidePlate"
    output.Description = "SidePlate which is welded to the leg & top plate"
    pOutputs.SetOutput output
    output.Reset

    output.name = "SidePlate2"
    output.Description = "SidePlate2 which is welded to the leg & top plate"
    pOutputs.SetOutput output
    output.Reset

    output.name = "SideAngle1"
    output.Description = "SideAngle1 which is welded to the leg & top plate"
    pOutputs.SetOutput output
    output.Reset

    output.name = "SideAngle2"
    output.Description = "SideAngle2 which is welded to the leg & top plate"
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
    Dim PlateLength As Double
    Dim PlateWidth As Double
    Dim PlateThickness As Double
    Dim PlateEdgeClearance As Double
    Dim LegAngleWidth As Double
    Dim LegAngleThickness As Double
    Dim PadThickness As Double, PadRadius As Double
    Dim dblWithPads As Double
    Dim WithPads As Boolean

    Dim SidePlateWidth As Double, SidePlateThickness As Double
    Dim height As Double, FoundationType As Double
    Dim Offset As Double
    Offset = 0.01 'This value has been taken from the SAMSUNG drawing
    Dim LegLength As Double
    
    Dim LegSection As String, LegSecStandard As String
   
    Dim oVec As IJDVector
    Dim elems As IJElements
    Dim pObj As Object
    Dim OutStr As String
    Dim ii As Integer
    
    Dim pRepDG As IJDRepresentationDuringGame
    Set pRepDG = pIRepSCM
    Dim pOC As IJDOutputCollection
    Set pOC = pRepDG.outputCollection
    
    'remove all outputs
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep
    oOutputs.RemoveAllOutput
    
    Dim pInputs As IJDInputs
    Set pInputs = pRepDG.definition.IJDInputs
        
    Dim tmpLineString As IngrGeom3D.LineString3d
    Dim oProjection As IngrGeom3D.Projection3d
    'assign to meaningful variables from the input array
    PlateEdgeClearance = pInputs.GetInputByIndex(PLATEEDGECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    PlateLength = pInputs.GetInputByIndex(PLATELENGTH_INDEX).IJDInputDuringGame.Result.UomValue + PlateEdgeClearance
    PlateWidth = pInputs.GetInputByIndex(PLATEWIDTH_INDEX).IJDInputDuringGame.Result.UomValue + PlateEdgeClearance
    PlateThickness = pInputs.GetInputByIndex(PLATETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
        
    LegAngleWidth = pInputs.GetInputByIndex(LEGANGLEWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    LegAngleThickness = pInputs.GetInputByIndex(LEGANGLETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue

    PadThickness = pInputs.GetInputByIndex(PADTHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    PadRadius = pInputs.GetInputByIndex(PADRADIUS_INDEX).IJDInputDuringGame.Result.UomValue
    dblWithPads = pInputs.GetInputByIndex(WITHPAD_INDEX).IJDInputDuringGame.Result.UomValue
    
    If dblWithPads = 0# Then
        WithPads = False
    Else
        WithPads = True
    End If

    SidePlateWidth = pInputs.GetInputByIndex(SIDEPLATEWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    SidePlateThickness = pInputs.GetInputByIndex(SIDEPLATETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue

    height = pInputs.GetInputByIndex(HT_INDEX).IJDInputDuringGame.Result.UomValue

    LegLength = height - PlateThickness
    If WithPads Then
        LegLength = LegLength - PadThickness
    End If

    '=================
    'BUILD TOP PLATE
    '=================
    Dim TopPlate(0 To 14) As Double
    Dim initialX As Double, initialY As Double, initialZ As Double
    initialX = 0 - PlateWidth / 2
    initialY = 0 - PlateLength / 2
    initialZ = 0#
    
    TopPlate(0) = initialX
    TopPlate(1) = initialY
    TopPlate(2) = initialZ

    TopPlate(3) = initialX + PlateWidth
    TopPlate(4) = initialY
    TopPlate(5) = initialZ

    TopPlate(6) = TopPlate(3)
    TopPlate(7) = initialY + PlateLength
    TopPlate(8) = initialZ

    TopPlate(9) = initialX
    TopPlate(10) = TopPlate(7)
    TopPlate(11) = initialZ

    TopPlate(12) = initialX
    TopPlate(13) = initialY
    TopPlate(14) = initialZ
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory

    Set tmpLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, TopPlate)
    OutStr = "TopPlate"
    ii = 0
    Set oVec = New DVector
    oVec.Set 0, 0, -1
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, PlateThickness)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
   Next
    Set elems = Nothing
    
    '=================
    'BUILD FIRST ANGLE
    '=================
    Dim angle1(0 To 20) As Double
    
    initialX = 0 - PlateWidth / 2 + Offset
    initialY = 0 - PlateLength / 2 + Offset

    
    angle1(0) = initialX
    angle1(1) = initialY
    angle1(2) = -PlateThickness
    
    angle1(3) = initialX + LegAngleWidth
    angle1(4) = initialY
    angle1(5) = -PlateThickness
    
    angle1(6) = angle1(3)
    angle1(7) = initialY + LegAngleWidth
    angle1(8) = -PlateThickness
    
    angle1(9) = angle1(3) - LegAngleThickness
    angle1(10) = angle1(7)
    angle1(11) = -PlateThickness
    
    angle1(12) = angle1(9)
    angle1(13) = initialY + LegAngleThickness
    angle1(14) = -PlateThickness
    
    angle1(15) = initialX
    angle1(16) = angle1(13)
    angle1(17) = -PlateThickness
    
    angle1(18) = angle1(0)
    angle1(19) = angle1(1)
    angle1(20) = angle1(2)

    tmpLineString.SetPoints 7, angle1
    OutStr = "Angle1"
    ii = 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, LegLength)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
    Next
    Set elems = Nothing
    
    If WithPads Then
        InitPad PadRadius, LegAngleWidth, -(height - PadThickness)
        AddPad pOC, "angle1Pad", PadThickness, 90#, angle1(3), angle1(4)
    End If

    '=================
    'BUILD SECOND ANGLE
    '=================
    Dim angle2(0 To 20) As Double
    
    initialX = 0 - PlateWidth / 2 + PlateWidth - Offset
    initialY = 0 - PlateLength / 2 + Offset
    
    angle2(0) = initialX
    angle2(1) = initialY
    angle2(2) = -PlateThickness
    
    angle2(3) = angle2(0)
    angle2(4) = initialY + LegAngleThickness
    angle2(5) = -PlateThickness
    
    angle2(6) = angle2(0) - LegAngleWidth + LegAngleThickness
    angle2(7) = angle2(4)
    angle2(8) = -PlateThickness
    
    angle2(9) = angle2(6)
    angle2(10) = angle2(1) + LegAngleWidth
    angle2(11) = -PlateThickness
    
    angle2(12) = angle2(9) - LegAngleThickness
    angle2(13) = angle2(10)
    angle2(14) = -PlateThickness
    
    angle2(15) = angle2(12)
    angle2(16) = angle2(1)
    angle2(17) = -PlateThickness
    
    angle2(18) = angle2(0)
    angle2(19) = angle2(1)
    angle2(20) = angle2(2)

    tmpLineString.SetPoints 7, angle2
    OutStr = "Angle2"
    ii = 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, LegLength)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
    Next
    Set elems = Nothing
    
    If WithPads Then
        AddPad pOC, "angle2Pad", PadThickness, 0#, angle2(15), angle2(16)
    End If
    
    '=================
    'BUILD THIRD ANGLE
    '=================
    Dim angle3(0 To 20) As Double
    initialX = 0 - PlateWidth / 2 + PlateWidth - Offset
    initialY = 0 - PlateLength / 2 + PlateLength - Offset
    
    angle3(0) = initialX
    angle3(1) = initialY
    angle3(2) = -PlateThickness
    
    angle3(3) = angle2(12)
    angle3(4) = angle3(1)
    angle3(5) = -PlateThickness
    
    angle3(6) = angle3(3)
    angle3(7) = angle3(4) - LegAngleWidth
    angle3(8) = -PlateThickness
    
    angle3(9) = angle2(6)
    angle3(10) = angle3(7)
    angle3(11) = -PlateThickness
    
    angle3(12) = angle3(9)
    angle3(13) = angle3(1) - LegAngleThickness
    angle3(14) = -PlateThickness
    
    angle3(15) = angle3(0)
    angle3(16) = angle3(13)
    angle3(17) = -PlateThickness
    
    angle3(18) = angle3(0)
    angle3(19) = angle3(1)
    angle3(20) = angle3(2)

    tmpLineString.SetPoints 7, angle3
    OutStr = "Angle3"
    ii = 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, LegLength)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
    Next
    Set elems = Nothing
    
    If WithPads Then
        AddPad pOC, "angle3Pad", PadThickness, 270#, angle3(3), angle3(4)
    End If

    '=================
    'BUILD FOURTH ANGLE
    '=================
    Dim angle4(0 To 20) As Double
    
    angle4(0) = angle1(0)
    angle4(1) = angle3(1)
    angle4(2) = -PlateThickness
    
    angle4(3) = angle4(0)
    angle4(4) = angle4(1) - LegAngleThickness
    angle4(5) = -PlateThickness
    
    angle4(6) = angle1(9)
    angle4(7) = angle4(4)
    angle4(8) = -PlateThickness
    
    angle4(9) = angle4(6)
    angle4(10) = angle4(1) - LegAngleWidth
    angle4(11) = -PlateThickness
    
    angle4(12) = angle1(3)
    angle4(13) = angle4(10)
    angle4(14) = -PlateThickness
    
    angle4(15) = angle4(12)
    angle4(16) = angle4(1)
    angle4(17) = -PlateThickness
    
    angle4(18) = angle4(0)
    angle4(19) = angle4(1)
    angle4(20) = angle4(2)

    tmpLineString.SetPoints 7, angle4
    OutStr = "Angle4"
    ii = 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, LegLength)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
   Next
    Set elems = Nothing
    
    If WithPads Then
        AddPad pOC, "angle4Pad", PadThickness, 180#, angle4(15), angle4(16)
    End If
    
    '=================
    'BUILD SIDE PLATE1
    '=================
    Dim SidePlateCS1(0 To 14) As Double
    initialY = Offset - PlateLength / 2
    SidePlateCS1(0) = angle1(3)
    SidePlateCS1(1) = initialY
    SidePlateCS1(2) = -(PlateThickness + SidePlateWidth) '-ve sign indicates into the plane of the screen
    
    SidePlateCS1(3) = angle2(12)
    SidePlateCS1(4) = initialY 'angle2(16)
    SidePlateCS1(5) = -(PlateThickness + SidePlateWidth)
    
    SidePlateCS1(6) = SidePlateCS1(3)
    SidePlateCS1(7) = initialY 'SidePlateCS1(4)
    SidePlateCS1(8) = -PlateThickness
    
    SidePlateCS1(9) = SidePlateCS1(0)
    SidePlateCS1(10) = initialY 'SidePlateCS1(1)
    SidePlateCS1(11) = -PlateThickness
    
    SidePlateCS1(12) = SidePlateCS1(0)
    SidePlateCS1(13) = SidePlateCS1(1)
    SidePlateCS1(14) = SidePlateCS1(2)

    tmpLineString.SetPoints 5, SidePlateCS1
    OutStr = "SidePlate"
    ii = 0
    oVec.Set 0, 1, 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, SidePlateThickness)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
    Next
    Set elems = Nothing
    
    '=================
    'BUILD SIDE PLATE2
    '=================
    Dim SidePlateCS2(0 To 14) As Double

    SidePlateCS2(0) = SidePlateCS1(3)
    SidePlateCS2(1) = angle3(4)
    SidePlateCS2(2) = -(PlateThickness + SidePlateWidth) '-ve sign indicates into the plane of the screen
    
    SidePlateCS2(3) = SidePlateCS1(0)
    SidePlateCS2(4) = SidePlateCS2(1)
    SidePlateCS2(5) = -(PlateThickness + SidePlateWidth)
    
    SidePlateCS2(6) = SidePlateCS2(3)
    SidePlateCS2(7) = SidePlateCS2(4)
    SidePlateCS2(8) = -PlateThickness
    
    SidePlateCS2(9) = SidePlateCS2(0)
    SidePlateCS2(10) = SidePlateCS2(1)
    SidePlateCS2(11) = -PlateThickness
    
    SidePlateCS2(12) = SidePlateCS2(0)
    SidePlateCS2(13) = SidePlateCS2(1)
    SidePlateCS2(14) = SidePlateCS2(2)

    tmpLineString.SetPoints 5, SidePlateCS2
    OutStr = "SidePlate2"
    ii = 0
    oVec.Set 0, -1, 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, SidePlateThickness)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
    Next
    Set elems = Nothing
    
    'Build horizontal SideAngle1
    'points clockwise beginning with line below top plate
    'in X-Z plane, project along +Y
    Dim pts(0 To 20) As Double
    initialX = Offset - PlateWidth / 2
    initialY = Offset + LegAngleThickness - PlateLength / 2
    
    pts(0) = initialX
    pts(1) = initialY
    pts(2) = -PlateThickness

    pts(3) = initialX + LegAngleWidth
    pts(4) = pts(1)
    pts(5) = -PlateThickness
    
    pts(6) = initialX + LegAngleWidth
    pts(7) = pts(1)
    pts(8) = -(PlateThickness + LegAngleWidth)
    
    pts(9) = initialX + LegAngleWidth - LegAngleThickness
    pts(10) = pts(1)
    pts(11) = -(PlateThickness + LegAngleWidth)
    
    pts(12) = initialX + LegAngleWidth - LegAngleThickness
    pts(13) = pts(1)
    pts(14) = -(PlateThickness + LegAngleThickness)
    
    pts(15) = initialX
    pts(16) = pts(1)
    pts(17) = -(PlateThickness + LegAngleThickness)
    
    pts(18) = pts(0)
    pts(19) = pts(1)
    pts(20) = pts(2)

    tmpLineString.SetPoints 7, pts
    OutStr = "SideAngle1"
    ii = 0
    oVec.Set 0, 1, 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, PlateLength - 2# * Offset - 2# * LegAngleThickness)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
    Next
    Set elems = Nothing
    
    'Build horizontal SideAngle2
    'points clockwise beginning with line below top plate
    'in X-Z plane, project along +Y
    initialX = PlateWidth - Offset - LegAngleWidth - PlateWidth / 2
    
    pts(0) = initialX 'PlateWidth - Offset - LegAngleWidth
    pts(2) = -PlateThickness

    pts(3) = initialX + LegAngleWidth ' PlateWidth - Offset
    pts(5) = -PlateThickness
    
    pts(6) = initialX + LegAngleWidth 'PlateWidth - Offset
    pts(8) = -(PlateThickness + LegAngleThickness)
    
    pts(9) = initialX + LegAngleThickness 'PlateWidth - Offset - LegAngleWidth + LegAngleThickness
    pts(11) = -(PlateThickness + LegAngleThickness)
    
    pts(12) = initialX + LegAngleThickness 'PlateWidth - Offset - LegAngleWidth + LegAngleThickness
    pts(14) = -(PlateThickness + LegAngleWidth)
    
    pts(15) = initialX  'PlateWidth - Offset - LegAngleWidth
    pts(17) = -(PlateThickness + LegAngleWidth)
    
    pts(18) = pts(0)
    pts(20) = pts(2)

    tmpLineString.SetPoints 7, pts
    OutStr = "SideAngle2"
    ii = 0
    Set elems = CreateSolidAsPlanes(pOC.ResourceManager, tmpLineString, oVec, PlateLength - 2# * Offset - 2# * LegAngleThickness)
    For Each pObj In elems
         pOC.AddOutput OutStr, pObj
         ii = ii + 1
         OutStr = OutStr & Trim$(Str$(ii))
         Call InitNewOutput(pOC, OutStr)
   Next
    
    Set elems = Nothing
    Set oVec = Nothing
    Set tmpLineString = Nothing
    Set pOC = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Private Sub InitPad(PadRadius As Double, LegAngleWidth As Double, z As Double)
Const METHOD = "InitPad"
On Error GoTo ErrorHandler
    ' create a template complex curve, as for an angle whose outside
    ' corner is at (0,0), one leg goes along pos x and one along pos y
    '
    '   horizontal line to right below the x-axis
    '   135 degree sweep arc ccw around the right end
    '   line going from lower-right to upper-left
    '   135 degree sweep arc ccw around the upper end
    '   line going from upper-left downward
    '   90 degree sweep arc ccw back to start
    '
    ' create the transform object and set to identity

    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim iElements As IJElements
    Dim psq2 As Double             ' PadRadius * squareRoot(two) / 2

    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    
    psq2 = PadRadius * Sqr(2#) / 2#

    x1 = 0#
    y1 = -PadRadius
    x2 = LegAngleWidth
    y2 = -PadRadius
    
    Dim oGeometryFactory As New IngrGeom3D.GeometryFactory
    
    Set oLine = oGeometryFactory.Lines3d.CreateBy2Points(Nothing, x1, y1, z, x2, y2, z)
    Set iElements = New JObjectCollection  'IMSElements.DynElements
    iElements.Add oLine
    Set m_oComplex = oGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing

    x1 = x2
    y1 = y2
    x2 = LegAngleWidth + PadRadius
    y2 = 0#
    x3 = LegAngleWidth + psq2
    y3 = psq2
    Set oArc = oGeometryFactory.Arcs3d.CreateBy3Points(Nothing, x1, y1, z, x2, y2, z, x3, y3, z)
    m_oComplex.AddCurve oArc, True

    x1 = x3
    y1 = y3
    x2 = psq2
    y2 = LegAngleWidth + psq2
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_oComplex.AddCurve oLine, True
    
    x1 = x2
    y1 = y2
    x2 = 0#
    y2 = LegAngleWidth + PadRadius
    x3 = -PadRadius
    y3 = LegAngleWidth
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
    m_oComplex.AddCurve oArc, True
    
    x1 = x3
    y1 = y3
    x2 = -PadRadius
    y2 = 0#
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_oComplex.AddCurve oLine, True
    
    x1 = x2
    y1 = y2
    x2 = -psq2
    y2 = -psq2
    x3 = 0#
    y3 = -PadRadius
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
    m_oComplex.AddCurve oArc, True
    
    Set oArc = Nothing
    Set oLine = Nothing
    
    Set m_Transform = New DT4x4
    m_Transform.LoadIdentity
    Set oGeometryFactory = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub AddPad(pOC As IJDOutputCollection, padName As String, PadThickness As Double, _
            rotAngle As Double, posX As Double, posY As Double)
Const METHOD = "AddPad"
On Error GoTo ErrorHandler

    Dim iVector As IJDVector
    Dim oProj As IngrGeom3D.Projection3d

    Dim oGeometryFactory As New IngrGeom3D.GeometryFactory

    Set iVector = New DVector
    iVector.x = posX
    iVector.y = posY
    iVector.z = 0#
    m_Transform.Translate iVector
    iVector.x = 0#
    iVector.y = 0#
    iVector.z = 1#
    m_Transform.Rotate rotAngle * 3.14159265358979 / 180#, iVector
    
    m_oComplex.Transform m_Transform

    Dim oCurve As IJCurve
    m_oComplex.GetCurve 1, oCurve
    
    Dim sX As Double, sY As Double, sZ As Double
    Dim eX As Double, eY As Double, eZ As Double
    oCurve.EndPoints sX, sY, sZ, eX, eY, eZ
    Set oCurve = Nothing
    
    ' create a trace curve
    Dim oTrace As Line3d
    Set oTrace = oGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, sX, sY, sZ, 0#, 0#, -1#, PadThickness)
                                                            
    ' create solid using singlesweep
    ' use break both trace and cross section option and with caps
    Dim stNorm() As Double, endNorm() As Double
    Dim numCaps As Long
    
    Dim padEles As IJElements
    Set padEles = oGeometryFactory.GeometryServices.CreateBySingleSweepWCapsOpts(pOC.ResourceManager, _
                                                        oTrace, m_oComplex, _
                                                        CircularCorner, BreakPathAndCrossSection, _
                                                        StartAtTraceBeg, SkinningCrossSectionOrientation.TraditionalOrientation, stNorm, endNorm, _
                                                       True, numCaps)
    Dim pObj As Object
    Dim ii As Integer
    ii = 0
    For Each pObj In padEles
        pOC.AddOutput padName, pObj
        ii = ii + 1
        padName = padName & Trim$(Str$(ii))
    Next
    Set padEles = Nothing
    Set oTrace = Nothing
    
    m_Transform.Invert
    m_oComplex.Transform m_Transform
    
    m_Transform.LoadIdentity

    Set oProj = Nothing
    Set iVector = Nothing
    Set oGeometryFactory = Nothing
    Exit Sub
    
ErrorHandler:
       HandleError MODULE, METHOD

End Sub


