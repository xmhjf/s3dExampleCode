VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PierAndSlabFtgSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    PierAndSlabFtgSym.cls
'
'Author
'       28th January 2003        Aniket Patil
'
'Description
'
'Notes
'
'History:
'
'   19-jan-07       SS      TR#113234 - Made member as an optional input to the symbol. Otherwise
'                           when member is deleted footing gets deleted too instead of going to TDL
'
'*******************************************************************
Option Explicit
Private Const MODULE = "PierAndSlabFtgSym"
Const m_ItemProgId As String = "SPSFootingMacros.PierAndSlabFtgSym"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"

Private Enum InputIndex
SUPPORTEDMEMB_INDEX = 1
WITHGROUTPAD_INDEX
GLOBALDELTA_INDEX
GROUTSHAPE_INDEX
GROUTSIZINGRULE_INDEX
GROUTORIENTATION_INDEX
GROUTROTATIONANGLE_INDEX
GROUTEDGECLEARANCE_INDEX
GROUTLENGTH_INDEX
GROUTWIDTH_INDEX
GROUTHEIGHT_INDEX
GROUTSPSMATERIAL_INDEX
GROUTSPSGRADE_INDEX
PIERSHAPE_INDEX
PIERSIZINGRULE_INDEX
PIERORIENTATION_INDEX
PIERROTATIONANGLE_INDEX
PIEREDGECLEARANCE_INDEX
PIERSIZEINCREMENT_INDEX
PIERCHAMFERED_INDEX
PIERCHAMFERSIZE_INDEX
PIERSPSMATERIAL_INDEX
PIERSPSGRADE_INDEX
PIERLENGTH_INDEX
PIERWIDTH_INDEX
PIERHEIGHT_INDEX
SLABSHAPE_INDEX
SLABSIZINGRULE_INDEX
SLABORIENTATION_INDEX
SLABROTATIONANGLE_INDEX
SLABEDGECLEARANCE_INDEX
SLABSIZEINCREMENT_INDEX
SLABSPSMATERIAL_INDEX
SLABSPSGRADE_INDEX
SLABLENGTH_INDEX
SLABWIDTH_INDEX
SLABHEIGHT_INDEX
DETAILREFERENCE_INDEX
End Enum

Implements IJDUserSymbolServices
Implements IJSymbolVersion
'Implements IJDDefinitionPlayerExOverride
'Private Sub IJDDefinitionPlayerExOverride_BindInputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
'
'End Sub
'
'Private Sub IJDDefinitionPlayerExOverride_CloseGame(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
'
'End Sub
'
'Private Sub IJDDefinitionPlayerExOverride_ComputeConverter(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
'
'End Sub
'
'Private Sub IJDDefinitionPlayerExOverride_ComputeMatrix(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
'
'     CMFootingPosition pSymbolDefinition
'
'End Sub
'
'Private Sub IJDDefinitionPlayerExOverride_ConvertInputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
'
'End Sub
'
'Private Sub IJDDefinitionPlayerExOverride_OpenGame(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pPlayingSymbol As Object)
'
'End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
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
 
ErrorHandler:     HandleError MODULE, METHOD
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
     
'     Dim mthDesc As IMSSymbolEntities.IJDMethodDescription
'     Set mthDesc = New DMethodDescription
'
'     mthDesc.name = "ComputeMatrix"  'Has to be this name as "ComputeMatrix"
'     mthDesc.Properties = imsMETHOD_OVERRIDE
'     mthDesc.Cookie = imsCOOKIE_IS_NULL
'     mthDesc.Library = imsCOOKIE_ID_USS_LIB  ' the current USS acts a library
'     iUM.SetMethod mthDesc
'     Set mthDesc = Nothing
     
     
     Dim libDesc As New DLibraryDescription
     Dim mCookie As Long
     Dim libCookie As Long
     Dim methodCookie As Long
     
     
     libDesc.name = "mySelfAsLib"
     libDesc.Type = imsLIBRARY_IS_ACTIVEX
     libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
     libDesc.Source = m_ItemProgId
     
     pSymbolDefinition.IJDUserMethods.SetLibrary libDesc
     
     libCookie = libDesc.Cookie
    
    Dim ChecklibDesc As New DLibraryDescription
    Dim ChecklibCookie As Long
    Dim NegativeNumCheck As Long
    
    ChecklibDesc.name = "CMCheckLib"
    ChecklibDesc.Type = imsLIBRARY_IS_ACTIVEX
    ChecklibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    ChecklibDesc.Source = CheckProgId
    pSymbolDefinition.IJDUserMethods.SetLibrary ChecklibDesc
    ChecklibCookie = ChecklibDesc.Cookie

     NegativeNumCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("NegativeNum", ChecklibCookie)
      
     Dim pIJDInput As IMSSymbolEntities.IJDInput
     Set pIJDInput = New IMSSymbolEntities.DInput
     
     Dim oInput As IMSSymbolEntities.IJDInput
     Set oInput = New IMSSymbolEntities.DInput
     Dim PC As IMSSymbolEntities.IJDParameterContent
     Set PC = New IMSSymbolEntities.DParameterContent
     PC.Type = igValue

     Dim inputsProp As IMSDescriptionProperties
     inputsProp = pSymbolDefinition.IJDInputs.Property
     pSymbolDefinition.IJDInputs.Property = inputsProp Or igCOLLECTION_VARIABLE

     oInput.name = "SupportedMember"
     oInput.Description = "Supported"
'     methodCookie = iUM.GetMethodCookie("CMCacheSupported", libCookie)
'     oInput.IJDInputStdCustomMethod.SetCMCache imsCOOKIE_ID_USS_LIB, methodCookie
     oInput.Properties = igDESCRIPTION_OPTIONAL
     iInputs.SetInput oInput, SUPPORTEDMEMB_INDEX
     oInput.Reset
      
     oInput.name = WITH_GROUT_PAD
     oInput.Description = WITH_GROUT_PAD
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, WITHGROUTPAD_INDEX
     oInput.Reset
     PC.Reset
     
     oInput.name = "GlobalDelta"
     oInput.Description = "GlobalDelta"
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GLOBALDELTA_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_SHAPE
     oInput.Description = GROUT_SHAPE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTSHAPE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_SIZE_RULE
     oInput.Description = GROUT_SIZE_RULE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTSIZINGRULE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_ORIENTATION
     oInput.Description = GROUT_ORIENTATION
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 3
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTORIENTATION_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_ROTATION_ANGLE
     oInput.Description = GROUT_ROTATION_ANGLE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTROTATIONANGLE_INDEX
     oInput.Reset
     PC.Reset


     oInput.name = GROUT_EDGE_CLEARANCE
     oInput.Description = GROUT_EDGE_CLEARANCE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTEDGECLEARANCE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_LENGTH
     oInput.Description = GROUT_LENGTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 16
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTLENGTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_WIDTH
     oInput.Description = GROUT_WIDTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 16
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTWIDTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_HEIGHT
     oInput.Description = GROUT_HEIGHT
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTHEIGHT_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_MATERIAL
     oInput.Description = GROUT_MATERIAL
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Grout"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTSPSMATERIAL_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = GROUT_GRADE
     oInput.Description = GROUT_GRADE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "High Strength"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, GROUTSPSGRADE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_SHAPE
     oInput.Description = PIER_SHAPE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERSHAPE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_SIZING_RULE
     oInput.Description = PIER_SIZING_RULE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERSIZINGRULE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_ORIENTATION
     oInput.Description = PIER_ORIENTATION
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERORIENTATION_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_ROTATION_ANGLE
     oInput.Description = PIER_ROTATION_ANGLE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERROTATIONANGLE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_EDGE_CLEARANCE
     oInput.Description = PIER_EDGE_CLEARANCE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIEREDGECLEARANCE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_SIZE_INC
     oInput.Description = PIER_SIZE_INC
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, NegativeNumCheck
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERSIZEINCREMENT_INDEX
     oInput.Reset
     PC.Reset


     oInput.name = "PierChamfered"
     oInput.Description = "PierChamfered"
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERCHAMFERED_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = "PierChamferSize"
     oInput.Description = "PierChamferSize"
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERCHAMFERSIZE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_MATERIAL
     oInput.Description = PIER_MATERIAL
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Concrete"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERSPSMATERIAL_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_GRADE
     oInput.Description = PIER_GRADE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Fc 3000"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERSPSGRADE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_LENGTH
     oInput.Description = PIER_LENGTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 18
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERLENGTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_WIDTH
     oInput.Description = PIER_WIDTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 18
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERWIDTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_HEIGHT
     oInput.Description = PIER_HEIGHT
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 24
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERHEIGHT_INDEX
     oInput.Reset
     PC.Reset
     
     oInput.name = SLAB_SHAPE
     oInput.Description = SLAB_SHAPE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABSHAPE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_SIZING_RULE
     oInput.Description = SLAB_SIZING_RULE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABSIZINGRULE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_ORIENTATION
     oInput.Description = SLAB_ORIENTATION
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABORIENTATION_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_ROTATION_ANGLE
     oInput.Description = SLAB_ROTATION_ANGLE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABROTATIONANGLE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_EDGE_CLEARANCE
     oInput.Description = SLAB_EDGE_CLEARANCE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 24
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABEDGECLEARANCE_INDEX
     oInput.Reset
     PC.Reset
 
     oInput.name = SLAB_SIZE_INC
     oInput.Description = SLAB_SIZE_INC
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, NegativeNumCheck
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABSIZEINCREMENT_INDEX
     oInput.Reset
     PC.Reset
 
     oInput.name = SLAB_MATERIAL
     oInput.Description = SLAB_MATERIAL
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Concrete"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABSPSMATERIAL_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_GRADE
     oInput.Description = SLAB_GRADE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Fc 3000"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABSPSGRADE_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_LENGTH
     oInput.Description = SLAB_LENGTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 5
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABLENGTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_WIDTH
     oInput.Description = SLAB_WIDTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 5
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABWIDTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_HEIGHT
     oInput.Description = SLAB_HEIGHT
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 16
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, SLABHEIGHT_INDEX
     oInput.Reset
     PC.Reset
     
     oInput.name = "DetailReference"
     oInput.Description = "DetailReference"
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = " "
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, DETAILREFERENCE_INDEX
     oInput.Reset
     PC.Reset
          
          
     Dim pIReps As IMSSymbolEntities.IJDRepresentations
     Set pIReps = pSymbolDefinition
     Dim pIRep As IMSSymbolEntities.IJDRepresentation
     Set pIRep = New IMSSymbolEntities.DRepresentation
     
     pIRep.name = "Physical"
     pIRep.Description = "Physical representation"
'     pIRep.Properties = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
     pIRep.RepresentationId = SimplePhysical
     mCookie = iUM.GetMethodCookie("Physical", libCookie)
     pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
     
     Dim pOutputs As IMSSymbolEntities.IJDOutputs
     Set pOutputs = pIRep
     pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
     
     Dim output As IMSSymbolEntities.IJDOutput
     Set output = New IMSSymbolEntities.DOutput
     
     output.name = "Rectangle"
     output.Description = "Rectangle Solid"
     pOutputs.SetOutput output
     output.Reset
       
     pIReps.SetRepresentation pIRep 'Add representation to definition
     
    pIRep.name = "DetailPhysical"
    pIRep.Description = "DetailPhysical representation"
    'pIRep.Properties = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIRep.RepresentationId = DetailPhysical
    mCookie = iUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    Set pOutputs = pIRep
    pIReps.SetRepresentation pIRep 'Add representation to definition
        
'    'For Reference Geometry Representation
    pOutputs.RemoveAllOutput
    Set output = New DOutput

    pIRep.name = "RefGeometry"
    pIRep.Description = "ReferenceGeometry  representation"
    pIRep.Properties = igREPRESENTATION_ISVBFUNCTION
    'pIRep.Properties = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIRep.RepresentationId = ReferenceGeometry
    mCookie = iUM.GetMethodCookie("RefGeometry", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    
    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    Set output = New IMSSymbolEntities.DOutput
    output.name = "ReferenceGeometry1"
    output.Description = "Reference Geometry of Pier and Slab Footing"
    pOutputs.SetOutput output
    
    output.Reset
    pIReps.SetRepresentation pIRep 'Add representation to definition
    
    ' Set definition properties
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_SHARED  ' The definition will use the symbol cache machinery

    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

         
Exit Sub
ErrorHandler:    HandleError MODULE, METHOD
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
     Dim pRepDG As IJDRepresentationDuringGame
     Set pRepDG = pIRepSCM
     Dim pOC As IJDOutputCollection
     Set pOC = pRepDG.outputCollection
     
     Dim pInputs As IJDInputs
     Set pInputs = pRepDG.definition.IJDInputs
     
    
     Dim oRep As IJDRepresentation
     Dim oOutputs As IJDOutputs
     Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
     Set oOutputs = oRep
     oOutputs.RemoveAllOutput
'    Dim oMembPart As ISPSMemberPartPrismatic
     Dim withGroutPad As Boolean
     Dim GroutLength As Double, GroutWidth As Double, GroutHeight As Double
     Dim PierLength As Double, PierWidth As Double, PierHeight As Double
     Dim SlabLength As Double, SlabWidth As Double, SlabHeight As Double
     Dim GroutShape As Long, PierShape As Long, SlabShape As Long
     Dim GroutEdgeClearance As Double, PierEdgeClearance As Double, SlabEdgeClearance As Double
     Dim GroutRotationAngle As Double, PierRotationAngle As Double, SlabRotationAngle As Double
     Dim GroutOrientation As Long, PierOrientation As Long, SlabOrientation As Long
     Dim GlobalDelta As Double
 
     GlobalDelta = pInputs.GetInputByIndex(GLOBALDELTA_INDEX).IJDInputDuringGame.Result.UomValue
     withGroutPad = pInputs.GetInputByIndex(WITHGROUTPAD_INDEX).IJDInputDuringGame.Result.UomValue
     GroutLength = pInputs.GetInputByIndex(GROUTLENGTH_INDEX).IJDInputDuringGame.Result.UomValue
     GroutWidth = pInputs.GetInputByIndex(GROUTWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
     GroutHeight = pInputs.GetInputByIndex(GROUTHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
     GroutShape = pInputs.GetInputByIndex(GROUTSHAPE_INDEX).IJDInputDuringGame.Result.UomValue
     GroutEdgeClearance = pInputs.GetInputByIndex(GROUTEDGECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
     GroutRotationAngle = pInputs.GetInputByIndex(GROUTROTATIONANGLE_INDEX).IJDInputDuringGame.Result.UomValue
     GroutOrientation = pInputs.GetInputByIndex(GROUTORIENTATION_INDEX).IJDInputDuringGame.Result.UomValue
     
     CheckForUndefinedValueAndRaiseError pRepDG, GroutShape, PRISMATIC_FOOTING_SHAPES, 121
     CheckForUndefinedValueAndRaiseError pRepDG, GroutOrientation, STRUCT_COORD_SYS_REF, 123
          
     PierLength = pInputs.GetInputByIndex(PIERLENGTH_INDEX).IJDInputDuringGame.Result.UomValue
     PierWidth = pInputs.GetInputByIndex(PIERWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
     PierHeight = pInputs.GetInputByIndex(PIERHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
     PierShape = pInputs.GetInputByIndex(PIERSHAPE_INDEX).IJDInputDuringGame.Result.UomValue
     PierEdgeClearance = pInputs.GetInputByIndex(PIEREDGECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
     PierRotationAngle = pInputs.GetInputByIndex(PIERROTATIONANGLE_INDEX).IJDInputDuringGame.Result.UomValue
     PierOrientation = pInputs.GetInputByIndex(PIERORIENTATION_INDEX).IJDInputDuringGame.Result.UomValue
     
     CheckForUndefinedValueAndRaiseError pRepDG, PierShape, PRISMATIC_FOOTING_SHAPES, 124
     CheckForUndefinedValueAndRaiseError pRepDG, PierOrientation, STRUCT_COORD_SYS_REF, 126
     
     SlabLength = pInputs.GetInputByIndex(SLABLENGTH_INDEX).IJDInputDuringGame.Result.UomValue
     SlabWidth = pInputs.GetInputByIndex(SLABWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
     SlabHeight = pInputs.GetInputByIndex(SLABHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
     SlabShape = pInputs.GetInputByIndex(SLABSHAPE_INDEX).IJDInputDuringGame.Result.UomValue
     SlabEdgeClearance = pInputs.GetInputByIndex(SLABEDGECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
     SlabRotationAngle = pInputs.GetInputByIndex(SLABROTATIONANGLE_INDEX).IJDInputDuringGame.Result.UomValue
     SlabOrientation = pInputs.GetInputByIndex(SLABORIENTATION_INDEX).IJDInputDuringGame.Result.UomValue
     
     CheckForUndefinedValueAndRaiseError pRepDG, SlabShape, PRISMATIC_FOOTING_SHAPES, 127
     CheckForUndefinedValueAndRaiseError pRepDG, SlabOrientation, STRUCT_COORD_SYS_REF, 129
          
     Dim pts(15) As Double
     Dim octpts(27) As Double
     Dim pIJLineString As IJLineString
     Set pIJLineString = New LineString3d
     
     Dim elems As IJElements
     Set elems = New JObjectCollection ' IMSElements.DynElements
     Dim i As Integer
     Dim Matrix As IJDT4x4
     Set Matrix = New DT4x4
     Dim Vec As DVector
     Set Vec = New DVector
     Vec.Set 0, 0, 1
     Dim plane As IngrGeom3D.Plane3d
     Dim oPoint3d As IngrGeom3D.Point3d
     Dim oGeomFactory As New IngrGeom3D.GeometryFactory
     Dim proj As IJProjection
     Dim oCircle As New Circle3d
'     Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, 0)
'     pOC.AddOutput "Point1", oPoint3d

     If withGroutPad Then
    'GroutPad Geometry
         Matrix.LoadIdentity
         If GroutOrientation = 1 Then '1 is Global
         
            
            Matrix.Rotate GroutRotationAngle, Vec
         ElseIf GroutOrientation = 2 Then '2 is Local to the SPSMemberPart
            GroutRotationAngle = GroutRotationAngle + GlobalDelta
            Matrix.Rotate GroutRotationAngle, Vec
         End If
         
         If GroutShape = 3 Then
             oCircle.Radius = (GroutWidth / 2) '+ (GroutEdgeClearance * 5)'TR#72794 - Edge clearance is already considered incalculation of width
             Set proj = oGeomFactory.Projections3d.CreateByCurve(pOC.ResourceManager, oCircle, 0, 0, -1, GroutHeight, False)
             proj.Transform Matrix
             pOC.AddOutput "Projection", proj
             
             'Top plane
             oCircle.SetCenterPoint 0, 0, 0
             Set plane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
             plane.Transform Matrix
             pOC.AddOutput "Grout Top plane", plane
                          
             'Bottom plane
             oCircle.SetCenterPoint 0, 0, -GroutHeight
             oCircle.SetNormal 0, 0, -1 'make normal outwards
             Set plane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
             plane.Transform Matrix
             pOC.AddOutput "Grout Bottom plane", plane
             
         ElseIf GroutShape = 2 Then
            InitRectCurvePoints pts, GroutLength, GroutWidth, 0 'Build points in local XY plane at the centroid of the rectangle
            pIJLineString.SetPoints 5, pts    'Init Points
            Set elems = CreateSolidbyPlanes(pOC.ResourceManager, pIJLineString, GroutHeight)
             
            For i = 1 To elems.Count
                Set plane = elems.Item(i)
                plane.Transform Matrix
                pOC.AddOutput "Grout Rectangle" & i, plane
            Next i
            
         ElseIf GroutShape = 4 Then 'octagon
        
            InitOctCurvePoints octpts, GroutWidth, 0 'Build points in local XY plane at the centroid of the octagon
            
            pIJLineString.SetPoints 9, octpts    'Init Points
            Set elems = CreateSolidbyPlanes(pOC.ResourceManager, pIJLineString, GroutHeight)

            For i = 1 To elems.Count
                Set plane = elems.Item(i)
                plane.Transform Matrix
                pOC.AddOutput "Octagon" & i, plane
            Next i

         End If
     
'        Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, -GroutHeight)
'        pOC.AddOutput "Point2", oPoint3d
     Else
        GroutHeight = 0
     End If
     'Pier Geometry
     Matrix.LoadIdentity
     If PierOrientation = 1 Then
        
        Matrix.Rotate PierRotationAngle, Vec
     ElseIf PierOrientation = 2 Then
        PierRotationAngle = PierRotationAngle + GlobalDelta
        Matrix.Rotate PierRotationAngle, Vec
     End If
     If PierShape = 3 Then
         oCircle.Radius = (PierWidth / 2) ' + (PierEdgeClearance * 5)'TR#72794 - Edge clearance is already considered incalculation of width
         oCircle.SetCenterPoint 0, 0, -GroutHeight
         Set proj = oGeomFactory.Projections3d.CreateByCurve(pOC.ResourceManager, oCircle, 0, 0, -1, PierHeight, False)
         proj.Transform Matrix
         pOC.AddOutput "Pier Rectangle", proj
         
        'Top plane
        oCircle.SetCenterPoint 0, 0, -GroutHeight
        Set plane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        plane.Transform Matrix
        pOC.AddOutput "Pier Top plane", plane
                     
        'Bottom plane
        oCircle.SetCenterPoint 0, 0, -(GroutHeight + PierHeight)
        oCircle.SetNormal 0, 0, -1 'make normal outwards
        Set plane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        plane.Transform Matrix
        pOC.AddOutput "Pier Bottom plane", plane
     
     ElseIf PierShape = 2 Then
        InitRectCurvePoints pts, PierLength, PierWidth, -GroutHeight    'Build points in local XY plane at the centroid of the rectangle
        pIJLineString.SetPoints 5, pts    'Init Points
        Set elems = CreateSolidbyPlanes(pOC.ResourceManager, pIJLineString, PierHeight)
        For i = 1 To elems.Count
            Set plane = elems.Item(i)
            plane.Transform Matrix
            pOC.AddOutput "Pier" & i, plane
        Next i
    End If
     
'     Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, -GroutHeight - PierHeight)
'     pOC.AddOutput "Point3", oPoint3d
     
'     Slab Geometry
     Matrix.LoadIdentity
     If SlabOrientation = 1 Then
        Matrix.Rotate SlabRotationAngle, Vec
     ElseIf SlabOrientation = 2 Then
        SlabRotationAngle = SlabRotationAngle + GlobalDelta
        Matrix.Rotate SlabRotationAngle, Vec
     End If
     
    If SlabShape = 3 Then
        oCircle.Radius = (SlabWidth / 2) '+ (SlabEdgeClearance * 5)
        oCircle.SetCenterPoint 0, 0, -(GroutHeight + PierHeight)
        Set proj = oGeomFactory.Projections3d.CreateByCurve(pOC.ResourceManager, oCircle, 0, 0, -(GroutHeight + PierHeight), SlabHeight, False)
        proj.Transform Matrix
        pOC.AddOutput "Slab Rectangle", proj
        
        'Top plane
        oCircle.SetCenterPoint 0, 0, -(GroutHeight + PierHeight)
        Set plane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        plane.Transform Matrix
        pOC.AddOutput "Slab Top plane", plane
                
        'Bottom plane
        oCircle.SetCenterPoint 0, 0, -(GroutHeight + PierHeight + SlabHeight)
        oCircle.SetNormal 0, 0, -1 'make normal outwards
        Set plane = oGeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        plane.Transform Matrix
        pOC.AddOutput "Slab Bottom plane", plane
        
    ElseIf SlabShape = 2 Then
        InitRectCurvePoints pts, SlabLength, SlabWidth, -GroutHeight - PierHeight 'Build points in local XY plane at the centroid of the rectangle
        pIJLineString.SetPoints 5, pts    'Init Points4
        Set elems = CreateSolidbyPlanes(pOC.ResourceManager, pIJLineString, SlabHeight)
        For i = 1 To elems.Count
            Set plane = elems.Item(i)
            plane.Transform Matrix
            pOC.AddOutput "Slab" & i, plane
        Next i
    End If
'     Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, -GroutHeight - PierHeight - SlabHeight)
'     pOC.AddOutput "Point4", oPoint3d
    
     
Exit Sub
ErrorHandler:
    If Err.Description = "Undefined Value" Then
        Err.Raise E_FAIL
    Else
        HandleError MODULE, METHOD
    End If
End Sub
Public Sub CMCacheSupported(pInput As IJDInputStdCustomMethod, bArgToCache As Boolean, pArgument As Object, ppCache As Object)

 If bArgToCache Then

        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        oPC.Type = igString
        oPC.String = "Supported"
        Set ppCache = oPC
        Set oPC = Nothing
    End If
Exit Sub
End Sub
Public Sub RefGeometry(pIRepSCM As IJDRepresentationStdCustomMethod)
Const METHOD = "RefGeometry"
On Error GoTo ErrorHandler
    Dim pRepDG As IJDRepresentationDuringGame
    Dim pOC As IJDOutputCollection
    Dim oPoint3d As IngrGeom3D.Point3d
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim pInputs As IJDInputs
    Dim oOutputs As IJDOutputs
    Dim oRep As IJDRepresentation
   
    Set pRepDG = pIRepSCM
    Set pInputs = pRepDG.definition.IJDInputs
    Set pOC = pRepDG.outputCollection
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("RefGeometry")
    Set oOutputs = oRep
    oOutputs.RemoveAllOutput
    
    Dim withGroutPad As Boolean
    Dim GroutHeight As Double, PierHeight As Double, SlabHeight As Double
    
    withGroutPad = pInputs.GetInputByIndex(WITHGROUTPAD_INDEX).IJDInputDuringGame.Result.UomValue
    GroutHeight = pInputs.GetInputByIndex(GROUTHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    PierHeight = pInputs.GetInputByIndex(PIERHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    SlabHeight = pInputs.GetInputByIndex(SLABHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
         
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, 0)
    pOC.AddOutput "Point1", oPoint3d
    
    If withGroutPad Then
        Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, -GroutHeight)
        pOC.AddOutput "Point2", oPoint3d
    Else
        GroutHeight = 0
    End If
    
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, -GroutHeight - PierHeight)
    pOC.AddOutput "Point3", oPoint3d
     
    Set oPoint3d = oGeomFactory.Points3d.CreateByPoint(pOC.ResourceManager, _
                   0, 0, -GroutHeight - PierHeight - SlabHeight)
    pOC.AddOutput "Point4", oPoint3d
    
    
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IJSymbolVersion_GetSymbolVersion() As String
    IJSymbolVersion_GetSymbolVersion = "10.1.0.0"
End Function
