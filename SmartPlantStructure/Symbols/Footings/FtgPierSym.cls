VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FtgPierSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    FtgPierSym.cls
'
'Author
'       28th January 2003        Aniket Patil
'
'Description
'
'Notes
'
'History:

'*******************************************************************
Option Explicit
Private Const MODULE = "FtgPierSym"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"
Const m_ItemProgId As String = "SPSFootingMacros.FtgPierSym"
Implements IJDUserSymbolServices
Implements IJSymbolVersion
Private Enum InputIndex
SUPPORTIONGPLANE_INDEX = 1
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
End Enum

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
     IJDUserSymbolServices_GetDefinitionName = m_ItemProgId
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
     
     Dim mthDesc As IMSSymbolEntities.IJDMethodDescription
     Set mthDesc = New DMethodDescription
     
     
     
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
    Dim NegativeNumCheck As Long 'check for negative number entry
    
     
    ChecklibDesc.name = "CMCheckLib"
    ChecklibDesc.Type = imsLIBRARY_IS_ACTIVEX
    ChecklibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    ChecklibDesc.Source = CheckProgId
    pSymbolDefinition.IJDUserMethods.SetLibrary ChecklibDesc
    ChecklibCookie = ChecklibDesc.Cookie
    
    GTZeroCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("GTZero", ChecklibCookie)
    NegativeNumCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("NegativeNum", ChecklibCookie)
     
     ' Get the lib/method cookie
     libCookie = libDesc.Cookie
      
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
 
     oInput.name = SUPPORTING_PLANE
     oInput.Description = "Supporting"
     oInput.Properties = igDESCRIPTION_OPTIONAL
     iInputs.SetInput oInput, SUPPORTIONGPLANE_INDEX
     oInput.Reset
     
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
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
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
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 18
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERLENGTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_WIDTH
     oInput.Description = PIER_WIDTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 18
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERWIDTH_INDEX
     oInput.Reset
     PC.Reset

     oInput.name = PIER_HEIGHT
     oInput.Description = PIER_HEIGHT
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 24
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, PIERHEIGHT_INDEX
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
     
     output.name = "Rectangle"
     output.Description = "Rectangle Solid"
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
   
     pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_SHARED
     
    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
     
       
Exit Sub
ErrorHandler:    HandleError MODULE, METHOD
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
ErrorHandler:     HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub
Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)
Const METHOD = "Physical"

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
     
    'Get the Inputs
    Dim PierLength As Double, PierWidth As Double, PierHeight As Double
    Dim PierShape As Long, PierSizingRule As Long, PierOrientation As Long
    Dim PierRotationAngle As Double, PierEdgeClearance As Double, PierSizeIncrement As Double
    Dim PierChamfered As Boolean, PierChamferSize As Double
    
    PierShape = pInputs.GetInputByIndex(PIERSHAPE_INDEX).IJDInputDuringGame.Result.UomValue
    CheckForUndefinedValueAndRaiseError pRepDG, PierShape, PRISMATIC_FOOTING_SHAPES, 124
    
    PierLength = pInputs.GetInputByIndex(PIERLENGTH_INDEX).IJDInputDuringGame.Result.UomValue
    PierWidth = pInputs.GetInputByIndex(PIERWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    PierHeight = pInputs.GetInputByIndex(PIERHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    PierEdgeClearance = pInputs.GetInputByIndex(PIEREDGECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    
    Dim pSuppPlane As IJPlane
    Dim IsPlane As Boolean
    IsPlane = False
    Dim oClipSurfWrapper As New SP3DSPSClipGeomService.SPSClipSurfWrapper
    Dim oinverseMatrix As IJDT4x4
    Set oinverseMatrix = New DT4x4
    Dim oMatrix As IJDT4x4
    
    Dim GeomFactory As New IngrGeom3D.GeometryFactory
    
    Dim plane As IngrGeom3D.Plane3d
    Dim pClipSurfs As IJElements
    Dim iClipSurfCount As Integer, ii As Integer
    Dim pSurface As ijdObject
    If PierShape = 3 Then
        Dim proj As IJProjection
        Dim oCircle As New Circle3d
        oCircle.Radius = (PierWidth / 2) '+ (PierEdgeClearance * 5) 'TR#72794 - Edge clearance is already considered incalculation of width
        
        Set proj = GeomFactory.Projections3d.CreateByCurve(pOC.ResourceManager, oCircle, 0, 0, -1, PierHeight, False)
        pOC.AddOutput "Rectangle", proj
        
        'Top plane
        oCircle.SetCenterPoint 0, 0, 0
        Set plane = GeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        pOC.AddOutput "Top plane", plane
                
        'Bottom plane
        oCircle.SetCenterPoint 0, 0, -(PierHeight)
        oCircle.SetNormal 0, 0, -1 'make normal outwards
        Set plane = GeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        pOC.AddOutput "Bottom plane", plane

     ElseIf PierShape = 2 Then
     
          Dim pts(15) As Double
          Dim pIJLineString As IJLineString
          Set pIJLineString = New LineString3d
          Dim elems As IJElements
          Set elems = New JObjectCollection 'IMSElements.DynElements
          Dim i As Integer
          
   'Pier Geometry
          InitRectCurvePoints pts, PierLength, PierWidth, 0 'Build points in local XY plane at the centroid of the rectangle
          pIJLineString.SetPoints 5, pts    'Init Points
          
         Set elems = CreateSolidbyPlanes(pOC.ResourceManager, pIJLineString, PierHeight)
         For i = 1 To elems.Count
              pOC.AddOutput "Rectangle" & i, elems.Item(i)
         Next i
     End If
     
End Sub
Public Sub CMCacheSupported(pInput As IJDInputStdCustomMethod, bArgToCache As Boolean, pArgument As Object, ppCache As Object)

 If bArgToCache Then

        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        oPC.Type = igString
        oPC.String = "Supporting"
        Set ppCache = oPC
        Set oPC = Nothing
    End If
Exit Sub
End Sub

Private Function IJSymbolVersion_GetSymbolVersion() As String
    IJSymbolVersion_GetSymbolVersion = "10.1.0.0"
End Function
