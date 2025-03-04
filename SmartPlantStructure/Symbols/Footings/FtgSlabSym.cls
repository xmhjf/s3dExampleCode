VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FtgSlabSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    FtgSlabSym.cls
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
Private Const MODULE = "FtgSlabSym"
Const m_ItemProgId As String = "SPSFootingMacros.FtgSlabSym"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"
Implements IJDUserSymbolServices
Implements IJSymbolVersion

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
    Dim NegativeNumCheck As Long
    
    
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
'     mCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheForPart", libCookie)
     
     Dim pIJDInput As IMSSymbolEntities.IJDInput
     Set pIJDInput = New IMSSymbolEntities.DInput
     
     Dim oInput As IMSSymbolEntities.IJDInput
     Set oInput = New IMSSymbolEntities.DInput
     Dim PC As IMSSymbolEntities.IJDParameterContent
     Set PC = New IMSSymbolEntities.DParameterContent
     PC.Type = igValue
      
     oInput.name = SLAB_SHAPE
     oInput.Description = SLAB_SHAPE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 1
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_SIZING_RULE
     oInput.Description = SLAB_SIZING_RULE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 2
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_ORIENTATION
     oInput.Description = SLAB_ORIENTATION
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 3
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_ROTATION_ANGLE
     oInput.Description = SLAB_ROTATION_ANGLE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 4
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_EDGE_CLEARANCE
     oInput.Description = SLAB_EDGE_CLEARANCE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 24
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 5
     oInput.Reset
     PC.Reset
 
     oInput.name = SLAB_SIZE_INC
     oInput.Description = SLAB_SIZE_INC
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, NegativeNumCheck
     PC.UomValue = 1
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 6
     oInput.Reset
     PC.Reset
 
     oInput.name = SLAB_MATERIAL
     oInput.Description = SLAB_MATERIAL
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Concrete"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 7
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_GRADE
     oInput.Description = SLAB_GRADE
     oInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Fc 3000"
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 8
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_LENGTH
     oInput.Description = SLAB_LENGTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 5
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 9
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_WIDTH
     oInput.Description = SLAB_WIDTH
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 5
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 10
     oInput.Reset
     PC.Reset

     oInput.name = SLAB_HEIGHT
     oInput.Description = SLAB_HEIGHT
     oInput.Properties = igINPUT_IS_A_PARAMETER
     oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 16
     oInput.DefaultParameterValue = PC
     iInputs.SetInput oInput, 11
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
ErrorHandler: HandleError MODULE, METHOD
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
 
ErrorHandler: HandleError MODULE, METHOD

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
     Dim SlabLength As Double, SlabWidth As Double, SlabHeight As Double
     Dim SlabShape As Long, SlabSizingRule As Long, SlabOrientation As Long
     Dim SlabRotationAngle As Double, SlabEdgeClearance As Double, SlabSizeIncrement As Double
     
     
     SlabLength = pInputs.GetInputByIndex(9).IJDInputDuringGame.Result.UomValue
     SlabWidth = pInputs.GetInputByIndex(10).IJDInputDuringGame.Result.UomValue
     SlabHeight = pInputs.GetInputByIndex(11).IJDInputDuringGame.Result.UomValue
     SlabShape = pInputs.GetInputByIndex(1).IJDInputDuringGame.Result.UomValue
     
     CheckForUndefinedValueAndRaiseError pRepDG, SlabShape, PRISMATIC_FOOTING_SHAPES, 127
     
     Dim plane As IngrGeom3D.Plane3d
      
    If SlabShape = 3 Then
        Dim GeomFactory As New IngrGeom3D.GeometryFactory
        Dim proj As IJProjection
        Dim oCircle As New Circle3d
        oCircle.Radius = (SlabWidth / 2)
        
        Set proj = GeomFactory.Projections3d.CreateByCurve(pOC.ResourceManager, oCircle, 0, 0, -1, SlabHeight, False)
        pOC.AddOutput "Rectangle", proj
           
        'Top plane
        oCircle.SetCenterPoint 0, 0, 0
        Set plane = GeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        pOC.AddOutput "Top plane", plane
                
        'Bottom plane
        oCircle.SetCenterPoint 0, 0, -(SlabHeight)
        oCircle.SetNormal 0, 0, -1 'make normal outwards
        Set plane = GeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
        pOC.AddOutput "Bottom plane", plane


     ElseIf SlabShape = 2 Then
          Dim pts(15) As Double
          Dim Vec As DVector
          Dim pIJLineString As IJLineString
          Set pIJLineString = New LineString3d
          Set Vec = New DVector
          Dim elems As IJElements
          Set elems = New JObjectCollection ' IMSElements.DynElements
          Dim i As Integer
          
          
          'GroutPad Geometry
          InitRectCurvePoints pts, SlabLength, SlabWidth, 0 'Build points in local XY plane at the centroid of the rectangle
          pIJLineString.SetPoints 5, pts    'Init Points
          Vec.Set 0, 0, -1
          Vec.length = SlabHeight
          Set elems = CreateSolidbyPlanes(pOC.ResourceManager, pIJLineString, SlabHeight)
          For i = 1 To elems.Count
               pOC.AddOutput "Rectangle" & i, elems.Item(i)
          Next i
     End If
End Sub

Private Function IJSymbolVersion_GetSymbolVersion() As String
    IJSymbolVersion_GetSymbolVersion = "10.1.0.0"
End Function
