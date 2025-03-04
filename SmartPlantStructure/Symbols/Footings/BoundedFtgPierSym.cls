VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoundedFtgPierSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    BoundedFtgPierSym.cls
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
'  06-Mar-07    RS & SS     CR#41094 - Changes for placing a footing in space
'                           and dis/reconnect to members
'
'*******************************************************************
Option Explicit
Private Const MODULE = "BoundedFtgPierSym"
Const m_ItemProgId As String = "SPSFootingMacros.BoundedFtgPierSym"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"
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
    
    ' CustomMethodCheck for negative num
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
     oInput.Description = "Supporting Plane"
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
   
     pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
       
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
    Dim pOcc As IJDOccurrence
    Dim oMatrix As IJDT4x4
    Dim pSymbol As IJDSymbol
    Dim pDefPlayerEx As IJDDefinitionPlayerEx
    Dim Arguments As IJDArguments
    Dim pArg1 As IJDArgument
    Set pDefPlayerEx = pRepDG.definition.IJDDefinitionPlayerEx
    Set Arguments = pDefPlayerEx.SymbolInputs
   
    Set pSymbol = pDefPlayerEx.PlayingSymbol
    
    On Error Resume Next
    Set pOcc = pSymbol
    
    Set oMatrix = New DT4x4
    Set oMatrix = pOcc.Matrix
    Dim x As Double, y As Double, z As Double
    Dim Rootx As Double, Rooty As Double, Rootz As Double
    Set pArg1 = Arguments.Item(1)
    Set pSuppPlane = pArg1.Entity
    
    Dim GeomFactory As New IngrGeom3D.GeometryFactory
    
    Dim GeomServices As IGeometryServices
    Set GeomServices = GeomFactory
    If Not pSuppPlane Is Nothing Then
        pSuppPlane.GetRootPoint x, y, z
        pSuppPlane.GetNormal Rootx, Rooty, Rootz
        
        Dim DummyFace As New Plane3d
        Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, Rootx, Rooty, Abs(Rootz)) 'TR 46074

        IsPlane = True
        Set oinverseMatrix = oMatrix.Clone
        oinverseMatrix.Invert
        DummyFace.Transform oinverseMatrix
    End If
    
    Dim pClipSurfs As IJElements
    Dim iClipSurfCount As Integer, ii As Integer
    Dim pSurface As ijdObject
    Dim proj As IJProjection
    Dim oTopPlane As New Plane3d
    Dim oBottomPlane As New Plane3d
    Dim tmpSurf As IJSurface
    Set tmpSurf = DummyFace
    Dim Geomconst As Geom3dIntersectConstants
    Dim projectionHeight As Double
          
    If PierShape = 3 Then
           Dim oCircle As New Circle3d
           Set oCircle = GeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, 0, 0, 0, 1, (PierWidth / 2)) 'TR#72794 - Edge clearance is already considered incalculation of width
           
           If IsPlane Then
                'Getting the distance from the circle which is to be projected to the surface then we are adding the
                'PierHeight that way we can get an intersection between the projected curve and the surface.  This change
                'is part of TR 208333
                projectionHeight = GetDistanceToSurface(tmpSurf, oCircle) + PierHeight
                Set proj = GeomFactory.Projections3d.CreateByCurve(Nothing, oCircle, 0, 0, -1, projectionHeight, True)
                 'Top plane
                 Set oTopPlane = GeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, oCircle)
                 pOC.AddOutput "Top Plane", oTopPlane
                 
                 'Bottom Plane
                 Call tmpSurf.Intersect(proj, pClipSurfs, Geomconst)
                 
                If Not pClipSurfs Is Nothing Then
                    If pClipSurfs.Count = 1 Then
                       Set oBottomPlane = GeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, pClipSurfs.Item(1))
                       GeomServices.SetOutwardNormal pOC.ResourceManager, oBottomPlane, 1, 0, 0, -1
                       pOC.AddOutput "Bottom Plane", oBottomPlane
                    End If
                    Set pClipSurfs = Nothing
                End If
                
                'TR-CP·49367  Equipment Foundations should offer correct surface normals for relationships
                proj.ReverseNormal = True 'Reverse the normal so that the planes we get back has normals outward
                Call oClipSurfWrapper.GetClipSurfaceWithPlanes(pOC.ResourceManager, _
                                                    proj, _
                                                    DummyFace, _
                                                    0.000001, _
                                                    pClipSurfs)
                iClipSurfCount = pClipSurfs.Count
                If iClipSurfCount >= 1 Then
                    For ii = 1 To iClipSurfCount
                        Set pSurface = pClipSurfs(ii)
                        pOC.AddOutput "Pier" & ii, pSurface
                        Set pSurface = Nothing
                    Next ii
                End If
            Else
                Set proj = GeomFactory.Projections3d.CreateByCurve(pOC.ResourceManager, oCircle, 0, 0, -1, PierHeight, True)
                pOC.AddOutput "Rectangle", proj
           End If
           

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
          
          If IsPlane Then
                 projectionHeight = GetDistanceToSurface(tmpSurf, pIJLineString) + PierHeight
                 Set proj = GeomFactory.Projections3d.CreateByCurve(Nothing, pIJLineString, 0, 0, -1, projectionHeight, True)
                 'Top plane
                 Set oTopPlane = GeomFactory.Planes3d.CreateByPoints(pOC.ResourceManager, 5, pts)
                 pOC.AddOutput "Top Plane", oTopPlane
                 'Bottom Plane
                 Call tmpSurf.Intersect(proj, pClipSurfs, Geomconst)
                 
                 
                 If pClipSurfs.Count = 1 Then
                    Set oBottomPlane = GeomFactory.Planes3d.CreateByOuterBdry(pOC.ResourceManager, pClipSurfs.Item(1))
                    GeomServices.SetOutwardNormal pOC.ResourceManager, oBottomPlane, 1, 0, 0, -1
                    pOC.AddOutput "Bottom Plane", oBottomPlane
                 End If
                 Set pClipSurfs = Nothing
                 
                 proj.ReverseNormal = True 'Reverse the normal so that the planes we get back has normals outward
                 Call oClipSurfWrapper.GetClipSurfaceWithPlanes(pOC.ResourceManager, _
                                                    proj, _
                                                    DummyFace, _
                                                    0.000001, _
                                                    pClipSurfs)
                    iClipSurfCount = pClipSurfs.Count
                    If iClipSurfCount >= 1 Then
                        For ii = 1 To iClipSurfCount
                            Set pSurface = pClipSurfs(ii)
                            pOC.AddOutput "Pier" & ii, pSurface
                            Set pSurface = Nothing
                        Next ii
                    End If
            Else
                Set elems = CreateSolidbyPlanes(pOC.ResourceManager, pIJLineString, PierHeight)
                For i = 1 To elems.Count
                     pOC.AddOutput "Rectangle" & i, elems.Item(i)
                Next i
                
          End If
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
'######################################################################
'This method returns the distance between the surface and the source
'obeject.
'######################################################################
Private Function GetDistanceToSurface(surface As IJSurface, sourceObject As Object) As Double

    Dim minDistance As Double, srcX As Double, Srcy As Double, Srcz As Double, inx As Double, iny As Double, inz As Double
    Dim numOfPoints As Long
    Dim points1() As Double, points2() As Double, pars1() As Double, pars2() As Double
    
    surface.DistanceBetween sourceObject, minDistance, srcX, Srcy, Srcz, inx, iny, inz, numOfPoints, points1, points2, pars1, pars2
    
    GetDistanceToSurface = minDistance

End Function
