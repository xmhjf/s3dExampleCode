VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AOWithNoSupports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' object to delegate the standard implementation to
'
Dim m_oAdvancedOpeningHelper As Object
'
' definition of the error codes
'
Private Enum Errors
    UNABLE_TO_PROJECT = 1
End Enum
'
' implemented interfaces
'
Implements IJGeometricConstructionDefinitionService
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJGCMirror
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
Implements IJGeometricConstructionDynamic_ReStart
Implements IJGeometricConstructionDynamic_IsEnabled
Implements IJGCGeometricElement
Implements IJGCToDoDelegate
'
' some definitions for debug purpose
'
Dim m_oAdvancedDebug As New AdvancedDebug
Private Property Get Source() As String
    Let Source = "GCAdvancedOpening.AdvancedOpeningImp"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property
'
' constructor/destructor of the class
'
Private Sub Class_Initialize()
    Call m_oAdvancedDebug.StartSource(Source)
    
    Set m_oAdvancedOpeningHelper = CreateObject("GCAdvancedOpeningHelper.AdvancedOpeningHelper")
End Sub
Private Sub Class_Terminate()
    Set m_oAdvancedOpeningHelper = Nothing
    
    Call m_oAdvancedDebug.StopSource(Source)
    Set m_oAdvancedDebug = Nothing
End Sub

Private Sub IJGCGeometricElement_TransformElement(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pMatrix As AutoMath.IJDT4x4)

Dim oRevMgrMultiTransCopy As IJRevMgrMultiTransCopy
Set oRevMgrMultiTransCopy = New REVISIONLib.JRevision
Dim eCmdCtxt As MultiTransCmdContexts
oRevMgrMultiTransCopy.GetCommandContext eCmdCtxt
If eCmdCtxt = CmdContext_CopyByFamily Then
    Dim oPoint As IJPoint
    Set oPoint = New Point3d
    oPoint.SetPoint pGC.Parameter("X"), pGC.Parameter("Y"), pGC.Parameter("Z")
    oPoint.Transform pMatrix

    Dim x As Double
    Dim y As Double
    Dim z As Double
    oPoint.GetPoint x, y, z

    pGC.Parameter("X") = x
    pGC.Parameter("Y") = y
    pGC.Parameter("Z") = z

    Set oPoint = Nothing
End If
Set oRevMgrMultiTransCopy = Nothing
End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGCDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    Call m_oAdvancedOpeningHelper.PreInitialize(pGCDefinition)
    Call Initialize(pGCDefinition)
    Call m_oAdvancedOpeningHelper.PostInitialize(pGCDefinition)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
    
    Dim bEvaluateGeometry As Boolean
    Call m_oAdvancedOpeningHelper.PreEvaluate(pGC, pPOM, bEvaluateGeometry)
    Call m_oAdvancedDebug.EnterMethod(Method("EvaluateGeometry"))
    If bEvaluateGeometry Then
        Call Evaluate(pGC, pPOM)
    Else
        Call m_oAdvancedDebug.ShowMsg("Skip")
    End If
    Call m_oAdvancedDebug.ExitMethod
    Call m_oAdvancedOpeningHelper.PostEvaluate(pGC, pPOM)

    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Dim lErrorNumber As Long: Let lErrorNumber = Err.Number
    Call m_oAdvancedDebug.ExitMethod
    Call m_oAdvancedOpeningHelper.ProcessError(pGC, lErrorNumber)
End Sub
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PostConnectionAdded"))
    
    Call m_oAdvancedOpeningHelper.PostConnectionAdded(oRelationship)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDynamic_IsEnabled_PropertyValue(ByVal Name As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, IsEnabled As Boolean)
    Dim pGCPrivateAccess As IJGCPrivateAccess: Set pGCPrivateAccess = pGC
    Dim pGCType As IJGeometricConstructionType: Set pGCType = pGCPrivateAccess.GeometricConstructionType
    If TypeOf pGCType Is IJGeometricConstructionItem Then
        IsEnabled = True
        Select Case Name
              Case "Length", "Width", "Radius", "RoundedRad":
                 If pGC.Parameter("Custom") = 0 Then IsEnabled = False
              Case "Shape":
                IsEnabled = False
        End Select
   End If
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PreConnectionRemoved"))
    
    Call m_oAdvancedOpeningHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMigrate_Migrate"))

    Call m_oAdvancedOpeningHelper.Migrate(MyGC, pMigrateHelper)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMirror interface
'
Private Sub IJGCMirror_Adapt(ByVal pGCOfOriginalGC As IJGeometricConstruction, ByVal pGCOfMirroredGC As IJGeometricConstruction, _
                             ByVal pPlaneOfMirrorPlane As IJPlane, ByVal pT4x4OfMirrorTransformation As IJDT4x4, _
                             ByVal pElementsOfGCsFromInitalCopySet As IJElements)

    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMirror_Adapt"))
    Dim oRevMgrMultiTransCopy As IJRevMgrMultiTransCopy
    Set oRevMgrMultiTransCopy = New REVISIONLib.JRevision
    Dim eCmdCtxt As MultiTransCmdContexts
    oRevMgrMultiTransCopy.GetCommandContext eCmdCtxt
    If eCmdCtxt <> CmdContext_ModelDataReuse Then

        Dim oPoint As IJPoint
        Set oPoint = New Point3d
        oPoint.SetPoint pGCOfOriginalGC.Parameter("X"), pGCOfOriginalGC.Parameter("Y"), pGCOfOriginalGC.Parameter("Z")
        oPoint.Transform pT4x4OfMirrorTransformation
        
        Dim x As Double
        Dim y As Double
        Dim z As Double
        oPoint.GetPoint x, y, z
        
        pGCOfMirroredGC.Parameter("X") = x
        pGCOfMirroredGC.Parameter("Y") = y
        pGCOfMirroredGC.Parameter("Z") = z
        
        Set oPoint = Nothing
      End If
      
    Call m_oAdvancedOpeningHelper.Adapt(pGCOfOriginalGC, pGCOfMirroredGC, _
                                     pPlaneOfMirrorPlane, pT4x4OfMirrorTransformation, _
                                     pElementsOfGCsFromInitalCopySet, _
                                     "CoordinateSystem", "Location")
     
    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGeometricConstructionDynamic_IsOnRibbonBar interface
'
Public Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As GeometricConstruction, ByVal vInfo As Variant, ByRef bIsOnRibbonBar As Boolean)
    Call m_oAdvancedOpeningHelper.IsOnRibbonBar(sName, pGC, vInfo, bIsOnRibbonBar)
     If pGC.Inputs("BasePlate").Count = 1 Then
        pGC.Input("PlateSystem") = pGC.Input("BasePlate")
    End If
    
      
    Select Case sName
        Case "Length"
            bIsOnRibbonBar = CInt(pGC.Parameter("Shape")) = 1 Or CInt(pGC.Parameter("Shape")) = 3
        Case "Width"
            bIsOnRibbonBar = CInt(pGC.Parameter("Shape")) = 1 Or CInt(pGC.Parameter("Shape")) = 3
        Case "Radius"
            bIsOnRibbonBar = CInt(pGC.Parameter("Shape")) = 2
        Case "RoundedRad"
            bIsOnRibbonBar = CInt(pGC.Parameter("Shape")) = 3
        Case "Custom"
            bIsOnRibbonBar = False
    End Select
     
End Sub
'
' management of error
'
Private Sub GCEvaluate(pGeometricConstructionToCatchForError As IJGeometricConstruction, pGeometricConstructionToPutInError As IJGeometricConstruction, lErrorToCatch As Long, lErrorToRaise As Long)
    Call m_oAdvancedOpeningHelper.GCEvaluate(pGeometricConstructionToCatchForError, pGeometricConstructionToPutInError, lErrorToCatch, lErrorToRaise)
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCToDoDelegate_ToDoDelegate"))

    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    On Error Resume Next
    Set IJGCToDoDelegate_ToDoDelegate = m_oAdvancedOpeningHelper.ToDoDelegate(pGC)
    On Error GoTo 0

    Call m_oAdvancedDebug.ExitMethod
End Property
'
' to be overriden by PrintGraph
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInputGUIinfo("CoordinateSystem", "Select a coordinate system", , "Select a coordinate system", , , "ActiveCoordinateSystem")
   
    Call pGeometricConstructionDefinition.AddInput("BasePlate", "Select a PlateSystem", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInputGUIinfo("BasePlate", "Modify the PlateSystem", , "Modify the PlateSystem", , , "PlateSystemForOpening")
  
    Call pGeometricConstructionDefinition.AddControlledInput("Port_2_1_2_1")
    Call pGeometricConstructionDefinition.AddControlledInput("_Parameters_3_1", "IJGeometricConstruction")
    Call pGeometricConstructionDefinition.AddOutput(7, "Contours")
    Call pGeometricConstructionDefinition.AddParameter("X", "X", 8, 1, 59, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("Y", "Y", 8, 1, 59, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("Z", "Z", 8, 1, 59, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("ProjectionDirection", "ProjDirection", 4, 0, 0, 0, 0, 4)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along x", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along y", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along z", 3)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Normal", 4)
    Call pGeometricConstructionDefinition.AddParameter("Angle", "Enter angle", 8, 2, 78, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("Shape", "Shape", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Shape", "Oval", 1)
       Call pGeometricConstructionDefinition.AddParameterValue("Shape", "Circle", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Shape", "RoundedRect", 3)
    Call pGeometricConstructionDefinition.AddParameter("Length", "Length", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Radius", "Radius", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Width", "Width", 8, 1, 59, 0, 0, 0.8)
    Call pGeometricConstructionDefinition.AddParameter("RoundedRad", "RoundedRad", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Custom", "Custom", GCLong, 0, 0, 0, 0, 1, False)

    Call pGeometricConstructionDefinition.AddErrorValue(Errors.UNABLE_TO_PROJECT, "Not able to project the point on to the surface", "Not able to project the point on to the surface")
           
    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    On Error Resume Next
    Dim oPointFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS1 = oGCFactory.CreateEntity("PointFromCS", pPOM, "0001-PointFromCS")
    oPointFromCS1.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oPointFromCS1.Parameter("X") = pGeometricConstruction.Parameter("X")
    oPointFromCS1.Parameter("Y") = pGeometricConstruction.Parameter("Y")
    oPointFromCS1.Parameter("Z") = pGeometricConstruction.Parameter("Z")
    oPointFromCS1.Evaluate


    Dim oCSByLines4 As SP3DGeometricConstruction.GeometricConstruction
      
    If pGeometricConstruction.Parameter("ProjectionDirection") = 4 Then
    
        Dim oFacePortExtractor02 As SP3DGeometricConstruction.GeometricConstruction
        Set oFacePortExtractor02 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, "0002-FacePortExtractor0")
        oFacePortExtractor02.Inputs("Connectable").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
        oFacePortExtractor02.Parameter("Offset") = 0#
        oFacePortExtractor02.Parameter("GeometrySelector") = 1
        oFacePortExtractor02.Evaluate
            
        
                    
        Dim oVectorNormalToSurface2 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorNormalToSurface2 = oGCFactory.CreateEntity("VectorNormalToSurface", pPOM, "0002-VectorNormalToSurface")
        oVectorNormalToSurface2.Inputs("Surface").Add oFacePortExtractor02, "1"
        oVectorNormalToSurface2.Inputs("Point").Add oPointFromCS1
        oVectorNormalToSurface2.Parameter("Range") = 1#
        oVectorNormalToSurface2.Parameter("Orientation") = 1
        oVectorNormalToSurface2.Parameter("TrackFlag") = 1
        oVectorNormalToSurface2.Evaluate
        
         Dim oPointAtCurveEnd1 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveEnd1 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, "0001-PointAtCurveEnd")
        oPointAtCurveEnd1.Inputs("Curve").Add oVectorNormalToSurface2, "1"
        oPointAtCurveEnd1.Evaluate
        
        Dim oPointByProjectOnSurf1 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByProjectOnSurf1 = oGCFactory.CreateEntity("PointByProjectOnSurf", pPOM, "0001-PointByProjectOnSurf")
        oPointByProjectOnSurf1.Inputs("Point").Add oPointFromCS1, "1"
        oPointByProjectOnSurf1.Inputs("Surface").Add oFacePortExtractor02, "1"
        oPointByProjectOnSurf1.Inputs("Line").Add oVectorNormalToSurface2, "1"
        oPointByProjectOnSurf1.Parameter("TrackFlag") = 1
        oPointByProjectOnSurf1.Evaluate
                    
        Dim oCSByCS1 As SP3DGeometricConstruction.GeometricConstruction
        Set oCSByCS1 = oGCFactory.CreateEntity("CSByCS", pPOM, "0011-CSByCS")
        oCSByCS1.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
        oCSByCS1.Inputs("Origin").Add oPointByProjectOnSurf1, "1"
        oCSByCS1.Parameter("AxesRoles") = 1
        oCSByCS1.Parameter("CSOrientation") = 1
        oCSByCS1.Parameter("TrackFlag") = 1
        oCSByCS1.Evaluate
                    
        Dim oPlaneByPositionsAndCS1 As SP3DGeometricConstruction.GeometricConstruction
        Set oPlaneByPositionsAndCS1 = oGCFactory.CreateEntity("PlaneByPositionsAndCS", pPOM, "0001-PlaneByPositionsAndCS")
        oPlaneByPositionsAndCS1.Inputs("CoordinateSystem").Add oCSByCS1, "1"
        oPlaneByPositionsAndCS1.Inputs("1st Point").Add oPointByProjectOnSurf1
        oPlaneByPositionsAndCS1.Inputs("2nd Point").Add oPointAtCurveEnd1, "1"
        oPlaneByPositionsAndCS1.Parameter("Range") = 10#
        oPlaneByPositionsAndCS1.Parameter("X1") = 0#
        oPlaneByPositionsAndCS1.Parameter("Y1") = 0#
        oPlaneByPositionsAndCS1.Parameter("Z1") = 0#
        oPlaneByPositionsAndCS1.Parameter("X2") = 1#
        oPlaneByPositionsAndCS1.Parameter("Y2") = 0#
        oPlaneByPositionsAndCS1.Parameter("Z2") = 0#
        oPlaneByPositionsAndCS1.Parameter("X3") = 0#
        oPlaneByPositionsAndCS1.Parameter("Y3") = 1#
        oPlaneByPositionsAndCS1.Parameter("Z3") = 0#
        oPlaneByPositionsAndCS1.Evaluate
    
        If Err.Number <> 0 Then
            Err.Clear
            oPlaneByPositionsAndCS1.Inputs("1st Point").Clear
            oPlaneByPositionsAndCS1.Inputs("2nd Point").Clear
            oPlaneByPositionsAndCS1.Evaluate
        End If
       
    
        Dim oCurveByIntersection3 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByIntersection3 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0003-CurveByIntersection")
        oCurveByIntersection3.Inputs("Surface1").Add oFacePortExtractor02, "1"
        oCurveByIntersection3.Inputs("Surface2").Add oPlaneByPositionsAndCS1, "1"
        oCurveByIntersection3.Inputs("CoordinateSystem").Add oCSByCS1, "1"
        oCurveByIntersection3.Inputs("TrackPoint").Add oPointByProjectOnSurf1
        oCurveByIntersection3.Parameter("TrackFlag") = 1
        oCurveByIntersection3.Parameter("ConstructionSurface") = 0
        oCurveByIntersection3.Evaluate
    
        
        If Not TypeOf oCurveByIntersection3.Output Is Line3d Then
             Dim oPlaneByPositionsAndCS2 As SP3DGeometricConstruction.GeometricConstruction
            Set oPlaneByPositionsAndCS2 = oGCFactory.CreateEntity("PlaneByPositionsAndCS", pPOM, "0002-PlaneByPositionsAndCS")
            oPlaneByPositionsAndCS2.Inputs("CoordinateSystem").Add oCSByCS1, "1"
            oPlaneByPositionsAndCS2.Inputs("1st Point").Add oPointByProjectOnSurf1
            oPlaneByPositionsAndCS2.Inputs("2nd Point").Add oPointAtCurveEnd1, "1"
            oPlaneByPositionsAndCS2.Parameter("Range") = 0.1
            oPlaneByPositionsAndCS2.Parameter("X1") = 0#
            oPlaneByPositionsAndCS2.Parameter("Y1") = 0#
            oPlaneByPositionsAndCS2.Parameter("Z1") = 0#
            oPlaneByPositionsAndCS2.Parameter("X2") = 1#
            oPlaneByPositionsAndCS2.Parameter("Y2") = 0#
            oPlaneByPositionsAndCS2.Parameter("Z2") = 0#
            oPlaneByPositionsAndCS2.Parameter("X3") = 0#
            oPlaneByPositionsAndCS2.Parameter("Y3") = 1#
            oPlaneByPositionsAndCS2.Parameter("Z3") = 0#
            oPlaneByPositionsAndCS2.Evaluate
        
            If Err.Number <> 0 Then
                Err.Clear
                oPlaneByPositionsAndCS2.Inputs("1st Point").Clear
                oPlaneByPositionsAndCS2.Inputs("2nd Point").Clear
                oPlaneByPositionsAndCS2.Evaluate
            End If
            
            Dim oCurveByIntersection4 As SP3DGeometricConstruction.GeometricConstruction
            Set oCurveByIntersection4 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, "0004-CurveByIntersection")
            oCurveByIntersection4.Inputs("Surface1").Add oFacePortExtractor02, "1"
            oCurveByIntersection4.Inputs("Surface2").Add oPlaneByPositionsAndCS2, "1"
            oCurveByIntersection4.Inputs("CoordinateSystem").Add oCSByCS1, "1"
            oCurveByIntersection4.Inputs("TrackPoint").Add oPointByProjectOnSurf1
            oCurveByIntersection4.Parameter("TrackFlag") = 1
            oCurveByIntersection4.Parameter("ConstructionSurface") = 0
            oCurveByIntersection4.Evaluate
            
             Set oCSByLines4 = oGCFactory.CreateEntity("CSByLines", pPOM, "0004-CSByLines")
            oCSByLines4.Inputs("AxisLine1").Add oVectorNormalToSurface2, "1"
            oCSByLines4.Inputs("AxisLine2").Add oCurveByIntersection4, "1"
            oCSByLines4.Parameter("AxesRoles") = 5
            oCSByLines4.Parameter("CSOrientation") = 1
            oCSByLines4.Parameter("TrackFlag") = 1
            oCSByLines4.Evaluate
        Else
            Set oCSByLines4 = oGCFactory.CreateEntity("CSByLines", pPOM, "0004-CSByLines")
            oCSByLines4.Inputs("AxisLine1").Add oVectorNormalToSurface2, "1"
            oCSByLines4.Inputs("AxisLine2").Add oCurveByIntersection3, "1"
            oCSByLines4.Parameter("AxesRoles") = 5
            oCSByLines4.Parameter("CSOrientation") = 1
            oCSByLines4.Parameter("TrackFlag") = 1
            oCSByLines4.Evaluate
        End If
      
       
    End If
    
    


    Dim oCSAndPointByProj2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSAndPointByProj2 = oGCFactory.CreateEntity("CSAndPointByProj", pPOM, "0002-CSAndPointByProj")
    oCSAndPointByProj2.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oCSAndPointByProj2.Inputs("BasePlate").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCSAndPointByProj2.Inputs("RefPosition").Add oPointFromCS1, "1"
    If pGeometricConstruction.Parameter("ProjectionDirection") = 4 Then
        oCSAndPointByProj2.Inputs("LocalCoordinateSystem").Add oCSByLines4
    End If
    oCSAndPointByProj2.Parameter("ProjectionDirection") = pGeometricConstruction.Parameter("ProjectionDirection")
    oCSAndPointByProj2.Parameter("Angle") = pGeometricConstruction.Parameter("Angle")
    oCSAndPointByProj2.Evaluate
    
    If Err.Number <> 0 Then
        MsgBox "The given point cannot be projected on to the surface, change the projection direction and try again"
        Err.Raise Errors.UNABLE_TO_PROJECT
        Exit Sub
    End If
   
'    pGeometricConstruction.ControlledInputs("Port_2_1_2_1").Clear
'    pGeometricConstruction.ControlledInputs("Port_2_1_2_1").Add oCSAndPointByProj2.ControlledInputs("Port_2_1").Item("1")
    Dim oOpeningShapes3 As SP3DGeometricConstruction.GeometricConstruction
    Set oOpeningShapes3 = oGCFactory.CreateEntity("OpeningShapes", pPOM, "0003-OpeningShapes")
    oOpeningShapes3.Inputs("Origin").Add oCSAndPointByProj2.Output("ProjectedPoint", 1), "1"
    oOpeningShapes3.Inputs("CoordinateSystem").Add oCSAndPointByProj2.Output("Co-ordinateSystem", 1), "1"
    oOpeningShapes3.Parameter("Shape") = pGeometricConstruction.Parameter("Shape")
    oOpeningShapes3.Parameter("Length") = pGeometricConstruction.Parameter("Length")
    oOpeningShapes3.Parameter("Radius") = pGeometricConstruction.Parameter("Radius")
    oOpeningShapes3.Parameter("Width") = pGeometricConstruction.Parameter("Width")
    oOpeningShapes3.Parameter("RoundedRad") = pGeometricConstruction.Parameter("RoundedRad")
    oOpeningShapes3.Evaluate
    

'    pGeometricConstruction.ControlledInputs("_Parameters_3_1").Clear
'    pGeometricConstruction.ControlledInputs("_Parameters_3_1").Add oOpeningShapes3.ControlledInputs("_Parameters").Item("Length")
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Contours", 1) = oOpeningShapes3.Output("Curve", 1)
End Sub

Private Sub IJGeometricConstructionDynamic_ReStart_PropertyValue(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal firstStart As Boolean, ActiveInput4ReStart As String)
    
    
    If pGC.Inputs("CoordinateSystem").Count = 0 Then
        ActiveInput4ReStart = "CoordinateSystem"
    ElseIf pGC.Inputs("BasePlate").Count = 0 Then
        ActiveInput4ReStart = "BasePlate"
    End If

     If pGC.Inputs("BasePlate").Count <> 0 And firstStart = True Then
        Dim oSurfaceBody As IJSurfaceBody
        Set oSurfaceBody = pGC.Inputs("BasePlate").Item(1)
        Dim oCOG As IJDPosition
        oSurfaceBody.GetCenterOfGravity oCOG
        pGC.Parameter("X") = oCOG.x
        pGC.Parameter("Y") = oCOG.y
        pGC.Parameter("Z") = oCOG.z
        Set oSurfaceBody = Nothing
    End If

End Sub


