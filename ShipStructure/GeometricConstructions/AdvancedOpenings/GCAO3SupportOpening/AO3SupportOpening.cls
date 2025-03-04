VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AO3SupportOpening"
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
    EVALUATE_FAILED = 1
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
Implements IJGeometricConstructionDynamic_IsReadyForPreview
Implements IJGeometricConstructionStatic_ParamGraphicSelection
Implements IJGCToDoDelegate
Implements IJGCEnumValues
'
' 20-mar-2019 : PBL : HF-CP-343654 - Split PlateSystem should not recompute position of GCPoints
'   -analyze assoc flags to skip the recompute
'

'
' some definitions for debug purpose
'
Dim m_oAdvancedDebug As New AdvancedDebug
Private Property Get Source() As String
    Let Source = "GCAO3SupportOpening.AO3SupportOpening"
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
'
' implementation of the IJGCEnumValues interface
'
Private Function IJGCEnumValues_SkipEnumeration(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Boolean
    ' prepare result
    Dim bSkipEnumeration As Boolean
    
    Dim bEvaluateGeometry As Boolean: bEvaluateGeometry = False
    Call m_oAdvancedOpeningHelper.PreEvaluate(pGC, Nothing, bEvaluateGeometry)
    bSkipEnumeration = Not bEvaluateGeometry
    'MsgBox "SkipEnumeration= " + CStr(bSkipEnumeration)
    
    ' return result
    IJGCEnumValues_SkipEnumeration = bSkipEnumeration
End Function
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
    
    Dim bEvaluateGeometry As Boolean: bEvaluateGeometry = True
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

    Call m_oAdvancedOpeningHelper.Adapt(pGCOfOriginalGC, pGCOfMirroredGC, _
                                     pPlaneOfMirrorPlane, pT4x4OfMirrorTransformation, _
                                     pElementsOfGCsFromInitalCopySet, _
                                     "CoordinateSystem", "Location")
                                     
                                    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    On Error Resume Next
    Dim x As Double
    Dim y As Double
    Dim z As Double
    Dim oBasePlatePlane As IJPlane
    Set oBasePlatePlane = pGCOfOriginalGC.Inputs("BasePlate").Item(1)
    Dim dDotNorm As Double
    dDotNorm = 0
    If Not oBasePlatePlane Is Nothing Then
        Dim oBasePlateNormVec As IJDVector
        Set oBasePlateNormVec = New DVector
        oBasePlatePlane.GetNormal x, y, z
        oBasePlateNormVec.Set x, y, z
        Dim MirrorPlaneVec As IJDVector
        Set MirrorPlaneVec = New DVector
        
        pPlaneOfMirrorPlane.GetNormal x, y, z
        MirrorPlaneVec.Set x, y, z
        dDotNorm = MirrorPlaneVec.Dot(oBasePlateNormVec)
        Set MirrorPlaneVec = Nothing
        Set oBasePlateNormVec = Nothing
    End If
       
        
    If dDotNorm = 0 Then ' if the miror plane and the base plate are not parallel
    
        Dim oCurveFromSupport1 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveFromSupport1 = oGCFactory.CreateEntity("CurveFromSupport2", Nothing, "0001-CurveFromSupport2")
        oCurveFromSupport1.Inputs("Support").Add pGCOfOriginalGC.Inputs(sSUPPORT1).Item(1), "2"
        oCurveFromSupport1.Inputs("PlateSystem").Add pGCOfOriginalGC.Inputs("BasePlate").Item(1), "1"
        oCurveFromSupport1.Parameter("Ambiguity") = 1
        oCurveFromSupport1.Parameter("Geometry") = 2
        oCurveFromSupport1.Evaluate
    
                
        Dim oPointAtCurveMiddle2 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAtCurveMiddle2 = oGCFactory.CreateEntity("PointAtCurveMiddle", Nothing, "0002-PointAtCurveMiddle")
        oPointAtCurveMiddle2.Inputs("Curve").Add oCurveFromSupport1.Output("Curve", 1), "2"
        oPointAtCurveMiddle2.Evaluate
            
        Dim oVectorTangentToCurve3 As SP3DGeometricConstruction.GeometricConstruction
        Set oVectorTangentToCurve3 = oGCFactory.CreateEntity("VectorTangentToCurve", Nothing, "0003-VectorTangentToCurve")
        oVectorTangentToCurve3.Inputs("Curve").Add oCurveFromSupport1.Output("Curve", 1), "1"
        oVectorTangentToCurve3.Inputs("Point").Add oPointAtCurveMiddle2, "1"
        oVectorTangentToCurve3.Inputs("Surface").Add pGCOfOriginalGC.Inputs("BasePlate").Item(1), "1"
        oVectorTangentToCurve3.Parameter("Range") = 1#
        oVectorTangentToCurve3.Parameter("Orientation") = 1
        oVectorTangentToCurve3.Parameter("TrackFlag") = 1
        oVectorTangentToCurve3.Evaluate
        
        Dim OrigLine As IJLine
        Set OrigLine = oVectorTangentToCurve3
        Dim oOrigVect As IJDVector
        Set oOrigVect = New DVector
        OrigLine.GetDirection x, y, z
        oOrigVect.Set x, y, z
        
        Dim oTransformedLine As IJLine
        Set oTransformedLine = OrigLine
        Dim TransVect As IJDVector
        Set TransVect = New DVector
        
        
        oTransformedLine.Transform pT4x4OfMirrorTransformation
        oTransformedLine.GetDirection x, y, z
        TransVect.Set x, y, z
        If TransVect.Dot(oOrigVect) > 0 Then
            pGCOfMirroredGC.Parameter("DistanceDir") = IIf(pGCOfOriginalGC.Parameter("DistanceDir") = 1, 2, 1)
        End If
    End If
    
    Set OrigLine = Nothing
    Set oTransformedLine = Nothing
    Set oOrigVect = Nothing
    Set TransVect = Nothing
  
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
        Case "Ambiguity"
            Dim dNumAmbig As Integer
            dNumAmbig = 0
            If pGC.Inputs(sSUPPORT2).Count > 0 And pGC.Inputs(sSUPPORT1).Count > 0 And pGC.Inputs(sSUPPORT3).Count > 0 And pGC.Inputs("BasePlate").Count > 0 Then
                GetAmguityInfo pGC, dNumAmbig, Nothing
            End If
            If dNumAmbig < 2 Then bIsOnRibbonBar = False
            
        
    End Select
     
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
    Call pGeometricConstructionDefinition.AddInput(sSUPPORT1, "Select the middle support", "IJPort AND [GCSHPSTRDefinitions.ShipFilters,IsEdgePortOnRootPlateSystem] OR IHFrame OR IJStructPlane3d OR IJSeam OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] or IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, sROOT_SHIP_SYSTEMS_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sSUPPORT1, "Select the middle support", , "Select the middle support", , , "Support1ForOpening")
    Call pGeometricConstructionDefinition.AddInput(sSUPPORT2, "Select the second support", "IJPort AND [GCSHPSTRDefinitions.ShipFilters,IsEdgePortOnRootPlateSystem] OR IHFrame OR IJStructPlane3d OR IJSeam OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] or IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, sROOT_SHIP_SYSTEMS_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sSUPPORT2, "Select the second support", , "Select the second support", , , "Support2ForOpening")
    Call pGeometricConstructionDefinition.AddInput(sSUPPORT3, "Select the third support", "IJPort AND [GCSHPSTRDefinitions.ShipFilters,IsEdgePortOnRootPlateSystem] OR IHFrame OR IJStructPlane3d OR IJSeam OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] or IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, sROOT_SHIP_SYSTEMS_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sSUPPORT3, "Select the third support", , "Select the third support", , , "Support3ForOpening")
    Call pGeometricConstructionDefinition.AddInput("BasePlate", "Select a Plate System)", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInputGUIinfo("BasePlate", "Modify the PlateSystem", , "Modify the PlateSystem", , , "PlateSystemForOpening")
    
    ' obsolete controlled inputs. Kept here for compatibility with already placed openings
    Call pGeometricConstructionDefinition.AddControlledInput("Port1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port2")
    Call pGeometricConstructionDefinition.AddControlledInput("Port3")
    Call pGeometricConstructionDefinition.AddControlledInput("Port4", "IJPlate IJDPlateMoldedConventions IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port5", "IJPlate IJDPlateMoldedConventions IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port6", "IJPlate IJDPlateMoldedConventions IJGeometry")
           
    Call pGeometricConstructionDefinition.AddParameter("DistanceFrom", "DistanceFrom", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("DistanceFrom", "Edge", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("DistanceFrom", "Center", 2)
    
    Call pGeometricConstructionDefinition.AddParameter("Dist", "Distance", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("DimensionType", "DimensionType", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("DimensionType", "Molded", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("DimensionType", "Thickness", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("DimensionType", "Clearance", 3)
    
    Call pGeometricConstructionDefinition.AddParameter("ProjectionDirection", "ProjDirection", 4, 0, 0, 0, 0, 4)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along X", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along Y", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along Z", 3)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Normal", 4)
    
    Call pGeometricConstructionDefinition.AddParameter("Angle", "Angle", 8, 2, 78, 0, 0, 0#)
    Call pGeometricConstructionDefinition.AddParameter("Shape", "Shape", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Shape", "Oval", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Shape", "Circle", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("Shape", "RoundedRect", 3)
    
    Call pGeometricConstructionDefinition.AddParameter("Length", "Length", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("Width", "Width", 8, 1, 59, 0, 0, 0.8)
    Call pGeometricConstructionDefinition.AddParameter("Radius", "Radius", 8, 1, 59, 0, 0, 1#)
    Call pGeometricConstructionDefinition.AddParameter("RoundedRad", "RoundedRad", 8, 1, 59, 0, 0, 0.1)
    Call pGeometricConstructionDefinition.AddParameter("Ambiguity", "Ambiguity", 4, 0, 0, 0, 0, 1, True)
    Call pGeometricConstructionDefinition.AddParameterValue("Ambiguity", "Solution 1", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("Ambiguity", "Solution 2", 2)

    Call pGeometricConstructionDefinition.AddParameter("Custom", "Custom", GCLong, 0, 0, 0, 0, 1, False)
    Call pGeometricConstructionDefinition.AddParameter("DistanceDir", "DistanceDir", 4, 0, 0, 0, 0, 1, True)
    Call pGeometricConstructionDefinition.AddParameterValue("DistanceDir", "Solution 1", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("DistanceDir", "Solution 2", 2)
    Call pGeometricConstructionDefinition.AddParameterGUIinfo("DistanceDir", , , , , True)
    
    ' controlled inputs inherited from the CurveFromSupport, ClosestSurfaceOnPlate and PlaneBtw2Supports components
    Call pGeometricConstructionDefinition.AddControlledInput("Trigger_BasePlate", "")
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_SUPPORT_1, "")
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_SUPPORT_2, "")
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_SUPPORT_3, "")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_BasePlate")
    Call pGeometricConstructionDefinition.AddControlledInput(sPORT_SUPPORT_1)
    Call pGeometricConstructionDefinition.AddControlledInput(sPORT_SUPPORT_2)
    Call pGeometricConstructionDefinition.AddControlledInput(sPORT_SUPPORT_3)
    Call pGeometricConstructionDefinition.AddControlledInput(sGEOMETRY_SUPPORT_1)
    Call pGeometricConstructionDefinition.AddControlledInput(sGEOMETRY_SUPPORT_2)
    Call pGeometricConstructionDefinition.AddControlledInput(sGEOMETRY_SUPPORT_3)
    Call pGeometricConstructionDefinition.AddControlledInput(sROOT_PLATE_SYSTEM_SUPPORT_1, sROOT_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sROOT_PLATE_SYSTEM_SUPPORT_2, sROOT_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sROOT_PLATE_SYSTEM_SUPPORT_3, sROOT_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PLATE_SYSTEM_SUPPORT_1, sLEAF_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PLATE_SYSTEM_SUPPORT_2, sLEAF_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PLATE_SYSTEM_SUPPORT_3, sLEAF_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PROFILE_SYSTEM_SUPPORT_1, sLEAF_STIFFENER_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PROFILE_SYSTEM_SUPPORT_2, sLEAF_STIFFENER_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PROFILE_SYSTEM_SUPPORT_3, sLEAF_STIFFENER_SYSTEM_COMPUTE_IIDS)

    Call pGeometricConstructionDefinition.AddOutput(7, "Contours")
    
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.EVALUATE_FAILED, "Failed to place opening, check if supports are valid", "Failed to place opening, check if supports are valid")


    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim iBodyIndex As Integer: iBodyIndex = 0
    Dim lErrorNumber As Long
    On Error GoTo ErrorHandler:

    Dim ogCMacro As IJGeometricConstructionMacro
    Set ogCMacro = pGeometricConstruction
   
    ' remove obsolete controlled inputs
    Call pGeometricConstruction.ControlledInputs("Port1").Clear
    Call pGeometricConstruction.ControlledInputs("Port2").Clear
    Call pGeometricConstruction.ControlledInputs("Port3").Clear
    Call pGeometricConstruction.ControlledInputs("Port4").Clear
    Call pGeometricConstruction.ControlledInputs("Port5").Clear
    Call pGeometricConstruction.ControlledInputs("Port6").Clear

    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", pPOM, GetFullName(iBodyIndex, "ParamDistConstant"))
    oParamDistConstant1.Parameter("Value") = IIf(pGeometricConstruction.Parameter("DistanceDir") = 1, Abs(pGeometricConstruction.Parameter("Dist")), Abs(pGeometricConstruction.Parameter("Dist")) * -1)
    oParamDistConstant1.Evaluate

    Dim bIsPlanarPlate As Boolean
    bIsPlanarPlate = True
    Dim oBasePlatePlane As IJPlane
    
    On Error Resume Next
    Set oBasePlatePlane = pGeometricConstruction.Inputs("BasePlate").Item(1)
    If Err.Number <> 0 Then
        bIsPlanarPlate = False
        Err.Clear
    Else
        If oBasePlatePlane Is Nothing Then bIsPlanarPlate = False
    End If
    On Error GoTo ErrorHandler
    
    'If bIsPlanarPlate = False Then MsgBox ("Non planar")
    
    Dim oFacePortExtractor02 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor02 = oGCFactory.CreateEntity("FacePortExtractor0", pPOM, GetFullName(iBodyIndex, "FacePortExtractor0"))
    oFacePortExtractor02.Inputs("Connectable").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oFacePortExtractor02.Parameter("Offset") = 0#
    oFacePortExtractor02.Parameter("GeometrySelector") = 4
    oFacePortExtractor02.Evaluate
    
    Dim oPlaneBtw2Supports3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneBtw2Supports3 = oGCFactory.CreateEntity("PlaneBtw2Supports2", pPOM, GetFullName(iBodyIndex, "PlaneBtw2Supports2"))
    oPlaneBtw2Supports3.Inputs("BasePlate").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oPlaneBtw2Supports3.Inputs(sSUPPORT1).Add pGeometricConstruction.Inputs(sSUPPORT2).Item(1), "1"
    oPlaneBtw2Supports3.Inputs(sSUPPORT2).Add pGeometricConstruction.Inputs(sSUPPORT3).Item(1), "1"
    oPlaneBtw2Supports3.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oPlaneBtw2Supports3.Parameter("DimensionType") = pGeometricConstruction.Parameter("DimensionType")
    oPlaneBtw2Supports3.Parameter("Ambiguity") = pGeometricConstruction.Parameter("Ambiguity")
    ' propagate triggers forward
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sTRIGGER_SUPPORT_2, sTRIGGER_SUPPORT_1, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sTRIGGER_SUPPORT_3, sTRIGGER_SUPPORT_2, GCForward)
    
    ' evaluate the GCMacro
    On Error Resume Next
    oPlaneBtw2Supports3.Evaluate
    If Err.Number <> 0 Then
        lErrorNumber = Err.Number
        ' propagate triggers backward
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sTRIGGER_SUPPORT_2, sTRIGGER_SUPPORT_1, GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sTRIGGER_SUPPORT_3, sTRIGGER_SUPPORT_2, GCBackward)
        Err.Raise lErrorNumber
    End If
    On Error GoTo ErrorHandler
    
    ' propagate triggers backward
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sTRIGGER_SUPPORT_2, sTRIGGER_SUPPORT_1, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sTRIGGER_SUPPORT_3, sTRIGGER_SUPPORT_2, GCBackward)
    ' update controlled inputs
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, "Port_BasePlate", "Port_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sPORT_SUPPORT_2, sPORT_SUPPORT_1, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sPORT_SUPPORT_3, sPORT_SUPPORT_2, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sGEOMETRY_SUPPORT_2, sGEOMETRY_SUPPORT_1, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sGEOMETRY_SUPPORT_3, sGEOMETRY_SUPPORT_2, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sROOT_PLATE_SYSTEM_SUPPORT_2, sROOT_PLATE_SYSTEM_SUPPORT_1, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sROOT_PLATE_SYSTEM_SUPPORT_3, sROOT_PLATE_SYSTEM_SUPPORT_2, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sLEAF_PLATE_SYSTEM_SUPPORT_2, sLEAF_PLATE_SYSTEM_SUPPORT_1, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sLEAF_PLATE_SYSTEM_SUPPORT_3, sLEAF_PLATE_SYSTEM_SUPPORT_2, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sLEAF_PROFILE_SYSTEM_SUPPORT_2, sLEAF_PROFILE_SYSTEM_SUPPORT_1, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oPlaneBtw2Supports3, sLEAF_PROFILE_SYSTEM_SUPPORT_3, sLEAF_PROFILE_SYSTEM_SUPPORT_2, GCBackward)
    
    Dim oPointAtCurveMiddle4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle4 = oGCFactory.CreateEntity("PointAtCurveMiddle", pPOM, GetFullName(iBodyIndex, "PointAtCurveMiddle"))
    oPointAtCurveMiddle4.Inputs("Curve").Add oPlaneBtw2Supports3.Output("Sup1ToSup2Line", 1), "1"
    oPointAtCurveMiddle4.Evaluate

    Dim oPointByCurveAndSurface5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface5 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, GetFullName(iBodyIndex, "PointByCurveAndSurface"))
    oPointByCurveAndSurface5.Inputs("Curve1").Add oPlaneBtw2Supports3.Output("Sup1ToSup2Line", 1), "1"
    oPointByCurveAndSurface5.Inputs("Surface2").Add oPlaneBtw2Supports3.Output("Plane", 1), "1"
    oPointByCurveAndSurface5.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface5.Evaluate

    Dim oCurveFromSupport6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport6 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport6.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT1).Item(1), "1"
    oCurveFromSupport6.Inputs("PlateSystem").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport6.Parameter("Ambiguity") = pGeometricConstruction.Parameter("Ambiguity")
    ' propagate triggers forward
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, sTRIGGER_SUPPORT_1, "Trigger_Support", GCForward)

    ' evaluate the GCMacro
    On Error Resume Next
    oCurveFromSupport6.Evaluate
    If Err.Number <> 0 Then
        lErrorNumber = Err.Number
        ' propagate triggers backward
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, sTRIGGER_SUPPORT_1, "Trigger_Support", GCBackward)
        Err.Raise lErrorNumber
    End If
    On Error GoTo ErrorHandler

    ' propagate triggers backward
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, sTRIGGER_SUPPORT_1, "Trigger_Support", GCBackward)
    ' update controlled inputs
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, "Port_BasePlate", "Port_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, sPORT_SUPPORT_1, "Port_Support", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport6, sGEOMETRY_SUPPORT_1, "Geometry_Support", GCBackward)
      
    Dim oPointByCurveAndSurface7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface7 = oGCFactory.CreateEntity("PointByCurveAndSurface", pPOM, GetFullName(iBodyIndex, "PointByCurveAndSurface"))
    oPointByCurveAndSurface7.Inputs("Curve1").Add oCurveFromSupport6.Output("Curve", 1), "1"
    oPointByCurveAndSurface7.Inputs("Surface2").Add oPlaneBtw2Supports3.Output("Plane", 1), "1"
    oPointByCurveAndSurface7.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface7.Evaluate

    Dim oCurveByIntersection8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection8 = oGCFactory.CreateEntity("CurveByIntersection", pPOM, GetFullName(iBodyIndex, "CurveByIntersection"))
    oCurveByIntersection8.Inputs("Surface1").Add oPlaneBtw2Supports3.Output("Plane", 1), "1"
    oCurveByIntersection8.Inputs("Surface2").Add oFacePortExtractor02, "1"
    oCurveByIntersection8.Parameter("TrackFlag") = pGeometricConstruction.Parameter("Ambiguity")
    oCurveByIntersection8.Parameter("ConstructionSurface") = 0
    oCurveByIntersection8.Evaluate
    
    Dim oPointAlongCurve9 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve9 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, GetFullName(iBodyIndex, "PointAlongCurve"))
    oPointAlongCurve9.Inputs("Curve").Add oCurveByIntersection8, "1"
    oPointAlongCurve9.Inputs("Point").Add oPointByCurveAndSurface7, "1"
    oPointAlongCurve9.Parameter("Distance") = oParamDistConstant1
    oPointAlongCurve9.Parameter("TrackFlag") = 1
    oPointAlongCurve9.Evaluate
    
    Dim oPointBySurfaces11 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointBySurfaces11 = oGCFactory.CreateEntity("PointBySurfaces", pPOM, GetFullName(iBodyIndex, "PointBySurfaces"))
    If (pGeometricConstruction.Parameter("DimensionType") <> 1) And (TypeOf pGeometricConstruction.Inputs(sSUPPORT1).Item(1) Is IJPlate Or TypeOf pGeometricConstruction.Inputs(sSUPPORT1).Item(1) Is IJProfile) Then
        Dim oClosestSurfaceOnPlate10 As SP3DGeometricConstruction.GeometricConstruction
        Set oClosestSurfaceOnPlate10 = oGCFactory.CreateEntity("ClosestSurfaceOnPlate2", pPOM, GetFullName(iBodyIndex, "ClosestSurfaceOnPlate2"))
        oClosestSurfaceOnPlate10.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT1).Item(1), "1"
        oClosestSurfaceOnPlate10.Inputs("BasePlate").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
        oClosestSurfaceOnPlate10.Inputs("RefPosition").Add oPointAlongCurve9, "2"
        oClosestSurfaceOnPlate10.Parameter("DimensionType") = pGeometricConstruction.Parameter("DimensionType")
        ' propagate triggers forward
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, sTRIGGER_SUPPORT_1, "Trigger_Support", GCForward)

        ' evaluate the GCMacro
        On Error Resume Next
        oClosestSurfaceOnPlate10.Evaluate
        If Err.Number <> 0 Then
            lErrorNumber = Err.Number
            ' propagate triggers backward
            Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
            Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, sTRIGGER_SUPPORT_1, "Trigger_Support", GCBackward)
            Err.Raise lErrorNumber
        End If
        On Error GoTo ErrorHandler

        ' propagate triggers backward
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, sTRIGGER_SUPPORT_1, "Trigger_Support", GCBackward)
        ' update controlled inputs
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, "Port_BasePlate", "Port_BasePlate", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, sPORT_SUPPORT_1, "Port_Support", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, sROOT_PLATE_SYSTEM_SUPPORT_1, "RootPlateSystem_Support", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, sLEAF_PLATE_SYSTEM_SUPPORT_1, "LeafPlateSystem_Support", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oClosestSurfaceOnPlate10, sLEAF_PROFILE_SYSTEM_SUPPORT_1, "LeafProfileSystem_Support", GCBackward)
           
        oPointBySurfaces11.Inputs("Surface1").Add oFacePortExtractor02, "1"
        If TypeOf pGeometricConstruction.Inputs(sSUPPORT1).Item(1) Is IJProfile Then
            
            Dim oCSByPlane10 As SP3DGeometricConstruction.GeometricConstruction
            Set oCSByPlane10 = oGCFactory.CreateEntity("CSByPlane", pPOM, GetFullName(iBodyIndex, "CSByPlane"))
            oCSByPlane10.Inputs("Plane").Add oClosestSurfaceOnPlate10.Output("Surface", 1), "1"
            oCSByPlane10.Evaluate
        
            Dim oPlaneFromCS11 As SP3DGeometricConstruction.GeometricConstruction
            Set oPlaneFromCS11 = oGCFactory.CreateEntity("PlaneFromCS", pPOM, GetFullName(iBodyIndex, "PlaneFromCS"))
            oPlaneFromCS11.Inputs("CoordinateSystem").Add oCSByPlane10, "2"
            oPlaneFromCS11.Parameter("LookingAxis") = 3
            oPlaneFromCS11.Parameter("Offset") = 0#
            oPlaneFromCS11.Parameter("Range") = 100#
            oPlaneFromCS11.Evaluate
            
            oPointBySurfaces11.Inputs("Surface2").Add oPlaneFromCS11, "1"
        Else
            oPointBySurfaces11.Inputs("Surface2").Add oClosestSurfaceOnPlate10.Output("Surface", 1), "1"
        End If
        oPointBySurfaces11.Inputs("Surface3").Add oPlaneBtw2Supports3.Output("Plane", 1), "1"
        oPointBySurfaces11.Inputs("TrackPoint").Add oPointByCurveAndSurface7, "1"
        oPointBySurfaces11.Parameter("TrackFlag") = 1
        oPointBySurfaces11.Evaluate
    Else
        Set oPointBySurfaces11 = oPointByCurveAndSurface7
    End If

    Dim oPointAlongCurve13 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve13 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, GetFullName(iBodyIndex, "PointAlongCurve"))
    oPointAlongCurve13.Inputs("Curve").Add oCurveByIntersection8, "2"
    oPointAlongCurve13.Inputs("Point").Add oPointBySurfaces11, "1"
    oPointAlongCurve13.Inputs("Surface").Add oFacePortExtractor02, "1"
    oPointAlongCurve13.Parameter("Distance") = oParamDistConstant1
    oPointAlongCurve13.Parameter("TrackFlag") = 1
    oPointAlongCurve13.Evaluate

    On Error Resume Next
    Dim oPointByCurves14 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurves14 = oGCFactory.CreateEntity("PointByCurves", pPOM, GetFullName(iBodyIndex, "PointByCurves"))
    oPointByCurves14.Inputs("Curve1").Add oPlaneBtw2Supports3.Output("Sup1ToSup2Line", 1), "1"
    oPointByCurves14.Inputs("Curve2").Add oCurveByIntersection8, "1"
    oPointByCurves14.Inputs("Surface").Add oFacePortExtractor02, "1"
    oPointByCurves14.Parameter("TrackFlag") = 1
    oPointByCurves14.Evaluate

    'TR267836  Could not place 3S GC openings on curved plates(hull plates and LEP's)
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler

        Dim oCurveByProjection1 As SP3DGeometricConstruction.GeometricConstruction
        Set oCurveByProjection1 = oGCFactory.CreateEntity("CurveByProjection", pPOM, GetFullName(iBodyIndex, "CurveByProjection"))
        oCurveByProjection1.Inputs("Curve").Add oPlaneBtw2Supports3.Output("Sup1ToSup2Line", 1), "1"
        oCurveByProjection1.Inputs("Surface").Add oFacePortExtractor02, "1"
        oCurveByProjection1.Parameter("TrackFlag") = 1
        oCurveByProjection1.Parameter("ConstructionSurface") = 0
        oCurveByProjection1.Evaluate
        
        Dim oPointByCurves21 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves21 = oGCFactory.CreateEntity("PointByCurves", pPOM, GetFullName(iBodyIndex, "PointByCurves"))
        oPointByCurves21.Inputs("Curve1").Add oCurveByProjection1, "1"
        oPointByCurves21.Inputs("Curve2").Add oCurveByIntersection8, "1"
        oPointByCurves21.Inputs("Surface").Add oFacePortExtractor02, "1"
        oPointByCurves21.Parameter("TrackFlag") = 1
        oPointByCurves21.Evaluate
       
        Set oPointByCurves14 = oPointByCurves21
        
    End If
    On Error GoTo ErrorHandler
    
    Dim oPointAtCurveEnd15 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveEnd15 = oGCFactory.CreateEntity("PointAtCurveEnd", pPOM, GetFullName(iBodyIndex, "PointAtCurveEnd"))
    oPointAtCurveEnd15.Inputs("Curve").Add oPlaneBtw2Supports3.Output("Sup1ToSup2Line", 1), "1"
    oPointAtCurveEnd15.Evaluate
    
    Dim oPointAlongCurve16 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAlongCurve16 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, GetFullName(iBodyIndex, "PointAlongCurve"))
    oPointAlongCurve16.Inputs("Curve").Add oCurveByIntersection8, "2"
    oPointAlongCurve16.Inputs("Point").Add oPointByCurves14, "1"
    oPointAlongCurve16.Inputs("Surface").Add oFacePortExtractor02, "1"
    oPointAlongCurve16.Parameter("Distance") = oParamDistConstant1
    oPointAlongCurve16.Parameter("TrackFlag") = 1
    oPointAlongCurve16.Evaluate
        
    Dim oCSByPoints17 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByPoints17 = oGCFactory.CreateEntity("CSByPoints", pPOM, GetFullName(iBodyIndex, "CSByPoints"))
    oCSByPoints17.Inputs("OriginPoint").Add oPointByCurves14, "1"
    oCSByPoints17.Inputs("AxisPoint1").Add oPointAlongCurve16, "2"
    oCSByPoints17.Inputs("AxisPoint2").Add oPointAtCurveEnd15, "1"
    oCSByPoints17.Parameter("AxesRoles") = 1
    oCSByPoints17.Parameter("CSOrientation") = 1
    oCSByPoints17.Evaluate

    Dim oCSByCS16 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSByCS16 = oGCFactory.CreateEntity("CSByCS", pPOM, GetFullName(iBodyIndex, "CSByCS"))
    oCSByCS16.Inputs("CoordinateSystem").Add oCSByPoints17, "1"
    oCSByCS16.Inputs("Origin").Add oPointBySurfaces11, "1"
    oCSByCS16.Parameter("AxesRoles") = 1
    oCSByCS16.Parameter("CSOrientation") = 1
    oCSByCS16.Parameter("TrackFlag") = 1
    oCSByCS16.Evaluate

    Dim bSupport2IsNonPenetrating As Boolean
    Dim bSupport3IsNonPenetrating As Boolean

    bSupport2IsNonPenetrating = CheckIfSupportIsPlateOrChildOfOpeningParent(pGeometricConstruction.Inputs("BasePlate").Item(1), pGeometricConstruction.Inputs(sSUPPORT2).Item(1))
    bSupport3IsNonPenetrating = CheckIfSupportIsPlateOrChildOfOpeningParent(pGeometricConstruction.Inputs("BasePlate").Item(1), pGeometricConstruction.Inputs(sSUPPORT3).Item(1))

    Dim dDistance As Double
    Dim dAng As Double

    dDistance = pGeometricConstruction.Parameter("Dist")
    If bSupport2IsNonPenetrating = False And bSupport3IsNonPenetrating = False Then
        dAng = pGeometricConstruction.Parameter("Angle") * IIf(dDistance > 0.000001, 1, -1)
     Else
         dAng = pGeometricConstruction.Parameter("Angle")
    End If
    Dim oCSAndPointByProj17 As SP3DGeometricConstruction.GeometricConstruction
    Set oCSAndPointByProj17 = oGCFactory.CreateEntity("CSAndPointByProj2", pPOM, GetFullName(iBodyIndex, "CSAndPointByProj2"))
    oCSAndPointByProj17.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oCSAndPointByProj17.Inputs("BasePlate").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCSAndPointByProj17.Inputs("RefPosition").Add oPointAlongCurve13, "1"
    oCSAndPointByProj17.Inputs("LocalCoordinateSystem").Add oCSByCS16, "1"
    oCSAndPointByProj17.Parameter("ProjectionDirection") = pGeometricConstruction.Parameter("ProjectionDirection")
    oCSAndPointByProj17.Parameter("Angle") = dAng
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCSAndPointByProj17, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)

    ' evaluate the GCMacro
    On Error Resume Next
    oCSAndPointByProj17.Evaluate
    If Err.Number <> 0 Then
        lErrorNumber = Err.Number
        ' propagate triggers backward
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCSAndPointByProj17, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
        Err.Raise lErrorNumber
    End If
    On Error GoTo ErrorHandler

    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCSAndPointByProj17, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCSAndPointByProj17, "Port_BasePlate", "Port_BasePlate", GCBackward)

    'If Err.Number <> 0 Then Exit Sub
    Dim ogCMacro1 As IJGeometricConstructionMacro
    Set ogCMacro1 = oCSAndPointByProj17
    If ogCMacro1.Outputs("Co-ordinateSystem").Count = 0 Then Exit Sub

    Dim oOpeningShapes18 As SP3DGeometricConstruction.GeometricConstruction
    Set oOpeningShapes18 = oGCFactory.CreateEntity("OpeningShapes", pPOM, GetFullName(iBodyIndex, "OpeningShapes"))
    oOpeningShapes18.Inputs("Origin").Add oPointAlongCurve13, "1"
    oOpeningShapes18.Inputs("CoordinateSystem").Add oCSAndPointByProj17.Output("Co-ordinateSystem", 1), "1"
    oOpeningShapes18.Parameter("Shape") = pGeometricConstruction.Parameter("Shape")
    oOpeningShapes18.Parameter("Length") = pGeometricConstruction.Parameter("Length")
    oOpeningShapes18.Parameter("Radius") = pGeometricConstruction.Parameter("Radius")
    oOpeningShapes18.Parameter("Width") = pGeometricConstruction.Parameter("Width")
    oOpeningShapes18.Parameter("RoundedRad") = pGeometricConstruction.Parameter("RoundedRad")
    oOpeningShapes18.Evaluate

    'TR-CP-267836  Could not place 3S GC openings on curved plates(hull plates and LEP's)
    'When the openings is place at curved area of a plate, the curve is going through the center of the circle.
    'It has no intersection with the circule curve. Project the opening onto the base surface
    If bIsPlanarPlate = False Then
        Dim oOpeningCurveProj18p As SP3DGeometricConstruction.GeometricConstruction
        If oBasePlatePlane Is Nothing Then
            Set oOpeningCurveProj18p = oGCFactory.CreateEntity("CurveByProjection", pPOM, GetFullName(iBodyIndex, "CurveByProjection"))
            oOpeningCurveProj18p.Inputs("Curve").Add oOpeningShapes18.Output("Curve", 1), "1"
            oOpeningCurveProj18p.Inputs("Surface").Add oFacePortExtractor02
            oOpeningCurveProj18p.Parameter("TrackFlag") = 1
            oOpeningCurveProj18p.Evaluate
        End If
    End If
    
    If pGeometricConstruction.Parameter("DistanceFrom") = 2 Then
        ogCMacro.Output("Contours", 1) = oOpeningShapes18.Output("Curve", 1)
    Else
        Dim oPointByCurves19 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointByCurves19 = oGCFactory.CreateEntity("PointByCurves", pPOM, GetFullName(iBodyIndex, "PointByCurves"))
        If bIsPlanarPlate = False And Not oOpeningCurveProj18p Is Nothing Then
            oPointByCurves19.Inputs("Curve1").Add oOpeningCurveProj18p
        Else
            oPointByCurves19.Inputs("Curve1").Add oOpeningShapes18.Output("Curve", 1), "1"
        End If
        
        oPointByCurves19.Inputs("Curve2").Add oCurveByIntersection8, "1"
        oPointByCurves19.Inputs("TrackPoint").Add oPointBySurfaces11, "1"
        oPointByCurves19.Inputs("Surface").Add oFacePortExtractor02, "1"
        oPointByCurves19.Parameter("TrackFlag") = 1
        oPointByCurves19.Evaluate
    
        Dim oParamDistMeasureParallel20 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistMeasureParallel20 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
        oParamDistMeasureParallel20.Inputs("Graphics").Add oPointByCurveAndSurface7, "1"
        oParamDistMeasureParallel20.Inputs("Graphics").Add oPointByCurves19, "2"
        oParamDistMeasureParallel20.Parameter("Value") = 0.500000000000001
        oParamDistMeasureParallel20.Evaluate
    
        Dim oParamDistFunctionWith2Args22 As SP3DGeometricConstruction.GeometricConstruction
        Set oParamDistFunctionWith2Args22 = oGCFactory.CreateEntity("ParamDistFunctionWith2Args", pPOM, GetFullName(iBodyIndex, "ParamDistFunctionWith2Args"))
           
        If oParamDistMeasureParallel20.Parameter("Value") > 0.000001 Then
            Dim oLineByPoints20 As SP3DGeometricConstruction.GeometricConstruction
            Set oLineByPoints20 = oGCFactory.CreateEntity("LineByPoints", pPOM, GetFullName(iBodyIndex, "LineByPoints"))
            oLineByPoints20.Inputs("StartPoint").Add oPointByCurveAndSurface7, "1"
            oLineByPoints20.Inputs("EndPoint").Add oPointByCurves19, "1"
            oLineByPoints20.Evaluate
            
            Dim oParamDistMeasureParallel21 As SP3DGeometricConstruction.GeometricConstruction
            Set oParamDistMeasureParallel21 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
            oParamDistMeasureParallel21.Inputs("Graphics").Add oPointByCurves19, "1"
            oParamDistMeasureParallel21.Inputs("Graphics").Add oPointAlongCurve13, "2"
            If TypeOf oCurveByIntersection8 Is IJLine Then
                oParamDistMeasureParallel21.Inputs("OrientationLine").Add oCurveByIntersection8, "1"
             End If
            oParamDistMeasureParallel21.Parameter("Value") = 0.500000000000001
            oParamDistMeasureParallel21.Evaluate
        
            If Not TypeOf oCurveByIntersection8 Is IJLine Then
              oParamDistMeasureParallel21.Parameter("Value") = oParamDistMeasureParallel21.Parameter("Value") * oParamDistConstant1.Parameter("Value") / Abs(oParamDistConstant1.Parameter("Value"))
            End If
            
            oParamDistFunctionWith2Args22.Inputs("Arg1").Add oParamDistConstant1, "1"
            oParamDistFunctionWith2Args22.Inputs("Arg2").Add oParamDistMeasureParallel21, "1"
            oParamDistFunctionWith2Args22.Parameter("Value") = 1.5
            oParamDistFunctionWith2Args22.Parameter("Operation") = 1
            oParamDistFunctionWith2Args22.Evaluate
        Else
            oParamDistFunctionWith2Args22.Inputs("Arg1").Add oParamDistConstant1, "1"
            oParamDistFunctionWith2Args22.Inputs("Arg2").Add oParamDistConstant1, "1"
            oParamDistFunctionWith2Args22.Parameter("Value") = 1.5
            oParamDistFunctionWith2Args22.Parameter("Operation") = 1
            oParamDistFunctionWith2Args22.Evaluate
        End If
       
        Dim oPointAlongCurve23 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve23 = oGCFactory.CreateEntity("PointAlongCurve", pPOM, GetFullName(iBodyIndex, "PointAlongCurve"))
        oPointAlongCurve23.Inputs("Curve").Add oCurveByIntersection8, "1"
        oPointAlongCurve23.Inputs("Point").Add oPointBySurfaces11, "1"
        oPointAlongCurve23.Inputs("Surface").Add oFacePortExtractor02, "1"
        oPointAlongCurve23.Parameter("Distance") = oParamDistFunctionWith2Args22
        oPointAlongCurve23.Parameter("TrackFlag") = 1
        oPointAlongCurve23.Evaluate
    
        Dim oOpeningShapes24 As SP3DGeometricConstruction.GeometricConstruction
        Set oOpeningShapes24 = oGCFactory.CreateEntity("OpeningShapes", pPOM, GetFullName(iBodyIndex, "OpeningShapes"))
        oOpeningShapes24.Inputs("Origin").Add oPointAlongCurve23, "1"
        oOpeningShapes24.Inputs("CoordinateSystem").Add oCSAndPointByProj17.Output("Co-ordinateSystem", 1), "1"
        oOpeningShapes24.Parameter("Shape") = pGeometricConstruction.Parameter("Shape")
        oOpeningShapes24.Parameter("Length") = pGeometricConstruction.Parameter("Length")
        oOpeningShapes24.Parameter("Radius") = pGeometricConstruction.Parameter("Radius")
        oOpeningShapes24.Parameter("Width") = pGeometricConstruction.Parameter("Width")
        oOpeningShapes24.Parameter("RoundedRad") = pGeometricConstruction.Parameter("RoundedRad")
        oOpeningShapes24.Evaluate
            
        ogCMacro.Output("Contours", 1) = oOpeningShapes24.Output("Curve", 1)
     End If
     
    ' remove duplicate controlled inputs
    Call GeometricConstruction_RemoveDuplicateInputsOrControlledInputs(pGeometricConstruction, GetArrayOfStrings(sROOT_PLATE_SYSTEM_SUPPORT_1, sROOT_PLATE_SYSTEM_SUPPORT_2, sROOT_PLATE_SYSTEM_SUPPORT_3))
    Call GeometricConstruction_RemoveDuplicateInputsOrControlledInputs(pGeometricConstruction, GetArrayOfStrings(sLEAF_PLATE_SYSTEM_SUPPORT_1, sLEAF_PLATE_SYSTEM_SUPPORT_2, sLEAF_PLATE_SYSTEM_SUPPORT_3))
    Exit Sub
ErrorHandler:
    Err.Raise Errors.EVALUATE_FAILED
End Sub
Private Function CheckIfSupportIsPlateOrChildOfOpeningParent(ByVal oSupportPlate As Object, ByVal oSupport As Object) As Boolean
    CheckIfSupportIsPlateOrChildOfOpeningParent = False
    If Not TypeOf oSupport Is IJProfile Then
        CheckIfSupportIsPlateOrChildOfOpeningParent = True
        Exit Function
    Else
        Dim oDesignChild As IJDesignChild
        Set oDesignChild = oSupport
        If oDesignChild.GetParent Is oSupportPlate Then
            CheckIfSupportIsPlateOrChildOfOpeningParent = True
        Else
            Dim oProfileUtils As IJProfileAttributes
            Set oProfileUtils = New ProfileUtils
            Dim Profile As IJProfile
            Set Profile = oSupport
           On Error Resume Next
            Dim oPoint As AutoMath.IJDPosition
            Set oPoint = oProfileUtils.GetProfileIntersectionPoint(Profile, oSupportPlate)
            If Not oPoint Is Nothing Then
                CheckIfSupportIsPlateOrChildOfOpeningParent = False
            Else
                CheckIfSupportIsPlateOrChildOfOpeningParent = True
            End If
            Set Profile = Nothing
        End If

    End If
End Function
'
'
Private Function CompareVAxisWithGlobalAxis(oCS As Object, oTangentVec As Object, pPOM As IJDPOM) As Boolean

    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim iBodyIndex As Integer: iBodyIndex = 300

    Dim oLineFromCS1 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS1 = oGCFactory.CreateEntity("LineFromCS", pPOM, GetFullName(iBodyIndex, "LineFromCS"))
    oLineFromCS1.Inputs("CoordinateSystem").Add oCS, "1"
    oLineFromCS1.Parameter("LookingAxis") = 1
    oLineFromCS1.Parameter("Length") = 1#
    oLineFromCS1.Parameter("CSOrientation") = 1
    oLineFromCS1.Parameter("LineJustification") = 1
    oLineFromCS1.Evaluate

    Dim oLineFromCS2 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS2 = oGCFactory.CreateEntity("LineFromCS", pPOM, GetFullName(iBodyIndex, "LineFromCS"))
    oLineFromCS2.Inputs("CoordinateSystem").Add oCS, "1"
    oLineFromCS2.Parameter("LookingAxis") = 2
    oLineFromCS2.Parameter("Length") = 1#
    oLineFromCS2.Parameter("CSOrientation") = 1
    oLineFromCS2.Parameter("LineJustification") = 1
    oLineFromCS2.Evaluate

    Dim oLineFromCS3 As SP3DGeometricConstruction.GeometricConstruction
    Set oLineFromCS3 = oGCFactory.CreateEntity("LineFromCS", pPOM, GetFullName(iBodyIndex, "LineFromCS"))
    oLineFromCS3.Inputs("CoordinateSystem").Add oCS, "1"
    oLineFromCS3.Parameter("LookingAxis") = 3
    oLineFromCS3.Parameter("Length") = 1#
    oLineFromCS3.Parameter("CSOrientation") = 1
    oLineFromCS3.Parameter("LineJustification") = 1
    oLineFromCS3.Evaluate

    Dim oPointFromCS4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointFromCS4 = oGCFactory.CreateEntity("PointFromCS", pPOM, GetFullName(iBodyIndex, "PointFromCS"))
    oPointFromCS4.Inputs("CoordinateSystem").Add oCS, "1"
    oPointFromCS4.Parameter("X") = 0.05
    oPointFromCS4.Parameter("Y") = 0.025
    oPointFromCS4.Parameter("Z") = 0.025
    oPointFromCS4.Evaluate

    Dim oVectorX As IJDVector
    Dim oVectorY As IJDVector
    Dim oVectorZ As IJDVector

    Set oVectorX = Vector_FromLine(oLineFromCS1)
    Set oVectorY = Vector_FromLine(oLineFromCS2)
    Set oVectorZ = Vector_FromLine(oLineFromCS3)

    Dim oSupportVector As IJDVector
    Set oSupportVector = Vector_FromLine(oTangentVec)

    Dim dXdot As Double
    Dim dYdot As Double
    Dim dZdot As Double

    dXdot = oVectorX.Dot(oSupportVector)
    dYdot = oVectorY.Dot(oSupportVector)
    dZdot = oVectorZ.Dot(oSupportVector)
    Dim dblBool As Double
    If Abs(dXdot) > Abs(dYdot) And Abs(dXdot) > Abs(dZdot) Then dblBool = dXdot / Abs(dXdot)
    If Abs(dYdot) > Abs(dXdot) And Abs(dYdot) > Abs(dZdot) Then dblBool = dYdot / Abs(dYdot)
    If Abs(dZdot) > Abs(dYdot) And Abs(dZdot) > Abs(dXdot) Then dblBool = dZdot / Abs(dZdot)

    CompareVAxisWithGlobalAxis = (dblBool > 0)
End Function
Private Function Vector_FromLine(ByVal pLine As IJLine) As IJDVector
    ' get vector
    Dim u As Double, v As Double, w As Double
    Call pLine.GetDirection(u, v, w)

    ' create result
    Dim pVector As New DVector
    Call pVector.Set(u, v, w)

    ' return result
   Set Vector_FromLine = pVector

End Function
Private Sub IJGeometricConstructionDynamic_ReStart_PropertyValue(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal firstStart As Boolean, ActiveInput4ReStart As String)
    If pGC.Inputs("CoordinateSystem").Count = 0 Then
        ActiveInput4ReStart = "CoordinateSystem"
    ElseIf pGC.Inputs("BasePlate").Count = 0 Then
        ActiveInput4ReStart = "BasePlate"
    Else
        ActiveInput4ReStart = sSUPPORT1
    End If
    
   If pGC.Inputs(sSUPPORT3).Count > 0 Then
  
    pGC.Input(sSUPPORT2) = pGC.Input(sSUPPORT3)
    pGC.Inputs(sSUPPORT3).Clear
   End If
   If firstStart = False Then _
    ActiveInput4ReStart = sSUPPORT3
        
End Sub
Private Sub GetAmguityInfo(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, _
                               dNumAmbig As Integer, pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim iBodyIndex As Integer: iBodyIndex = 100
'    If CheckIfBasePlateIsAmbigous(pGeometricConstruction.Inputs("BasePlate").Item(1)) Then
'        dNumAmbig = 2
'        Exit Sub
'    End If
    dNumAmbig = 0
  Dim oCurveFromSupport1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport1 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport1.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT2).Item(1), "1"
    oCurveFromSupport1.Inputs("PlateSystem").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport1.Parameter("Ambiguity") = 1
    oCurveFromSupport1.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, sTRIGGER_SUPPORT_2, "Trigger_Support", GCForward)
    oCurveFromSupport1.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, sTRIGGER_SUPPORT_2, "Trigger_Support", GCBackward)
    
    Dim oCurveFromSupport2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport2 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport2.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT2).Item(1), "1"
    oCurveFromSupport2.Inputs("PlateSystem").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport2.Parameter("Ambiguity") = 2
    oCurveFromSupport2.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, sTRIGGER_SUPPORT_2, "Trigger_Support", GCForward)
    oCurveFromSupport2.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, sTRIGGER_SUPPORT_2, "Trigger_Support", GCBackward)
    
        
    Dim oParamDistMeasureParallel3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel3 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
    oParamDistMeasureParallel3.Inputs("Graphics").Add oCurveFromSupport1.Output("Curve", 1), "1"
    oParamDistMeasureParallel3.Inputs("Graphics").Add oCurveFromSupport2.Output("Curve", 1), "2"
    oParamDistMeasureParallel3.Parameter("Value") = 2.84631048735548
    oParamDistMeasureParallel3.Evaluate
        
    Dim oCurveFromSupport4 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport4 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport4.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT3).Item(1), "1"
    oCurveFromSupport4.Inputs("PlateSystem").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport4.Parameter("Ambiguity") = 1
    oCurveFromSupport4.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, sTRIGGER_SUPPORT_3, "Trigger_Support", GCForward)
    oCurveFromSupport4.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, sTRIGGER_SUPPORT_3, "Trigger_Support", GCBackward)
    
    Dim oCurveFromSupport5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport5 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport5.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT3).Item(1), "1"
    oCurveFromSupport5.Inputs("PlateSystem").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport5.Parameter("Ambiguity") = 2
    oCurveFromSupport5.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, sTRIGGER_SUPPORT_3, "Trigger_Support", GCForward)
    oCurveFromSupport5.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, sTRIGGER_SUPPORT_3, "Trigger_Support", GCBackward)
      
    Dim oParamDistMeasureParallel6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel6 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
    oParamDistMeasureParallel6.Inputs("Graphics").Add oCurveFromSupport4.Output("Curve", 1), "1"
    oParamDistMeasureParallel6.Inputs("Graphics").Add oCurveFromSupport5.Output("Curve", 1), "2"
    oParamDistMeasureParallel6.Parameter("Value") = 2.84661048765548
    oParamDistMeasureParallel6.Evaluate
    
    
    Dim oCurveFromSupport7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport7 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport7.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT1).Item(1), "1"
    oCurveFromSupport7.Inputs("PlateSystem").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport7.Parameter("Ambiguity") = 1
    oCurveFromSupport7.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, sTRIGGER_SUPPORT_1, "Trigger_Support", GCForward)
    oCurveFromSupport7.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, sTRIGGER_SUPPORT_1, "Trigger_Support", GCBackward)
    
    Dim oCurveFromSupport8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport8 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport8.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT1).Item(1), "1"
    oCurveFromSupport8.Inputs("PlateSystem").Add pGeometricConstruction.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport8.Parameter("Ambiguity") = 2
    oCurveFromSupport8.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, "Trigger_BasePlate", "Trigger_BasePlate", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, sTRIGGER_SUPPORT_1, "Trigger_Support", GCForward)
    oCurveFromSupport8.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, "Trigger_BasePlate", "Trigger_BasePlate", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, sTRIGGER_SUPPORT_1, "Trigger_Support", GCBackward)
    
        
    Dim oParamDistMeasureParallel9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel9 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
    oParamDistMeasureParallel9.Inputs("Graphics").Add oCurveFromSupport7.Output("Curve", 1), "1"
    oParamDistMeasureParallel9.Inputs("Graphics").Add oCurveFromSupport8.Output("Curve", 1), "2"
    oParamDistMeasureParallel9.Parameter("Value") = 2.84631048735548
    oParamDistMeasureParallel9.Evaluate
    
    If (Abs(oParamDistMeasureParallel9.Parameter("Value")) > 0.000001 Or Abs(oParamDistMeasureParallel3.Parameter("Value")) > 0.000001 Or Abs(oParamDistMeasureParallel6.Parameter("Value")) > 0.000001) Then dNumAmbig = 2
    
End Sub
Private Sub IJGeometricConstructionDynamic_IsReadyForPreview_PropertyValue(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, IsReadyForPreview As Boolean)
    Dim ogCMacro As IJGeometricConstructionMacro
    If CheckIfBasePlateIsAmbigous(pGC.Inputs("BasePlate").Item(1)) Then
            If pGC.Parameter("Ambiguity") = 1 Then
                pGC.Parameter("Ambiguity") = 2
            Else
                pGC.Parameter("Ambiguity") = 1
            End If
    End If
    Evaluate pGC, Nothing
    Set ogCMacro = pGC

    If ogCMacro.Outputs("Contours").Count < 1 Then
       MsgBox "Make sure ProjDirection and the DistanceDir are valid and try again"
       IsReadyForPreview = False
       Exit Sub
    End If
    IsReadyForPreview = True
End Sub
Private Function CheckIfBasePlateIsAmbigous(oPlate As Object) As Boolean
    On Error Resume Next
    CheckIfBasePlateIsAmbigous = False
     Dim oStructApplyOperation As IJDStructApplyOperation
     Set oStructApplyOperation = oPlate
     Dim oSolutions As IJElements
     Set oSolutions = New MonikerElements
     oStructApplyOperation.GetAmbiguityList "PlateEntity.PlateBound_AE.1", oSolutions
      
     If oSolutions.Count > 1 Then CheckIfBasePlateIsAmbigous = True
     
     Set oSolutions = Nothing
     Set oStructApplyOperation = Nothing
End Function
Private Sub IJGeometricConstructionStatic_ParamGraphicSelection_AdditionalGraphics(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal graphicList As IJElements)
End Sub
Private Sub IJGeometricConstructionStatic_ParamGraphicSelection_DisplayInfo(ByVal Name As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal Value As Variant, xWorldPosition As Double, yWorldPosition As Double, zWorldPosition As Double)
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim iBodyIndex As Integer: iBodyIndex = 200

    Dim oParamDistConstant1 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistConstant1 = oGCFactory.CreateEntity("ParamDistConstant", Nothing, GetFullName(iBodyIndex, "ParamDistConstant"))
    oParamDistConstant1.Parameter("Value") = pGC.Parameter("Dist")
    oParamDistConstant1.Evaluate

    Dim oFacePortExtractor02 As SP3DGeometricConstruction.GeometricConstruction
    Set oFacePortExtractor02 = oGCFactory.CreateEntity("FacePortExtractor0", Nothing, GetFullName(iBodyIndex, "FacePortExtractor0"))
    oFacePortExtractor02.Inputs("Connectable").Add pGC.Inputs("BasePlate").Item(1), "1"
    oFacePortExtractor02.Parameter("Offset") = 0#
    oFacePortExtractor02.Parameter("GeometrySelector") = 4
    oFacePortExtractor02.Evaluate

    Dim oPlaneBtw2Supports3 As SP3DGeometricConstruction.GeometricConstruction
    Set oPlaneBtw2Supports3 = oGCFactory.CreateEntity("PlaneBtw2Supports2", Nothing, GetFullName(iBodyIndex, "PlaneBtw2Supports2"))
    oPlaneBtw2Supports3.Inputs("BasePlate").Add pGC.Inputs("BasePlate").Item(1), "1"
    oPlaneBtw2Supports3.Inputs(sSUPPORT1).Add pGC.Inputs(sSUPPORT2).Item(1), "1"
    oPlaneBtw2Supports3.Inputs(sSUPPORT2).Add pGC.Inputs(sSUPPORT3).Item(1), "1"
    oPlaneBtw2Supports3.Inputs("CoordinateSystem").Add pGC.Inputs("CoordinateSystem").Item(1), "1"
    oPlaneBtw2Supports3.Parameter("DimensionType") = pGC.Parameter("DimensionType")
    oPlaneBtw2Supports3.Parameter("Ambiguity") = pGC.Parameter("Ambiguity")
    oPlaneBtw2Supports3.Evaluate

    Dim oPointAtCurveMiddle4 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointAtCurveMiddle4 = oGCFactory.CreateEntity("PointAtCurveMiddle", Nothing, GetFullName(iBodyIndex, "PointAtCurveMiddle"))
    oPointAtCurveMiddle4.Inputs("Curve").Add oPlaneBtw2Supports3.Output("Sup1ToSup2Line", 1), "1"
    oPointAtCurveMiddle4.Evaluate

    Dim oPointByCurveAndSurface5 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface5 = oGCFactory.CreateEntity("PointByCurveAndSurface", Nothing, GetFullName(iBodyIndex, "PointByCurveAndSurface"))
    oPointByCurveAndSurface5.Inputs("Curve1").Add oPlaneBtw2Supports3.Output("Sup1ToSup2Line", 1), "1"
    oPointByCurveAndSurface5.Inputs("Surface2").Add oPlaneBtw2Supports3.Output("Plane", 1), "1"
    oPointByCurveAndSurface5.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface5.Evaluate

    Dim oCurveFromSupport6 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport6 = oGCFactory.CreateEntity("CurveFromSupport2", Nothing, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport6.Inputs("Support").Add pGC.Inputs(sSUPPORT1).Item(1), "1"
    oCurveFromSupport6.Inputs("PlateSystem").Add pGC.Inputs("BasePlate").Item(1), "1"
    oCurveFromSupport6.Parameter("Ambiguity") = pGC.Parameter("Ambiguity")
    oCurveFromSupport6.Evaluate
     
    Dim oPointByCurveAndSurface7 As SP3DGeometricConstruction.GeometricConstruction
    Set oPointByCurveAndSurface7 = oGCFactory.CreateEntity("PointByCurveAndSurface", Nothing, GetFullName(iBodyIndex, "PointByCurveAndSurface"))
    oPointByCurveAndSurface7.Inputs("Curve1").Add oCurveFromSupport6.Output("Curve", 1), "1"
    oPointByCurveAndSurface7.Inputs("Surface2").Add oPlaneBtw2Supports3.Output("Plane", 1), "1"
    oPointByCurveAndSurface7.Parameter("TrackFlag") = 1
    oPointByCurveAndSurface7.Evaluate

    Dim oCurveByIntersection8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveByIntersection8 = oGCFactory.CreateEntity("CurveByIntersection", Nothing, GetFullName(iBodyIndex, "CurveByIntersection"))
    oCurveByIntersection8.Inputs("Surface1").Add oPlaneBtw2Supports3.Output("Plane", 1), "1"
    oCurveByIntersection8.Inputs("Surface2").Add oFacePortExtractor02, "1"
    oCurveByIntersection8.Parameter("TrackFlag") = 1
    oCurveByIntersection8.Parameter("ConstructionSurface") = 0
    oCurveByIntersection8.Evaluate

    Dim oPoint As IJPoint
    If Value = 1 Then
         Dim oPointAlongCurve9 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve9 = oGCFactory.CreateEntity("PointAlongCurve", Nothing, GetFullName(iBodyIndex, "PointAlongCurve"))
        oPointAlongCurve9.Inputs("Curve").Add oCurveByIntersection8, "1"
        oPointAlongCurve9.Inputs("Point").Add oPointByCurveAndSurface7, "1"
        oPointAlongCurve9.Parameter("Distance") = oParamDistConstant1
        oPointAlongCurve9.Parameter("TrackFlag") = 1
        oPointAlongCurve9.Evaluate
        
        Set oPoint = oPointAlongCurve9.Output
    Else
         Dim oPointAlongCurve10 As SP3DGeometricConstruction.GeometricConstruction
        Set oPointAlongCurve10 = oGCFactory.CreateEntity("PointAlongCurve", Nothing, GetFullName(iBodyIndex, "PointAlongCurve"))
        oPointAlongCurve10.Inputs("Curve").Add oCurveByIntersection8, "1"
        oPointAlongCurve10.Inputs("Point").Add oPointByCurveAndSurface7, "1"
        oPointAlongCurve10.Parameter("Distance") = oParamDistConstant1.Parameter("Value") * -1
        oPointAlongCurve10.Parameter("TrackFlag") = 1
        oPointAlongCurve10.Evaluate
        
        Set oPoint = oPointAlongCurve10.Output
    End If
    oPoint.GetPoint xWorldPosition, yWorldPosition, zWorldPosition
End Sub
