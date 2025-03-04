VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AO4SupportOpening"
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
    Let Source = "GCAO4SupportOpening.AO4SupportOpening"
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

 

'
' implementation of the IJGeometricConstructionDynamic_IsOnRibbonBar interface
'
Public Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As GeometricConstruction, ByVal vInfo As Variant, ByRef bIsOnRibbonBar As Boolean)
    Call m_oAdvancedOpeningHelper.IsOnRibbonBar(sName, pGC, vInfo, bIsOnRibbonBar)
    
    If pGC.Inputs(sBASE_PLATE).Count = 1 Then
        pGC.Input("PlateSystem") = pGC.Input(sBASE_PLATE)
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
            If pGC.Inputs(sSUPPORT1).Count > 0 And pGC.Inputs(sSUPPORT2).Count > 0 And pGC.Inputs(sSUPPORT3).Count > 0 And pGC.Inputs(sSUPPORT4).Count > 0 And pGC.Inputs(sBASE_PLATE).Count > 0 Then
                GetAmguityInfo pGC, dNumAmbig, Nothing
            End If
            If dNumAmbig < 2 Then bIsOnRibbonBar = False
        
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
    Call pGeometricConstructionDefinition.AddInput(sSUPPORT1, "Select the first support", "IJPort AND [GCSHPSTRDefinitions.ShipFilters,IsEdgePortOnRootPlateSystem] OR IHFrame OR  IJStructPlane3D OR IJSeam OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] or IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, sROOT_SHIP_SYSTEMS_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sSUPPORT1, "Select a First support", , "Select a First support", , , "Support1ForOpening")
    Call pGeometricConstructionDefinition.AddInput(sSUPPORT2, "Select a Second support Opposite to First", "IJPort AND [GCSHPSTRDefinitions.ShipFilters,IsEdgePortOnRootPlateSystem] OR IHFrame OR  IJStructPlane3D OR IJSeam OR IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem] OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem]", 1, 1, sROOT_SHIP_SYSTEMS_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sSUPPORT2, "Select a Second support Opposite to First", , "Select a Second support Opposite to First", , , "Support2ForOpening")
    Call pGeometricConstructionDefinition.AddInput(sSUPPORT3, "Select the third support", "IJPort AND [GCSHPSTRDefinitions.ShipFilters,IsEdgePortOnRootPlateSystem] OR IHFrame OR  IJStructPlane3D OR IJSeam OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] or IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, sROOT_SHIP_SYSTEMS_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sSUPPORT3, "Select a third support", , "Select a third support", , , "Support3ForOpening")
    Call pGeometricConstructionDefinition.AddInput(sSUPPORT4, "Select the fourth support Opposite to third", "IJPort AND [GCSHPSTRDefinitions.ShipFilters,IsEdgePortOnRootPlateSystem] OR IHFrame OR  IJStructPlane3D OR IJSeam OR IJStiffenerSystem AND [GCFilters.Filters,IsRootStiffenerSystem] or IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, sROOT_SHIP_SYSTEMS_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sSUPPORT4, "Select the fourth support Opposite to third", , "Select the fourth support Opposite to third", , , "Support4ForOpening")
    Call pGeometricConstructionDefinition.AddInput(sBASE_PLATE, "Select a PlateSystem", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInputGUIinfo(sBASE_PLATE, "Modify the PlateSystem", , "Modify the PlateSystem", , , "PlateSystemForOpening")
    
    Call pGeometricConstructionDefinition.AddParameter("DimensionType", "DimensionType", 4, 0, 0, 0, 0, 1)
    Call pGeometricConstructionDefinition.AddParameterValue("DimensionType", "Molded", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("DimensionType", "Thickness", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("DimensionType", "Clearance", 3)
    Call pGeometricConstructionDefinition.AddParameter("ProjectionDirection", "ProjectionDirection", 4, 0, 0, 0, 0, 4)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along x", 1)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along y", 2)
    Call pGeometricConstructionDefinition.AddParameterValue("ProjectionDirection", "Along z", 3)
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

    ' controlled inputs inherited from the CenterOf4Supports component
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_BASE_PLATE, "")
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_SUPPORT_1, "")
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_SUPPORT_2, "")
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_SUPPORT_3, "")
    Call pGeometricConstructionDefinition.AddControlledInput(sTRIGGER_SUPPORT_4, "")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_BasePlate")
    Call pGeometricConstructionDefinition.AddControlledInput(sPORT_SUPPORT_1)
    Call pGeometricConstructionDefinition.AddControlledInput(sPORT_SUPPORT_2)
    Call pGeometricConstructionDefinition.AddControlledInput(sPORT_SUPPORT_3)
    Call pGeometricConstructionDefinition.AddControlledInput(sPORT_SUPPORT_4)
    Call pGeometricConstructionDefinition.AddControlledInput(sGEOMETRY_SUPPORT_1)
    Call pGeometricConstructionDefinition.AddControlledInput(sGEOMETRY_SUPPORT_2)
    Call pGeometricConstructionDefinition.AddControlledInput(sGEOMETRY_SUPPORT_3)
    Call pGeometricConstructionDefinition.AddControlledInput(sGEOMETRY_SUPPORT_4)
    Call pGeometricConstructionDefinition.AddControlledInput(sROOT_PLATE_SYSTEM_SUPPORT_1, sROOT_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sROOT_PLATE_SYSTEM_SUPPORT_2, sROOT_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sROOT_PLATE_SYSTEM_SUPPORT_3, sROOT_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sROOT_PLATE_SYSTEM_SUPPORT_4, sROOT_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PLATE_SYSTEM_SUPPORT_1, sLEAF_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PLATE_SYSTEM_SUPPORT_2, sLEAF_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PLATE_SYSTEM_SUPPORT_3, sLEAF_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PLATE_SYSTEM_SUPPORT_4, sLEAF_PLATE_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PROFILE_SYSTEM_SUPPORT_1, sLEAF_STIFFENER_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PROFILE_SYSTEM_SUPPORT_2, sLEAF_STIFFENER_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PROFILE_SYSTEM_SUPPORT_3, sLEAF_STIFFENER_SYSTEM_COMPUTE_IIDS)
    Call pGeometricConstructionDefinition.AddControlledInput(sLEAF_PROFILE_SYSTEM_SUPPORT_4, sLEAF_STIFFENER_SYSTEM_COMPUTE_IIDS)

    ' obsolete controlled inputs. Kept here for compatibility with already placed openings
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1_1_1_1_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1_2_1_1_1_1_1", "IJPlate IJDPlateMoldedConventions IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_2_1_2_1_1_1_1_1")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1_3_1_1_1_1_1", "IJPlate IJDPlateMoldedConventions IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1_2_1_2_1_1_1", "IJPlate IJDPlateMoldedConventions IJGeometry")
    Call pGeometricConstructionDefinition.AddControlledInput("Port_1_1_3_1_2_1_1_1", "IJPlate IJDPlateMoldedConventions IJGeometry")
    
    Call pGeometricConstructionDefinition.AddOutput(7, "Contours")

    pGeometricConstructionDefinition.Behavior = 0
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim iBodyIndex As Integer: iBodyIndex = 0
    Dim lErrorNumber As Long
    On Error GoTo ErrorHandler
    
    ' remove obsolete controlled inputs
    Call pGeometricConstruction.ControlledInputs("Port_1_1_1_1_1_1").Clear
    Call pGeometricConstruction.ControlledInputs("Port_1_1_2_1_1_1_1_1").Clear
    Call pGeometricConstruction.ControlledInputs("Port_2_1_2_1_1_1_1_1").Clear
    Call pGeometricConstruction.ControlledInputs("Port_1_1_3_1_1_1_1_1").Clear
    Call pGeometricConstruction.ControlledInputs("Port_1_1_2_1_2_1_1_1").Clear
    Call pGeometricConstruction.ControlledInputs("Port_1_1_3_1_2_1_1_1").Clear
    
    Dim oCenterOf4Supports1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCenterOf4Supports1 = oGCFactory.CreateEntity("CenterOf4Supports2", pPOM, GetFullName(iBodyIndex, "CenterOf4Supports2"))
    oCenterOf4Supports1.Inputs("CoordinateSystem").Add pGeometricConstruction.Inputs("CoordinateSystem").Item(1), "1"
    oCenterOf4Supports1.Inputs(sBASE_PLATE).Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCenterOf4Supports1.Inputs(sSUPPORT1).Add pGeometricConstruction.Inputs(sSUPPORT1).Item(1), "1"
    oCenterOf4Supports1.Inputs(sSUPPORT2).Add pGeometricConstruction.Inputs(sSUPPORT2).Item(1), "1"
    oCenterOf4Supports1.Inputs(sSUPPORT3).Add pGeometricConstruction.Inputs(sSUPPORT3).Item(1), "1"
    oCenterOf4Supports1.Inputs(sSUPPORT4).Add pGeometricConstruction.Inputs(sSUPPORT4).Item(1), "1"
    oCenterOf4Supports1.Parameter("DimensionType") = pGeometricConstruction.Parameter("DimensionType")
    oCenterOf4Supports1.Parameter("ProjectionDirection") = pGeometricConstruction.Parameter("ProjectionDirection")
    oCenterOf4Supports1.Parameter("Angle") = pGeometricConstruction.Parameter("Angle")
    oCenterOf4Supports1.Parameter("Ambiguity") = pGeometricConstruction.Parameter("Ambiguity")
    ' propagate triggers forward
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_BASE_PLATE, "", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_1, "", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_2, "", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_3, "", GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_4, "", GCForward)

    ' evaluate the GCMacro
    On Error Resume Next
    oCenterOf4Supports1.Evaluate
    If Err.Number <> 0 Then
        lErrorNumber = Err.Number
        ' propagate triggers backward
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_BASE_PLATE, "", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_1, "", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_2, "", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_3, "", GCBackward)
        Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_4, "", GCBackward)
        Err.Raise lErrorNumber
    End If
    On Error GoTo ErrorHandler

    ' propagate triggers backward
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_BASE_PLATE, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_1, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_2, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_3, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sTRIGGER_SUPPORT_4, "", GCBackward)
    ' update controlled inputs
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, "Port_BasePlate", "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sPORT_SUPPORT_1, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sPORT_SUPPORT_2, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sPORT_SUPPORT_3, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sPORT_SUPPORT_4, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sGEOMETRY_SUPPORT_1, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sGEOMETRY_SUPPORT_2, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sGEOMETRY_SUPPORT_3, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sGEOMETRY_SUPPORT_4, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sROOT_PLATE_SYSTEM_SUPPORT_1, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sROOT_PLATE_SYSTEM_SUPPORT_2, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sROOT_PLATE_SYSTEM_SUPPORT_3, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sROOT_PLATE_SYSTEM_SUPPORT_4, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PLATE_SYSTEM_SUPPORT_1, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PLATE_SYSTEM_SUPPORT_2, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PLATE_SYSTEM_SUPPORT_3, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PLATE_SYSTEM_SUPPORT_4, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PROFILE_SYSTEM_SUPPORT_1, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PROFILE_SYSTEM_SUPPORT_2, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PROFILE_SYSTEM_SUPPORT_3, "", GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCenterOf4Supports1, sLEAF_PROFILE_SYSTEM_SUPPORT_4, "", GCBackward)

'    If Err.Number <> 0 Then
'        If CheckIfBasePlateIsAmbigous(pGeometricConstruction.Inputs(sBASE_PLATE).Item(1)) Then
'            MsgBox "Check your ambiguity and make sure projection direction is valid and try again"
'        Else
'             MsgBox "Make sure projection direction is valid and try again"
'        End If
'        Exit Sub
'    End If

    Dim oOpeningShapes2 As SP3DGeometricConstruction.GeometricConstruction
    Set oOpeningShapes2 = oGCFactory.CreateEntity("OpeningShapes", pPOM, GetFullName(iBodyIndex, "OpeningShapes"))
    oOpeningShapes2.Inputs("CoordinateSystem").Add oCenterOf4Supports1.Output("Co-ordinateSystem", 1), "1"
    oOpeningShapes2.Inputs("Origin").Add oCenterOf4Supports1.Output("CenterPoint", 1), "1"
    oOpeningShapes2.Parameter("Shape") = pGeometricConstruction.Parameter("Shape")
    oOpeningShapes2.Parameter("Length") = pGeometricConstruction.Parameter("Length")
    oOpeningShapes2.Parameter("Width") = pGeometricConstruction.Parameter("Width")
    oOpeningShapes2.Parameter("Radius") = pGeometricConstruction.Parameter("Radius")
    oOpeningShapes2.Parameter("RoundedRad") = pGeometricConstruction.Parameter("RoundedRad")
    oOpeningShapes2.Evaluate
    
    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction

    oGCMacro.Output("Contours", 1) = oOpeningShapes2.Output("Curve", 1)

    ' remove duplicate controlled inputs
    Call GeometricConstruction_RemoveDuplicateInputsOrControlledInputs(pGeometricConstruction, GetArrayOfStrings(sROOT_PLATE_SYSTEM_SUPPORT_1, sROOT_PLATE_SYSTEM_SUPPORT_2, sROOT_PLATE_SYSTEM_SUPPORT_3, sROOT_PLATE_SYSTEM_SUPPORT_4))
    Call GeometricConstruction_RemoveDuplicateInputsOrControlledInputs(pGeometricConstruction, GetArrayOfStrings(sLEAF_PLATE_SYSTEM_SUPPORT_1, sLEAF_PLATE_SYSTEM_SUPPORT_2, sLEAF_PLATE_SYSTEM_SUPPORT_3, sLEAF_PLATE_SYSTEM_SUPPORT_4))
    Exit Sub
ErrorHandler:
    Err.Raise Errors.EVALUATE_FAILED
End Sub

Private Sub IJGeometricConstructionDynamic_IsReadyForPreview_PropertyValue(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, IsReadyForPreview As Boolean)

        Dim oGCMacro As IJGeometricConstructionMacro

        If CheckIfBasePlateIsAmbigous(pGC.Inputs(sBASE_PLATE).Item(1)) Then
                If pGC.Parameter("Ambiguity") = 1 Then
                    pGC.Parameter("Ambiguity") = 2
                Else
                    pGC.Parameter("Ambiguity") = 1
                End If
        End If
        Evaluate pGC, Nothing
        Set oGCMacro = pGC

        If oGCMacro.Outputs("Contours").Count < 1 Then
           MsgBox "Make sure projection direction is valid and try again"
           IsReadyForPreview = False
           Exit Sub
        End If
        IsReadyForPreview = True
End Sub

Private Sub IJGeometricConstructionDynamic_ReStart_PropertyValue(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal firstStart As Boolean, ActiveInput4ReStart As String)
    
    
    If pGC.Inputs("CoordinateSystem").Count = 0 Then
        ActiveInput4ReStart = "CoordinateSystem"
    ElseIf pGC.Inputs(sBASE_PLATE).Count = 0 Then
        ActiveInput4ReStart = sBASE_PLATE
    ElseIf pGC.Inputs(sSUPPORT1).Count = 0 Then
       ActiveInput4ReStart = sSUPPORT1
    ElseIf pGC.Inputs(sSUPPORT2).Count = 0 Then
        ActiveInput4ReStart = sSUPPORT2
   ElseIf pGC.Inputs(sSUPPORT3).Count = 0 Then
       ActiveInput4ReStart = sSUPPORT3
    ElseIf pGC.Inputs(sSUPPORT4).Count = 0 Then
       ActiveInput4ReStart = sSUPPORT4
    End If
     
   If pGC.Inputs(sSUPPORT4).Count > 0 Then

    pGC.Input(sSUPPORT3) = pGC.Input(sSUPPORT4)
    pGC.Inputs(sSUPPORT4).Clear
       
   End If
   If firstStart = False Then _
    ActiveInput4ReStart = sSUPPORT4
End Sub
Private Sub GetAmguityInfo(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, _
                               dNumAmbig As Integer, pPOM As IJDPOM)
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory
    Set oGCFactory = New GeometricConstructionEntitiesFactory
    Dim iBodyIndex As Integer: iBodyIndex = 1000
    
'    If CheckIfBasePlateIsAmbigous(pGeometricConstruction.Inputs(sBASE_PLATE).Item(1)) Then
'        dNumAmbig = 2
'        Exit Sub
'    End If
    dNumAmbig = 0
    
    Dim oCurveFromSupport1 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport1 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport1.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT1).Item(1), "1"
    oCurveFromSupport1.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport1.Parameter("Ambiguity") = 1
    oCurveFromSupport1.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, sTRIGGER_SUPPORT_1, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport1.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport1, sTRIGGER_SUPPORT_1, sTRIGGER_SUPPORT, GCBackward)
    
    Dim oCurveFromSupport2 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport2 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport2.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT1).Item(1), "1"
    oCurveFromSupport2.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport2.Parameter("Ambiguity") = 2
    oCurveFromSupport2.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, sTRIGGER_SUPPORT_1, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport2.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport2, sTRIGGER_SUPPORT_1, sTRIGGER_SUPPORT, GCBackward)
    
        
    Dim oParamDistMeasureParallel3 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel3 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
    oParamDistMeasureParallel3.Inputs("Graphics").Add oCurveFromSupport1.Output("Curve", 1), "1"
    oParamDistMeasureParallel3.Inputs("Graphics").Add oCurveFromSupport2.Output("Curve", 1), "2"
    oParamDistMeasureParallel3.Parameter("Value") = 2.84631048735548
    oParamDistMeasureParallel3.Evaluate
        
    Dim oCurveFromSupport4 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport4 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport4.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT2).Item(1), "1"
    oCurveFromSupport4.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport4.Parameter("Ambiguity") = 1
    oCurveFromSupport4.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, sTRIGGER_SUPPORT_2, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport4.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport4, sTRIGGER_SUPPORT_2, sTRIGGER_SUPPORT, GCBackward)
    
    Dim oCurveFromSupport5 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport5 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport5.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT2).Item(1), "1"
    oCurveFromSupport5.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport5.Parameter("Ambiguity") = 2
    oCurveFromSupport5.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, sTRIGGER_SUPPORT_2, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport5.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport5, sTRIGGER_SUPPORT_2, sTRIGGER_SUPPORT, GCBackward)
      
    Dim oParamDistMeasureParallel6 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel6 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
    oParamDistMeasureParallel6.Inputs("Graphics").Add oCurveFromSupport4.Output("Curve", 1), "1"
    oParamDistMeasureParallel6.Inputs("Graphics").Add oCurveFromSupport5.Output("Curve", 1), "2"
    oParamDistMeasureParallel6.Parameter("Value") = 2.84661048765548
    oParamDistMeasureParallel6.Evaluate
    
    Dim oCurveFromSupport7 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport7 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport7.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT3).Item(1), "1"
    oCurveFromSupport7.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport7.Parameter("Ambiguity") = 1
    oCurveFromSupport7.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, sTRIGGER_SUPPORT_3, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport7.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport7, sTRIGGER_SUPPORT_3, sTRIGGER_SUPPORT, GCBackward)
    
    Dim oCurveFromSupport8 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport8 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport8.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT3).Item(1), "1"
    oCurveFromSupport8.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport8.Parameter("Ambiguity") = 2
    oCurveFromSupport8.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, sTRIGGER_SUPPORT_3, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport8.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport8, sTRIGGER_SUPPORT_3, sTRIGGER_SUPPORT, GCBackward)
      
    Dim oParamDistMeasureParallel9 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel9 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
    oParamDistMeasureParallel9.Inputs("Graphics").Add oCurveFromSupport7.Output("Curve", 1), "1"
    oParamDistMeasureParallel9.Inputs("Graphics").Add oCurveFromSupport8.Output("Curve", 1), "2"
    oParamDistMeasureParallel9.Parameter("Value") = 2.84991048795548
    oParamDistMeasureParallel9.Evaluate
    
    Dim oCurveFromSupport10 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport10 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport10.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT4).Item(1), "1"
    oCurveFromSupport10.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport10.Parameter("Ambiguity") = 1
    oCurveFromSupport10.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport10, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport10, sTRIGGER_SUPPORT_4, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport10.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport10, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport10, sTRIGGER_SUPPORT_4, sTRIGGER_SUPPORT, GCBackward)
    
    Dim oCurveFromSupport11 As SP3DGeometricConstruction.GeometricConstruction
    Set oCurveFromSupport11 = oGCFactory.CreateEntity("CurveFromSupport2", pPOM, GetFullName(iBodyIndex, "CurveFromSupport2"))
    oCurveFromSupport11.Inputs("Support").Add pGeometricConstruction.Inputs(sSUPPORT4).Item(1), "1"
    oCurveFromSupport11.Inputs("PlateSystem").Add pGeometricConstruction.Inputs(sBASE_PLATE).Item(1), "1"
    oCurveFromSupport11.Parameter("Ambiguity") = 2
    oCurveFromSupport11.Parameter("Geometry") = 2
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport11, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCForward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport11, sTRIGGER_SUPPORT_4, sTRIGGER_SUPPORT, GCForward)
    oCurveFromSupport11.Evaluate
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport11, sTRIGGER_BASE_PLATE, sTRIGGER_BASE_PLATE, GCBackward)
    Call GeometricConstruction_PropagateControlledInput(pGeometricConstruction, oCurveFromSupport11, sTRIGGER_SUPPORT_4, sTRIGGER_SUPPORT, GCBackward)
      
    Dim oParamDistMeasureParallel12 As SP3DGeometricConstruction.GeometricConstruction
    Set oParamDistMeasureParallel12 = oGCFactory.CreateEntity("ParamDistMeasureParallel", pPOM, GetFullName(iBodyIndex, "ParamDistMeasureParallel"))
    oParamDistMeasureParallel12.Inputs("Graphics").Add oCurveFromSupport10.Output("Curve", 1), "1"
    oParamDistMeasureParallel12.Inputs("Graphics").Add oCurveFromSupport11.Output("Curve", 1), "2"
    oParamDistMeasureParallel12.Parameter("Value") = 2.84121210487126
    oParamDistMeasureParallel12.Evaluate
    
        
    If Abs(oParamDistMeasureParallel3.Parameter("Value") > 0.000001 Or Abs(oParamDistMeasureParallel6.Parameter("Value")) > 0.000001 Or Abs(oParamDistMeasureParallel9.Parameter("Value")) > 0.000001 Or Abs(oParamDistMeasureParallel12.Parameter("Value")) > 0.000001) Then
       dNumAmbig = 2
       
    End If
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
