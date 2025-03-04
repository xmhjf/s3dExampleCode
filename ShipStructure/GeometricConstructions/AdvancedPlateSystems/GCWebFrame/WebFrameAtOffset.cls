VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebFrameAtOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJGeometricConstructionDefinitionService
Implements IJGeometricConstructionDynamic_IsOnRibbonBar
Implements IJGeometricConstructionDynamic_IsOptional
Implements IJGeometricConstructionDynamic_ReStart
Implements IJGCSemanticConnection
Implements IJGCMigrate
Implements IJGCMirror
Implements IJGCToDoDelegate
Implements IJGCConnectedElementSC

Dim m_oGenericPlateHelper As Object
Dim m_oAdvancedDebug As New AdvancedDebug

Const iFIRST_GC = 0
Const iMAX_GCS = 13
Private m_sGCTypes(iFIRST_GC To iMAX_GCS) As String
Private m_pGCsDefinitionsServices(iFIRST_GC To iMAX_GCS) As IJGeometricConstructionDefinitionService
Private m_pGCsDefinitions(iFIRST_GC To iMAX_GCS) As IJGeometricConstructionDefinition
Private m_sSubstitutionsList(iFIRST_GC To iMAX_GCS) As String
Private m_pGCs(iFIRST_GC To iMAX_GCS) As IJGeometricConstruction
Private m_pGCsEntitiesFactory As IJGeometricConstructionEntitiesFactory

Private Const sFIRST As String = "First"
Private Const sLAST As String = "Last"

Private Enum eComponents
    CS_AS_CS = 0
    PLANE_AT_OFFSET_FROM_FCBO = 1
    POINTS_FOR_WEB_FRAME = 2
    SELECT_FREE_EDGE = 3
    CURVE_GENERATOR = 4
    SELECT_POINT_TYPE1 = 5
    POINT_AT_MIN_DIST_EX1 = 6
    POINT_AT_OFF_KEY_PT_EX1 = 7
    SELECT_POINT_TYPE2 = 8
    POINT_AT_MIN_DIST_EX2 = 9
    POINT_AT_OFF_KEY_PT_EX2 = 10
    FREE_EDGE1 = 11
    FREE_EDGE2 = 12
    FREE_EDGE3 = 13
End Enum

Private Enum Errors
    PLANE_AT_OFFSET_FROM_FCBO = 1
    POINTS_FOR_WEB_FRAME_FAILED = 2
    CURVE_GENERATOR_FAILED = 4
    POINT_AT_MIN_DIST_EX1_FAILED = 6
    POINT_AT_OFF_KEY_PT_EX1_FAILED = 7
    POINT_AT_MIN_DIST_EX2_FAILED = 9
    POINT_AT_OFF_KEY_PT_EX2_FAILED = 10
    FREE_EDGE1_FAILED = 11
    FREE_EDGE2_FAILED = 12
    FREE_EDGE3_FAILED = 13
End Enum

Private Property Get Source() As String
    Let Source = "GCAPSWithTrimBack3.GenericPlateImp"
End Property
Private Property Get Method(sMethod As String) As String
    Let Method = Source + "::" + sMethod
End Property

Private Sub Class_Initialize()
    Set m_pGCsEntitiesFactory = Nothing
    Set m_pGCs(0) = Nothing
    
    Call m_oAdvancedDebug.StartSource(Source)
    Set m_oGenericPlateHelper = CreateObject("GCGenericPlateHelper.GenericPlateHelper")
End Sub
Private Sub Class_Terminate()
    Call ReleaseCacheOfGCDefinitions
    
    Set m_oGenericPlateHelper = Nothing
    Call m_oAdvancedDebug.StopSource(Source)
    Set m_oAdvancedDebug = Nothing
End Sub

Private Sub IJGCConnectedElementSC_PostDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PostDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PostDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod


End Sub

Private Sub IJGCConnectedElementSC_PreDisconnectExternalRels(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pInfo As REVISIONLib.IJTransformGraphInfo)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCConnectedElementSC_PreDisconnectExternalRels"))
    
    Call m_oGenericPlateHelper.PreDisconnectExternalRels(pGC, pInfo)
    
    Call m_oAdvancedDebug.ExitMethod

End Sub

'
' implementation of the IJGeometricConstructionDefinitionService interface
'
Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGCDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Initialize"))
    
    Call m_oGenericPlateHelper.PreInitialize(pGCDefinition)
    Call Initialize(pGCDefinition)
    Call m_oGenericPlateHelper.PostInitialize(pGCDefinition)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGC As IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    Call m_oAdvancedDebug.EnterMethod(Method("IJGeometricConstructionDefinitionService_Evaluate"))
        
    Dim bEvaluateGeometry As Boolean
    Call m_oGenericPlateHelper.PreEvaluate(pGC, pPOM, bEvaluateGeometry)
    Call m_oAdvancedDebug.EnterMethod(Method("EvaluateGeometry"))
    If bEvaluateGeometry Then
      Call Evaluate(pGC, pPOM)
    Else
        Call m_oAdvancedDebug.ShowMsg("Skip")
    End If
    
    Call m_oAdvancedDebug.ExitMethod
    Call m_oGenericPlateHelper.PostEvaluate(pGC, pPOM)

'''    Dim oAPS As Object:
'''    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
'''        Set oAPS = pGC.ControlledInputs("AdvancedPlateSystem")(1)
'''        Dim pPlateMoldedConventions As IJDPlateMoldedConventions: Set pPlateMoldedConventions = oAPS
'''        pPlateMoldedConventions.plateThicknessDirection = WithNormalDir
'''    End If
    
    Call m_oAdvancedDebug.ExitMethod
    Exit Sub
ErrorHandler:
    Dim iError As Long: Let iError = Err.Number
    Call m_oAdvancedDebug.ExitMethod
    Call m_oGenericPlateHelper.ProcessError(pGC, iError)
End Sub
'
' implementation of the IJGCSemanticConnection interface
'
Private Sub IJGCSemanticConnection_PostConnectionAdded(ByVal oRelationship As Object)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PostConnectionAdded"))
    
    Call m_oGenericPlateHelper.PostConnectionAdded(oRelationship)
    
    Call m_oAdvancedDebug.ExitMethod
End Sub
Private Sub IJGCSemanticConnection_PreConnectionRemoved(ByVal oRelationship As Object, ByVal bIsOriginDeleted As Boolean, ByVal bIsDestinationDeleted As Boolean)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCSemanticConnection_PreConnectionRemoved"))
    
    Call m_oGenericPlateHelper.PreConnectionRemoved(oRelationship, bIsOriginDeleted, bIsDestinationDeleted)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMigrate interface
'
Private Sub IJGCMigrate_Migrate(ByVal MyGC As IJGeometricConstruction, ByVal pMigrateHelper As IJGCMigrateHelper)
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMigrate_Migrate"))

    Call m_oGenericPlateHelper.Migrate(MyGC, pMigrateHelper)

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCMirror interface
'
Private Sub IJGCMirror_Adapt(ByVal pGCOfOriginalGC As IJGeometricConstruction, ByVal pGCOfMirroredGC As IJGeometricConstruction, _
                             ByVal pPlaneOfMirrorPlane As IJPlane, ByVal pT4x4OfMirrorTransformation As IJDT4x4, _
                             ByVal pElementsOfGCsFromInitalCopySet As IJElements)

    Call m_oAdvancedDebug.EnterMethod(Method("IJGCMirror_Adapt"))

    Call m_oGenericPlateHelper.Adapt(pGCOfOriginalGC, pGCOfMirroredGC, _
                                     pPlaneOfMirrorPlane, pT4x4OfMirrorTransformation, _
                                     pElementsOfGCsFromInitalCopySet, _
                                     "CoordinateSystem1", "Side_2")

    Call m_oAdvancedDebug.ExitMethod
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    Call m_oAdvancedDebug.EnterMethod(Method("IJGCToDoDelegate_ToDoDelegate"))

    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    On Error Resume Next
    Set IJGCToDoDelegate_ToDoDelegate = m_oGenericPlateHelper.ToDoDelegate(pGC)
    On Error GoTo 0

    Call m_oAdvancedDebug.ExitMethod
End Property
'
' implementation of the IJGeometricConstructionDefinition interface
'
Private Sub Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call InitializeCacheOfGCDefinitions
    
    ' define inputs
    Call GeometricConstructionDefinition_CopyInputs(pGeometricConstructionDefinition, m_pGCsDefinitions(0))
    Call GeometricConstructionDefinition_CopyInputs(pGeometricConstructionDefinition, m_pGCsDefinitions(1), "AuxiliaryFCBO_1|AuxiliaryFCBO,|CoordinateSystem")
    Call GeometricConstructionDefinition_CopyInputs(pGeometricConstructionDefinition, m_pGCsDefinitions(2), "|SketchingPlane,|CoordinateSystem")
    
    Call pGeometricConstructionDefinition.AddParameter("BehaviorOfAPS", "BehaviorOfAPS", 1, 0, 0, 0, 0, "PlateSolveAmbiguity", 0)
    
    ' copy controlled inputs
    Dim i As Integer
    For i = iFIRST_GC To iMAX_GCS
        Call GeometricConstructionDefinition_CopyControlledInputs(pGeometricConstructionDefinition, m_pGCsDefinitions(i), GetSuffixOfComponent(i))
    Next

    ' define parameters
    Call pGeometricConstructionDefinition.AddParameter(sFIRST, sFIRST, GCBoolean, 0, 0, 0, 0, False, False)
    For i = iFIRST_GC To iMAX_GCS
        Call GeometricConstructionDefinition_CopyParameters(pGeometricConstructionDefinition, m_pGCsDefinitions(i), GetSuffixOfComponent(i))
    Next
    Call pGeometricConstructionDefinition.AddParameter(sLAST, sLAST, GCBoolean, 0, 0, 0, 0, False, False)
    
    ' define output
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Support")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCSurfaceBody2, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCGTypePoint3d, "Node")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCLocalCoordinateSystem, "CoordinateSystem")
    Call pGeometricConstructionDefinition.AddOutput(GCOutputType.GCGTypeLine3d, "OrientationLine")
    
    ' define errors
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.PLANE_AT_OFFSET_FROM_FCBO, "ErrorValue1", "Evaluation of plane at offset from FCBO failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.POINTS_FOR_WEB_FRAME_FAILED, "ErrorValue2", "Evaluation of positioning points failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.CURVE_GENERATOR_FAILED, "ErrorValue4", "Evaluation of curve to be offseted failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.POINT_AT_MIN_DIST_EX1_FAILED, "ErrorValue6", "Evaluation of lower anchor point failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.POINT_AT_OFF_KEY_PT_EX1_FAILED, "ErrorValue7", "Evaluation of lower anchor point failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.POINT_AT_MIN_DIST_EX2_FAILED, "ErrorValue9", "Evaluation of upper anchor point failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.POINT_AT_OFF_KEY_PT_EX2_FAILED, "ErrorValue10", "Evaluation of upper anchor point failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.FREE_EDGE1_FAILED, "ErrorValue11", "Evaluation of free edge 1 failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.FREE_EDGE2_FAILED, "ErrorValue12", "Evaluation of free edge 2 failed")
    Call pGeometricConstructionDefinition.AddErrorValue(Errors.FREE_EDGE3_FAILED, "ErrorValue13", "Evaluation of free edge 3 failed")

End Sub
Private Sub Evaluate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As RESPOMLib.IJDPOM)
    ' share the same value for the 3 Offset parameters
    pGC.Parameter("Offset_12") = pGC.Parameter("Offset_4")
    pGC.Parameter("Offset_13") = pGC.Parameter("Offset_4")
   
    Call InitializeCacheOfGCDefinitions
    Call InitializeCacheOfGCs(pGC, True, pPOM)
    Dim i As Integer
    For i = iFIRST_GC To iMAX_GCS
        If Not m_pGCs(i) Is Nothing Then
            Call GeometricConstruction_PropagateControlledInputs(pGC, m_pGCs(i), GetSuffixOfComponent(i))
        Else
            Call GeometricConstruction_ClearControlledInputs(pGC, m_pGCsDefinitions(i), GetSuffixOfComponent(i))
        End If
    Next
    
    Dim pGCMacro As IJGeometricConstructionMacro: Set pGCMacro = pGC

    ' collect outputs
    pGC.Output("Support", 1) = m_pGCs(eComponents.POINTS_FOR_WEB_FRAME).Output("Support", 1)
    pGC.Output("OrientationLine", 1) = m_pGCs(eComponents.POINTS_FOR_WEB_FRAME).Output("OrientationLine", 1)
    pGC.Output("Boundary", "MiddleConnectable") = m_pGCs(eComponents.POINTS_FOR_WEB_FRAME).Output("Boundaries", "MiddleConnectable")
    
    If Not m_pGCs(eComponents.SELECT_POINT_TYPE1) Is Nothing Then
        If CDbl(pGC.Parameter("BoundaryOffset_6")) = 0 Then
            Call GeometricConstructionCMacro_RemoveBoundary(pGCMacro, "LowerBoundary")
            pGC.Output("Boundary", "LowerConnectable") = m_pGCs(eComponents.SELECT_POINT_TYPE1).Output("Boundary", 1)
        Else
            Call GeometricConstructionCMacro_RemoveBoundary(pGCMacro, "LowerConnectable")
            pGC.Output("Boundary", "LowerBoundary") = m_pGCs(eComponents.SELECT_POINT_TYPE1).Output("Boundary", 1)
        End If
    Else
        pGC.Output("Boundary", "LowerConnectable") = m_pGCs(eComponents.POINTS_FOR_WEB_FRAME).Output("Boundaries", "LowerConnectable")
    End If

    ' only a second boundary if the 2 boundaries are not the same
    If pGC.Parameter("BoundaryCount_2") = 2 Then
        If Not m_pGCs(eComponents.SELECT_POINT_TYPE2) Is Nothing Then
            If CDbl(pGC.Parameter("BoundaryOffset_9")) = 0 Then
                Call GeometricConstructionCMacro_RemoveBoundary(pGCMacro, "UpperBoundary")
                pGC.Output("Boundary", "UpperConnectable") = m_pGCs(eComponents.SELECT_POINT_TYPE2).Output("Boundary", 1)
            Else
                Call GeometricConstructionCMacro_RemoveBoundary(pGCMacro, "UpperConnectable")
                pGC.Output("Boundary", "UpperBoundary") = m_pGCs(eComponents.SELECT_POINT_TYPE2).Output("Boundary", 1)
            End If
        Else
            pGC.Output("Boundary", "UpperConnectable") = m_pGCs(eComponents.POINTS_FOR_WEB_FRAME).Output("Boundaries", "UpperConnectable")
        End If
    End If
  
    Dim pGCMacro2 As IJGeometricConstructionMacro: Set pGCMacro2 = m_pGCs(eComponents.SELECT_FREE_EDGE)
    For i = 1 To pGCMacro2.Outputs("Boundary").Count
        pGC.Output("Boundary", "FreeEdge" + CStr(i)) = pGCMacro2.Outputs("Boundary")(CStr(i))
    Next
    
    If pGCMacro2.Outputs("Boundary").Count = 1 _
    And pGCMacro.Outputs("Boundary").Count = 3 + 2 Then
        Dim pObject As IJDObject: Set pObject = pGCMacro.Outputs("Boundary")("FreeEdge2")
        Call pGCMacro.Outputs("Boundary").Remove("FreeEdge2")
        Call pObject.Remove
    End If
    
    pGC.Output("Node", 1) = m_pGCs(eComponents.POINTS_FOR_WEB_FRAME).Output("Points", "Middle")
    pGC.Output("CoordinateSystem", 1) = m_pGCs(eComponents.POINTS_FOR_WEB_FRAME).Output("CoordinateSystem", 1)
    
    Call ReleaseCacheOfGCs
End Sub
'
' implementation of the IJGeometricConstructionDynamic_IsOnRibbonBar interface
'
Private Sub IJGeometricConstructionDynamic_IsOnRibbonBar_PropertyValue(ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, bIsOnRibbonBar As Boolean)
   If pGC.Parameter("BoundaryCount_2") = 1 And pGC.Inputs("LowerConnectable").Count = 1 Then pGC.Input("UpperConnectable") = pGC.Input("LowerConnectable")
   
   If sName = "UpperConnectable" Then sName = "UpperConnectable_2"
   Call IsPropertyValue("IsOnRibbonBar", sName, pGC, info, bIsOnRibbonBar)
End Sub
Private Sub IJGeometricConstructionDynamic_IsOptional_PropertyValue(ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, bIsOptional As Boolean)
    If pGC.Parameter("BoundaryCount_2") = 1 And pGC.Inputs("LowerConnectable").Count = 1 Then pGC.Input("UpperConnectable") = pGC.Input("LowerConnectable")
   
    Call IsPropertyValue("IsOptional", sName, pGC, info, bIsOptional)
End Sub
Private Sub IsPropertyValue(ByVal sPropertyName As String, ByVal sName As String, ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal info As Variant, bIsProperyValue As Boolean)
    Select Case sName
        Case "CoordinateSystem1", sFIRST:
            If m_pGCs(0) Is Nothing Then Call InitializeCacheOfGCs(pGC, False, Nothing)
        Case sLAST:
            Call ReleaseCacheOfGCs
        Case "Offset_12", "Offset_13":
            bIsProperyValue = False
        Case Else:
            Dim iIndex As Integer
            Dim sShortName As String
            If True Then
                Dim iPositionOfUnderScore As Integer: Let iPositionOfUnderScore = InStr(sName, "_")
                If iPositionOfUnderScore <> 0 Then iIndex = Val(Mid$(sName, iPositionOfUnderScore + 1)) Else iIndex = 0
                If iPositionOfUnderScore <> 0 Then sShortName = Mid$(sName, 1, iPositionOfUnderScore - 1) Else sShortName = sName
            End If
            
            If iIndex <> 0 Then
                If Not m_pGCs(iIndex) Is Nothing Then
                    Select Case sPropertyName
                        Case "IsOnRibbonBar":
                            If TypeOf m_pGCsDefinitionsServices(iIndex) Is IJGeometricConstructionDynamic_IsOnRibbonBar Then
                                Dim pIsOnRibbonBar As IJGeometricConstructionDynamic_IsOnRibbonBar: Set pIsOnRibbonBar = m_pGCsDefinitionsServices(iIndex)
                                Call pIsOnRibbonBar.PropertyValue(sShortName, m_pGCs(iIndex), info, bIsProperyValue)
                            Else
                                bIsProperyValue = sShortName <> "Choice"
                            End If
                        Case "IsOptional":
                            If TypeOf m_pGCsDefinitionsServices(iIndex) Is IJGeometricConstructionDynamic_IsOptional Then
                                Dim pIsOptional As IJGeometricConstructionDynamic_IsOptional: Set pIsOptional = m_pGCsDefinitionsServices(iIndex)
                                Call pIsOptional.PropertyValue(sShortName, m_pGCs(iIndex), info, bIsProperyValue)
                            Else
                                bIsProperyValue = sShortName <> "Choice"
                            End If
                    End Select
                Else
                    bIsProperyValue = False
                End If
            End If
    End Select
End Sub
'
' implementation of the IJGeometricConstructionDynamic_ReStart interface
'
Private Sub IJGeometricConstructionDynamic_ReStart_PropertyValue(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction, ByVal bFirstStart As Boolean, sActiveInput4ReStart As String)
    If bFirstStart Then sActiveInput4ReStart = "FCBO" Else sActiveInput4ReStart = "MiddleConnectable"
End Sub
'
' manage cache of GCDefinitions
'
Private Sub InitializeCacheOfGCDefinitions()
    If m_pGCsEntitiesFactory Is Nothing Then
        ' define a set of GCTypes
        Dim i As Integer
        i = eComponents.CS_AS_CS: m_sGCTypes(i) = "CSAsCS": m_sSubstitutionsList(i) = ""
        i = eComponents.PLANE_AT_OFFSET_FROM_FCBO: m_sGCTypes(i) = "PlaneAtOffsetFromFCBO": m_sSubstitutionsList(i) = "CoordinateSystem|0.CoordinateSystem"
        i = eComponents.POINTS_FOR_WEB_FRAME: m_sGCTypes(i) = "PointsForWebFrame": m_sSubstitutionsList(i) = "SketchingPlane|1.,CoordinateSystem|0.CoordinateSystem"
        i = eComponents.SELECT_FREE_EDGE: m_sGCTypes(i) = "SelectFreeEdge": m_sSubstitutionsList(i) = ""
        i = eComponents.CURVE_GENERATOR: m_sGCTypes(i) = "CurveGenerator": m_sSubstitutionsList(i) = "SketchingPlane|1.,MiddleBoundary|2.Boundaries.MiddleConnectable,LowerCurve|2.Curves.Lower,UpperCurve|2.Curves.Upper,MiddlePoint|2.Points.Middle,LowerPoint|2.Points.Lower,UpperPoint|2.Points.Upper,RootPlateSystem|MiddleConnectable,CoordinateSystem|2.CoordinateSystem"
        i = eComponents.SELECT_POINT_TYPE1: m_sGCTypes(i) = "SelectPointType": m_sSubstitutionsList(i) = "Connectable1|MiddleConnectable,Connectable2|LowerConnectable"
        i = eComponents.POINT_AT_MIN_DIST_EX1: m_sGCTypes(i) = "PointAtMinDistEx": m_sSubstitutionsList(i) = "SketchingPlane|1.,Connectable1|MiddleConnectable,Connectable2|LowerConnectable,CoordinateSystem|2.CoordinateSystem,TracePoint|2.Points.Lower,SidePoint|2.Points.Upper"
        i = eComponents.POINT_AT_OFF_KEY_PT_EX1: m_sGCTypes(i) = "PointAtOffKeyPtEx": m_sSubstitutionsList(i) = "SketchingPlane|1.,Connectable|LowerConnectable,CoordinateSystem|2.CoordinateSystem,TracePoint|2.Points.Lower,SidePoint|2.Points.Upper"
        i = eComponents.SELECT_POINT_TYPE2: m_sGCTypes(i) = "SelectPointType": m_sSubstitutionsList(i) = "Connectable1|MiddleConnectable,Connectable2|UpperConnectable"
        i = eComponents.POINT_AT_MIN_DIST_EX2: m_sGCTypes(i) = "PointAtMinDistEx": m_sSubstitutionsList(i) = "SketchingPlane|1.,Connectable1|MiddleConnectable,Connectable2|UpperConnectable,CoordinateSystem|2.CoordinateSystem,TracePoint|2.Points.Upper,SidePoint|2.Points.Lower"
        i = eComponents.POINT_AT_OFF_KEY_PT_EX2: m_sGCTypes(i) = "PointAtOffKeyPtEx": m_sSubstitutionsList(i) = "SketchingPlane|1.,Connectable|UpperConnectable,CoordinateSystem|2.CoordinateSystem,TracePoint|2.Points.Upper,SidePoint|2.Points.Lower"
        i = eComponents.FREE_EDGE1: m_sGCTypes(i) = "FreeEdge1": m_sSubstitutionsList(i) = "SketchingPlane|1.,Point1|5.Point,CoordinateSystem1|5.CoordinateSystem,Point2|8.Point,CoordinateSystem2|8.CoordinateSystem"
        i = eComponents.FREE_EDGE2: m_sGCTypes(i) = "FreeEdge2": m_sSubstitutionsList(i) = "SketchingPlane|1.,Curve|4.Curve,CoordinateSystem|2.CoordinateSystem"
        i = eComponents.FREE_EDGE3: m_sGCTypes(i) = "FreeEdge3": m_sSubstitutionsList(i) = "SketchingPlane|1.,Curve|4.Curve,Point1|5.Point,CoordinateSystem1|5.CoordinateSystem,Point2|8.Point,CoordinateSystem2|8.CoordinateSystem,CoordinateSystem|2.CoordinateSystem"
        
        ' load corresponding GCDefinitions
        For i = iFIRST_GC To iMAX_GCS
            Dim bIsDefinitionAlreadyLoaded As Boolean: Let bIsDefinitionAlreadyLoaded = False
            Dim j As Integer
            If i > 1 Then
                For j = 1 To i - 1
                    If m_sGCTypes(i) = m_sGCTypes(j) Then
                        bIsDefinitionAlreadyLoaded = True
                        Exit For
                    End If
                Next
            End If
            
            If Not bIsDefinitionAlreadyLoaded Then
                Dim sProgid As String
                If m_sGCTypes(i) = "PlaneAtOffsetFromFCBO" Then
                    sProgid = "GCSHPSTRDefinitions.PlaneAtOffsetFromFCBODef"
                Else
                    sProgid = "GCWebFrame." + m_sGCTypes(i)
                End If

                Set m_pGCsDefinitionsServices(i) = SP3DCreateObject(sProgid)
                Set m_pGCsDefinitions(i) = New GeometricConstructionDefinition
                Call m_pGCsDefinitionsServices(i).Initialize(m_pGCsDefinitions(i))
            Else
                Set m_pGCsDefinitionsServices(i) = m_pGCsDefinitionsServices(j)
                Set m_pGCsDefinitions(i) = m_pGCsDefinitions(j)
            End If
        Next
         
        Set m_pGCsEntitiesFactory = New GeometricConstructionEntitiesFactory
    End If
End Sub
Private Sub ReleaseCacheOfGCDefinitions()
    If Not m_pGCsEntitiesFactory Is Nothing Then
        Dim i As Integer
        For i = iFIRST_GC To iMAX_GCS
            Set m_pGCsDefinitionsServices(i) = Nothing
            Set m_pGCsDefinitions(i) = Nothing
        Next
         
        Set m_pGCsEntitiesFactory = Nothing
    End If
End Sub
'
' manage cache of GCs
'
Private Sub InitializeCachedGC(i As Integer, pGC As IJGeometricConstruction, bEvaluate As Boolean, pPOM As IJDPOM)
    Dim sName As String: sName = "000" + CStr(i) + "-" + m_sGCTypes(i)
    Set m_pGCs(i) = m_pGCsEntitiesFactory.CreateEntity(m_sGCTypes(i), pPOM, sName)
    Call GeometricConstruction_CopyInputs(m_pGCs(i), pGC, m_pGCs, m_sSubstitutionsList(i))
    Call GeometricConstruction_CopyParameters(m_pGCs(i), pGC, GetSuffixOfComponent(i))
    If bEvaluate Then
        On Error Resume Next
        Call m_pGCs(i).Evaluate
        If Err.Number <> 0 Then
            On Error GoTo 0
            'MsgBox "Error on GC # " + CStr(i)
            Err.Raise i
        End If
    End If
End Sub
Private Sub InitializeCacheOfGCs(pGC As GeometricConstruction, bEvaluate As Boolean, pPOM As IJDPOM)
    Call InitializeCachedGC(eComponents.CS_AS_CS, pGC, bEvaluate, pPOM)
    Call InitializeCachedGC(eComponents.PLANE_AT_OFFSET_FROM_FCBO, pGC, bEvaluate, pPOM)
    Call InitializeCachedGC(eComponents.POINTS_FOR_WEB_FRAME, pGC, bEvaluate, pPOM)
    
    Call InitializeCachedGC(eComponents.SELECT_FREE_EDGE, pGC, True, pPOM)
    Select Case CInt(m_pGCs(eComponents.SELECT_FREE_EDGE).Parameter("Choice"))
        Case 2, 3:
            Call InitializeCachedGC(eComponents.CURVE_GENERATOR, pGC, bEvaluate, pPOM)
    End Select
    Select Case CInt(m_pGCs(eComponents.SELECT_FREE_EDGE).Parameter("Choice"))
        Case 1, 3:
            Call InitializeCachedGC(eComponents.SELECT_POINT_TYPE1, pGC, True, pPOM)
            Select Case CInt(m_pGCs(eComponents.SELECT_POINT_TYPE1).Parameter("Choice"))
                Case 1:
                    Call InitializeCachedGC(eComponents.POINT_AT_MIN_DIST_EX1, pGC, bEvaluate, pPOM)
                    If bEvaluate Then Call GeometricConstruction_CopyOutputs(m_pGCs(eComponents.SELECT_POINT_TYPE1), m_pGCs(eComponents.POINT_AT_MIN_DIST_EX1), "")
                Case 2:
                    Call InitializeCachedGC(eComponents.POINT_AT_OFF_KEY_PT_EX1, pGC, bEvaluate, pPOM)
                    If bEvaluate Then Call GeometricConstruction_CopyOutputs(m_pGCs(eComponents.SELECT_POINT_TYPE1), m_pGCs(eComponents.POINT_AT_OFF_KEY_PT_EX1), "")
            End Select
            
            Call InitializeCachedGC(eComponents.SELECT_POINT_TYPE2, pGC, True, pPOM)
            Select Case CInt(m_pGCs(eComponents.SELECT_POINT_TYPE2).Parameter("Choice"))
                Case 1:
                    Call InitializeCachedGC(eComponents.POINT_AT_MIN_DIST_EX2, pGC, bEvaluate, pPOM)
                    If bEvaluate Then Call GeometricConstruction_CopyOutputs(m_pGCs(eComponents.SELECT_POINT_TYPE2), m_pGCs(eComponents.POINT_AT_MIN_DIST_EX2), "")
                Case 2:
                    Call InitializeCachedGC(eComponents.POINT_AT_OFF_KEY_PT_EX2, pGC, bEvaluate, pPOM)
                    If bEvaluate Then Call GeometricConstruction_CopyOutputs(m_pGCs(eComponents.SELECT_POINT_TYPE2), m_pGCs(eComponents.POINT_AT_OFF_KEY_PT_EX2), "")
            End Select
    End Select
   
    Select Case CInt(m_pGCs(eComponents.SELECT_FREE_EDGE).Parameter("Choice"))
        Case 1:
            Call InitializeCachedGC(eComponents.FREE_EDGE1, pGC, bEvaluate, pPOM)
            If bEvaluate Then Call GeometricConstruction_CopyOutputs(m_pGCs(eComponents.SELECT_FREE_EDGE), m_pGCs(eComponents.FREE_EDGE1), "")
        Case 2:
            Call InitializeCachedGC(eComponents.FREE_EDGE2, pGC, bEvaluate, pPOM)
            If bEvaluate Then Call GeometricConstruction_CopyOutputs(m_pGCs(eComponents.SELECT_FREE_EDGE), m_pGCs(eComponents.FREE_EDGE2), "")
        Case 3:
            Call InitializeCachedGC(eComponents.FREE_EDGE3, pGC, bEvaluate, pPOM)
            If bEvaluate Then Call GeometricConstruction_CopyOutputs(m_pGCs(eComponents.SELECT_FREE_EDGE), m_pGCs(eComponents.FREE_EDGE3), "")
    End Select
End Sub
Private Sub ReleaseCacheOfGCs()
    ' release transient GCs
    Dim i As Integer
    For i = iFIRST_GC To iMAX_GCS
        Set m_pGCs(i) = Nothing
    Next
End Sub


