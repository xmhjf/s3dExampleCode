VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BucklingPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum Errors
    MISSING_TRANSITION_PLATE1 = 1
    MISSING_TRANSITION_PLATE2 = 2
    MISSING_CONTINUITY_PLATE1 = 3
    DUPLICATE_INPUTS = 4
End Enum

Implements IJGeometricConstructionDefinitionService
Implements IJGCToDoDelegate

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    ' located input
    Call pGeometricConstructionDefinition.AddInput("ContinuityPlate1", "Select the first positioning Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate1", "Select the first limiting Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("TransitionPlate2", "Select the second limiting Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 1, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("ContinuityPlate2", "Select a second positioning Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate1", "Select a first lateral bounding Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddInput("SidePlate2", "Select a second lateral bounding Plate", "IJPlateSystem AND [GCFilters.Filters,IsRootPlateSystem]", 0, 1, "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    
    ' controlled inputs
    Call pGeometricConstructionDefinition.AddControlledInput("AdvancedPlateSystem", "ISPSPartPrismaticDesignNotify ISPSDesignedMemberDesignNotify")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfTransitionPlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfTransitionPlate2")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfContinuityPlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfContinuityPlate2")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfSidePlate1")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfSidePlate2")
    Call pGeometricConstructionDefinition.AddControlledInput("PortOfMemberPart")
       
    ' error codes
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_TRANSITION_PLATE1, "MissingTransitionPlate1", "Missing the transition plate #1")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_TRANSITION_PLATE2, "MissingTransitionPlate2", "Missing the transition plate #2")
    Call pGeometricConstructionDefinition.AddErrorValue(MISSING_CONTINUITY_PLATE1, "MissingContinuityPlate1", "Missing the continuity plate #1")
    Call pGeometricConstructionDefinition.AddErrorValue(DUPLICATE_INPUTS, "DuplicateInputs", "2 inputs are the same")
    
    ' outputs
    Call pGeometricConstructionDefinition.AddOutput(6, "Support")
    Call pGeometricConstructionDefinition.AddOutput(6, "Boundary")
    Call pGeometricConstructionDefinition.AddOutput(GCLocalCoordinateSystem, "CoordinateSystem")
End Sub
Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As IJDPOM)
    On Error GoTo ErrorHandler
    
    Call Evaluate(pGeometricConstruction, False)
    
    Exit Sub
ErrorHandler:
    Call GCProcessError(pGeometricConstruction, , Err.Number)
End Sub
Private Sub Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal bIsAPSDeleting As Boolean)
    ' check inputs
    If pGeometricConstruction.Inputs("TransitionPlate1").Count = 0 Then Err.Raise MISSING_TRANSITION_PLATE1
    If pGeometricConstruction.Inputs("TransitionPlate2").Count = 0 Then Err.Raise MISSING_TRANSITION_PLATE2
    If pGeometricConstruction.Inputs("ContinuityPlate1").Count = 0 Then Err.Raise MISSING_CONTINUITY_PLATE1
    If True Then
        Dim pElements As IJElements: Set pElements = New JObjectCollection
        Call pElements.Add(pGeometricConstruction.Input("TransitionPlate1"))
        If pElements.Contains(pGeometricConstruction.Input("TransitionPlate2")) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input("TransitionPlate2"))
        If pElements.Contains(pGeometricConstruction.Input("ContinuityPlate1")) Then Err.Raise DUPLICATE_INPUTS
        Call pElements.Add(pGeometricConstruction.Input("ContinuityPlate1"))
        If pGeometricConstruction.Inputs("ContinuityPlate2").Count > 0 Then
            If pElements.Contains(pGeometricConstruction.Input("ContinuityPlate2")) Then Err.Raise DUPLICATE_INPUTS
            Call pElements.Add(pGeometricConstruction.Input("ContinuityPlate2"))
        End If
        If pGeometricConstruction.Inputs("SidePlate1").Count > 0 Then
            If pElements.Contains(pGeometricConstruction.Input("SidePlate1")) Then Err.Raise DUPLICATE_INPUTS
            Call pElements.Add(pGeometricConstruction.Input("SidePlate1"))
        End If
        If pGeometricConstruction.Inputs("SidePlate2").Count > 0 Then
            If pElements.Contains(pGeometricConstruction.Input("SidePlate2")) Then Err.Raise DUPLICATE_INPUTS
        End If
    End If
    
    Dim oGCFactory As IJGeometricConstructionEntitiesFactory: Set oGCFactory = New GeometricConstructionEntitiesFactory
    
    Dim oBucklingPlate As GeometricConstruction
    If pGeometricConstruction.Inputs("ContinuityPlate1").Count = 1 _
    And pGeometricConstruction.Inputs("ContinuityPlate2").Count = 1 Then
        Set oBucklingPlate = oGCFactory.CreateEntity("BucklingPlate2", Nothing)
        oBucklingPlate.Inputs("ContinuityPlate1").Add pGeometricConstruction.Inputs("ContinuityPlate1").Item(1)
        oBucklingPlate.Inputs("ContinuityPlate2").Add pGeometricConstruction.Inputs("ContinuityPlate2").Item(1)
        oBucklingPlate.Inputs("TransitionPlate1").Add pGeometricConstruction.Inputs("TransitionPlate1").Item(1)
        oBucklingPlate.Inputs("TransitionPlate2").Add pGeometricConstruction.Inputs("TransitionPlate2").Item(1)
        oBucklingPlate.Inputs("SidePlate1").AddElements pGeometricConstruction.Inputs("SidePlate1")
        oBucklingPlate.Inputs("SidePlate2").AddElements pGeometricConstruction.Inputs("SidePlate2")
        oBucklingPlate.ControlledInputs("AdvancedPlateSystem").AddElements pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")
        oBucklingPlate.ControlledInputs("PortOfContinuityPlate1").AddElements pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1")
        oBucklingPlate.ControlledInputs("PortOfContinuityPlate2").AddElements pGeometricConstruction.ControlledInputs("PortOfContinuityPlate2")
        oBucklingPlate.ControlledInputs("PortOfTransitionPlate1").AddElements pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1")
        oBucklingPlate.ControlledInputs("PortOfTransitionPlate2").AddElements pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2")
        oBucklingPlate.ControlledInputs("PortOfSidePlate1").AddElements pGeometricConstruction.ControlledInputs("PortOfSidePlate1")
        oBucklingPlate.ControlledInputs("PortOfSidePlate2").AddElements pGeometricConstruction.ControlledInputs("PortOfSidePlate2")
        oBucklingPlate.Evaluate
        pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1").Clear
        pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1").AddElements oBucklingPlate.ControlledInputs("PortOfContinuityPlate1")
        pGeometricConstruction.ControlledInputs("PortOfContinuityPlate2").Clear
        pGeometricConstruction.ControlledInputs("PortOfContinuityPlate2").AddElements oBucklingPlate.ControlledInputs("PortOfContinuityPlate2")
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").Clear
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").AddElements oBucklingPlate.ControlledInputs("PortOfTransitionPlate1")
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2").Clear
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2").AddElements oBucklingPlate.ControlledInputs("PortOfTransitionPlate2")
        pGeometricConstruction.ControlledInputs("PortOfSidePlate1").Clear
        pGeometricConstruction.ControlledInputs("PortOfSidePlate1").AddElements oBucklingPlate.ControlledInputs("PortOfSidePlate1")
        pGeometricConstruction.ControlledInputs("PortOfSidePlate2").Clear
        pGeometricConstruction.ControlledInputs("PortOfSidePlate2").AddElements oBucklingPlate.ControlledInputs("PortOfSidePlate2")
    ElseIf pGeometricConstruction.Inputs("ContinuityPlate1").Count = 1 Then
        Set oBucklingPlate = oGCFactory.CreateEntity("BucklingPlate1", Nothing)
        oBucklingPlate.Inputs("ContinuityPlate").Add pGeometricConstruction.Inputs("ContinuityPlate1").Item(1)
        oBucklingPlate.Inputs("TransitionPlate1").Add pGeometricConstruction.Inputs("TransitionPlate1").Item(1)
        oBucklingPlate.Inputs("TransitionPlate2").Add pGeometricConstruction.Inputs("TransitionPlate2").Item(1)
        oBucklingPlate.Inputs("SidePlate1").AddElements pGeometricConstruction.Inputs("SidePlate1")
        oBucklingPlate.Inputs("SidePlate2").AddElements pGeometricConstruction.Inputs("SidePlate2")
        oBucklingPlate.ControlledInputs("AdvancedPlateSystem").AddElements pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")
        oBucklingPlate.ControlledInputs("PortOfContinuityPlate").AddElements pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1")
        oBucklingPlate.ControlledInputs("PortOfTransitionPlate1").AddElements pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1")
        oBucklingPlate.ControlledInputs("PortOfTransitionPlate2").AddElements pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2")
        oBucklingPlate.ControlledInputs("PortOfSidePlate1").AddElements pGeometricConstruction.ControlledInputs("PortOfSidePlate1")
        oBucklingPlate.ControlledInputs("PortOfSidePlate2").AddElements pGeometricConstruction.ControlledInputs("PortOfSidePlate2")
        oBucklingPlate.Evaluate
        pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1").Clear
        pGeometricConstruction.ControlledInputs("PortOfContinuityPlate1").AddElements oBucklingPlate.ControlledInputs("PortOfContinuityPlate")
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").Clear
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate1").AddElements oBucklingPlate.ControlledInputs("PortOfTransitionPlate1")
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2").Clear
        pGeometricConstruction.ControlledInputs("PortOfTransitionPlate2").AddElements oBucklingPlate.ControlledInputs("PortOfTransitionPlate2")
        pGeometricConstruction.ControlledInputs("PortOfSidePlate1").Clear
        pGeometricConstruction.ControlledInputs("PortOfSidePlate1").AddElements oBucklingPlate.ControlledInputs("PortOfSidePlate1")
        pGeometricConstruction.ControlledInputs("PortOfSidePlate2").Clear
        pGeometricConstruction.ControlledInputs("PortOfSidePlate2").AddElements oBucklingPlate.ControlledInputs("PortOfSidePlate2")
    End If

    Dim oGCMacro As IJGeometricConstructionMacro
    Set oGCMacro = pGeometricConstruction
    
    oGCMacro.Output("Support", 1) = oBucklingPlate.Output("Support", 1)
    oGCMacro.Output("Boundary", "PortOfTransitionPlate1") = oBucklingPlate.Output("Boundary", "PortOfTransitionPlate1")
    oGCMacro.Output("Boundary", "PortOfTransitionPlate2") = oBucklingPlate.Output("Boundary", "PortOfTransitionPlate2")
    oGCMacro.Output("Boundary", "PortOfSidePlate1") = oBucklingPlate.Output("Boundary", "PortOfSidePlate1")
    oGCMacro.Output("Boundary", "PortOfSidePlate2") = oBucklingPlate.Output("Boundary", "PortOfSidePlate2")
    oGCMacro.Output("CoordinateSystem", 1) = oBucklingPlate.Output("CoordinateSystem", 1)
    
    If pGeometricConstruction.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        ' add new boundaries
        Call PlateSystem_AddNewBoundaries(pGeometricConstruction.ControlledInputs("AdvancedPlateSystem")(1), oGCMacro)
    End If
End Sub
'
' implementation of the IJGCToDoDelegate interface
'
Private Property Get IJGCToDoDelegate_ToDoDelegate(ByVal pGC As SP3DGeometricConstruction.IJGeometricConstruction) As Object
    ' Delegate to the APS when exists
    Set IJGCToDoDelegate_ToDoDelegate = Nothing
    If pGC.ControlledInputs("AdvancedPlateSystem").Count = 1 Then
        Set IJGCToDoDelegate_ToDoDelegate = pGC.ControlledInputs("AdvancedPlateSystem")(1)
    End If
End Property
