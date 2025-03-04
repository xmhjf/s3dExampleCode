VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJGeometricConstructionDefinitionService
Implements IJGeometricConstructionDefinitionTemplate

Private Sub IJGeometricConstructionDefinitionService_Initialize(ByVal pGeometricConstructionDefinition As SP3DGeometricConstruction.IJGeometricConstructionDefinition)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystem", "Select a coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("CoordinateSystemBis", "Select a second coordinate system", "IJDCoordinateSystem", 1, 1)
    Call pGeometricConstructionDefinition.AddInput("Point", "Select a Point", "IJPoint", 1, 1)
    pGeometricConstructionDefinition.AddOutput GCGTypePoint3d, "Point"
    pGeometricConstructionDefinition.AddOutput GCGTypePoint3d, "SuperPoint"
End Sub

Private Sub IJGeometricConstructionDefinitionService_Evaluate(ByVal pGeometricConstruction As SP3DGeometricConstruction.IJGeometricConstruction, ByVal pPOM As RESPOMLib.IJDPOM)
                
    
    Dim lInputCount As Long: lInputCount = pGeometricConstruction.Definition.InputCount
    Dim i As Long
    For i = 1 To lInputCount
        Dim sSetName As String: Dim sSetPrompt As String: Dim sSetFilter As String: Dim lSetMin As Long: Dim lSetMax As Long: Dim sSetComputeIIDs As String
        pGeometricConstruction.Definition.GetInputInfoByIndex i, sSetName, sSetPrompt, sSetFilter, lSetMin, lSetMax, sSetComputeIIDs
            
        Dim lTemplateInputCount As Long: lTemplateInputCount = pGeometricConstruction.Definition(True).InputCount
        Dim j As Long
        For j = 1 To lTemplateInputCount
            Dim sName As String: Dim sPrompt As String: Dim sFilter As String: Dim lMin As Long: Dim lMax As Long: Dim sComputeIIDs As String
            pGeometricConstruction.Definition(True).GetInputInfoByIndex j, sName, sPrompt, sFilter, lMin, lMax, sComputeIIDs

            ' create a filter from the template filter string
            Dim oFilter As IJDStepFilter: Set oFilter = New StepFilter: oFilter.AddCriteria sFilter
            
            ' if the select input object passes the filter criteria, add to the possible values
            ' try to find an input object on the set that pass the current filter criteria from the template
            If oFilter.PassCriteria(pGeometricConstruction.Inputs(sSetName)(1)) = 1 Then
            
                Dim oGCSet As IJGeometricConstructionSet: Set oGCSet = pGeometricConstruction
                oGCSet.OverrideInputInfoByName sSetName, sName, sPrompt, sFilter, lMin, lMax, sComputeIIDs
                
                Exit For
            End If
        Next
    Next
End Sub


