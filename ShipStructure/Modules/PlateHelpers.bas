Attribute VB_Name = "PlateHelpers"
Option Explicit
Private Const sModule As String = "PlateHelper"
Public Function PlateSystem_GetMaterialName(ByVal oPlateSystem As Object) As String
    Const sMethod As String = "PlateSystem_GetMaterialName"
    On Error GoTo ErrorHandler
    
    ' prepare result
    Dim sMaterialName As String
    
    ' retrieve material name
    Dim pStructureMaterial As IJStructureMaterial: Set pStructureMaterial = oPlateSystem
    sMaterialName = pStructureMaterial.MaterialName
    
    ' return result
    PlateSystem_GetMaterialName = sMaterialName
    Exit Function
ErrorHandler:
    Call LogAndRaiseError(sModule, sMethod)
End Function
Public Function PlateSystem_GetMaterialGrade(ByVal oPlateSystem As Object) As String
    Const sMethod As String = "PlateSystem_GetMaterialGrade"
    On Error GoTo ErrorHandler
    
    ' prepare result
    Dim sMaterialGrade As String
    
    ' retrieve material grade
    Dim pStructureMaterial As IJStructureMaterial: Set pStructureMaterial = oPlateSystem
    sMaterialGrade = pStructureMaterial.MaterialGrade
    
    ' return result
    PlateSystem_GetMaterialGrade = sMaterialGrade
    Exit Function
ErrorHandler:
    Call LogAndRaiseError(sModule, sMethod)
End Function
Public Function PlateSystem_GetThickness(ByVal oPlateSystem As Object) As Double
    Const sMethod As String = "PlateSystem_GetThickness"
    On Error GoTo ErrorHandler
    
    ' prepare result
    Dim dThickness As Double
    
    ' retrieve thickness
    Dim pPlate As IJPlate: Set pPlate = oPlateSystem
    dThickness = pPlate.thickness
    
    ' return result
    PlateSystem_GetThickness = dThickness
    Exit Function
ErrorHandler:
    Call LogAndRaiseError(sModule, sMethod)
End Function
Public Sub PlateSystem_SetMaterial(ByVal oPlateSystem As Object, ByVal sMaterialName As String, ByVal sMaterialGrade As String)
    Const sMethod As String = "PlateSystem_SetMaterial"
    On Error GoTo ErrorHandler
    
    ' find material object
    Dim pPlateSystem As IJPlateSystem: Set pPlateSystem = oPlateSystem
    Dim pStructQuery2 As IJDStructQuery2: Set pStructQuery2 = pPlateSystem.MoldedFormSpec
    Dim pMaterial As IJDMaterial: Call pStructQuery2.GetMaterial(sMaterialName, sMaterialGrade, pMaterial)
    
    ' set material object
    Dim pStructureMaterial As IJStructureMaterial: Set pStructureMaterial = oPlateSystem
    pStructureMaterial.Material = pMaterial
    Exit Sub
ErrorHandler:
    Call LogAndRaiseError(sModule, sMethod)
End Sub
Public Sub PlateSystem_SetThickness(ByVal oPlateSystem As Object, ByVal dExpectedThickness As Double, ByRef dRealThickness As Double)
    Const sMethod As String = "PlateSystem_SetThickness"
    On Error GoTo ErrorHandler
    
    ' retrieve the material name and grade
    Dim sMaterialName As String: sMaterialName = PlateSystem_GetMaterialName(oPlateSystem)
    Dim sMaterialGrade As String: sMaterialGrade = PlateSystem_GetMaterialGrade(oPlateSystem)
        
    ' retrieve the list of valid thickness
    Dim dArrayOfValidThickness() As Double
    Dim pPlateSystem As IJPlateSystem: Set pPlateSystem = oPlateSystem
    Dim pStructQuery2 As IJDStructQuery2: Set pStructQuery2 = pPlateSystem.MoldedFormSpec
    Call pStructQuery2.GetPlateThicknesses(sMaterialName, sMaterialGrade, dArrayOfValidThickness)

    ' retrieve the thickness, which is the closer from the expected thickness
    Dim i As Integer
    For i = LBound(dArrayOfValidThickness, 1) To UBound(dArrayOfValidThickness, 1)
        If dExpectedThickness <= dArrayOfValidThickness(i) Then
            dRealThickness = dArrayOfValidThickness(i)
            Exit For
        End If
    Next

    ' build a dimensions object
    Dim pPlateDimensions As IJDPlateDimensions
    Call pStructQuery2.GetPlateDimension(sMaterialName, sMaterialGrade, dRealThickness, pPlateDimensions)
    
    ' set the dimensions of the plate
    Dim pPlate As IJPlate: Set pPlate = oPlateSystem
    pPlate.Dimensions = pPlateDimensions
    Exit Sub
ErrorHandler:
    Call LogAndRaiseError(sModule, sMethod)
End Sub
Public Sub LogAndRaiseError(sModule As String, sMethod As String)
    ' retrieve usefull information
    Dim lErrorNumber As Long: lErrorNumber = Err.Number
    Dim sErrorDescription As String: sErrorDescription = Err.Description
    
    ' build good error message
    Dim sMessage As String: sMessage = "Unexpected error was occurred at the following module: " & sModule & " in method: " & sModule & "::" & sMethod
    
    ' log error to error.log
    Dim pEditErrors As IJEditErrors: Set pEditErrors = New JServerErrors
    Call pEditErrors.Write(lErrorNumber, sModule, sErrorDescription, sMessage)

    Err.Raise lErrorNumber, sModule + "::" + sMethod, sErrorDescription
End Sub
