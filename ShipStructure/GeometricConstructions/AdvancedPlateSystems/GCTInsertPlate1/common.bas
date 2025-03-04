Attribute VB_Name = "Common"
Public Function GetPlateDimension(ByVal oPlate As Object, dThickness As Double, ByRef dRealThickness As Double) As IJDPlateDimensions
Const METHOD = "GetPlateDimension"

    On Error GoTo ErrorHandler
    
    Dim oSDHelper As New StructDetailHelper

    'Get Plate Part from APS
    '------------------------------------------------------------------------------
    Dim oEnumPartsUnk As IEnumUnknown
    oSDHelper.GetPartsDerivedFromSystem oPlate, oEnumPartsUnk, True

    Dim oCollectionOfParts As Collection

    Dim ConvertUtils As CCollectionConversions
    Set ConvertUtils = New CCollectionConversions
    ConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPartsUnk, oCollectionOfParts

    Dim oSD_PlatePart As New StructDetailObjects.PlatePart
    Set oSD_PlatePart.object = oCollectionOfParts.Item(1)
    
    dRealThickness = oSD_PlatePart.PlateThickness
    '----------------------------------------------------------------------------------
    
    'Get Material and Grad
    'That means, this is the necessary action for adding the plate thickness
    '---------------------------------------------------------------------------------
    Dim sGrade As String
    Dim sMaterial As String

    Dim oMaterial As IJDMaterial
    Dim oPlateSpec As IJDMoldedFormSpec
    Dim oPlateSystem As IJPlateSystem
    Dim oPlateDimensions As IJDPlateDimensions

    Set oMaterial = oSD_PlatePart.Material
    
    sGrade = oMaterial.materialGrade
    sMaterial = oMaterial.MaterialType
    '----------------------------------------------------------------------------------
    
    'Get Plate Dimension object based on Material and Grade from Plate Spec
    '---------------------------------------------------------------------------------
    Set oPlateSystem = oPlate
    Set oPlateSpec = oPlateSystem.MoldedFormSpec

    Dim oStructQuery2 As IJDStructQuery2
    Set oStructQuery2 = oPlateSpec

    '----------------------------------------------------------------------------------
    ' if No valid Thickness value was greater then the requested value
    ' use the largest value found
'    If dBestThickness < 0# Then
'        dBestThickness = dMaxThickness
'    End If
    
    oStructQuery2.GetPlateDimension sMaterial, sGrade, dThickness, _
                                    oPlateDimensions
    '---------------------------------------------------------------------------------
    
    Set GetPlateDimension = oPlateDimensions
        
Exit Function

ErrorHandler:
'    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


Public Function GetPlateDimension1(ByVal oPlate As Object, dThickness As Double, ByRef dRealThickness As Double) As IJDPlateDimensions
Const METHOD = "GetPlateDimension1"

    On Error GoTo ErrorHandler
    
    Dim oSDHelper As New StructDetailHelper

    'Get Plate Part from APS
    '------------------------------------------------------------------------------
    Dim oEnumPartsUnk As IEnumUnknown
    oSDHelper.GetPartsDerivedFromSystem oPlate, oEnumPartsUnk, True

    Dim oCollectionOfParts As Collection

    Dim ConvertUtils As CCollectionConversions
    Set ConvertUtils = New CCollectionConversions
    ConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPartsUnk, oCollectionOfParts

    Dim oSD_PlatePart As New StructDetailObjects.PlatePart
    Set oSD_PlatePart.object = oCollectionOfParts.Item(1)
    
'    dRealThickness = oSD_PlatePart.PlateThickness
    '----------------------------------------------------------------------------------
    
    'Get Material and Grad
    'That means, this is the necessary action for adding the plate thickness
    '---------------------------------------------------------------------------------
    Dim sGrade As String
    Dim sMaterial As String
    Dim iIndex As Long
    Dim lLowerBound As Long
    Dim lUpperBound As Long
    
    Dim dMaxThickness As Double
    Dim dBestThickness As Double
    Dim dListThickness() As Double

    Dim oMaterial As IJDMaterial
    Dim oPlateSpec As IJDMoldedFormSpec
    Dim oPlateSystem As IJPlateSystem
    Dim oPlateDimensions As IJDPlateDimensions

    Set oMaterial = oSD_PlatePart.Material
    
    sGrade = oMaterial.materialGrade
    sMaterial = oMaterial.MaterialType
    '----------------------------------------------------------------------------------
    
    'Get Plate Dimension object based on Material and Grade from Plate Spec
    '---------------------------------------------------------------------------------
    Set oPlateSystem = oPlate
    Set oPlateSpec = oPlateSystem.MoldedFormSpec

    Dim oStructQuery2 As IJDStructQuery2
    Set oStructQuery2 = oPlateSpec
    oStructQuery2.GetPlateThicknesses sMaterial, sGrade, dListThickness

    '---------------------------------------------------------------------------------
    ' Loop thru the available Thickness values to find the Thcikness value
    ' is greater then the requested thickness value

    lLowerBound = LBound(dListThickness, 1)
    lUpperBound = UBound(dListThickness, 1)

    For iIndex = lLowerBound To lUpperBound

    If dThickness <= dListThickness(iIndex) Then
     dThickness = dListThickness(iIndex)

      Exit For
    End If

    Next iIndex

    '----------------------------------------------------------------------------------
    ' if No valid Thickness value was greater then the requested value
    ' use the largest value found
'    If dBestThickness < 0# Then
'        dBestThickness = dMaxThickness
'    End If

    oStructQuery2.GetPlateDimension sMaterial, sGrade, dThickness, _
                                    oPlateDimensions
    '---------------------------------------------------------------------------------
    
    Set GetPlateDimension1 = oPlateDimensions
        
Exit Function

ErrorHandler:
'    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


    
    



    
    



