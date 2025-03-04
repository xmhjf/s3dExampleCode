Attribute VB_Name = "PlateShapeUtils"
Option Explicit

Const MODULE As String = "PlateShapeUtils"

Public Function GetCustomAttributeFromInputObject(oObject As Object, _
                                                  CustomIntfName As String, _
                                                  CustomAttrName As String) As IJDAttribute
    Const METHOD = "GetCustomAttributeFromInputObject"
    On Error GoTo ErrorHandler
    
    Set GetCustomAttributeFromInputObject = Nothing
    
    If Not TypeOf oObject Is IJDAttributeMetaData Then Exit Function
    If Not TypeOf oObject Is IJDAttributes Then Exit Function
    
    ' Get the interface IID of the custom interface
    Dim oAttributeMetaData As IJDAttributeMetaData
    Set oAttributeMetaData = oObject
    
    ' This call might fail if the custom I/F is not bulkloaded
    On Error Resume Next
        Dim varIID As Variant
        varIID = oAttributeMetaData.IID(CustomIntfName)
    On Error GoTo ErrorHandler
    
    If vbEmpty = VarType(varIID) Or vbNull = VarType(varIID) Then Exit Function

    Dim oAttributes As IJDAttributes
    Set oAttributes = oObject
    
    Dim oAttributesCol As IJDAttributesCol
    Set oAttributesCol = oAttributes.CollectionOfAttributes(varIID)
    
    Set GetCustomAttributeFromInputObject = oAttributesCol.Item(CustomAttrName)
    
CleanUp:
    Set oAttributeMetaData = Nothing
    Set oAttributes = Nothing
    Set oAttributesCol = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Function

Public Function IsSwage(oPlate As IJPlatePart) As Boolean
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlate, "IJUASwagePlate", "SwagePlate")
    
    If Not oAttr Is Nothing Then
        IsSwage = (oAttr.Value > 0)
    Else
        IsSwage = False
    End If
    
    Set oAttr = Nothing
End Function

Public Function IsSwadge(oPlate As IJPlatePart) As Boolean
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlate, "IJUAProductionPlateInfo", "CustomPlate")
    
    If Not oAttr Is Nothing Then
        IsSwadge = (100 = oAttr.Value) ' 100 is value for "Swadge" in codelist "CustomProdInfoPlateType"
    Else
        IsSwadge = False
    End If
    
    Set oAttr = Nothing
End Function

Public Function IsCorrugated(oPlate As IJPlatePart) As Boolean
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlate, "IJUAProductionPlateInfo", "CustomPlate")
    
    If Not oAttr Is Nothing Then
        IsCorrugated = (200 = oAttr.Value) ' 200 is value for "Corrugated" in codelist "CustomProdInfoPlateType"
    Else
        IsCorrugated = False
    End If
    
    Set oAttr = Nothing
End Function

Public Function IsBilgeHopper(oPlate As IJPlatePart) As Boolean
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlate, "IJUAProductionPlateInfo", "CustomPlate")
    
    If Not oAttr Is Nothing Then
        IsBilgeHopper = (201 = oAttr.Value) ' 201 is value for "Bilge Hopper" in codelist "CustomProdInfoPlateType"
    Else
        IsBilgeHopper = False
    End If
    
    Set oAttr = Nothing
End Function

Public Function IsLBH(oPlate As IJPlatePart) As Boolean
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlate, "IJUAProductionPlateInfo", "CustomPlate")
    
    If Not oAttr Is Nothing Then
        IsLBH = (202 = oAttr.Value) ' 202 is value for "LBH" in codelist "CustomProdInfoPlateType"
    Else
        IsLBH = False
    End If
    
    Set oAttr = Nothing
End Function

Public Function IsShoulderTank(oPlate As IJPlatePart) As Boolean
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlate, "IJUAProductionPlateInfo", "CustomPlate")
    
    If Not oAttr Is Nothing Then
        IsShoulderTank = (203 = oAttr.Value) ' 203 is value for "Shoulder Tank" in codelist "CustomProdInfoPlateType"
    Else
        IsShoulderTank = False
    End If
    
    Set oAttr = Nothing
End Function

Public Function IsCladPlate(oPlate As IJPlatePart) As Boolean
    Dim oAttr As IJDAttribute
    Set oAttr = GetCustomAttributeFromInputObject(oPlate, "IJUAProductionPlateInfo", "CustomPlate")
    
    If Not oAttr Is Nothing Then
        IsCladPlate = (300 = oAttr.Value) ' 300 is value for "Clad Plate" in codelist "CustomProdInfoPlateType"
    Else
        IsCladPlate = False
    End If
    
    Set oAttr = Nothing
End Function

Public Function IsCylidrical(oPlatePart As IJPlatePart) As Boolean
    Const METHOD As String = "IsCylidrical"
    On Error GoTo ErrorHandler

    IsCylidrical = False
    
    Dim ePlateCurvType As PlateCurvature
    ePlateCurvType = oPlatePart.Curved
        
    If ePlateCurvType = PLATE_CURVATURE_DoubleCurvature Or _
       ePlateCurvType = PLATE_CURVATURE_DoubleCurvature_Knuckled Or _
       ePlateCurvType = PLATE_CURVATURE_Flat Or _
       ePlateCurvType = PLATE_CURVATURE_SingleCurvature_Knuckled Or _
       ePlateCurvType = PLATE_CURVATURE_Knuckled _
    Then
        IsCylidrical = False
        Exit Function
    End If
    
    Dim PlateUtil As IJPlateAttributes
    Set PlateUtil = New GSCADCreateModifyUtilities.PlateUtils
  
    Dim eSurfType As SurGeoType
                
    Dim oPlatePartHlpr  As MfgRuleHelpers.PlatePartHlpr
    Set oPlatePartHlpr = New MfgRuleHelpers.PlatePartHlpr
    
    Set oPlatePartHlpr.object = oPlatePart
                    
    Dim oPlateSystem    As IJPlateSystem
    Set oPlateSystem = oPlatePartHlpr.GetRootSystem
                
    eSurfType = PlateUtil.SurfaceGeoType(oPlateSystem)
        
    If eSurfType = Revolution Then
        IsCylidrical = True
        GoTo CleanUp
    End If
    
    Dim oConnectable As IJStructConnectable
    Set oConnectable = oPlatePart
    
    Dim BasePort As IJPort
    oConnectable.GetBaseOffsetLateralTransientPorts vbNullString, True, _
                                                    BasePort, Nothing, Nothing
                                                    
    Dim oBaseSurf As IJSurfaceBody
    Set oBaseSurf = BasePort.Geometry
    
    Dim oTopoTool As IJDTopologyToolBox
    Set oTopoTool = New DGeomOpsToolBox
    
    Dim SurfFaces As IJElements
    oTopoTool.ExplodeSurfaceBodyByFaces Nothing, oBaseSurf, SurfFaces
    
    Dim i As Long
    For i = 1 To SurfFaces.Count
        Dim oSurfFace As IJSurfaceBody
        Set oSurfFace = SurfFaces.Item(i)
        
        Dim IsPlanar As Boolean
        oSurfFace.GetProperties IsPlanar
        Set oSurfFace = Nothing
        
        If IsPlanar Then
            IsCylidrical = False
            GoTo CleanUp
        End If
    Next
    
    IsCylidrical = True ' If this assessment of cylindrical turns out to be wrong,
                        ' we have to figure out what to do non-planar face(s).

CleanUp:
    Set oPlatePartHlpr = Nothing
    Set oPlateSystem = Nothing
    Set PlateUtil = Nothing
    Set oConnectable = Nothing
    Set BasePort = Nothing
    Set oBaseSurf = Nothing
    Set oTopoTool = Nothing
    If Not SurfFaces Is Nothing Then SurfFaces.Clear
    Set SurfFaces = Nothing
    
    Exit Function

ErrorHandler:

    Err.Raise StrMfgLogError(Err, MODULE, METHOD)
    GoTo CleanUp
End Function

