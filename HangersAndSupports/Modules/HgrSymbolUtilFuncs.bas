Attribute VB_Name = "HgrSymbolUtilFuncs"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   HgrSymbolUtilFuncs.bas
'   ProgID:
'   Author:         Suresh
'   Creation Date:  10.Dec.2002
'   Description:
'    Utility procedures for the symbols
'
'   Change History:
'       10.Dec.2002     Suresh            Creation Date
'       27.Oct.2011     Ramya             TR 200039  Warnings found in error log when placing HS_S3DPlate
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "HgrSymbolUtilFuncs"  'Used for error messages

Public Function GetAttributeFromObject(pObject As Object, strName As String) As Variant
Const METHOD = "GetAttributeFromObject"
On Error GoTo ErrHandler

    Dim oAttributes             As IMSAttributes.IJDAttributes
    Dim oAttributesCol          As IJDAttributesCol
    Dim oAttribute              As IJDAttribute
    Dim oAttributeInfo          As IJDAttributeInfo
    Dim iID                     As Variant
    Dim iValue                  As Integer
    iValue = 0

    Set oAttributes = pObject
    For Each iID In oAttributes
        Set oAttributesCol = oAttributes.CollectionOfAttributes(iID)
        If oAttributesCol.Count > 0 Then
            For Each oAttribute In oAttributesCol
                Set oAttributeInfo = oAttribute.AttributeInfo
                If UCase(oAttributeInfo.name) = UCase(strName) Then
                    GetAttributeFromObject = oAttribute.Value
                    'cleanup
                    Set oAttributeInfo = Nothing
                    Set oAttribute = Nothing
                    Set oAttributesCol = Nothing
                    Exit Function
                End If
                Set oAttributeInfo = Nothing
                Set oAttribute = Nothing
            Next
        End If
        Set oAttributesCol = Nothing
    Next iID
    
    Set GetAttributeFromObject = Nothing
    ' Log which attribute is not found
    LogError Err, MODULE, METHOD, "Attribute '" & strName & "' not found"
    
Exit Function
ErrHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Function


Public Sub CalcWCG(oSymbol As Object, ByRef Weight As Double, ByRef CogX As Double, ByRef CogY As Double, ByRef CogZ As Double)
Const METHOD = "CalcWCG"
On Error GoTo ErrHandler

    Dim oIJDSymbol As IJDSymbol
    Dim oSymbolDef As IJDSymbolDefinition
    Dim oRep As IJDRepresentation
    Dim oReps As IJDRepresentations
    Dim oOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    
    Set oIJDSymbol = oSymbol
    Set oSymbolDef = oIJDSymbol.IJDSymbolDefinition(1)
    Set oReps = oSymbolDef.IJDRepresentations
    Set oRep = oReps.GetRepresentationByName("Symbolic")
    Set oOutputs = oRep
    
    Dim oProj As Projection3d
    Dim oCurve As IJCurve
    Dim curveArea As Double
    Dim curveCG As DPosition
    Set curveCG = New DPosition
    Dim projCG As DPosition
    Set projCG = New DPosition
    Dim projVolume As Double
    Dim AccumCG As DPosition
    Set AccumCG = New DPosition
    Dim AccumVolume As Double
    Dim projVecx As Double, projVecy As Double, projVecz As Double
    Dim projVecLength As Double
    Dim x As Double, y As Double, z As Double

    Dim ind As Integer
    For ind = 1 To oOutputs.OutputCount
        ' get the symbol out put and check if it is of type projection.
        ' If out is not of type projection then we exclude that from WCG calculation.
        ' We need to find a way to calculate the WCG for other types also.
        Set oOutput = oOutputs.GetOutputAtIndex(ind)
        
        On Error Resume Next
        Set oProj = Nothing ' free the interface if something is already stored
        Set oProj = oIJDSymbol.BindToOutput("Symbolic", oOutput.name)
        
        If Not oProj Is Nothing Then
            oProj.GetProjection projVecx, projVecy, projVecz
            Set oCurve = oProj.curve

            curveArea = oCurve.Area
            oCurve.Centroid x, y, z
            curveCG.Set x, y, z
            projCG.x = curveCG.x + (0.5 * projVecx)
            projCG.y = curveCG.y + (0.5 * projVecy)
            projCG.z = curveCG.z + (0.5 * projVecz)
            projVecLength = Sqr((projVecx * projVecx) + (projVecy * projVecy) + (projVecz * projVecz))
            
            projVolume = projVecLength * curveArea
    
            AccumVolume = AccumVolume + projVolume
            If AccumVolume > 0 Then
                AccumCG.x = AccumCG.x + (projCG.x - AccumCG.x) * projVolume / AccumVolume
                AccumCG.y = AccumCG.y + (projCG.y - AccumCG.y) * projVolume / AccumVolume
                AccumCG.z = AccumCG.z + (projCG.z - AccumCG.z) * projVolume / AccumVolume
            End If
        Else
            'something besides projection..
        End If

    Next
    
'    Dim oPartOcc As IJPartOcc
'    Dim oHgrPart As IJHgrPart
'
'    Dim oIJDMaterial As IJDMaterial
'    Dim density As Variant
'
'    Set oPartOcc = oSymbol
'    oPartOcc.GetPart oHgrPart
'
'    Set oIJDMaterial = oHgrPart.material
'    On Error Resume Next
'
'    If Not oIJDMaterial Is Nothing Then
'        density = oIJDMaterial.density
'    Else
'        density = 7800
'    End If
    
    Dim oCrossSection As Object
    Dim oSupportComp As IJHgrSupportComponent
    Set oSupportComp = oSymbol
    Set oCrossSection = oSupportComp.GetPartCrossSection
    
    ' get the weight per unit length attribute from cross section
    Dim WeightPerUnitLength As Double
    WeightPerUnitLength = GetAttributeFromObject(oCrossSection, "UnitWeight")
    Weight = WeightPerUnitLength * projVecLength
'    Weight = AccumVolume * density
    CogX = AccumCG.x
    CogY = AccumCG.y
    CogZ = AccumCG.z
    
    Set oSymbolDef = Nothing
    Set oRep = Nothing
    Set oReps = Nothing
    Set oOutputs = Nothing
    Set oOutput = Nothing
    
Exit Sub
ErrHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub

Public Function GetModelResourceManager() As Object
    Dim jContext As IJContext
    'Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set GetModelResourceManager = oConnectMiddle.GetResourceManager(strModelDBID)
End Function
Public Function GetCatalogResourceManager() As Object
    Dim jContext As IJContext
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strCatalogDBID As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strCatalogDBID = oDBTypeConfig.get_DataBaseFromDBType("Catalog")
    Set GetCatalogResourceManager = oConnectMiddle.GetResourceManager(strCatalogDBID)
End Function
Public Function GetPortByNameFromCollection(oPortElems As IJElements, strName As String) As Object
Const METHOD = "GetPortByNameFromCollection"
    On Error Resume Next
    
    
    Dim oHgrPort As IJHgrPort
    
    Set GetPortByNameFromCollection = Nothing
    
    For Each oHgrPort In oPortElems
        If oHgrPort.name = strName Then
            Set GetPortByNameFromCollection = oHgrPort
            Exit Function
        End If
    Next
    
    Exit Function
    
ErrHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
