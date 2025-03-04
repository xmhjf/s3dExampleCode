Attribute VB_Name = "CustomPlateCommon"
Option Explicit
'*******************************************************************
'
'Copyright (C) 2004 Intergraph Corporation. All rights reserved.
'
'File : CustomPlateCommon
'
'Author :
'
'Description :
'    SmartPlant Structural common custom plate part funtions
'
'History:
'
' 05/21/03   J.Schwartz     Replaced the Msgbox in HandleError with
'                           JServerErrors Add function
' 11/23/10   G. Gu     		DI-177040 Added new function to return proxy for IJDPlateDimensions using IJDMaterial and thickness as inputs
'
'********************************************************************
Private Const MODULE = "CustomPlateCommon"
Private Const CUSTOMPLATEPROGID = "StructCustomPlatePart.StructCustomPlatePart"

'*************************************************************************
'Function
'SetPlateDimensions
'
'Abstract
'Returns a PlateThickness value using the plate part's material property, IJDPlateDimensions.Thickness
'   that is greater than or equal to the value of PlateThickness when input
'Also sets oPlate.Dimensions as the found thickness value
'
'Arguments
'PlatePart is input to the function
'PlateThickness is input, and used to compare with materials' thickness
'PlateThickness is the returned double value
'
'Return
'Errors are written to the error log file and cleared.
'
'Exceptions
'
'***************************************************************************

Public Sub SetPlateDimensions(oPlate As IJStructPlate, ByRef PlateThickness As Double)
    Const METHOD = "SetPlateDimensions"

    On Error GoTo ErrorHandler

    Dim oIJDMaterial As IJDMaterial
    Dim oPlateMaterial As IJStructMaterial
    Dim matlThickCol As IJDCollection
    Dim oRefDataQuery As IJDStructServices
    Dim iIndex As Integer, maxThickIndex As Integer, bestThickIndex As Integer
    Dim tmpThickness As Double, dMaxThickness As Double, dBestThickness As Double
    Dim oPlateDims As IJDPlateDimensions
    Dim oCatalogPOM As IJDPOM
    Dim iJDObject As iJDObject
    Dim oModelPOM As IJDPOM
    Dim iProxy As IJDProxy
    
    Set oCatalogPOM = GetCatalogResourceManager()
    Set oPlateMaterial = oPlate
    Set iJDObject = oPlate
    Set oModelPOM = iJDObject.ResourceManager
    Set oIJDMaterial = oPlateMaterial.StructMaterial
    Set oRefDataQuery = New StructServices
    Set matlThickCol = oRefDataQuery.GetPlateDimensions(oCatalogPOM, oIJDMaterial.MaterialType, oIJDMaterial.MaterialGrade)

    ' Loop thru the available Thickness values to find the Thcikness value
    ' is greater then the requested thickness value
    dMaxThickness = -1#
    dBestThickness = -1#

    For iIndex = 1 To matlThickCol.Size
        Set oPlateDims = matlThickCol.Item(iIndex)
        tmpThickness = oPlateDims.Thickness
        If tmpThickness >= dMaxThickness Then
            dMaxThickness = tmpThickness
            maxThickIndex = iIndex
        End If
        
        If tmpThickness >= PlateThickness Then
            ' if not initialized yet, or smaller value found found ...
            If dBestThickness < 0# Or dBestThickness > tmpThickness Then
                dBestThickness = tmpThickness
                bestThickIndex = iIndex
            End If
        End If
    Next iIndex
    
    ' if No valid Thickness value was greater then the requested value
    ' use the largest value found
    If dBestThickness < 0# Then
        dBestThickness = dMaxThickness
        bestThickIndex = maxThickIndex
    End If
    
    ' Set the Part's Thickness value
    If dBestThickness > 0# Then
        Set oPlateDims = matlThickCol.Item(bestThickIndex)
        Set iProxy = oModelPOM.GetProxy(oPlateDims, True)
        oPlate.Dimensions = iProxy
        PlateThickness = dBestThickness
    End If
    
    Set oPlateMaterial = Nothing
    Set oIJDMaterial = Nothing
    Set oRefDataQuery = Nothing
    Set oPlateDims = Nothing
    Set matlThickCol = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'*************************************************************************
'Function
'GetPlateDimensionsProxy
'
'Abstract
'Returns a IJDPlateDimensions proxy object using the plate part's material property, IJDPlateDimensions.Thickness
'   The thickness in the object is greater than or equal to the input value of PlateThickness
'
'Arguments
'oModelObj is input to the function
'oIJDMaterial is input object contain metarial information
'PlateThickness is input, and used to compare with materials' thickness
'IJDProxy is the returned object type which represent the PlateDimensions
'
'Return
'Errors are written to the error log file and cleared.
'
'Exceptions
'
'***************************************************************************

Public Function GetPlateDimensionsProxy(oModelObj As Object, oIJDMaterial As IJDMaterial, ByVal PlateThickness As Double) As IJDProxy
    Const METHOD = "GetPlateDimensionsProxy"

    On Error GoTo ErrorHandler

    Dim matlThickCol As IJDCollection
    Dim oRefDataQuery As IJDStructServices
    Dim iIndex As Integer, maxThickIndex As Integer, bestThickIndex As Integer
    Dim tmpThickness As Double, dMaxThickness As Double, dBestThickness As Double
    Dim oPlateDims As IJDPlateDimensions
    Dim oCatalogPOM As IJDPOM
    Dim iJDObject As iJDObject
    Dim oModelPOM As IJDPOM
    
    Set oCatalogPOM = GetCatalogResourceManager()
    Set iJDObject = oModelObj
    Set oModelPOM = iJDObject.ResourceManager
    Set oRefDataQuery = New StructServices
    Set matlThickCol = oRefDataQuery.GetPlateDimensions(oCatalogPOM, oIJDMaterial.MaterialType, oIJDMaterial.MaterialGrade)

    ' Loop thru the available Thickness values to find the Thcikness value
    ' is greater then the requested thickness value
    dMaxThickness = -1#
    dBestThickness = -1#

    For iIndex = 1 To matlThickCol.Size
        Set oPlateDims = matlThickCol.Item(iIndex)
        tmpThickness = oPlateDims.Thickness
        If tmpThickness >= dMaxThickness Then
            dMaxThickness = tmpThickness
            maxThickIndex = iIndex
        End If
        
        If tmpThickness >= PlateThickness Then
            ' if not initialized yet, or smaller value found found ...
            If dBestThickness < 0# Or dBestThickness > tmpThickness Then
                dBestThickness = tmpThickness
                bestThickIndex = iIndex
            End If
        End If
    Next iIndex
    
    ' if No valid Thickness value was greater then the requested value
    ' use the largest value found
    If dBestThickness < 0# Then
        dBestThickness = dMaxThickness
        bestThickIndex = maxThickIndex
    End If
    
    ' Set the Part's Thickness value
    If dBestThickness > 0# Then
        Set oPlateDims = matlThickCol.Item(bestThickIndex)
        Set GetPlateDimensionsProxy = oModelPOM.GetProxy(oPlateDims, True)
    End If
    
    Set oRefDataQuery = Nothing
    Set oPlateDims = Nothing
    Set matlThickCol = Nothing

    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function


'*************************************************************************
'Function
'SetPlateMaterial
'
'Abstract
'Establishes a relation between the Material definition and the plate part.
'
'Arguments
'PlatePart is input, the plate part
'strType is the name of the material
'strGrade is the grade of the material
'
'Return
'Errors are written to the error log file and cleared.
'
'Exceptions
'
'***************************************************************************

Public Sub SetPlateMaterial(oPlate As IJStructCustomPlatePart, strType As String, strGrade As String)
    Const METHOD = "SetPlateMaterial"
    On Error GoTo ErrorHandler

    Dim oMatlObj As IJDMaterial
    Dim iProxy As IJDProxy
    Dim oRefDataQuery As IJDStructServices
    Dim oPlateMaterial As IJStructMaterial
    Dim oCatalogPOM As IJDPOM
    Dim oModelPOM As IJDPOM
    Dim iJDObject As iJDObject
    
    Set oRefDataQuery = New StructServices
    Set oCatalogPOM = GetCatalogResourceManager()
    Set oMatlObj = oRefDataQuery.GetMaterialFromGradeAndType(oCatalogPOM, strType, strGrade)

    Set iJDObject = oPlate
    Set oModelPOM = iJDObject.ResourceManager
    Set iProxy = oModelPOM.GetProxy(oMatlObj, True)

    Set oPlateMaterial = oPlate
    oPlateMaterial.StructMaterial = iProxy

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
    Err.Clear
End Sub

'*************************************************************************
'Function
'GenerateNameForPlate
'
'Abstract
'Establishes a name rule for the given plate
'
'Arguments
'PlatePart is input, the plate part
'
'Return
'Errors are written to the error log file and cleared.
'
'Exceptions
'
'***************************************************************************

Public Sub GenerateNameForPlate(oPlate As IJStructCustomPlatePart)
    Const METHOD = "GenerateNameForPlate"
    On Error GoTo ErrorHandler

    Dim NamingRules As IJElements
    Dim oNameRuleHolder As IJDNameRuleHolder

    Dim oNameRuleHlpr As IJDNamingRulesHelper
    Set oNameRuleHlpr = New NamingRulesHelper

    oNameRuleHlpr.GetEntityNamingRulesGivenProgID CUSTOMPLATEPROGID, NamingRules
    Dim oNameRuleAE As IJNameRuleAE

    If NamingRules.Count > 0 Then
        Set oNameRuleHolder = NamingRules.Item(1)
    End If

    Call oNameRuleHlpr.AddNamingRelations(oPlate, oNameRuleHolder, oNameRuleAE)
    
    Set NamingRules = Nothing
    Set oNameRuleAE = Nothing
    Set oNameRuleHolder = Nothing
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

