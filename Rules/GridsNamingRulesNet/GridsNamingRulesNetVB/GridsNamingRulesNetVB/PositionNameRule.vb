'***************************************************************************************
'  Copyright (C) 2009, Intergraph Corporation.  All rights reserved.
'
'  Project  : \GridSystem\SOM\Examples\Rules\GridsNamingRulesNet\GridsNamingRulesNetVB
'  Class    : PositionNameRule
'  Abstract : The file contains  implementation of Position NamingRule
'
'  Chaitanya   April '09       Creation
'  Chaitanya   09/07/2009      TR- 170922  Error is displayed when selecting .Net NameRule for a cylindrical plane. 
'***************************************************************************************

Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Grids.Middle

Public Class PositionNameRule
    Inherits NameRuleBase

    Private Const strDecimalFormat As String = "0.000"    'Precision of the Position 

    ''' <summary>
    '''  Creates a name for the object passed in. The name is based on the parents
    '''  name and object name.The Naming Parents are added in AddNamingParents().
    '''  Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <param name="oParents">Naming parents collection.</param>
    ''' <param name="oActiveEntity">Naming rules active entity on which the NamingParentsString is stored.</param>
    ''' <exception cref="ArgumentNullException">The Grids entity is null.</exception> 
    Public Overrides Sub ComputeName(ByVal oEntity As BusinessObject, ByVal oParents As ReadOnlyCollection(Of BusinessObject), ByVal oActiveEntity As BusinessObject)

        Dim oGridEntity As GridPlaneBase = Nothing
        Dim oGridCylinderEntity As GridCylinder = Nothing
        Dim oCS As CoordinateSystem

        Try
            If oEntity Is Nothing Then
                Throw New ArgumentNullException()
            End If

            Dim Axis As AxisType
            Dim CSType As CoordinateSystem.CoordinateSystemType

            If TypeOf (oEntity) Is GridPlaneBase Then
                oGridEntity = oEntity
                oCS = oGridEntity.Axis.CoordinateSystem
                Axis = oGridEntity.Axis.AxisType
                CSType = oCS.Type
            ElseIf TypeOf (oEntity) Is GridCylinder Then  ' RadialCylinder is to be handled seperately.
                oGridCylinderEntity = oEntity
                oCS = oGridCylinderEntity.Axis.CoordinateSystem
                Axis = oGridCylinderEntity.Axis.AxisType
            End If

            'Basing on the AxisType, set the name of the Plane.
            Select Case Axis

                Case AxisType.X
                    oEntity.SetPropertyValue(IIf(CSType = CoordinateSystem.CoordinateSystemType.Ship, "F", "E") + " " + (oGridEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name")
                Case AxisType.Y
                    oEntity.SetPropertyValue(IIf(CSType = CoordinateSystem.CoordinateSystemType.Ship, "L", "N") + " " + (oGridEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name")
                Case AxisType.Z
                    oEntity.SetPropertyValue(IIf(CSType = CoordinateSystem.CoordinateSystemType.Ship, "D", "El") + " " + (oGridEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name")
                Case AxisType.Radial
                    oEntity.SetPropertyValue("R" + " " + (oGridEntity.DistanceFromOrigin * 180 / 3.14159265358979).ToString(strDecimalFormat) + " deg", "IJNamedItem", "Name")
                Case AxisType.Cylindrical
                    oEntity.SetPropertyValue("C" + " " + (oGridCylinderEntity.DistanceFromOrigin).ToString(strDecimalFormat) + " m", "IJNamedItem", "Name")

            End Select

        Catch ex As Exception
            Throw New Exception("GridsNamingRulesNetVB.PositionNameRule.ComputeName " + ex.Message)
        End Try

        oGridEntity = Nothing
        oGridCylinderEntity = Nothing
        oCS = Nothing

    End Sub

    ''' <summary>
    ''' All the Naming Parents that need to participate in an objects naming are added here to the
    ''' Collection(Of BusinessObject). The parents added here are used in computing the name of the object in
    ''' ComputeName(). Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <returns> Collection of parents that participate </returns>

    Public Overrides Function GetNamingParents(ByVal oEntity As BusinessObject) As Collection(Of BusinessObject)
        Dim oParentsColl As New Collection(Of BusinessObject)

        GetNamingParents = oParentsColl
    End Function
End Class
