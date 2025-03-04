Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
	' Create a cached symbol with an output notification of IJSurfaceArea (cannot
    '	output IJGeometry / IJDAttributes because this is being output by
    '	the parent custom assembly, which causes a problem with Assoc where two
	'	semantics are outputting the same interfaces for the same objects.
    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    <OutputNotification("{6671BDD8-A919-11D4-B2A2-00104BCC2DC1}")> _
    Public Class CustomAsmComponent : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"

        Private Const CONST_IJGenericVolume As String = "IJGenericVolume"
        Private Const CONST_PropertyVolume As String = "Volume"

        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Component of snowman")> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblRadius As Double
            dblRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objSphere As Sphere3d

            objSphere = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_physicalAspect.Outputs(CONST_Component) = objSphere
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            Dim oBusinessObject As BusinessObject = Me.Occurrence

            '*********************************
            ' Compute/Set Volume
            '*********************************
            Dim dblRadius As Double
            dblRadius = m_dblDiameter.Value / 2.0

            Dim dVolume As Double = PI * dblRadius * dblRadius * dblRadius
            Dim oVolume As PropertyValueDouble = New PropertyValueDouble(CONST_IJGenericVolume, CONST_PropertyVolume, dVolume)
            oBusinessObject.SetPropertyValue(oVolume)
        End Sub

    End Class

End Namespace