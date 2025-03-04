Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithOnlyCachedSymbolOutputs : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"

        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
            Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        <SymbolOutput(CONST_Body, "Body of snowman")> _
        <SymbolOutput(CONST_Head, "Head of snowman")> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase

            '=================================================
            ' Construct body
            '=================================================
            Dim objBody As Sphere3d

            objBody = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius * 2.0 + dblBodyRadius), dblBodyRadius, True)

            m_physicalAspect.Outputs(CONST_Body) = objBody

            '=================================================
            ' Construct head
            '=================================================
            Dim objHead As Sphere3d

            objHead = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, (dblBaseRadius + dblBodyRadius) * 2.0 + dblHeadRadius), dblHeadRadius, True)

            m_physicalAspect.Outputs(CONST_Head) = objHead

        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()

            ' Check that we do not have assembly outputs
            If HasAssemblyOutputs Then
                Throw New CmnException("HasAssemblyOutputs returned 'true' when there were no outputs.")
            End If
        End Sub

    End Class

End Namespace