' 3 sample symbols. 
' ManyOutputsTestSymbol defines a variable number of outputs and creates each outputs as a separate geometrical object
' CompositeGeometryTestSymbol creates a composite geometry and put each geometrical object into that composite geometry
' TransientOutputTestSymbol creates a variable number of outputs. Eeach outptus is added as a transient object.
'
'
Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Collections.ObjectModel

Namespace Symbols
    ' Symbol with many outputs 
    '   non-cached symbol
    '   symbol version
    '   Variable outputs
    <CacheOption(CacheOptionType.NonCached)> _
    <VariableOutputs()> _
    <SymbolVersion("1.0.0.0")> _
    Public Class ManyOutputsTestSymbol : Inherits CustomSymbolDefinition
        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        Public m_oPhysicalAspect As AspectDefinition
        <Aspect("DetailedPhysical", "Detailed Physical Aspect", AspectID.DetailedPhysical)> _
        Public m_oDetailedPhysicalAspect As AspectDefinition
        <InputDouble(1, "NumberOfPlanes", "How many outputs planes to create", 10.0)> _
        Public m_nbPlane As InputDouble

        Protected Overrides Sub ConstructOutputs()

            Dim nbPlane As Integer

            nbPlane = Convert.ToInt32(m_nbPlane.Value)

            For i As Integer = 1 To nbPlane
                Dim oPointCollection As Collection(Of Position)
                oPointCollection = New Collection(Of Position)

                oPointCollection.Add(New Position(0.0, 0.1 * i, 0.0))
                oPointCollection.Add(New Position(0.1 * i, 0.0, 0.0))
                oPointCollection.Add(New Position(0.1 * i, 0.1, 0.0))
                oPointCollection.Add(New Position(0.1, 0.3 * i, 0.0))

                'Create regular outputs
                Dim oSP3DConnection As SP3DConnection = Me.OccurrenceConnection

                Dim oPlane3d As Plane3d
                oPlane3d = New Plane3d(oSP3DConnection, oPointCollection)

                Dim sOutputName As String
                sOutputName = String.Format("Plane{0}", i)

                'Add output to each aspect
                m_oPhysicalAspect.Outputs.Add(sOutputName, oPlane3d)

                m_oDetailedPhysicalAspect.Outputs.Add(sOutputName, oPlane3d)
                oPlane3d = Nothing
            Next

        End Sub
    End Class



    ' Declare: Symbol with a single composite geometry output
    '   non-cached symbol
    '   symbol version
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CompositeGeometryTestSymbol : Inherits CustomSymbolDefinition
        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput("CompositeGeometry", "Composite Geometry")> _
        Public m_oPhysicalAspect As AspectDefinition
        <Aspect("DetailedPhysical", "Detailed Physical Aspect", AspectID.DetailedPhysical)> _
        <SymbolOutput("CompositeGeometry", "Composite Geometry")> _
        Public m_oDetailedPhysicalAspect As AspectDefinition
        <InputDouble(1, "NumberOfPlanes", "How many outputs planes to create", 10.0)> _
        Public m_nbPlane As InputDouble

        Protected Overrides Sub ConstructOutputs()

            Dim nbPlane As Integer

            nbPlane = Convert.ToInt32(m_nbPlane.Value)

            'Create the composite geometry
            Dim oSP3DConnection As SP3DConnection = Me.OccurrenceConnection

            'Add a single composite geometry to each aspect
            Dim oCompositeGeometry As CompositeGeometry = New CompositeGeometry(oSP3DConnection)
            m_oPhysicalAspect.Outputs.Add("CompositeGeometry", oCompositeGeometry)

            Dim oCompositeGeometry2 As CompositeGeometry = New CompositeGeometry(oSP3DConnection)
            m_oDetailedPhysicalAspect.Outputs.Add("CompositeGeometry", oCompositeGeometry2)


            For i As Integer = 1 To nbPlane
                Dim oPointCollection As Collection(Of Position)
                oPointCollection = New Collection(Of Position)

                oPointCollection.Add(New Position(0.0, 0.1 * i, 0.0))
                oPointCollection.Add(New Position(0.1 * i, 0.0, 0.0))
                oPointCollection.Add(New Position(0.1 * i, 0.1, 0.0))
                oPointCollection.Add(New Position(0.1, 0.3 * i, 0.0))

                'Create a transient plane
                Dim oPlane3d As Plane3d
                oPlane3d = New Plane3d(oPointCollection)

                Dim sOutputName As String
                sOutputName = String.Format("Plane{0}", i)

                'Add to each aspect
                oCompositeGeometry.Add(sOutputName, oPlane3d)

                oCompositeGeometry2.Add(sOutputName, oPlane3d)
                oPlane3d = Nothing
            Next

        End Sub
    End Class


    ' Declare: Symbol with many transient geometrical outputs
    '   non-cached symbol
    '   symbol version
    '   Variable outputs
    <CacheOption(CacheOptionType.NonCached)> _
    <VariableOutputs()> _
    <SymbolVersion("1.0.0.0")> _
    Public Class TransientOutputTestSymbol : Inherits CustomSymbolDefinition
        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        Public m_oPhysicalAspect As AspectDefinition
        <Aspect("DetailedPhysical", "Detailed Physical Aspect", AspectID.DetailedPhysical)> _
        Public m_oDetailedPhysicalAspect As AspectDefinition
        <InputDouble(1, "NumberOfPlanes", "How many outputs planes to create", 10.0)> _
        Public m_nbPlane As InputDouble

        Protected Overrides Sub ConstructOutputs()

            Dim nbPlane As Integer

            nbPlane = Convert.ToInt32(m_nbPlane.Value)

            For i As Integer = 1 To nbPlane
                Dim oPointCollection As Collection(Of Position)
                oPointCollection = New Collection(Of Position)

                oPointCollection.Add(New Position(0.0, 0.1 * i, 0.0))
                oPointCollection.Add(New Position(0.1 * i, 0.0, 0.0))
                oPointCollection.Add(New Position(0.1 * i, 0.1, 0.0))
                oPointCollection.Add(New Position(0.1, 0.3 * i, 0.0))

                'Create a transient plane
                Dim oPlane3d As Plane3d
                oPlane3d = New Plane3d(oPointCollection)

                Dim sOutputName As String
                sOutputName = String.Format("Plane{0}", i)

                'Add to each aspect
                'Add output to each aspect
                m_oPhysicalAspect.Outputs.Add(sOutputName, oPlane3d)

                m_oDetailedPhysicalAspect.Outputs.Add(sOutputName, oPlane3d)
            Next

        End Sub
    End Class
End Namespace