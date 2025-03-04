Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithNoObjectInputs : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   5. "Diameter"

        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
            Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base")> _
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

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub

    End Class

    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithLineInput : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Line"
        '   3. "object"
        '   4. "Diameter"

        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
            Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base")> _
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

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub

    End Class

    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithLineInputs : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Line"
        '   3. "Line"
        '   4. "Diameter"

        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
            Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base")> _
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

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub

    End Class

    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithPointInput : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Point"
        '   3. "object"
        '   4. "Diameter"

        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
            Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base")> _
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

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub

    End Class


    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithPlaneInput : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Plane"
        '   3. "object"
        '   4. "Diameter"

        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
            Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base")> _
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

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub

    End Class

End Namespace