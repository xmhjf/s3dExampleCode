Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    ' Declare:
    '   non-cached symbol
    '   symbol version
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class DefaultDotNetTestSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble
        <InputObject(4, "StickAngleReferenceObject", "Reference object to set an angle for the stick", True)> _
        Public m_dStickAngleReferenceObject As InputObject

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class ValAddInputDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble
        <InputDouble(4, "NewInput", "This is a new input", 33.0)> _
        Public m_dNewInput As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class ValModifyDefaultValueDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 33.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class ValAddNewReprDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 33.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        <Aspect("Insulation", "Insulation Aspect", AspectID.Insulation)> _
        <SymbolOutput(CONST_Component, "Candy Ball Insulation")> _
        Public m_oInsulationAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class ValChgReqToOptDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 1.0, True)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 33.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class ValChgCacheToNonCacheDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 33.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class

    Public Class InvChngInputNameDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameterInvalid", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class InvChngReprNameDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameterInvalid", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("PhysicalInvalid", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class InvChngRemInputDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameterInvalid", "How big the candy part is", 1.0)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("PhysicalInvalid", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class
    Public Class InvChngRemReprDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 0.1)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class

    Public Class InvChngOptToReqDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 0.1)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class

    Public Class InvChngInputTypeDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 0.1)> _
        Public m_dBallDiameter As InputDouble
        <InputString(2, "StickDiameter", "How wide the stick part is", "String", True)> _
        Public m_dStickDiameter As InputString
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition


        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class

    Public Class DuplicateOutputsNameDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 0.1)> _
        Public m_dBallDiameter As InputDouble
        <InputString(2, "StickDiameter", "How wide the stick part is", "String", True)> _
        Public m_dStickDiameter As InputString
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition


        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall

            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)
            m_oPhysicalAspect.Outputs(CONST_Component) = oBall
        End Sub
    End Class


    Public Class MsgToDoListDotNetSymbol : Inherits CustomSymbolDefinition
        <InputDouble(1, "BallDiameter", "How big the candy part is", 0.1)> _
        Public m_dBallDiameter As InputDouble
        <InputDouble(2, "StickDiameter", "How wide the stick part is", 0.5, True)> _
        Public m_dStickDiameter As InputDouble
        <InputDouble(3, "StickLength", "How long the stick part is", 3.0)> _
        Public m_dStickLength As InputDouble
        <InputObject(4, "StickAngleReferenceObject", "Reference object to set an angle for the stick", True)> _
        Public m_dStickAngleReferenceObject As InputObject

        'Physical Aspect
        Private Const CONST_Component As String = "Component"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Component, "Candy Ball")> _
        Public m_oPhysicalAspect As AspectDefinition

        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(CONST_Component, "Candy Ball Maintenance")> _
        Public m_oMaintenanceAspect As AspectDefinition

        Protected Overrides Sub ConstructOutputs()
            Dim oBall As Sphere3d
            Dim dblRadius As Double
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            dblRadius = m_dBallDiameter.Value / 2.0
            oBall = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, 0.0), dblRadius, True)

            m_oPhysicalAspect.Outputs(CONST_Component) = oBall



            Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetTestSymbol", 42, "Test error msg")

        End Sub
    End Class

End Namespace