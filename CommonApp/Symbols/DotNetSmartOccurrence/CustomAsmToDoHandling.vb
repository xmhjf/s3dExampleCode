Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Exceptions
Imports Ingr.SP3D.Common.Middle.Services.Hidden
Imports SP3DPIA.SmartOccurrence
'********************************************************************************
'
'   ATP to test the ToDo handling of the Custom Assembly to verify that
'   the custom assembly developer can post their provided codelisted
'   message to the ToDo list.
'
'   For this test:
'       Diameter value of 1.1, 1.2, 1.3, 1.4, 1.5, 5.1, 5.2, 6.6 - Fails in the construct
'       Diameter value of 4.1, 4.2, 4.3, 4.4, 4.5, 5.1, 5.2, 5.3, 6.4 - Fails the EvaluateAssembly
'       Diameter value of 6.1, 6.3, 6.5, 6.6, 5.7 - Fails the PreConstructOutputs
'
'********************************************************************************
Namespace Symbols
    Module testGlobals
        Public _failedCount As Integer
    End Module

    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmToDoHandling : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_ErrorCodeList As String = "DotNetCustomAsmErrorCodeList"
        Private Const CONST_FailedConstructOutputs As Integer = 1
        Private Const CONST_FailedEvaluateAssembly As Integer = 2
        Private Const CONST_WarningConstructOutputs As Integer = 3
        Private Const CONST_WarningEvaluateAssembly As Integer = 4

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_PropertyDiameter As String = "Diameter"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            ' Get the connection
            Dim oSP3DConnection As SP3DConnection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase

            ' Diameter equals 1.0 throw an error
            If Abs(m_dblDiameter.Value - 1.1) < Math3d.DistanceTolerance Then
                Throw New SymbolErrorException(CONST_ErrorCodeList, CONST_FailedConstructOutputs, Occurrence)
            ElseIf Abs(m_dblDiameter.Value - 1.2) < Math3d.DistanceTolerance Then
                Throw New SymbolWarningException(CONST_ErrorCodeList, CONST_WarningConstructOutputs, Occurrence)
            ElseIf Abs(m_dblDiameter.Value - 1.3) < Math3d.DistanceTolerance Then
                ' Diameter equals 5.1 throw an unknown exception
                Throw New CmnException("Failed ConstructOutputs")
            ElseIf Abs(m_dblDiameter.Value - 1.4) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "DotNetSmartOccurrence", 1400, "TDR:Error from ConstructOutputs. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 1.5) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetSmartOccurrence", 1500, "TDR: Warning from ConstructOutputs. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 5.1) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetSmartOccurrence", 1500, "TDR: Warning from ConstructOutputs. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 5.2) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetSmartOccurrence", 1500, "TDR: Warning from ConstructOutputs. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 6.6) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "DotNetSmartOccurrence", 1400, "TDR:Error from ConstructOutputs. (MiddleTier CustomAssemblyATPs)")
            End If
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            Dim oBusinessObject As BusinessObject
            oBusinessObject = Me.Occurrence
            If oBusinessObject Is Nothing Then
                Throw New CmnException("Occurrence property is null!")
            End If

            Dim oConnection As SP3DConnection
            oConnection = oBusinessObject.DBConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Set where the origin of the body should be
            Dim origin As New Position(0.0, 0.0, dblBaseRadius * 2.0 + dblBodyRadius)

            If m_objBody.Output Is Nothing Then
                '**************************
                ' Create the body
                '**************************
                m_objBody.Output = New Sphere3d(oConnection, origin, dblBodyRadius, True)
            End If

            ' Set where the origin of the body should be
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            If m_objHead.Output Is Nothing Then
                '**************************
                ' Create the head
                '**************************
                m_objHead.Output = New Sphere3d(oConnection, origin, dblHeadRadius, True)
            End If

            ' Diameter equals 4.0 throw an error
            If Abs(m_dblDiameter.Value - 4.1) < Math3d.DistanceTolerance Then
                Throw New SymbolErrorException(CONST_ErrorCodeList, CONST_FailedEvaluateAssembly, Occurrence)
            ElseIf Abs(m_dblDiameter.Value - 4.2) < Math3d.DistanceTolerance Then
                Throw New SymbolWarningException(CONST_ErrorCodeList, CONST_WarningEvaluateAssembly, Occurrence)
            ElseIf Abs(m_dblDiameter.Value - 4.3) < Math3d.DistanceTolerance Then
                ' Diameter equals 4.1 throw an unknown exception
                Throw New CmnException("Failed EvaluateAssembly")
            ElseIf Abs(m_dblDiameter.Value - 4.4) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "DotNetSmartOccurrence", 4400, "TDR:Error from EvaluateAssembly. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 4.5) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetSmartOccurrence", 4500, "TDR: Warning from EvaluateAssembly. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 5.1) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "DotNetSmartOccurrence", 4400, "TDR:Error from EvaluateAssembly. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 5.2) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetSmartOccurrence", 4500, "TDR: Warning from EvaluateAssembly. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 5.3) < Math3d.DistanceTolerance Then
                ' This test scenario is to return a error from the assembly evaluate then force a re-compute so the PreConstruct is called again while
                '   there is a ToDo record attached to the .Net business object, This ToDo error should be cleared so the PreConstruct can be called and not
                '   assume there's a failure.
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "DotNetSmartOccurrence", 4400, "TDR: Error from EvaluateAssembly. Testing that the PreConstruct outputs clears this error on a graph re-catch.")
                testGlobals._failedCount = testGlobals._failedCount + 1
				' Repeat this failure several times so the semantic believes it is caught in a failure loop. The semantic
				'	will indicate to RevisionMgr to stop the graph re-catch. But the semantic should still be triggered by
				'	this change of marking the IJSmartOccurrence itnerface modified.
                If testGlobals._failedCount > 11 Then
                    testGlobals._failedCount = 0 ' Reset the count so if the ATP is re-run then the test can be repeated
                    ' Explicitly reset the diameter so it doesn't loop back through here again
                    Dim diameter As PropertyValueDouble = DirectCast(Occurrence.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)
                    diameter.PropValue = 5.35
                    Occurrence.SetPropertyValue(diameter)
                End If
                ForceSmartOccurrenceRecompute()
            Else
                Dim diameter As PropertyValueDouble = DirectCast(Occurrence.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)
                If Abs(diameter.PropValue.Value - 6.4) < Math3d.DistanceTolerance Then
                    ' value of 6.4 indicates that a ToDo warning should be logged but the PreConstructOutputs should have already provided a warning
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetSmartOccurrence", 4500, "TDR: Warning from EvaluateAssembly. (MiddleTier CustomAssemblyATPs)")
                End If
            End If
        End Sub

        ''' <summary>
        ''' Method added to CustomAssembly to test overriding PreConstructOutputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub PreConstructOutputs()
            ' To verify this method there are several use cases to consider:
            ' 1) Verify that this method is called which means when the client sets a certain value for the diameter
            '        then this method will change it to something else
            ' 2) verify this method can file a ToDo Warning
            ' 3) verify this method can file a ToDO Error
            ' 4) verify this method can file a ToWarning but have it overriden by a ToDo error in the Evaluate
            ' 5) exception produces an appropriate ToDo error emssage

            Dim diameter As PropertyValueDouble = DirectCast(Occurrence.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)

            If Not diameter.PropValue Is Nothing Then
                ' 6.1 menas that this method should change the value to something else (in this case 6.2)
                If Abs(diameter.PropValue.Value - 6.1) < Math3d.DistanceTolerance Then
                    diameter.PropValue = 6.2
                    Occurrence.SetPropertyValue(diameter)
                ElseIf Abs(diameter.PropValue.Value - 6.3) < Math3d.DistanceTolerance Then
                    ' value of 6.3 indicates that a ToDo warning should be logged
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "PreConstructOutputs warning")
                ElseIf Abs(diameter.PropValue.Value - 6.5) < Math3d.DistanceTolerance Then
                    ' value of 6.5 indicates that a ToDo error should be logged
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "PreConstructOutputs error")
                ElseIf Abs(diameter.PropValue.Value - 6.6) < Math3d.DistanceTolerance Then
                    ' value of 6.5 indicates a ToWarning overriden by a ToDo error in the Evaluate
                ElseIf Abs(diameter.PropValue.Value - 6.7) < Math3d.DistanceTolerance Then
                    ' value of 6.7 indicates an exception
                    Throw New CmnException("diameter value of 6.7 is bad!")
                End If
            End If
        End Sub

        Private Sub ForceSmartOccurrenceRecompute()
            Dim SmartOccurrence As IJSmartOccurrence = DirectCast(COMConverters.ConvertBOToCOMBO(Occurrence), IJSmartOccurrence)

            SmartOccurrence.Update()
        End Sub
    End Class

    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    <OutputNotification("IJDAttributes")> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmToDoHandlingForCachedSymbol : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_ErrorCodeList As String = "DotNetCustomAsmErrorCodeList"
        Private Const CONST_FailedConstructOutputs As Integer = 1
        Private Const CONST_WarningConstructOutputs As Integer = 3

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_PropertyDiameter As String = "Diameter"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            ' Get the connection
            Dim oSP3DConnection As SP3DConnection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase

            ' Diameter equals 1.2 create a warning ToDo record
            If Abs(m_dblDiameter.Value - 1.2) < Math3d.DistanceTolerance Then
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "DotNetSmartOccurrence", 1501, "TDR: Warning from ConstructOutputs. (MiddleTier CustomAssemblyATPs)")
            ElseIf Abs(m_dblDiameter.Value - 2.5) < Math3d.DistanceTolerance Then ' Diameter equals 2.5 create a warning ToDo record
                Try
                    Dim oBusinessObject As BusinessObject
                    oBusinessObject = Me.Occurrence
                Catch e As Ingr.SP3D.Common.Exceptions.OccurrenceNotAvailableException
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "DotNetSmartOccurrence", 1502, "TDR: error from ConstructOutputs." + e.Message.ToString)
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            Dim oBusinessObject As BusinessObject
            oBusinessObject = Me.Occurrence
            If oBusinessObject Is Nothing Then
                Throw New CmnException("Occurrence property is null!")
            End If

            Dim oConnection As SP3DConnection
            oConnection = oBusinessObject.DBConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Set where the origin of the body should be
            Dim origin As New Position(0.0, 0.0, dblBaseRadius * 2.0 + dblBodyRadius)

            If m_objBody.Output Is Nothing Then
                '**************************
                ' Create the body
                '**************************
                m_objBody.Output = New Sphere3d(oConnection, origin, dblBodyRadius, True)
            End If

            ' Set where the origin of the body should be
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            If m_objHead.Output Is Nothing Then
                '**************************
                ' Create the head
                '**************************
                m_objHead.Output = New Sphere3d(oConnection, origin, dblHeadRadius, True)
            End If

        End Sub
    End Class

    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    <OutputNotification("IJDAttributes")> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmToDoHandlingForBadCachedSymbol : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_ErrorCodeList As String = "DotNetCustomAsmErrorCodeList"
        Private Const CONST_FailedConstructOutputs As Integer = 1
        Private Const CONST_WarningConstructOutputs As Integer = 3

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_PropertyDiameter As String = "Diameter"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Throw New CmnException("ConstructOutputs thows exception!")
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
        End Sub
    End Class

    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    <OutputNotification("IJDAttributes")> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmToDoHandlingForBadNonCachedSymbol : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_ErrorCodeList As String = "DotNetCustomAsmErrorCodeList"
        Private Const CONST_FailedConstructOutputs As Integer = 1
        Private Const CONST_WarningConstructOutputs As Integer = 3

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_PropertyDiameter As String = "Diameter"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Throw New CmnException("ConstructOutputs thows exception!")
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
        End Sub
    End Class
End Namespace

