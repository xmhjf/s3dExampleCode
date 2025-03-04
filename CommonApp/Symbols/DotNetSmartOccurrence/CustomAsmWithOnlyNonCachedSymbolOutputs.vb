Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithOnlyNonCachedSymbolOutputs : Inherits CustomAssemblyDefinition
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
        End Sub

        Public Overrides Function IsPropertyReadOnly(ByVal assemblyOutputName As String, ByVal interfaceName As String, ByVal propertyName As String) As Boolean
            IsPropertyReadOnly = True
        End Function

        Public Overrides Function IsPropertyValid(ByVal assemblyOutputName As String, ByVal interfaceName As String, ByVal propertyName As String, ByVal propertyValue As Object, ByRef errorMessage As String) As Boolean
            IsPropertyValid = False
        End Function

        Public Overrides Sub OnPreLoad(ByVal businessObject As Common.Middle.BusinessObject, ByVal colAllDisplayedValues As System.Collections.ObjectModel.ReadOnlyCollection(Of Common.Middle.Services.PropertyDescriptor))
            ' Make certain that we can access the symbol inputs
            If (m_dblDiameter.Value < 0.0) Then
                Throw New CmnException("Did not expect diameter to be less than 0.")
            End If
        End Sub

        Public Overrides Function OnPropertyChange(ByVal businessObject As Common.Middle.BusinessObject, ByVal colAllDisplayedValues As System.Collections.ObjectModel.ReadOnlyCollection(Of Common.Middle.Services.PropertyDescriptor), ByVal propertyToChange As Common.Middle.Services.PropertyDescriptor, ByVal newPropertyValue As Common.Middle.PropertyValue, ByRef errorMessage As String) As Boolean
            ' Because we are overriding this method no properties should be read-only
            ' (i.e., the IsPropertyReadOnly will not get called)
            errorMessage = ""
            OnPropertyChange = True
            ' Make certain that we can access the symbol inputs
            If (m_dblDiameter.Value < 0.0) Then
                Throw New CmnException("Did not expect diameter to be less than 0.")
            End If
        End Function
    End Class

End Namespace
