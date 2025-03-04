Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmWithBothSymbolAndMemberOutputs : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

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


            ' Check that we recognize that we have assembly outputs
            If Not HasAssemblyOutputs Then
                Throw New CmnException("HasAssemblyOutputs returned 'false' when there are assembly outputs.")
            End If

            Dim oPropVal As PropertyValueDouble
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)
            Dim dblBaseRadius As Double
            dblBaseRadius = oPropVal.PropValue().Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Define the origin of the body
            Dim origin As New Position(0.0, 0.0, dblBaseRadius * 2.0 + dblBodyRadius)

            If m_objBody.Output Is Nothing Then
                '=================================================
                ' Construct body
                '=================================================
                m_objBody.Output = New Sphere3d(Me.OccurrenceConnection, origin, dblBodyRadius, True)
            End If

            ' Define the origin of the head
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            If m_objHead.Output Is Nothing Then
                '=================================================
                ' Construct head
                '=================================================
                m_objHead.Output = New Sphere3d(Me.OccurrenceConnection, origin, dblHeadRadius, True)
            End If
        End Sub

        Public Overrides Function IsPropertyReadOnly(ByVal assemblyOutputName As String, ByVal interfaceName As String, ByVal propertyName As String) As Boolean
            IsPropertyReadOnly = False
            If assemblyOutputName = m_objBody.Name Then
                If interfaceName = CONST_IJUATestDotNetSphere Then
                    If propertyName = CONST_PropertyDiameter Then
                        IsPropertyReadOnly = True
                    End If
                End If
            End If
        End Function

        Public Overrides Function IsPropertyValid(ByVal assemblyOutputName As String, ByVal interfaceName As String, ByVal propertyName As String, ByVal propertyValue As Object, ByRef errorMessage As String) As Boolean
            IsPropertyValid = True
            If assemblyOutputName = m_objBody.Name Then
                If interfaceName = CONST_IJUATestDotNetSphere Then
                    If propertyName = CONST_PropertyDiameter Then
                        If Not TypeOf (propertyValue) Is Double Then
                            Throw New CmnException("Bad property type")
                        End If
                        Dim oDiameterProp As Double = DirectCast(propertyValue, Double)
                        If oDiameterProp < 0.0 Then
                            IsPropertyValid = False
                            errorMessage = "Diameter must be > 0"
                        ElseIf oDiameterProp > 100.0 Then
                            IsPropertyValid = False
                            ' Intentionally don't set the error message which should trigger an exception
                        End If
                    End If
                End If
            End If
        End Function
    End Class
End Namespace

