Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    '
    ' This class is the same class as in CustomAssemblyRedefineMembers with the
    '   exception of a new output was added. This is to verify that a ToDo record
    '   is created but the outputs are still generated.

    ' Declare:
    '   non-cached symbol
    '   symbol version
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAssemblyRedefineMembers : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"
        Private Const CONST_Hat As String = "Hat"

        Private Const CONST_HatHeight As Double = 0.1

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_IJGenericVolume As String = "IJGenericVolume"
        Private Const CONST_PropertyVolume As String = "Volume"
        Private Const CONST_PropertyDiameter As String = "Diameter"
        Private Const CONST_PropertyOriginX As String = "OriginX"
        Private Const CONST_PropertyOriginY As String = "OriginY"
        Private Const CONST_PropertyOriginZ As String = "OriginZ"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        <AssemblyOutput(3, CONST_Hat)> _
        Public m_objHat As AssemblyOutput

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

            Dim origin As New Position(0.0, 0.0, 0.0)
            Dim oPropVal As PropertyValueDouble
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginX), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.X = oPropVal.PropValue().Value
            End If
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginY), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.Y = oPropVal.PropValue().Value
            End If
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginZ), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.Z = oPropVal.PropValue().Value
            End If

            ' Test updating the diameter property in the Evaluate
            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Define origin of body
            origin.Z = origin.Z + dblBaseRadius * 2.0 + dblBodyRadius

            If m_objBody.Output Is Nothing Then
                '********************************
                ' Create body
                '********************************
                m_objBody.Output = New Sphere3d(Me.OccurrenceConnection, origin, dblBodyRadius, True)
            Else
                '********************************
                ' Update body
                '********************************
                Dim oBody As Sphere3d
                oBody = DirectCast(m_objBody.Output, Sphere3d)
                oBody.Radius = dblBodyRadius
                oBody.Center() = origin
            End If

            ' Define origin of head
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            If m_objHead.Output Is Nothing Then
                '********************************
                ' Create head
                '********************************
                m_objHead.Output = New Sphere3d(Me.OccurrenceConnection, origin, dblBodyRadius, True)
            Else
                '********************************
                ' Update head
                '********************************
                Dim oHead As Sphere3d
                oHead = DirectCast(m_objHead.Output, Sphere3d)
                oHead.Radius = dblHeadRadius
                oHead.Center() = origin
            End If

            ' Define origin of hat (or cone in this case)
            origin.Z = origin.Z + dblHeadRadius

            If m_objHat.Output Is Nothing Then
                '********************************
                ' Create hat
                '********************************
                m_objHat.Output = New Cone3d(Me.OccurrenceConnection, _
                                  New Position(origin.X, origin.Y, origin.Z - (dblHeadRadius * 0.1)), _
                                  New Position(origin.X, origin.Y, origin.Z + CONST_HatHeight), _
                                  New Position(dblHeadRadius + origin.X, origin.Y, origin.Z - (dblHeadRadius * 0.1)), _
                                  New Position(dblHeadRadius + origin.X, origin.Y, origin.Z + CONST_HatHeight), True)
            Else
                '********************************
                ' Update hat
                '********************************
                Dim oHat As Cone3d
                oHat = DirectCast(m_objHat.Output, Cone3d)
                oHat.DefineBy4Pts(New Position(origin.X, origin.Y, origin.Z - (dblHeadRadius * 0.1)), _
                                  New Position(origin.X, origin.Y, origin.Z + CONST_HatHeight), _
                                  New Position(dblHeadRadius + origin.X, origin.Y, origin.Z - (dblHeadRadius * 0.1)), _
                                  New Position(dblHeadRadius + origin.X, origin.Y, origin.Z + CONST_HatHeight), True)
            End If
        End Sub
    End Class
End Namespace
