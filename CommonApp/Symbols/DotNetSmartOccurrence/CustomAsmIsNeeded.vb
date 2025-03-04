Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    ' Declare:
    '   IJGenericVolume interface declared as output
    '   IJSurfaceArea interface declared as output
    '   IJUATestDotNetSphere interface declared as output
    '   IJDGeometry interface declared as output
    <OutputNotification("{E897A3AF-E949-457C-968D-67E5DFDAD154}")> _
    <OutputNotification("{2A7D37F3-0440-4ACF-A731-90CE0613ABA0}")> _
    <OutputNotification("{5201D970-8E75-4245-BB3B-D932FA9CD534}")> _
    <OutputNotification("IJDGeometry")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public MustInherit Class CustomAsmIsNeededBaseClass : Inherits CustomAssemblyDefinition
    End Class

    ' Declare:
    '   non-cached symbol
    '   symbol version
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmIsNeeded : Inherits CustomAsmIsNeededBaseClass
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
        Private Const CONST_HatBrim As String = "HatBrim"
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

        <AssemblyOutput(4, CONST_HatBrim)> _
        Public m_objHatBrim As AssemblyOutput

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

            Dim oConnection As SP3DConnection
            oConnection = oBusinessObject.DBConnection

            Dim oPropVal As PropertyValueDouble
            Dim origin As New Position(0.0, 0.0, 0.0)
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginX), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.X = oPropVal.PropValue().Value
            End If
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginY), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                origin.Y = oPropVal.PropValue().Value
            End If
            Dim oZPropVal As PropertyValueDouble
            oZPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginZ), PropertyValueDouble)
            If oZPropVal.PropValue().HasValue Then
                origin.Z = oZPropVal.PropValue().Value
            End If

            ' Test updating the diameter property in the Evaluate
            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Set the position of the body relative to the base
            origin.Z = origin.Z + dblBaseRadius * 2.0 + dblBodyRadius

            If m_objBody.Output Is Nothing Then
                '********************************
                ' Construct body
                '********************************
                m_objBody.Output = New Sphere3d(oConnection, origin, dblBodyRadius, True)
            Else
                '********************************
                ' Update body
                '********************************
                Dim oBody As Sphere3d
                oBody = DirectCast(m_objBody.Output, Sphere3d)
                oBody.Radius = dblBodyRadius
                oBody.Center() = origin
            End If

            ' Set the position of the body relative to the base
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            If m_objHead.Output Is Nothing Then
                '********************************
                ' Construct body
                '********************************
                m_objHead.Output = New Sphere3d(oConnection, origin, dblHeadRadius, True)
            Else
                '********************************
                ' Update head
                '********************************
                Dim oHead As Sphere3d
                oHead = DirectCast(m_objHead.Output, Sphere3d)
                oHead.Radius = dblHeadRadius
                oHead.Center() = origin
            End If

            '*****************************************
            ' If the elevation is above 20 meters then
            '   we need a hat (with brim)
            '*****************************************
            Dim bNeedHat As Boolean
            bNeedHat = False
            If oZPropVal.PropValue.HasValue Then
                If oZPropVal.PropValue.Value >= 20.0 Then
                    bNeedHat = True
                End If
            End If

            If bNeedHat Then
                If m_objHat.Output Is Nothing Then
                    '*********************************
                    ' Create hat
                    '*********************************
                    m_objHat.Output = New Cone3d(oConnection, _
                                      New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 + CONST_HatHeight), _
                                      New Position(origin.X + dblHeadRadius, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X + dblHeadRadius, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 + CONST_HatHeight), True)
                Else
                    '*********************************
                    ' Update hat
                    '*********************************
                    Dim oHat As Cone3d
                    oHat = DirectCast(m_objHat.Output, Cone3d)
                    oHat.DefineBy4Pts(New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 + CONST_HatHeight), _
                                      New Position(origin.X + dblHeadRadius, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X + dblHeadRadius, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 + CONST_HatHeight), True)
                End If
                If m_objHatBrim.Output Is Nothing Then
                    '*********************************
                    ' Create hat brim
                    '*********************************
                    m_objHatBrim.Output = New Cone3d(oConnection, _
                                      New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1) + 0.001), _
                                      New Position(origin.X + dblHeadRadius + 0.1, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X + dblHeadRadius + 0.1, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1) + 0.001), True)
                Else
                    '*********************************
                    ' Update hat brim 
                    '*********************************
                    Dim oHatBrim As Cone3d
                    oHatBrim = DirectCast(m_objHatBrim.Output, Cone3d)
                    oHatBrim.DefineBy4Pts(New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1) + 0.001), _
                                      New Position(origin.X + dblHeadRadius + 0.1, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1)), _
                                      New Position(origin.X + dblHeadRadius + 0.1, origin.Y, origin.Z + (dblBaseRadius + dblBodyRadius + dblHeadRadius) * 2.0 - (dblHeadRadius * 0.1) + 0.001), True)
                End If
            Else
                ' if hat exists remove it
                If Not m_objHat.Output Is Nothing Then
                    m_objHat.Delete()
                End If
                ' if hat brim exists remove it
                If Not m_objHatBrim.Output Is Nothing Then
                    m_objHatBrim.Delete()
                End If
            End If

            '*********************************
            ' Compute/Set Volume
            '*********************************
            Dim dRadius As Double
            dRadius = m_dblDiameter.Value / 2.0
            Dim dVolume As Double = PI * dRadius * dRadius * dRadius

            dRadius = dblBodyRadius
            dVolume = dVolume + PI * dRadius * dRadius * dRadius

            dRadius = dblHeadRadius
            dVolume = dVolume + PI * dRadius * dRadius * dRadius

            Dim oVolume As PropertyValueDouble = New PropertyValueDouble(CONST_IJGenericVolume, CONST_PropertyVolume, dVolume)
            oBusinessObject.SetPropertyValue(oVolume)

        End Sub
    End Class
End Namespace


