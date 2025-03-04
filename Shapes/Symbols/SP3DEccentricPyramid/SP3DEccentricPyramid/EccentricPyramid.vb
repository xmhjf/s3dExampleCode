'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright (C) 2009, Intergraph PPO. All rights reserved.
'
'File
'  EccentricPyramid.vb
'
'Abstract
'	This is a Eccentric Truncated Pyramid symbol.Using this symbol we can place Truncated pyramids 
'   with some offset at the top plane. This class subclasses from CustomSymbolDefinition.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit On
Imports System
Imports System.Math
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.Common.Exceptions
Imports Ingr.SP3D.ReferenceData.Middle

'---------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Shapes.SP3DEccentricPyramid.
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'DefinitionName/ProgID of this symbol is "SP3DEccentricPyramid,Ingr.SP3D.Shapes.SP3DEccentricPyramid.SP3DEccentricPyramid"
'----------------------------------------------------------------------------------


Public Class EccentricPyramid : Inherits CustomSymbolDefinition


#Region "Definition of Inputs/Outputs"

    <InputCatalogPart(1)> _
    Public m_oPartInput As InputCatalogPart
    <InputDouble(2, "EccentricPyramidBottomLength", "IJUAEccentricPyramid::A::Eccentric Pyramid Bottom Length", 0.08)> _
    Public m_oBottomLength As InputDouble
    <InputDouble(3, "EccentricPyramidBottomWidth", "IJUAEccentricPyramid::B::Eccentric Pyramid Bottom Width", 0.06)> _
    Public m_oBottomWidth As InputDouble
    <InputDouble(4, "EccentricPyramidTopLength", "IJUAEccentricPyramid::C::Eccentric Pyramid Top Length", 0.04)> _
    Public m_oTopLength As InputDouble
    <InputDouble(5, "EccentricPyramidTopWidth", "IJUAEccentricPyramid::D::Eccentric Pyramid Top Width", 0.04)> _
    Public m_oTopWidth As InputDouble
    <InputDouble(6, "EccentricPyramidXOffset", "IJUAEccentricPyramid::E::Eccentric Pyramid X Offset", 0.01)> _
    Public m_oXoffset As InputDouble
    <InputDouble(7, "EccentricPyramidYOffset", "IJUAEccentricPyramid::F::Eccentric Pyramid Y Offset", 0.01)> _
    Public m_oYoffset As InputDouble
    <InputDouble(8, "EccentricPyramidHeight", "IJUAEccentricPyramid::G::Eccentric Pyramid Height", 0.04)> _
    Public m_oPyramidHeight As InputDouble

    ' Physical Aspect 
    Private Const CONST_EccentricPyramidBottomPlane As String = "Bottom Plane"
    Private Const CONST_EccentricPyramidTopPlane As String = "Top Plane"
    Private Const CONST_EccentricPyramidEastPlane As String = "East Plane"
    Private Const CONST_EccentricPyramidWestPlane As String = "West Plane"
    Private Const CONST_EccentricPyramidNorthPlane As String = "North Plane"
    Private Const CONST_EccentricPyramidSouthPlane As String = "South Plane"
    Private Const CONST_EccentricPyramidPoint1 As String = "Bottom Plane Point"
    Private Const CONST_EccentricPyramidPoint2 As String = "Top Plane Point"

    <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
    <SymbolOutput(CONST_EccentricPyramidBottomPlane, "Eccentric Pyramid Bottom Plane")> _
    <SymbolOutput(CONST_EccentricPyramidTopPlane, "Eccentric Pyramid Top Plane")> _
    <SymbolOutput(CONST_EccentricPyramidNorthPlane, "Eccentric Pyramid North Plane")> _
    <SymbolOutput(CONST_EccentricPyramidSouthPlane, "Eccentric Pyramid South Plane")> _
    <SymbolOutput(CONST_EccentricPyramidEastPlane, "Eccentric Pyramid East Plane")> _
    <SymbolOutput(CONST_EccentricPyramidWestPlane, "Eccentric Pyramid West Plane")> _
    <SymbolOutput(CONST_EccentricPyramidPoint1, "Eccentric Pyramid Bottom Plane Point")> _
    <SymbolOutput(CONST_EccentricPyramidPoint2, "Eccentric Pyramid Top Plane Point")> _
    Public m_oPhysicalAspect As AspectDefinition

#End Region

#Region "Symbol Output contruction/update methods"
    ''' <summary>
    ''' Construct symbol outputs/aspects.  
    ''' Geometry created   
    '''    In this we are creating a geometry for Eccentric Truncated Pyramid with Six planes as OutPut.   
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub ConstructOutputs()
        Try

            Dim oEccentriPyramidPart As Part, oSP3DConnection As SP3DConnection = Nothing
            Dim oWarningCollection As New Collection(Of SymbolWarningException)
            Dim dEPBottomLength As Double, dEPBottomWidth As Double
            Dim dEPTopLength As Double, dEPTopWidth As Double
            Dim dEPXOffset As Double, dEPYOffset As Double
            Dim dEPHeight As Double
            Dim oPlane As Plane3d
            Dim oTopPlane As Plane3d
            Dim oPoint3d As Point3d
            Dim oPosition As Position
          

            Dim oPosColl As New Collection(Of Position)
            Dim oPosition1 As Position, oPosition2 As Position, oPosition3 As Position, oPosition4 As Position
            Dim oPosition5 As Position, oPosition6 As Position, oPosition7 As Position, oPosition8 As Position
            'Get Input Values
            Try
                oEccentriPyramidPart = DirectCast(m_oPartInput.Value, Part)
                dEPBottomWidth = m_oBottomWidth.Value
                dEPBottomLength = m_oBottomLength.Value
                dEPTopWidth = m_oTopWidth.Value
                dEPTopLength = m_oTopLength.Value
                dEPYOffset = m_oYoffset.Value
                dEPXOffset = m_oXoffset.Value
                dEPHeight = m_oPyramidHeight.Value
                oSP3DConnection = Me.OccurrenceConnection

                If dEPBottomLength <= 0.0 Or dEPBottomWidth <= 0.0 Or dEPTopLength <= 0.0 Or _
                    dEPTopWidth <= 0.0 Or dEPHeight <= 0.0 Then
                    'Throwing a new SymInvalidInputsException
                    Throw New SymInvalidInputsException("Eccentric Pyramid Dimensions should be greater than zero")
                End If

            Catch ex As Exception
                If (TypeOf ex Is SymbolWarningException) Then
                    oWarningCollection.Add(ex)
                Else
                    Throw ex
                End If
            End Try

            '=================================================
            ' Construction of Physical Aspect
            '=================================================

            oPosition1 = New Position(0.0, dEPBottomLength / 2, -dEPBottomWidth / 2)
            oPosition2 = New Position(0.0, -dEPBottomLength / 2, -dEPBottomWidth / 2)
            oPosition3 = New Position(0.0, -dEPBottomLength / 2, dEPBottomWidth / 2)
            oPosition4 = New Position(0.0, dEPBottomLength / 2, dEPBottomWidth / 2)
            oPosition5 = New Position(dEPHeight, (-dEPXOffset + dEPBottomLength / 2), (dEPYOffset - dEPBottomWidth / 2))
            oPosition6 = New Position(dEPHeight, (-dEPXOffset + dEPBottomLength / 2 - dEPTopLength), (dEPYOffset - dEPBottomWidth / 2))
            oPosition7 = New Position(dEPHeight, (-dEPXOffset + dEPBottomLength / 2 - dEPTopLength), (dEPYOffset - dEPBottomWidth / 2 + dEPTopWidth))
            oPosition8 = New Position(dEPHeight, (-dEPXOffset + dEPBottomLength / 2), (dEPYOffset - dEPBottomWidth / 2 + dEPTopWidth))

            ''Creating output 1:EccentricPyramidBottomPlane
            oPosColl.Clear()
            oPosColl.Add(oPosition1)
            oPosColl.Add(oPosition2)
            oPosColl.Add(oPosition3)
            oPosColl.Add(oPosition4)

            oPlane = New Plane3d(oSP3DConnection, oPosColl)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidBottomPlane) = oPlane

            'Creating output 2:EccentricPyramidTopPlane
            oPosColl.Clear()
            oPlane = Nothing
            oPosColl.Add(oPosition5)
            oPosColl.Add(oPosition6)
            oPosColl.Add(oPosition7)
            oPosColl.Add(oPosition8)

            oPlane = New Plane3d(oSP3DConnection, oPosColl)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidTopPlane) = oPlane
            oTopPlane = oPlane

            'Creating output 3:EccentricPyramidSouthPlane
            oPosColl.Clear()
            oPlane = Nothing
            oPosColl.Add(oPosition1)
            oPosColl.Add(oPosition2)
            oPosColl.Add(oPosition6)
            oPosColl.Add(oPosition5)

            oPlane = New Plane3d(oSP3DConnection, oPosColl)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidSouthPlane) = oPlane

            'Creating output 4:EccentricPyramidEastPlane
            oPosColl.Clear()
            oPlane = Nothing
            oPosColl.Add(oPosition2)
            oPosColl.Add(oPosition3)
            oPosColl.Add(oPosition7)
            oPosColl.Add(oPosition6)

            oPlane = New Plane3d(oSP3DConnection, oPosColl)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidEastPlane) = oPlane

            'Creating output 5:EccentricPyramidNorthPlane
            oPosColl.Clear()
            oPlane = Nothing
            oPosColl.Add(oPosition3)
            oPosColl.Add(oPosition4)
            oPosColl.Add(oPosition8)
            oPosColl.Add(oPosition7)

            oPlane = New Plane3d(oSP3DConnection, oPosColl)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidNorthPlane) = oPlane

            'Creating output 6:EccentricPyramidWestPlane
            oPosColl.Clear()
            oPlane = Nothing
            oPosColl.Add(oPosition4)
            oPosColl.Add(oPosition1)
            oPosColl.Add(oPosition5)
            oPosColl.Add(oPosition8)

            oPlane = New Plane3d(oSP3DConnection, oPosColl)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidWestPlane) = oPlane

            oPosColl = Nothing
            oPlane = Nothing
            oPosition = Nothing

            'Creating output 7:EccentricPyramidPoint1
            oPosition = New Position(0.0, 0.0, 0.0)
            oPoint3d = New Point3d(oSP3DConnection, oPosition)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidPoint1) = oPoint3d
            oPoint3d = Nothing
            oPosition = Nothing


            'Creating output 8:EccentricPyramidPoint2 
            oPosition = New Position((oPosition5.X + oPosition7.X) / 2, (oPosition5.Y + oPosition7.Y) / 2, (oPosition5.Z + oPosition7.Z) / 2)
            oPoint3d = New Point3d(oSP3DConnection, oPosition)
            m_oPhysicalAspect.Outputs(CONST_EccentricPyramidPoint2) = oPoint3d
            oPoint3d = Nothing
            oPosition = Nothing

            If (oWarningCollection.Count > 0) Then
                Throw oWarningCollection.Item(0)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
