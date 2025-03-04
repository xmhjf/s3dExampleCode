''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright (C) 2010, Intergraph PPO. All rights reserved.
'
'File
'  cHgrBeam.vb
'
'Abstract
'	This is a sample Hanger Beam symbol.
'
'   History
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Who        Date         Description
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Prasanna   22/11/2011   TR 206236  .Net HgrBeams as not showing the graphics properly
'   Ramya      18/12/2012   TR 220360  3DAPI Sample Assembly fails to place  
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit On
Imports System.Math
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Exceptions
Imports Ingr.SP3D.Structure.Middle.Services
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.Support.Middle

'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Support.Content.Symbols
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------
Public Class HgrBeam : Inherits CustomSymbolDefinition
    Implements ICustomHgrWeightCG
    '----------------------------------------------------------------------------------
    'DefinitionName/ProgID of this symbol is "Structural_Sample,Ingr.SP3D.Support.Content.Symbols.HgrBeam"
    '----------------------------------------------------------------------------------
    Private Const DENSITY As Double = 7849  'kg/m^3
    Private dArea As Double

#Region "Definition of Inputs"

    <InputCatalogPart(1)> Public m_oPartInput As InputCatalogPart
    <InputDouble(2, "CardinalPoint", "Cardinal Point", 1)> Public m_iCP As InputDouble
    <InputDouble(3, "Length", "Length", 0.5)> Public m_dLength As InputDouble
    <InputString(4, "MaterialType", "Material Type", "Steel - Carbon")> Public m_sMaterialType As InputString
    <InputString(5, "MaterialGrade", "Material Grade", "A36")> Public m_sMaterialGrade As InputString
    <InputDouble(6, "Orientation", "Orientation", 0)> Public m_dOrientation As InputDouble
    <InputDouble(7, "BeginOverLength", "Begin Over Length", 0)> Public m_dBeginOverLength As InputDouble
    <InputDouble(8, "EndOverLength", "End Over Length", 0)> Public m_dEndOverLength As InputDouble
    <InputDouble(9, "BeginMiter", "Begin Miter", 0)> Public m_iBeginMiter As InputDouble
    <InputDouble(10, "EndMiter", "End Miter", 0)> Public m_iEndMiter As InputDouble
#End Region

#Region "Definitions of Aspects and their outputs"

    'Physical Aspect
    <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
    <SymbolOutput("BeginCap", "Begin Cap")> _
    <SymbolOutput("Neutral", "Neutral")> _
    <SymbolOutput("EndCap", "End Cap")> _
    <SymbolOutput("Section", "Section")> _
    <SymbolOutput("Section1", "Section1")> _
    <SymbolOutput("Section2", "Section2")> _
    <SymbolOutput("Section3", "Section3")> _
    <SymbolOutput("Section4", "Section4")> _
    <SymbolOutput("Section5", "Section5")> _
    <SymbolOutput("Section6", "Section6")> _
    <SymbolOutput("Section7", "Section7")> _
    Public m_oPhysicalAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

    Protected Overrides Sub ConstructOutputs()

        Try
            Dim oSectionPart As Part = Nothing, oConnection As SP3DConnection
            Dim dLength As Double, iCP As Integer, sMaterialType As String, sMaterialGrade As String
            Dim dBeginOverLength As Double, dEndOverLength As Double, dOrientation As Double
            Dim iBeginMiter As Integer, iEndMiter As Integer

            Dim oWarningColl As New Collection(Of SymbolWarningException)

            ' Get Input values
            Try
                oSectionPart = m_oPartInput.Value
                iCP = m_iCP.Value
                dLength = m_dLength.Value
                sMaterialType = m_sMaterialType.Value
                sMaterialGrade = m_sMaterialGrade.Value
                dOrientation = m_dOrientation.Value
                dBeginOverLength = m_dBeginOverLength.Value
                dEndOverLength = m_dEndOverLength.Value
                iBeginMiter = m_iBeginMiter.Value
                iEndMiter = m_iEndMiter.Value
            Catch oEx As Exception
                If (TypeOf oEx Is SymbolWarningException) Then
                    oWarningColl.Add(oEx)
                Else
                    Throw
                End If
            End Try

            oConnection = OccurrenceConnection ' Get the connection where outputs will be created.
            '=================================================
            ' Construction of Physical Aspect 
            '=================================================
            'Initialize SymbolGeometryHelper. Set the active position and orientation 
            Dim oSymbolGeomHlpr As New SymbolGeometryHelper()
            oSymbolGeomHlpr.ActivePosition = New Position(0, 0, 0)
            oSymbolGeomHlpr.SetOrientation(New Vector(0, 0, 1), New Vector(1, 0, 0))

            'Get the Cross Section from the Relationship between the HgrPart and the CrossSection
            Dim oHgrRelation As RelationCollection
            Dim oCrossSection As CrossSection
            Dim oCrossSectionServices As New CrossSectionServices()

            oHgrRelation = oSectionPart.GetRelationship("HgrCrossSection", "CrossSection")
            oCrossSection = oHgrRelation.TargetObjects.First()

            'Get Required Properties From Cross Section
            dArea = CType(oCrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Area"), PropertyValueDouble).PropValue

            'Add the SweepOption to CreateCaps and BreakCrossSection
            Dim eSweepOptions As SweepOptions
            eSweepOptions = SweepOptions.CreateCaps Or SweepOptions.BreakCrossSection

            'Create the graphical surfaces by projecting the cross section
            Dim oSurfaceColl As Collection(Of ISurface)
            oSurfaceColl = oCrossSectionServices.GetProjectionSurfacesFromCrossSection(oConnection, oCrossSection, New Line3d(New Position(0, 0, -dBeginOverLength), New Position(0, 0, dLength + dEndOverLength)), iCP, False, dOrientation, New Position(0, 0, -dBeginOverLength), New Vector(0, 0, 1), New Position(0, 0, dLength + dEndOverLength), New Vector(0, 0, 1), eSweepOptions)

            'Create Ports
            Dim oBeginCap = New Port(oConnection, oSectionPart, "BeginCap", New Position(0, 0, 0), New Vector(1, 0, 0), New Vector(0, 0, 1))
            m_oPhysicalAspect.Outputs("BeginCap") = oBeginCap

            Dim oNeutral = New Port(oConnection, oSectionPart, "Neutral", New Position(0, 0, dLength / 2), New Vector(1, 0, 0), New Vector(0, 0, 1))
            m_oPhysicalAspect.Outputs("Neutral") = oNeutral

            Dim oEndCap = New Port(oConnection, oSectionPart, "EndCap", New Position(0, 0, dLength), New Vector(1, 0, 0), New Vector(0, 0, 1))
            m_oPhysicalAspect.Outputs("EndCap") = oEndCap
            Dim iCount As Integer = 0
            Dim sOutputSection As String
            sOutputSection = "Section"

            For Each oSurface In oSurfaceColl
                m_oPhysicalAspect.Outputs(sOutputSection) = oSurface
                iCount = iCount + 1
                sOutputSection = "Section" & (iCount)
            Next
            'oSymbolGeomHlpr.CreateBox(Nothing, dLength, 0.5, 0.5)
            'oSymbolGeomHlpr.MoveAlongLine(New Position(0, 0, 0), New Position(0, 0, dLength), dLength)

        Catch Ex As Exception 'General Unhandled exception 
            Throw
        End Try

    End Sub

#End Region

#Region "ICustomHgrWeightCG Methods"

    Public Sub WeightCG(ByVal oSupportComponent As SupportComponent, ByRef Weight As Double, ByRef CogX As Double, ByRef CogY As Double, ByRef CogZ As Double) Implements ICustomHgrWeightCG.WeightCG

        Dim dLength = DirectCast(oSupportComponent.GetPropertyValue("IJUAHgrOccLength", "Length"), PropertyValueDouble).PropValue
        If (dLength Is Nothing) Then
            dLength = 0.5  ' All stretchable component's initial length will is 0. so set to default (0.5)
        End If
        Weight = dArea * dLength * DENSITY
        CogX = 0
        CogY = 0
        CogZ = 0

    End Sub

#End Region

End Class