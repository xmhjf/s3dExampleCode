
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright (C) 2009, Intergraph PPO. All rights reserved.
'
'File
'  CBallValve.vb
'
'Abstract
'	This is a Ball Valve symbol. This class subclasses from CustomSymbolDefinition.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit On
Imports System.Math
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Exceptions

'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Content.PipingSample.
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------
Public Class CBallValve : Inherits CustomSymbolDefinition

    Private Const NegligibleThk = 0.0001
    Private Const Face2FaceBasis = 1002 ' Face-to-face dimension basis, detailed representation

    '----------------------------------------------------------------------------------
    'DefinitionName/ProgID of this symbol is "SP3DBallValve,Ingr.SP3D.Content.PipingSample.CBallValve"
    '----------------------------------------------------------------------------------

#Region "Definition of Inputs"

    <InputCatalogPart(1)> _
        Public m_oPartInput As InputCatalogPart
    <InputDouble(2, "FacetoFace", "Face to Face", 0.31)> _
        Public m_dFaceToFaceInput As InputDouble
    <InputDouble(3, "ValCenLineToBot", "Valve Centerline to Bottom", 0, True)> _
       Public m_dValCenLineToBotInput As InputDouble
    <InputDouble(4, "OffsetFrmValCen", "Offset from Valve Centerline", 0, True)> _
       Public m_dOffsetFrmValCenInput As InputDouble
    <InputDouble(5, "InsulationThickness", "Insulation Thickness", 0.025)> _
        Public m_dInsulationThicknessInput As InputDouble

#End Region

#Region "Definitions of Aspects and their outputs"

    'Physical Aspect
    <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
    <SymbolOutput("ValveBody", "Valve Body")> _
    <SymbolOutput("BottomBody", "Bottom Body")> _
    <SymbolOutput("TopBody", "Top Body")> _
    <SymbolOutput("PNoz1", "PNoz1")> _
    <SymbolOutput("PNoz2", "PNoz2")> _
    <SymbolOutput("Operator", "Operator occurrence")> _
        Public m_oPhysicalAspect As AspectDefinition

    'Insulation Aspect
    <Aspect("Insulation", "Insulation Aspect", AspectID.Insulation)> _
    <SymbolOutput("ValveBodyInsulation", "Insulation of Valve Body")> _
    <SymbolOutput("BottomBodyInsulation", "Insulation of Bottom Body")> _
    <SymbolOutput("TopBodyInsulation", "Insulation of Top Body")> _
    <SymbolOutput("PNoz1Insulation", "Insulation of PNoz1")> _
    <SymbolOutput("PNoz2Insulation", "Insulation of PNoz2")> _
    Public m_oInsulationAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

    Protected Overrides Sub ConstructOutputs()        
        Try
            Dim oPipingPart As Part = Nothing, oConnection As SP3DConnection
            Dim dParFacetoFace As Double
            Dim dParValCenLineToBot As Double, dParOffsetFrmValCen As Double
            Dim dPartDatabasis As Double, oPropertyCodelist As PropertyValueCodelist
            Dim oWarningColl As New Collection(Of SymbolWarningException)

            ' Get Input values
            Try
                oPipingPart = m_oPartInput.Value
                If oPipingPart Is Nothing Then
                    Throw New CmnException("Unable to retrieve piping valve part.")
                End If
                dParFacetoFace = m_dFaceToFaceInput.Value
                dParValCenLineToBot = m_dValCenLineToBotInput.Value
                dParOffsetFrmValCen = m_dOffsetFrmValCenInput.Value
            Catch oEx As Exception
                If (TypeOf oEx Is SymbolWarningException) Then
                    oWarningColl.Add(oEx)
                Else
                    Throw
                End If
            End Try

            ' Get the PartDataBasis 
            oPropertyCodelist = oPipingPart.GetPropertyValue("IJDPipeComponent", "PartDataBasis")
            dPartDatabasis = oPropertyCodelist.PropValue

            oConnection = OccurrenceConnection ' Get the connection where outputs will be created.
            If oConnection Is Nothing Then
                Throw New CmnException("Unable to retrieve connection.")
            End If

            '=================================================
            ' Construction of Physical Aspect 
            '=================================================
            'RetrieveParameters
            Dim dMaxDia As Double, dValCenLineToBot As Double, dBallDiameter As Double
            Dim oBodyOD1 As Double, oBodyOD2 As Double
            Dim oPipePortDef1 As PipePortDef, oPipePortDef2 As PipePortDef

            oPipePortDef1 = oPipingPart.PortDefinitions.Item(0)
            oPipePortDef2 = oPipingPart.PortDefinitions.Item(1)

            'Get the value of pipe diameter Or flange diameter of Port 1 or Port 2, whcih ever is greater
            oBodyOD1 = IIf((oPipePortDef1.PipingOutsideDiameter > oPipePortDef1.FlangeOrHubOutsideDiameter), _
                            oPipePortDef1.PipingOutsideDiameter, oPipePortDef1.FlangeOrHubOutsideDiameter)

            oBodyOD2 = IIf((oPipePortDef2.PipingOutsideDiameter > oPipePortDef2.FlangeOrHubOutsideDiameter), _
                            oPipePortDef2.PipingOutsideDiameter, oPipePortDef2.FlangeOrHubOutsideDiameter)

            dMaxDia = IIf((oBodyOD1 > oBodyOD2 - Math3d.DistanceTolerance), oBodyOD1, oBodyOD2)

            'Calculate "Valve Centerline to Bottom" such that if it is not provided,
            'it will be 30% more than the pipe radius or flange radius which ever is greater
            If (Abs(dParValCenLineToBot) >= Math3d.DistanceTolerance) Then
                dValCenLineToBot = 1.3 * dMaxDia / 2
            Else
                dValCenLineToBot = dParValCenLineToBot
            End If

            'Compute Ball diameter such that it should be equal to
            'greater flange diameter in case of flanged ends and
            'in case of other ends it should be more than greater pipe diameter and less than
            '"Valve Center Line to Bottom"
            If (Abs(dMaxDia - oPipePortDef1.PipingOutsideDiameter) >= Math3d.DistanceTolerance) Or _
                (Abs(dMaxDia - oPipePortDef2.PipingOutsideDiameter) >= Math3d.DistanceTolerance) Then
                dBallDiameter = dMaxDia + 0.5 * (dValCenLineToBot - dMaxDia / 2)
            Else
                dBallDiameter = dMaxDia
            End If

            If (oPipePortDef1.FlangeOrHubThickness > oPipePortDef2.FlangeOrHubThickness - Math3d.DistanceTolerance) Then
                dBallDiameter = dBallDiameter - 2 * oPipePortDef1.FlangeOrHubThickness
            Else
                dBallDiameter = dBallDiameter - 2 * oPipePortDef2.FlangeOrHubThickness
            End If

            'Initialize SymbolGeometryHelper. Set the active position and orientation 
            Dim oSymbolGeomHlpr As New SymbolGeometryHelper()
            oSymbolGeomHlpr.ActivePosition = New Position(0, 0, 0)
            oSymbolGeomHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 1, 0))

            ' Create Center sphere   
            Dim oCenterSphere = oSymbolGeomHlpr.CreateSphere(oConnection, 0.5 * dBallDiameter)
            m_oPhysicalAspect.Outputs("ValveBody") = oCenterSphere

            ' Create Cylinder1 (Left Cylinder) 
            oSymbolGeomHlpr.ActivePosition = New Position(-0.5 * dParFacetoFace, 0, 0)
            Dim oCylinder1 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * oPipePortDef1.PipingOutsideDiameter, 0.5 * dParFacetoFace)
            m_oPhysicalAspect.Outputs("ValveBody") = oCylinder1

            ' Create Cylinder2 (Right Cylinder)
            ' The last createCylinder call the active position is moved to the other end of the cylinder 
            ' i.e. (0,0,0) in our case, the new cylinde start point will be (0,0,0) 
            Dim oCylinder2 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * oPipePortDef2.PipingOutsideDiameter, 0.5 * dParFacetoFace)
            m_oPhysicalAspect.Outputs("ValveBody") = oCylinder2

            Dim dOffsetFrmValCen As Double
            If (dPartDatabasis = Face2FaceBasis) Then
                ' Set the position and the orientation to create the Bottom and Top Projection 
                oSymbolGeomHlpr.ActivePosition = New Position(0, 0, -dValCenLineToBot)
                oSymbolGeomHlpr.SetOrientation(New Vector(0, 0, 1), New Vector(1, 0, 0))

                'Create the Bottom projection
                Dim dProjectionDia As Double

                dProjectionDia = 0.15 * dParFacetoFace
                Dim oBottomProj = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * dProjectionDia, dValCenLineToBot)
                'Set the Output
                m_oPhysicalAspect.Outputs("BottomBody") = oBottomProj

                'Create the Top projection
                dProjectionDia = 0.4 * dBallDiameter
                If (dParOffsetFrmValCen >= 0) And (dParOffsetFrmValCen <= 2 * Math3d.DistanceTolerance) Then
                    dOffsetFrmValCen = 1.6 * dMaxDia / 2
                Else
                    dOffsetFrmValCen = dParOffsetFrmValCen
                End If

                ' The last createCylinder call the active position is moved to the other end of the cylinder 
                ' i.e. (0,0,0) in our case, the new cylinde start point will be (0,0,0)
                Dim oTopProjection1 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * dProjectionDia, 0.9 * dOffsetFrmValCen)
                'Set the Output
                m_oPhysicalAspect.Outputs("TopBody") = oTopProjection1

                Dim oTopProjection2 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.7 * dProjectionDia, 0.1 * dOffsetFrmValCen)
                'Set the Output
                m_oPhysicalAspect.Outputs("TopBody") = oTopProjection2
            End If

            'Nozzle1 
            Dim oDir As New Vector(-1, 0, 0)
            Dim oPlacementPos As New Position(-0.5 * dParFacetoFace, 0, 0)
            Dim oNozzle As PipeNozzle

            ' Setting bIsFacePosition as true the constructor will internally calculate the exact position of the nozzle
            ' considering the CptOffset and Depth 
            oNozzle = New PipeNozzle(oPipingPart, oConnection, False, 1, oPlacementPos, oDir, 0.0, True)
            'Set the Output
            m_oPhysicalAspect.Outputs("PNoz1") = oNozzle

            ' Place Nozzle 2
            oDir.Set(1, 0, 0)
            oPlacementPos.Set(0.5 * dParFacetoFace, 0, 0)
            oNozzle = New PipeNozzle(oPipingPart, oConnection, False, 1, oPlacementPos, oDir, 0.0, True)
            'Set the Output
            m_oPhysicalAspect.Outputs("PNoz2") = oNozzle

            ' Operator
            Dim oPipeComp As IPipeComponent
            Dim oOperator As ComponentOcc, oOperatorPart As IPart

            oPipeComp = oPipingPart
            oOperatorPart = oPipeComp.ValveOperator
            oOperator = New ComponentOcc(oOperatorPart, oConnection)
            oOperator.Origin = New Position(0, 0, 0)
            m_oPhysicalAspect.Outputs("Operator") = oOperator

            '=================================================
            ' Construction of Insulation Aspect
            '=================================================
            Dim dParInsulationThickness As Double, dInsBallDiameter As Double

            dParInsulationThickness = m_dInsulationThicknessInput.Value
            dInsBallDiameter = dBallDiameter + 2 * dParInsulationThickness

            ' Set the Position and Orientation of the SymbolGeometryHelper 
            oSymbolGeomHlpr.ActivePosition = New Position(0, 0, 0)
            oSymbolGeomHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 1, 0))
            Dim oInsCenterSphere = oSymbolGeomHlpr.CreateSphere(oConnection, 0.5 * dInsBallDiameter)
            'Set the output
            m_oInsulationAspect.Outputs("ValveBodyInsulation") = oInsCenterSphere

            ' Create Insulation for Cylinder1 (Left Cylinder) 
            oSymbolGeomHlpr.ActivePosition = New Position(-0.5 * dParFacetoFace, 0, 0)
            Dim oInsCylinder1 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * oPipePortDef1.PipingOutsideDiameter + dParInsulationThickness, 0.5 * dParFacetoFace)
            m_oInsulationAspect.Outputs("ValveBodyInsulation") = oInsCylinder1

            ' Create Insulation for Cylinder2 (Right Cylinder)
            ' The last createCylinder call the active position is moved to the other end of the cylinder 
            ' i.e. (0,0,0) in our case, the new cylinde start point will be (0,0,0) 
            Dim oInsCylinder2 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * oPipePortDef2.PipingOutsideDiameter + dParInsulationThickness, 0.5 * dParFacetoFace)
            m_oInsulationAspect.Outputs("ValveBodyInsulation") = oInsCylinder2

            If (dPartDatabasis = Face2FaceBasis) Then
                ' Set the position and the orientation to create the Bottom and Top Projection 
                oSymbolGeomHlpr.ActivePosition = New Position(0, 0, -(dValCenLineToBot + dParInsulationThickness))
                oSymbolGeomHlpr.SetOrientation(New Vector(0, 0, 1), New Vector(1, 0, 0))

                'Create the Bottom projection
                Dim dInsProjectionDia As Double

                dInsProjectionDia = 0.15 * dParFacetoFace + 2 * dParInsulationThickness
                Dim oInsBotProjection = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * dInsProjectionDia, (dValCenLineToBot + dParInsulationThickness))
                'Set the Output
                m_oInsulationAspect.Outputs("BottomBodyInsulation") = oInsBotProjection

                'Create the Top projection
                dInsProjectionDia = 0.3 * dParFacetoFace + 2 * dParInsulationThickness

                Dim oInsTopProjection1 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * dInsProjectionDia, 0.9 * dOffsetFrmValCen)
                'Set the Output
                m_oInsulationAspect.Outputs("TopBodyInsulation") = oInsTopProjection1

                Dim oInsTopProjection2 = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.7 * dInsProjectionDia, 0.1 * dOffsetFrmValCen)
                'Set the Output
                m_oInsulationAspect.Outputs("TopBodyInsulation") = oInsTopProjection2
            End If

            'Create the Insulation for Ports
            Dim dInsThick1 As Double, dInsThick2 As Double    'Length of Insulation need apply on Port1 & Port 2
            Dim dInsulationDia1 As Double, dInsulationDia2 As Double

            'Calculate the insulation Diameter on Both sides of valve as per the ports
            dInsulationDia1 = IIf((oPipePortDef1.PipingOutsideDiameter >= oPipePortDef1.FlangeOrHubOutsideDiameter - Math3d.DistanceTolerance), _
                                  oPipePortDef1.PipingOutsideDiameter, oPipePortDef1.FlangeOrHubOutsideDiameter) + 2 * dParInsulationThickness
            dInsulationDia2 = IIf((oPipePortDef2.PipingOutsideDiameter >= oPipePortDef2.FlangeOrHubOutsideDiameter - Math3d.DistanceTolerance), _
                                  oPipePortDef2.PipingOutsideDiameter, oPipePortDef2.FlangeOrHubOutsideDiameter) + 2 * dParInsulationThickness

            'Insulation for port 1
            'Calculate the length of insulation for Port 1 and it should not be greater than Face 1 to Center
            If (Abs(oPipePortDef1.FlangeOrHubThickness) >= Math3d.DistanceTolerance) Then
                dInsThick1 = NegligibleThk
            Else
                dInsThick1 = dParInsulationThickness
            End If

            If (dInsThick1 > 0.5 * dParFacetoFace - Math3d.DistanceTolerance) Then
                dInsThick1 = 0.5 * dParFacetoFace
            End If

            oSymbolGeomHlpr.ActivePosition = New Position(-0.5 * dParFacetoFace, 0, 0)
            oSymbolGeomHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 1, 0))

            Dim oInsPort = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * dInsulationDia1, dInsThick1)
            'Set the Output
            m_oInsulationAspect.Outputs("PNoz1Insulation") = oInsPort

            'Insulation for port 2
            'Calculate the length of insulation for Port 2 and it should not be greater than Face 2 to Center
            If (Abs(oPipePortDef2.FlangeOrHubThickness) >= Math3d.DistanceTolerance) Then
                dInsThick2 = NegligibleThk
            Else
                dInsThick2 = dParInsulationThickness
            End If

            If (dInsThick2 > 0.5 * dParFacetoFace - Math3d.DistanceTolerance) Then
                dInsThick2 = 0.5 * dParFacetoFace
            End If

            oSymbolGeomHlpr.ActivePosition = New Position(0.5 * dParFacetoFace - dInsThick2, 0, 0)
            oInsPort = oSymbolGeomHlpr.CreateCylinder(oConnection, 0.5 * dInsulationDia2, dInsThick2)
            'Set the Output
            m_oInsulationAspect.Outputs("PNoz2Insulation") = oInsPort

            If (oWarningColl.Count > 0) Then
                Throw oWarningColl.Item(0)
            End If
        Catch Ex As Exception ' General Unhandled exception             
            Throw
        End Try

    End Sub

#End Region

   
End Class

