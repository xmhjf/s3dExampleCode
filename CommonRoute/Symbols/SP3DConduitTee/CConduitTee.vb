'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright (C) 2009, Intergraph PPO. All rights reserved.
'
'File
'  CConduitTee.vb
'
'Abstract
'	This is a Conduit Tee symbol. This class subclasses from CustomSymbolDefinition.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit On
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Common.Exceptions

'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Content.CounduitSample.
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------

Public Class CConduitTee : Inherits CustomSymbolDefinition

    Private Const CoverHeight = 0.002
    Private Const DefaultTee = 1
    Private Const TeeWithSideOver = 9032
    Private Const TeeWithBottomCover = 9033

    '----------------------------------------------------------------------------------
    'DefinitionName/ProgID of this symbol is "SP3DConduitTee,Ingr.SP3D.Content.CounduitSample.CConduitTee"
    '----------------------------------------------------------------------------------

#Region "Definition of Inputs"

    <InputCatalogPart(1)> _
        Public m_oPartInput As InputCatalogPart
    <InputDouble(2, "FacetoCenter", "Face to Face", 0.31, True)> _
           Public m_dFaceToCenterInput As InputDouble
    <InputDouble(3, "FacetoFace", "Face to Face", 0.31, True)> _
       Public m_dFaceToFaceInput As InputDouble
    <InputDouble(4, "Height", "Conduit Body Height", 0.31, True)> _
       Public m_dHeightInput As InputDouble
    <InputDouble(5, "Width", "Conduit Body Width", 0.31, True)> _
       Public m_dWidthInput As InputDouble
    <InputDouble(6, "CoverLength", "Cover Length", 0.31, True)> _
           Public m_dCoverLengthInput As InputDouble
    <InputDouble(7, "CoverWidth", "Cover Width", 0.31, True)> _
           Public m_dCoverWidthInput As InputDouble

#End Region

#Region "Definitions of Aspects and their outputs"

    ' Physical Aspect
    <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
    <SymbolOutput("Header", "Header Body")> _
    <SymbolOutput("Branch", "Branch Body")> _
    <SymbolOutput("Body", "Body")> _
    <SymbolOutput("Cover", "Cover")> _
    <SymbolOutput("Cylinder1", "Cylinder 1")> _
    <SymbolOutput("Cylinder2", "Cylinder 2")> _
    <SymbolOutput("Cylinder3", "Cylinder 3")> _
    <SymbolOutput("ConduitPort1", "Conduit Port 1")> _
    <SymbolOutput("ConduitPort2", "Conduit Port 2")> _
    <SymbolOutput("ConduitPort3", "Conduit Port 3")> _
        Public m_oPhysicalAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

    Protected Overrides Sub ConstructOutputs()        
        Try
            Dim oConduitpart As Part, oConnection As SP3DConnection
            Dim dPartDatabasis As Double, oPropertyCodelist As PropertyValueCodelist
            Dim oWarningColl As New Collection(Of SymbolWarningException)

            ' Get Input values
            Try
                oConduitpart = m_oPartInput.Value
                If oConduitpart Is Nothing Then
                    Throw New CmnException("Unable to retrieve conduit catalog part.")
                End If

                oPropertyCodelist = oConduitpart.GetPropertyValue("IJDPipeComponent", "PartDataBasis")
                dPartDatabasis = oPropertyCodelist.PropValue
                oConnection = OccurrenceConnection  ' Get the connection where outputs will be created.
                If oConnection Is Nothing Then
                    Throw New CmnException("Unable to retrieve connection.")
                End If
            Catch oEx As Exception
                If (TypeOf oEx Is SymbolWarningException) Then
                    oWarningColl.Add(oEx)
                Else
                    Throw
                End If
            End Try

            '=================================================
            ' Construction of Physical Aspect
            '=================================================
            Dim dParFaceToCenter As Double, dParFacetoFace As Double
            Dim dParHeight As Double, dParWidth As Double
            Dim dParCoverLength As Double, dParCoverWidth As Double
            Dim dWidth As Double, dHeight As Double

            If dPartDatabasis <= DefaultTee Then
                dParFaceToCenter = m_dFaceToCenterInput.Value
            ElseIf dPartDatabasis = TeeWithSideOver Or dPartDatabasis = TeeWithBottomCover Then
                dParFacetoFace = m_dFaceToFaceInput.Value
                dParHeight = m_dHeightInput.Value
                dParWidth = m_dWidthInput.Value
                dParCoverLength = m_dCoverLengthInput.Value
                dParCoverWidth = m_dCoverWidthInput.Value
            End If

            If dPartDatabasis = TeeWithSideOver Then       'Conduit Tee, with cover at side, specified
                dWidth = 0.7 * dParWidth         'by face-to-face, height, width, cover width
                dHeight = dParHeight             'and cover length

            ElseIf dPartDatabasis = TeeWithBottomCover Then   'Conduit Tee, with cover at the bottom, specified
                dWidth = dParWidth               'by face-to-face, height, width, cover
                dHeight = 0.7 * dParHeight       'width and cover length
            End If

            Dim oSymGeometryHlpr As New SymbolGeometryHelper() ' Initialize Symbol geometry helper 
            Dim oConduitPort1 As PipePortDef, oConduitPort2 As PipePortDef, oConduitPort3 As PipePortDef

            oConduitPort1 = oConduitpart.PortDefinitions.Item(0)
            oConduitPort2 = oConduitpart.PortDefinitions.Item(1)
            oConduitPort3 = oConduitpart.PortDefinitions.Item(2)

            If dPartDatabasis <= DefaultTee Then
                Dim dConduitOD1 As Double, dConduitOD3 As Double

                'Insert your code for output 1(Cylinder Body)
                If (oConduitPort1.FlangeOrHubOutsideDiameter >= oConduitPort1.PipingOutsideDiameter - Math3d.DistanceTolerance) Then
                    dConduitOD1 = oConduitPort1.FlangeOrHubOutsideDiameter
                Else
                    dConduitOD1 = oConduitPort1.PipingOutsideDiameter
                End If

                oSymGeometryHlpr.ActivePosition = New Position(-dParFaceToCenter, 0, 0)
                oSymGeometryHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 1, 0))
                Dim oTeeBody1 = oSymGeometryHlpr.CreateCylinder(oConnection, 0.5 * dConduitOD1, 2 * dParFaceToCenter)
                'Set the output
                m_oPhysicalAspect.Outputs("Header") = oTeeBody1

                'Insert your code for output 2(Branch Tee Body)
                If (oConduitPort3.FlangeOrHubOutsideDiameter >= oConduitPort3.PipingOutsideDiameter - Math3d.DistanceTolerance) Then
                    dConduitOD3 = oConduitPort3.FlangeOrHubOutsideDiameter
                Else
                    dConduitOD3 = oConduitPort3.PipingOutsideDiameter
                End If

                oSymGeometryHlpr.ActivePosition = New Position(0, 0, 0)
                oSymGeometryHlpr.SetOrientation(New Vector(0, 1, 0), New Vector(1, 0, 0))
                Dim oTeeBody3 = oSymGeometryHlpr.CreateCylinder(oConnection, 0.5 * dConduitOD3, dParFaceToCenter)
                'Set the output
                m_oPhysicalAspect.Outputs("Branch") = oTeeBody3
            Else
                'Insert your code for Body
                Dim oCurveColl As New Collection(Of ICurve)

                'Insert your code for Complex String
                oCurveColl.Add(New Line3d(New Position(-0.35 * dParFacetoFace + 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight), _
                                          New Position(0.35 * dParFacetoFace - 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight)))
                oCurveColl.Add(New Arc3d(New Position(0.35 * dParFacetoFace - 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight), _
                                         New Position(0.35 * dParFacetoFace, 0, 0.5 * dHeight), _
                                         New Position(0.35 * dParFacetoFace - 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight)))
                oCurveColl.Add(New Line3d(New Position(0.35 * dParFacetoFace - 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight), _
                                          New Position(-0.35 * dParFacetoFace + 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight)))
                oCurveColl.Add(New Arc3d(New Position(-0.35 * dParFacetoFace + 0.5 * dWidth, -0.5 * dWidth, 0.5 * dHeight), _
                                         New Position(-0.35 * dParFacetoFace, 0, 0.5 * dHeight), _
                                         New Position(-0.35 * dParFacetoFace + 0.5 * dWidth, 0.5 * dWidth, 0.5 * dHeight)))

                Dim oComplexStr3d As New ComplexString3d(oCurveColl)
                oSymGeometryHlpr.ActivePosition = New Position(0, 0, -dHeight)
                oSymGeometryHlpr.SetOrientation(New Vector(0, 0, 1), New Vector(0, -1, 0))
                Dim oBody = oSymGeometryHlpr.CreateProjectedPolygonFromCrossSection(oConnection, oComplexStr3d, dHeight, True)
                'Set the output
                m_oPhysicalAspect.Outputs("Body") = oBody

                'Insert your code for Cover
                oCurveColl.Clear()
                oCurveColl.Add(New Line3d(New Position(-dParCoverLength / 2 + dParCoverWidth / 2, dParCoverWidth / 2, 0.5 * dHeight), _
                               New Position(dParCoverLength / 2 - dParCoverWidth / 2, dParCoverWidth / 2, 0.5 * dHeight)))
                oCurveColl.Add(New Arc3d(New Position(dParCoverLength / 2 - dParCoverWidth / 2, dParCoverWidth / 2, 0.5 * dHeight), _
                                         New Position(dParCoverLength / 2, 0, 0.5 * dHeight), _
                                         New Position(dParCoverLength / 2 - dParCoverWidth / 2, -dParCoverWidth / 2, 0.5 * dHeight)))
                oCurveColl.Add(New Line3d(New Position(dParCoverLength / 2 - dParCoverWidth / 2, -dParCoverWidth / 2, 0.5 * dHeight), _
                                          New Position(-dParCoverLength / 2 + dParCoverWidth / 2, -dParCoverWidth / 2, 0.5 * dHeight)))
                oCurveColl.Add(New Arc3d(New Position(-dParCoverLength / 2 + dParCoverWidth / 2, -dParCoverWidth / 2, 0.5 * dHeight), _
                                         New Position(-dParCoverLength / 2, 0, 0.5 * dHeight), _
                                         New Position(-dParCoverLength / 2 + dParCoverWidth / 2, dParCoverWidth / 2, 0.5 * dHeight)))

                oComplexStr3d = New ComplexString3d(oCurveColl)
                oSymGeometryHlpr.SetOrientation(New Vector(0, 0, 1), New Vector(0, -1, 0))
                Dim oCover = oSymGeometryHlpr.CreateProjectedPolygonFromCrossSection(oConnection, oComplexStr3d, CoverHeight, True)
                'Set the output
                m_oPhysicalAspect.Outputs("Cover") = oCover

                'Insert your code for Cylinder at Port1
                oSymGeometryHlpr.ActivePosition = New Position(-0.5 * dParFacetoFace, 0, 0)
                oSymGeometryHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 1, 0))
                Dim oCylinder1 = oSymGeometryHlpr.CreateCylinder(oConnection, 0.55 * oConduitPort1.PipingOutsideDiameter, 0.3 * dParFacetoFace)
                'Set the Output
                m_oPhysicalAspect.Outputs("Cylinder1") = oCylinder1

                'Insert your code for Cylinder at Port2
                oSymGeometryHlpr.ActivePosition = New Position(0.2 * dParFacetoFace, 0, 0)
                Dim oCylinder2 = oSymGeometryHlpr.CreateCylinder(oConnection, 0.55 * oConduitPort2.PipingOutsideDiameter, 0.3 * dParFacetoFace)
                'Set the Output
                m_oPhysicalAspect.Outputs("Cylinder2") = oCylinder2

                'Insert your code for Cylinder at Port3
                Dim dLength As Double
                If dPartDatabasis = TeeWithSideOver Then
                    oSymGeometryHlpr.ActivePosition = New Position(0, 0.35 * dParWidth, 0)
                    oSymGeometryHlpr.SetOrientation(New Vector(0, 1, 0), New Vector(1, 0, 0))
                    dLength = 0.3 * dParWidth
                ElseIf dPartDatabasis = TeeWithBottomCover Then
                    oSymGeometryHlpr.ActivePosition = New Position(0, 0, -0.35 * dParHeight)
                    oSymGeometryHlpr.SetOrientation(New Vector(0, 0, -1), New Vector(0, 1, 0))
                    dLength = 0.3 * dParHeight
                End If
                Dim oCylinder3 = oSymGeometryHlpr.CreateCylinder(oConnection, 0.55 * oConduitPort3.PipingOutsideDiameter, dLength)
                'Set the Output
                m_oPhysicalAspect.Outputs("Cylinder3") = oCylinder3
            End If

            'Insert your code for output 3(Nozzle1)
            Dim oPlacePoint As New Position(), oDir As New Vector()

            If dPartDatabasis <= DefaultTee Then
                oPlacePoint.Set(-dParFaceToCenter - oConduitPort1.FlangeProjectionOrSocketOffset + oConduitPort1.SeatingOrGrooveOrSocketDepth, 0, 0)
                oDir.Set(-1, 0, 0)
            ElseIf dPartDatabasis = TeeWithSideOver Or dPartDatabasis = TeeWithBottomCover Then
                oPlacePoint.Set(-0.5 * dParFacetoFace - oConduitPort1.FlangeProjectionOrSocketOffset + oConduitPort1.SeatingOrGrooveOrSocketDepth, 0, 0)
                oDir.Set(-1, 0, 0)
            End If

            Dim oNozzle1 = New ConduitPort(oConduitpart, oConnection, 1, oPlacePoint, oDir)
            'Set the output
            m_oPhysicalAspect.Outputs("ConduitPort1") = oNozzle1

            'Insert your code for output 4(Nozzle2)
            If dPartDatabasis <= DefaultTee Then
                oPlacePoint.Set(dParFaceToCenter + oConduitPort2.FlangeProjectionOrSocketOffset - oConduitPort2.SeatingOrGrooveOrSocketDepth, 0, 0)
                oDir.Set(1, 0, 0)
            ElseIf dPartDatabasis = TeeWithSideOver Or dPartDatabasis = TeeWithBottomCover Then
                oPlacePoint.Set(0.5 * dParFacetoFace + oConduitPort2.FlangeProjectionOrSocketOffset - oConduitPort2.SeatingOrGrooveOrSocketDepth, 0, 0)
                oDir.Set(1, 0, 0)
            End If

            Dim oNozzle2 = New ConduitPort(oConduitpart, oConnection, 2, oPlacePoint, oDir)
            'Set the output
            m_oPhysicalAspect.Outputs("ConduitPort2") = oNozzle2

            'Insert your code for output 5(Nozzle3)
            If dPartDatabasis <= DefaultTee Then
                oPlacePoint.Set(0, dParFaceToCenter + oConduitPort3.FlangeProjectionOrSocketOffset - oConduitPort3.SeatingOrGrooveOrSocketDepth, 0)
                oDir.Set(0, 1, 0)
            ElseIf dPartDatabasis = TeeWithSideOver Then
                oPlacePoint.Set(0, 0.65 * dParWidth + oConduitPort3.FlangeProjectionOrSocketOffset - oConduitPort3.SeatingOrGrooveOrSocketDepth, 0)
                oDir.Set(0, 1, 0)
            ElseIf dPartDatabasis = TeeWithBottomCover Then
                oPlacePoint.Set(0, 0, -0.65 * dParHeight - oConduitPort3.FlangeProjectionOrSocketOffset + oConduitPort3.SeatingOrGrooveOrSocketDepth)
                oDir.Set(0, 0, -1)
            End If

            Dim oNozzle3 = New ConduitPort(oConduitpart, oConnection, 3, oPlacePoint, oDir)
            'Set the output
            m_oPhysicalAspect.Outputs("ConduitPort3") = oNozzle3

            If (oWarningColl.Count > 0) Then
                Throw oWarningColl.Item(0)
            End If


        Catch Ex As Exception ' General Unhandled exception 
            Throw
        End Try

    End Sub

#End Region

End Class
