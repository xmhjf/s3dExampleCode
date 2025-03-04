'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright (C) 2009, Intergraph PPO. All rights reserved.
'
'File
'  CHElbow.vb
'
'Abstract
'	This is a Horizontal elbow symbol. This class subclasses from CustomSymbolDefinition.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit On
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Common.Exceptions

'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Content.HVACSample.
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------

Public Class CHElbow : Inherits CustomSymbolDefinition

    Private Const DefaultElbow = 1
    Private Const ByThroatRadius = 137  'Elbow, specified by throat radius, width, depth and angle
    Private Const ByLegLengths = 138    'Elbow, specified by leg lengths, width, depth and angle

    '----------------------------------------------------------------------------------
    'DefinitionName/ProgID of this symbol is "SP3DHElbow,Ingr.SP3D.Content.HVACSample.CHElbow"
    '----------------------------------------------------------------------------------

#Region "Definition of Inputs"

    <InputCatalogPart(1)> _
        Public m_oPartInput As InputCatalogPart
    <InputDouble(2, "HVACShape", "HVAC shape of Elbow", 0.31)> _
        Public m_dHVACShapeInput As InputDouble
    <InputDouble(3, "Width", "Width of shape", 0.31)> _
        Public m_dWidthInput As InputDouble
    <InputDouble(4, "Depth", "Depth of shape", 0.31)> _
        Public m_dDepthInput As InputDouble
    <InputDouble(5, "Angle", "Angle of Elbow", 0.31)> _
        Public m_dAngleInput As InputDouble
    <InputDouble(6, "BWidth", "Width 2", 0.31, True)> _
        Public m_dBWidthInput As InputDouble
    <InputDouble(7, "LegLength1", "Leg Length 1", 0.31, True)> _
        Public m_dLegLength1Input As InputDouble
    <InputDouble(8, "LegLength2", "Leg Length 2", 0.31, True)> _
        Public m_dLegLength2Input As InputDouble
    <InputDouble(9, "PlaneOfTurn", "0 turn around depth, 1 turn around width side", 0.31, True)> _
        Public m_dPlanOfTurnInput As InputDouble
    <InputDouble(10, "CornerRadius", "Corner radius of shape", 0.31, True)> _
        Public m_dCornerRadiusInput As InputDouble
    <InputDouble(11, "ElbowThroatRadius", "", 0.31, True)> _
        Public m_dElbowThroatRadiusInput As InputDouble
    <InputDouble(12, "InsulationThickness", "Insulation thickness of body", 0.31)> _
    Public m_dInsThicknessInput As InputDouble

#End Region

#Region "Definitions of Aspects and their outputs"

    ' Physical Aspect 
    <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
    <SymbolOutput("OutElbow", "Outside Elbow")> _
    <SymbolOutput("HvacNozzle1", "Hvac Nozzle1")> _
    <SymbolOutput("HvacNozzle2", "Hvac Nozzle2")> _
        Public m_oPhysicalAspect As AspectDefinition

    ' Insulation Aspect 
    <Aspect("Insulation", "Insulation Aspect", AspectID.Insulation)> _
    <SymbolOutput("InsOutElbow", "Insulation of Elbow")> _
        Public m_oInsulationAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

    Protected Overrides Sub ConstructOutputs()
        Try
            ' Get Input values
            Dim oHVACPart As Part = Nothing, oConnection As SP3DConnection
            Dim dPartDatabasis As Double, oPropertyCodelist As PropertyValueCodelist
            Dim dParHVACShape As Double, dParWidth As Double, dParDepth As Double
            Dim dParAngle As Double, dParPlaneofTurn As Double
            Dim dParCornerRadius As Double, dParElbowThroatRadius As Double
            Dim oWarningColl As New Collection(Of SymbolWarningException)

            Try
                oHVACPart = m_oPartInput.Value
                If oHVACPart Is Nothing Then
                    Throw New CmnException("Unable to retrieve HVAC catalog part.")
                End If
                dParHVACShape = m_dHVACShapeInput.Value
                dParWidth = m_dWidthInput.Value
                dParDepth = m_dDepthInput.Value
                dParAngle = m_dAngleInput.Value
                dParPlaneofTurn = m_dPlanOfTurnInput.Value

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

            ' Get the PartDataBasis 
            oPropertyCodelist = oHVACPart.GetPropertyValue("IJDHvacPart", "PartDataBasis")
            dPartDatabasis = oPropertyCodelist.PropValue

            '=================================================
            ' Construction of Physical Aspect
            '=================================================

            Dim oCP As New Position   'arc center point
            Dim NO1 As New Vector   'orientation of the profile
            Dim NO2 As New Vector    'orientation of the profile
            Dim dLegLength As Double    ' leg length from the origin
            Dim dParBWidth As Double, dParLegLength1 As Double, dParLeglength2 As Double
            Dim dX As Double, dY As Double
            Dim oCurveColl As New Collection(Of ICurve)
            Dim oPointsColl As New Collection(Of Position)

            If (dPartDatabasis = DefaultElbow) Or (dPartDatabasis = ByThroatRadius) Then

                dParCornerRadius = m_dCornerRadiusInput.Value
                dParElbowThroatRadius = m_dElbowThroatRadiusInput.Value

            ElseIf (dPartDatabasis = ByLegLengths) Then

                dParBWidth = m_dBWidthInput.Value
                dParLegLength1 = m_dLegLength1Input.Value
                dParLeglength2 = m_dLegLength2Input.Value

                dX = dParLeglength2 / (Tan(dParAngle / 2))
                dY = dParLegLength1 / (Tan(dParAngle / 2))

            End If

            ' Initialize SymbolGeometryHelper 
            Dim oSymbolGeomHlpr As New SymbolGeometryHelper()

            Select Case dPartDatabasis
                Case Is <= DefaultElbow, ByThroatRadius

                    ' Insert your code for output 1(body of the elbow)
                    If (dParAngle >= -Math3d.DistanceTolerance) And (dParAngle <= Math3d.DistanceTolerance) Then
                        dParAngle = 0.1 / (180 / Math.PI)
                    End If

                    'if round make the depth = the width
                    If dParHVACShape = CrossSectionShapeTypes.Round Then dParDepth = dParWidth
                    'check to see that the elbow radius is larger than the profile so the profile doesn't cross the revolution axis
                    If dParPlaneofTurn = 1 Then
                        dParElbowThroatRadius = dParElbowThroatRadius + dParWidth / 2
                    Else
                        dParElbowThroatRadius = dParElbowThroatRadius + dParDepth / 2
                    End If
                    dLegLength = dParElbowThroatRadius * Tan(dParAngle / 2)  'Commented else condition as the formula is valid upto 180 degrees.

                    Dim oCV As New Vector(-1, 0, 0)  'rotation vector for rotation
                    Dim oAxis As New Vector(0, 0, 1)

                    'create the profile for the sweep
                    Dim oProfile As Curve3d = Nothing

                    oCP.Set(-dLegLength, 0, 0)
                    If dParHVACShape = CrossSectionShapeTypes.FlatOval Then
                        oProfile = CreateFlatOval(oCP, dParWidth, dParDepth, dParPlaneofTurn)
                    ElseIf dParHVACShape = CrossSectionShapeTypes.Rectangular Or dParHVACShape = 0 Then
                        oProfile = CreateRectangle(oCP, dParWidth, dParDepth, dParCornerRadius, dParPlaneofTurn)
                    ElseIf dParHVACShape = CrossSectionShapeTypes.Round Then
                        oProfile = PlaceTrCircleByCenter(oCP, oCV, dParWidth / 2)
                    End If

                    oCP.Y = dParElbowThroatRadius
                    oSymbolGeomHlpr.ActivePosition = oCP
                    oSymbolGeomHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 1, 0))
                    Dim oOutElbow = oSymbolGeomHlpr.CreateSurfaceofRevolution(oConnection, oProfile, dParAngle)
                    m_oPhysicalAspect.Outputs("OutElbow") = oOutElbow

                    NO1.Set(-1, 0, 0)
                    If dParPlaneofTurn = 0 Then
                        NO2.Set(0, 1, 0)
                    Else
                        NO2.Set(0, 0, 1)
                    End If

                    Dim oHVACNozzle1 As New HvacPort(oHVACPart, oConnection, False, 1, oCP, NO1, 0.0)
                    oHVACNozzle1.RadialVector = NO2
                    m_oPhysicalAspect.Outputs("HvacNozzle1") = oHVACNozzle1

                    'create nozzle 2
                    oCP.Set(dLegLength * Cos(dParAngle), dLegLength * Sin(dParAngle), 0)
                    NO1.Set(Cos(dParAngle), Sin(dParAngle), 0)

                    'set the orientation of the nozzle
                    If dParPlaneofTurn = 0 Then
                        NO2.Set(Cos(dParAngle + 90 / (180 / Math.PI)), Sin(dParAngle + 90 / (180 / Math.PI)), 0)
                    Else
                        NO2.Set(0, 0, 1)
                    End If

                    Dim oHVACNozzle2 As New HvacPort(oHVACPart, oConnection, False, 2, oCP, NO1, 0.0)
                    oHVACNozzle2.RadialVector = NO2
                    m_oPhysicalAspect.Outputs("HvacNozzle2") = oHVACNozzle2

                Case ByLegLengths

                    oPointsColl.Clear()
                    oPointsColl.Add(New Position(-dParLeglength2 * (dX + (dParWidth / 2)) / dX, 0.5 * dParWidth, 0))
                    oPointsColl.Add(New Position(-dParLeglength2 * (dX + 0.5 * dParWidth) / dX, -0.5 * dParWidth, 0))
                    oPointsColl.Add(New Position((dParLegLength1 * (dY + 0.5 * dParBWidth) / dY) * Cos(dParAngle) - (0.5 * dParBWidth * Sin(dParAngle)), _
                                                     -(dParLegLength1 * (dY + 0.5 * dParBWidth) / dY) * Sin(dParAngle) - (0.5 * dParBWidth * Cos(dParAngle)), 0))
                    oPointsColl.Add(New Position(oPointsColl.Item(2).X + dParBWidth * Sin(dParAngle), oPointsColl.Item(2).Y + dParBWidth * Cos(dParAngle), 0))
                    oPointsColl.Add(New Position(-((dParLeglength2 * (dX + 0.5 * dParWidth) / dX) - dParLeglength2), -0.5 * dParWidth, 0))

                    'For Arc
                    Dim dMRatio As Double, a As Double, b As Double
                    dX = dParLeglength2 / (Tan(dParAngle / 2))
                    b = dParWidth + dX
                    Dim dNum, dDen As Double
                    dNum = (oPointsColl.Item(3).X - oPointsColl.Item(0).X) ^ 2
                    dDen = 1 - (((oPointsColl.Item(3).Y + dX + oPointsColl.Item(0).Y) ^ 2) / (b ^ 2))
                    a = Math.Sqrt(dNum / dDen)
                    dMRatio = a / b

                    Dim oCen As New Position(oPointsColl(0).X, -(dX + (dParWidth / 2)), 0)

                    oCurveColl.Clear()
                    oCurveColl.Add(New EllipticalArc3d(oCen, New Vector(0, 0, -1), New Vector(0, dX + dParWidth, 0), dMRatio, 0, dParAngle))
                    oCurveColl.Add(New Line3d(oPointsColl(3), oPointsColl(2)))
                    oCurveColl.Add(New Line3d(oPointsColl(2), oPointsColl(4)))
                    oCurveColl.Add(New Line3d(oPointsColl(4), oPointsColl(1)))
                    oCurveColl.Add(New Line3d(oPointsColl(1), oPointsColl(0)))

                    Dim oComplexStr3d = New ComplexString3d(oCurveColl)

                    oSymbolGeomHlpr.ActivePosition = New Position(0, -0.5 * dParDepth, 0)
                    oSymbolGeomHlpr.SetOrientation(New Vector(0, 1, 0), New Vector(0, 0, 1))
                    Dim oOutElbow = oSymbolGeomHlpr.CreateProjectedPolygonFromCrossSection(oConnection, oComplexStr3d, dParDepth, True)
                    m_oPhysicalAspect.Outputs("OutElbow") = oOutElbow

                    'Create the nozzles
                    oCP.Set(-dParLeglength2 * (dX + (dParWidth / 2)) / dX, 0, 0)
                    NO1.Set(-1, 0, 0)

                    If dParPlaneofTurn = 0 Then
                        NO2.Set(0, 0, 1)
                    Else
                        NO2.Set(0, 1, 0)
                    End If

                    ' Create Nozzle 1 
                    Dim oHVACNozzle1 As New HvacPort(oHVACPart, oConnection, False, 1, oCP, NO1, 0.0)
                    oHVACNozzle1.RadialVector = NO2
                    m_oPhysicalAspect.Outputs("HvacNozzle1") = oHVACNozzle1

                    'Create Nozzle 2
                    oCP.Set((dParLegLength1 * (dY + (dParBWidth / 2)) / dY) * Cos(dParAngle), 0, _
                             -(dParLegLength1 * (dY + (dParBWidth / 2)) / dY) * Sin(dParAngle))
                    NO1.Set(Cos(dParAngle), 0, -Sin(dParAngle))

                    'set the orientation of the nozzle
                    If dParPlaneofTurn = 0 Then
                        NO2.Set(Cos(dParAngle + 90 / (180 / Math.PI)), 0, -Sin(dParAngle + 90 / (180 / Math.PI)))
                    Else
                        NO2.Set(0, 1, 0)
                    End If

                    Dim oHVACNozzle2 As New HvacPort(oHVACPart, oConnection, False, 2, oCP, NO1, 0.0)
                    oHVACNozzle2.RadialVector = NO2
                    m_oPhysicalAspect.Outputs("HvacNozzle2") = oHVACNozzle2
            End Select

            '=================================================
            ' Construction of Insulation Aspect
            '=================================================
            Dim dParInsulationThickness As Double
            dParInsulationThickness = m_dInsThicknessInput.Value

            Select Case dPartDatabasis
                Case Is <= DefaultElbow, ByThroatRadius

                    ' Insert your code for output 1(body of the elbow)
                    If (dParAngle >= -Math3d.DistanceTolerance) And (dParAngle <= Math3d.DistanceTolerance) Then
                        dParAngle = 0.1 / (180 / Math.PI)
                    End If

                    'if round make the depth = the width
                    If dParHVACShape = CrossSectionShapeTypes.Round Then dParDepth = dParWidth
                    'check to see that the elbow radius is larger than the profile so the profile doesn't cross the revolution axis
                    If dParPlaneofTurn = 1 Then
                        dParElbowThroatRadius = dParElbowThroatRadius + dParWidth / 2
                    Else
                        dParElbowThroatRadius = dParElbowThroatRadius + dParDepth / 2
                    End If
                    '   Add insulation thickness to Elbow dimensions
                    dParWidth = dParWidth + 2 * dParInsulationThickness
                    dParDepth = dParDepth + 2 * dParInsulationThickness
                    dParCornerRadius = dParCornerRadius + dParInsulationThickness
                    dLegLength = dParElbowThroatRadius * Tan(dParAngle / 2)  'Commented else condition as the formula is valid upto 180 degrees.

                    Dim oCV As New Vector(-1, 0, 0)  'rotation vector for rotation
                    Dim oAxis As New Vector(0, 0, 1)

                    'create the profile for the sweep
                    Dim oProfile As Curve3d = Nothing

                    oCP.Set(-dLegLength, 0, 0)
                    If dParHVACShape = CrossSectionShapeTypes.FlatOval Then
                        oProfile = CreateFlatOval(oCP, dParWidth, dParDepth, dParPlaneofTurn)
                    ElseIf dParHVACShape = CrossSectionShapeTypes.Rectangular Or dParHVACShape = 0 Then
                        oProfile = CreateRectangle(oCP, dParWidth, dParDepth, dParCornerRadius, dParPlaneofTurn)
                    ElseIf dParHVACShape = CrossSectionShapeTypes.Round Then    'Round=4
                        oProfile = PlaceTrCircleByCenter(oCP, oCV, dParWidth / 2)
                    End If

                    oCP.Y = dParElbowThroatRadius
                    oSymbolGeomHlpr.ActivePosition = oCP
                    oSymbolGeomHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 1, 0))
                    Dim oInsOutElbow = oSymbolGeomHlpr.CreateSurfaceofRevolution(oConnection, oProfile, dParAngle)
                    m_oInsulationAspect.Outputs("InsOutElbow") = oInsOutElbow

                Case ByLegLengths

                    'Add insulation thickness to Elbow dimensions
                    dParWidth = dParWidth + 2 * dParInsulationThickness
                    dParDepth = dParDepth + 2 * dParInsulationThickness
                    dParBWidth = dParBWidth + 2 * dParInsulationThickness

                    oPointsColl.Clear()
                    oPointsColl.Add(New Position(-dParLeglength2 * (dX + (dParWidth / 2)) / dX, 0.5 * dParWidth, 0))
                    oPointsColl.Add(New Position(-dParLeglength2 * (dX + 0.5 * dParWidth) / dX, -0.5 * dParWidth, 0))
                    oPointsColl.Add(New Position((dParLegLength1 * (dY + 0.5 * dParBWidth) / dY) * Cos(dParAngle) - (0.5 * dParBWidth * Sin(dParAngle)), _
                                                     -(dParLegLength1 * (dY + 0.5 * dParBWidth) / dY) * Sin(dParAngle) - (0.5 * dParBWidth * Cos(dParAngle)), 0))
                    oPointsColl.Add(New Position(oPointsColl.Item(2).X + dParBWidth * Sin(dParAngle), oPointsColl.Item(2).Y + dParBWidth * Cos(dParAngle), 0))
                    oPointsColl.Add(New Position(-((dParLeglength2 * (dX + 0.5 * dParWidth) / dX) - dParLeglength2), -0.5 * dParWidth, 0))

                    'For Arc
                    Dim dMRatio As Double, a As Double, b As Double
                    dX = dParLeglength2 / (Tan(dParAngle / 2))
                    b = dParWidth + dX
                    Dim dNum, dDen As Double
                    dNum = (oPointsColl.Item(3).X - oPointsColl.Item(0).X) ^ 2
                    dDen = 1 - (((oPointsColl.Item(3).Y + dX + oPointsColl.Item(0).Y) ^ 2) / (b ^ 2))
                    a = Math.Sqrt(dNum / dDen)
                    dMRatio = a / b

                    Dim oCen As New Position(oPointsColl(0).X, -(dX + (dParWidth / 2)), 0)

                    oCurveColl.Clear()
                    oCurveColl.Add(New EllipticalArc3d(oCen, New Vector(0, 0, -1), New Vector(0, dX + dParWidth, 0), dMRatio, 0, dParAngle))
                    oCurveColl.Add(New Line3d(oPointsColl(3), oPointsColl(2)))
                    oCurveColl.Add(New Line3d(oPointsColl(2), oPointsColl(4)))
                    oCurveColl.Add(New Line3d(oPointsColl(4), oPointsColl(1)))
                    oCurveColl.Add(New Line3d(oPointsColl(1), oPointsColl(0)))

                    Dim oComplexStr3d = New ComplexString3d(oCurveColl)

                    oSymbolGeomHlpr.ActivePosition = New Position(0, -0.5 * dParDepth, 0)
                    oSymbolGeomHlpr.SetOrientation(New Vector(0, 1, 0), New Vector(0, 0, 1))
                    Dim oInsOutElbow = oSymbolGeomHlpr.CreateProjectedPolygonFromCrossSection(oConnection, oComplexStr3d, dParDepth, True)
                    m_oInsulationAspect.Outputs("InsOutElbow") = oInsOutElbow

            End Select

            If (oWarningColl.Count > 0) Then
                Throw oWarningColl.Item(0)
            End If

        Catch Ex As Exception ' General Unhandled exception 
            Throw
        End Try

    End Sub

#End Region

#Region "Private Functions"

    Private Function CreateFlatOval(ByVal oCenterPoint As Position, _
                            ByVal dWidth As Double, _
                            ByVal dDepth As Double, _
                            ByVal dOrient As Double) _
                            As ComplexString3d

        Dim oPointsColl As New Collection(Of Position), oCurveColl As New Collection(Of ICurve)
        CreateFlatOval = Nothing

        Try
            If dOrient = 0 Then
                oPointsColl.Clear()
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dDepth / 2, oCenterPoint.Z - (dWidth - dDepth) / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dDepth / 2, oCenterPoint.Z + (dWidth - dDepth) / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y, oCenterPoint.Z + dWidth / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dDepth / 2, oCenterPoint.Z + (dWidth - dDepth) / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dDepth / 2, oCenterPoint.Z - (dWidth - dDepth) / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y, oCenterPoint.Z - dWidth / 2))
            Else
                oPointsColl.Clear()
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dWidth - dDepth) / 2, oCenterPoint.Z + dDepth / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dWidth - dDepth) / 2, oCenterPoint.Z + dDepth / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dWidth / 2, oCenterPoint.Z))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dWidth - dDepth) / 2, oCenterPoint.Z - dDepth / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dWidth - dDepth) / 2, oCenterPoint.Z - dDepth / 2))
                oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dWidth / 2, oCenterPoint.Z))
            End If

            oCurveColl.Clear()
            oCurveColl.Add(New Line3d(oPointsColl.Item(0), oPointsColl.Item(1)))
            oCurveColl.Add(New Arc3d(oPointsColl.Item(1), oPointsColl.Item(3), oPointsColl.Item(2)))
            oCurveColl.Add(New Line3d(oPointsColl.Item(3), oPointsColl.Item(4)))
            oCurveColl.Add(New Arc3d(oPointsColl.Item(4), oPointsColl.Item(0), oPointsColl.Item(5)))

            Dim oComplexStr3d = New ComplexString3d(oCurveColl)
            CreateFlatOval = oComplexStr3d

        Catch Ex As Exception
            Throw
        End Try

    End Function

    Private Function CreateRectangle(ByVal oCenterPoint As Position, _
                             ByVal dWidth As Double, _
                             ByVal dDepth As Double, _
                             ByVal dCornerRadius As Double, _
                             ByVal dOrient As Integer) _
                             As ComplexString3d

        CreateRectangle = Nothing
        Try
            Dim dHalfDepth As Double, dHalfWidth As Double
            Dim oPointsColl As New Collection(Of Position), oCurveColl As New Collection(Of ICurve)

            dHalfDepth = dDepth / 2
            dHalfWidth = dWidth / 2

            If (dCornerRadius <= Math3d.DistanceTolerance) Then
                If (dOrient = 0) Then
                    oPointsColl.Clear()
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfDepth, oCenterPoint.Z - dHalfWidth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfDepth, oCenterPoint.Z + dHalfWidth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfDepth, oCenterPoint.Z + dHalfWidth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfDepth, oCenterPoint.Z - dHalfWidth))
                Else
                    oPointsColl.Clear()
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfWidth, oCenterPoint.Z - dHalfDepth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfWidth, oCenterPoint.Z - dHalfDepth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfWidth, oCenterPoint.Z + dHalfDepth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfWidth, oCenterPoint.Z + dHalfDepth))
                End If

                oCurveColl.Add(New Line3d(oPointsColl.Item(0), oPointsColl.Item(1)))
                oCurveColl.Add(New Line3d(oPointsColl.Item(1), oPointsColl.Item(2)))
                oCurveColl.Add(New Line3d(oPointsColl.Item(2), oPointsColl.Item(3)))
                oCurveColl.Add(New Line3d(oPointsColl.Item(3), oPointsColl.Item(0)))

            Else
                Dim oNVector As New Vector(-1, 0, 0)
                If dOrient = 0 Then
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfDepth, oCenterPoint.Z - (dHalfWidth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfDepth - dCornerRadius), oCenterPoint.Z - (dHalfWidth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfDepth - dCornerRadius), oCenterPoint.Z - dHalfWidth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfDepth - dCornerRadius), oCenterPoint.Z - dHalfWidth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfDepth - dCornerRadius), oCenterPoint.Z - (dHalfWidth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfDepth, oCenterPoint.Z - (dHalfWidth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfDepth, oCenterPoint.Z + (dHalfWidth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfDepth - dCornerRadius), oCenterPoint.Z + (dHalfWidth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfDepth - dCornerRadius), oCenterPoint.Z + dHalfWidth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfDepth - dCornerRadius), oCenterPoint.Z + dHalfWidth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfDepth - dCornerRadius), oCenterPoint.Z + (dHalfWidth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfDepth, oCenterPoint.Z + (dHalfWidth - dCornerRadius)))
                Else
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfWidth, oCenterPoint.Z - (dHalfDepth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfWidth - dCornerRadius), oCenterPoint.Z - (dHalfDepth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfWidth - dCornerRadius), oCenterPoint.Z - dHalfDepth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfWidth - dCornerRadius), oCenterPoint.Z - dHalfDepth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfWidth - dCornerRadius), oCenterPoint.Z - (dHalfDepth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfWidth, oCenterPoint.Z - (dHalfDepth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - dHalfWidth, oCenterPoint.Z + (dHalfDepth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfWidth - dCornerRadius), oCenterPoint.Z + (dHalfDepth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y - (dHalfWidth - dCornerRadius), oCenterPoint.Z + dHalfDepth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfWidth - dCornerRadius), oCenterPoint.Z + dHalfDepth))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + (dHalfWidth - dCornerRadius), oCenterPoint.Z + (dHalfDepth - dCornerRadius)))
                    oPointsColl.Add(New Position(oCenterPoint.X, oCenterPoint.Y + dHalfWidth, oCenterPoint.Z + (dHalfDepth - dCornerRadius)))
                End If

                oCurveColl.Clear()
                oCurveColl.Add(New Arc3d(oPointsColl.Item(1), oNVector, oPointsColl.Item(0), oPointsColl.Item(2)))
                oCurveColl.Add(New Line3d(oPointsColl.Item(2), oPointsColl.Item(3)))
                oCurveColl.Add(New Arc3d(oPointsColl.Item(4), oNVector, oPointsColl.Item(3), oPointsColl.Item(5)))
                oCurveColl.Add(New Line3d(oPointsColl.Item(5), oPointsColl.Item(6)))
                oCurveColl.Add(New Arc3d(oPointsColl.Item(7), oNVector, oPointsColl.Item(6), oPointsColl.Item(8)))
                oCurveColl.Add(New Line3d(oPointsColl.Item(8), oPointsColl.Item(9)))
                oCurveColl.Add(New Arc3d(oPointsColl.Item(10), oNVector, oPointsColl.Item(9), oPointsColl.Item(11)))
                oCurveColl.Add(New Line3d(oPointsColl.Item(11), oPointsColl.Item(0)))
            End If

            Dim oComplexStr3d = New ComplexString3d(oCurveColl)
            CreateRectangle = oComplexStr3d

        Catch Ex As Exception
            Throw
        End Try

    End Function


    Private Function PlaceTrCircleByCenter(ByRef oCenterPoint As Position, _
                                         ByRef oNormalVec As Vector, _
                                         ByRef dRadius As Double) _
                                         As Circle3d

        PlaceTrCircleByCenter = Nothing
        Try
            ' Create Circle object
            Dim oCircle = New Circle3d(oCenterPoint, oNormalVec, dRadius)
            PlaceTrCircleByCenter = oCircle

        Catch Ex As Exception
            Throw
        End Try

    End Function

#End Region

End Class
