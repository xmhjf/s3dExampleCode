'******************************************************************************
' Copyright (C) 2004-10, Intergraph Corporation. All rights reserved.
'
'File
'   Vessels.vb
'
'Author
'   Haneef
'
'Description
'   This file contains the symbol CAD's  for HorizontalVessels type.
'History:
'   30.Sep.2010    GUK       Initial Creation
'   02.Dec.2010    Haneef    Recreated the entire symbol as per new design   
'   21.Jan.2011    Haneef    TR#192407. Theere is mismatch in the output names for vb and dot net.
'                            As prt of fix, made all output names match with the vb.
'  01 Jun 2011    Haneef              TR-CP-192721  EquipmentAssemblyDefinition should allow OnPreLoad to be override by the symbol  
'******************************************************************************
Option Explicit On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.Equipment.Middle.Services
Imports Ingr.SP3D.Common.Exceptions

'******************************************************************************
'<<AssemblyName,NameSpaceName.ClassName>> = 
'HorizontalVessels,Ingr.SP3D.Content.Equipment.StorageTank
'AssemblyName = HorizontalVessels
'NameSpaceName= Ingr.SP3D.Content.Equipment
'ClassName = StorageTank
'Symbol Config file entry: '<progid name="HorizontalVessels,Ingr.SP3D.Content.Equipment.StorageTank" clsid="{00000000-0000-0000-0000-000000000000}" dll="%OLE_SERVER%\bin\Equipment\Symbols\Release\SP3DNetTankAsm.dll" />
'******************************************************************************
Namespace Ingr.SP3D.Content.Equipment
    ''' <summary>
    ''' StorageTank contains 2 pipe nozzles, 1 Cable Port, 2 foundation ports and 1 DatumShape as AssembuOutPuts.
    ''' Class is defined in such a way that Nozzles and Datum shapes 
    ''' alone comes from Assembly definition, where as the geometry of the symbol related to Tank comes from symbol.
    ''' </summary>
    ''' <AssemblyName>HorizontalVessels</AssemblyName>>
    ''' <NameSpace>Ingr.SP3D.Content.Equipment</NameSpace>
    ''' <ClassName>StorageTank</ClassName>
    ''' <Progid>HorizontalVessels,Ingr.SP3D.Content.Equipment.StorageTank</Progid>
    ''' <remarks></remarks>
    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.3.0.0")> _
    Public Class StorageTank : Inherits EquipmentAssemblyDefinition
        Implements ICustomWeightCG

#Region "Constants delcaration for symbol input parameters and symbol outputs"
        'Inputs
        Private Const Input_VesselLength As String = "VesselLength"
        Private Const Input_VesselDiameter As String = "VesselDiameter"
        Private Const Input_SupportLength As String = "SupportLength"
        Private Const Input_SupportHeight As String = "SupportHeight"
        Private Const Input_SupportThickness As String = "SupportThickness"
        Private Const Input_InsulationThickness As String = "InsulationThickness"
        'Outputs
        'Simple physical Aspect
        Private Const Output_TankBody As String = "TankBodyCylinder"
        Private Const Output_TankLeftEndCap As String = "LeftEndCap"
        Private Const Output_TankRightEndCap As String = "RightEndCap"
        Private Const Output_LeftSupport As String = "Support1"
        Private Const Output_RightSupport As String = "Support2"
        Private Const Output_DefaultSurface1 As String = "DefaultSurface1"
        Private Const Output_DefaultSurface2 As String = "DefaultSurface2"
        Private Const Output_ConnectPort1 As String = "ConnectPort1"
        Private Const Output_ConnectPort2 As String = "ConnectPort2"
        'Insulation Aspect
        Private Const Output_TankBodyInsulation As String = "Cylinder"
        Private Const Output_TankLeftEndCapInsulation As String = "LeftEndCap"
        Private Const Output_TankRightEndCapInsulation As String = "RightEndCap"
        'Maintenance geometry
        Private Const Output_MaintenanceGeom As String = "TankEnvelope"
        '' output definitions for reference geometry aspects
        Private Const Output_TankControlPoint As String = "TankServicesControlPoint"
        ' CAD definition 
        Private Const Output_DatumShape As String = "DP1"
        Private Const Output_Nozzle1 As String = "NozzleSTNoz11"
        Private Const Output_Nozzle2 As String = "NozzleSTNoz22"
        Private Const Output_FoundationPort1 As String = "NozzleSTFndPort14"
        Private Const Output_FoundationPort2 As String = "NozzleSTFndPort25"
        Private Const Output_CABLEPORT As String = "NozzleSTNoz33"

#End Region

#Region "Private Variables"
        Dim m_oEquipment As Ingr.SP3D.Equipment.Middle.Equipment
#End Region

#Region "Symbol Definition"
        '''''''''''''''''Start of Symbol Geometry Definition Inputs '''''''''''''''''''''''''''''''''''
        'Declare the Catalog Part
        <InputCatalogPart(1)> Public m_CatalogPart As InputCatalogPart
        'Define 'VesselLength Input, default value = 4.0m'
        <InputDouble(2, Input_VesselLength, "Vessel Length", 4.0)> _
        Public m_dVesselLength As InputDouble
        'Define 'VesselDiameter Input, default value = 2.0m'
        <InputDouble(3, Input_VesselDiameter, "Vessel Diameter", 2.0)> _
        Public m_dVesselDiameter As InputDouble
        'Define 'SupportLength Input, default value = 1.7m'
        <InputDouble(4, Input_SupportLength, "Support Length", 1.7)> _
        Public m_dSupportLength As InputDouble
        'Define 'SupportHeight Input, default value = 0.35m'
        <InputDouble(5, Input_SupportHeight, "Support Height", 0.35)> _
        Public m_dSupportHeight As InputDouble
        'Define 'SupportThickness Input, default value = 0.15m'
        <InputDouble(6, Input_SupportThickness, "Support Thickness", 0.15)> _
        Public m_dSupportThickness As InputDouble
        'Define 'InsulationThickness Input, default value = 0.01m'
        <InputDouble(7, Input_InsulationThickness, "Insulation Thickness", 0.01)> _
        Public m_dInsulationThickness As InputDouble
        '''''''''''''''''End of Symbol Geometry Definition Inputs '''''''''''''''''''''''''''''''''''

        ''''''''''''''''' Start of Symbol Geometry Aspect definition and Outputs '''''''''''''''''''''''''''''''
        'Simple Physical Aspect
        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(Output_TankBody, "Tank Body Cylinder")> _
        <SymbolOutput(Output_TankLeftEndCap, "Left Elliptical End Cap")> _
        <SymbolOutput(Output_TankRightEndCap, "Right Elliptical End Cap")> _
        <SymbolOutput(Output_LeftSupport, "Support 1")> _
        <SymbolOutput(Output_RightSupport, "Support 2")> _
        <SymbolOutput(Output_DefaultSurface1, "Default Surface 1")> _
        <SymbolOutput(Output_DefaultSurface2, "Default Surface 2")> _
        <SymbolOutput(Output_ConnectPort1, "Connect Port 1")> _
        <SymbolOutput(Output_ConnectPort2, "Connect Port 2")> _
        Public m_PhysicalAspect As AspectDefinition
        'Insulation Aspect 'We try to keep the name of the 
        'outputs same since the collection it goes is different
        <Aspect("Insulation", "Insulation Aspect", AspectID.Insulation)> _
        <SymbolOutput(Output_TankBodyInsulation, "Cylinder")> _
        <SymbolOutput(Output_TankLeftEndCapInsulation, "Left End Cap")> _
        <SymbolOutput(Output_TankRightEndCapInsulation, "Right End Cap")> _
        Public m_InsulationAspect As AspectDefinition
        'Maintenance Aspect
        <Aspect("Maintenance", "Maintenance Aspect", AspectID.Maintenance)> _
        <SymbolOutput(Output_MaintenanceGeom, "Tank Envelope")> _
        Public m_Maitenance As AspectDefinition
        'Reference Geometry Aspect
        <Aspect("ReferenceGeometry", "ReferenceGeometry Aspect", AspectID.ReferenceGeometry)> _
        <SymbolOutput(Output_TankControlPoint, "Tank Services Control Point")> _
        Public m_ReferenceGeometryAspect As AspectDefinition
        '''''''''''''''''Ending of Symbol Geometry Aspect definition and Outputs '''''''''''''''''''''''''''''''
#End Region

#Region "Custom Assembly Definition"
        'Datum shape Assembly output
        <AssemblyOutput(1, Output_DatumShape)> _
        <CustomPropertyManagement("IJUADATUMSHAPE", CustomPropertyStatus.ReadOnly)> _
        Public m_DatumShape As AssemblyOutput
        'Nozzle1 Assembly output
        <AssemblyOutput(2, Output_Nozzle1)> _
        <CustomPropertyManagement("IJNOZZLEORIENTATION", CustomPropertyStatus.ByUser)> _
        Public m_SuctionNozzle As AssemblyOutput
        'Nozzle2 Assembly output
        <AssemblyOutput(3, Output_Nozzle2)> _
        <CustomPropertyManagement("IJNOZZLEORIENTATION", CustomPropertyStatus.ByUser)> _
        Public m_DischargeNozzle As AssemblyOutput
        'Cable Port Assembly output
        <AssemblyOutput(4, Output_CABLEPORT)> _
        <CustomPropertyManagement("IJNOZZLEORIENTATION", CustomPropertyStatus.ByUser)> _
        Public m_CablePort As AssemblyOutput
        'FoundationPort1 Assembly output
        <AssemblyOutput(5, Output_FoundationPort1)> _
        <CustomPropertyManagement("IJNOZZLEORIENTATION", CustomPropertyStatus.ReadOnly)> _
        Public m_FoundationPort1 As AssemblyOutput
        'FoundationPort2 Assembly output
        <AssemblyOutput(6, Output_FoundationPort2)> _
        <CustomPropertyManagement("IJNOZZLEORIENTATION", CustomPropertyStatus.ReadOnly)> _
        Public m_FoundationPort2 As AssemblyOutput
#End Region

#Region "Symbol Geometry Construction"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''' Construction of Symbol outputs  ''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' <summary>
        ''' Constructs the Symbol Geometry of the Storage Tank
        ''' </summary>
        ''' <remarks>Constructs the symbol geometry of the storage Tank for SP, Insulation and Maitennace aspects</remarks>
        Protected Overrides Sub ConstructOutputs()
            Try
                Dim oSymGeomHelper As SymbolGeometryHelper
                Dim oSp3dConnection As SP3DConnection
                Dim oWarningCollection As New Collection(Of SymbolWarningException)
                Dim dTempVal As Double
                Dim oProjection As Ingr.SP3D.Common.Middle.Projection3d
                Dim oPosPlacementPoint1 As Position
                Dim oPosPlacementPoint2 As Position
                Dim oPosSupport1Point As Position
                Dim oPosSupport2Point As Position
                Dim oXAxis As Vector
                Dim oNegXAxis As Vector
                Dim oZAxis As Vector
                Dim oYAxis As Vector

                'Validate the input parameters
                Try
                    If Not m_dVesselDiameter.Value > Math3d.DistanceTolerance OrElse Not m_dVesselLength.Value > Math3d.DistanceTolerance OrElse _
                    Not m_dSupportHeight.Value > Math3d.DistanceTolerance OrElse Not m_dSupportThickness.Value > Math3d.DistanceTolerance Then
                        'It means one of input parameter value is less than distance tolerence value.
                        'Symbol will fail with these input parameters. Hence raise SymbolErrorException
                        Throw New SymbolErrorException("EquipmentSymbolErrors", 1, Me.Occurrence)
                    End If
                Catch oEx As Exception
                    If (TypeOf oEx Is SymbolWarningException) Then
                        oWarningCollection.Add(oEx)
                    Else
                        'Trow all remaining exceptions as symbol computation will fail with these parameter values.
                        Throw
                    End If
                End Try


                'Initialize SymbolGeometryHelper object
                oSymGeomHelper = New SymbolGeometryHelper
                'Get the Connection
                oSp3dConnection = Me.OccurrenceConnection
                'Defining the placement points
                dTempVal = m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5
                'Define the First Placement posotion for the construction of VesselDiameter and LeftEndCap
                oPosPlacementPoint1 = New Position(0, 0, dTempVal)
                'Define the Second Placement position for the construction of RightEndCap
                oPosPlacementPoint2 = New Position(m_dVesselLength.Value, 0, dTempVal)
                'Define the vectors
                oXAxis = New Vector(1, 0, 0)
                oNegXAxis = New Vector(-1, 0, 0)
                oZAxis = New Vector(0, 0, 1)
                oYAxis = New Vector(0, 1, 0)

                'Creating the Physical Aspect outputs
                oSymGeomHelper.MoveToPoint(oPosPlacementPoint1)
                oSymGeomHelper.SetOrientation(oXAxis, oYAxis)

                ' output 1(TankBodyCylinder)
                oProjection = oSymGeomHelper.CreateCylinder(Nothing, m_dVesselDiameter.Value / 2.0, m_dVesselLength.Value)
                'Add the projection to symbol output in Simple physical Aspect
                m_PhysicalAspect.Outputs(Output_TankBody) = oProjection

                ' output 2 - Right Elliptical Head 
                'Create the Dishes.
                'Flatside dia should match with Vessel diameter.
                'Radius should be 1.5 times of flatSideDia
                Dim oEndCap As Ingr.SP3D.Common.Middle.Revolution3d
                Dim ellipticalArc As EllipticalArc3d
                ellipticalArc = New EllipticalArc3d(New Position(m_dVesselLength.Value, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), New Vector(0, -1, 0), New Vector(0, 0, m_dVesselDiameter.Value / 2), 0.5, 1.5 * Math.PI, Math.PI / 2)
                oEndCap = New Revolution3d(ellipticalArc, New Vector(1, 0, 0), New Position(0, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), 2 * Math.PI, True)
                m_PhysicalAspect.Outputs(Output_TankLeftEndCap) = oEndCap

                ' output 3 - Left Elliptical Head - reset position to start of cyliner and set direction in - x
                ellipticalArc = New EllipticalArc3d(New Position(0, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), New Vector(0, 1, 0), New Vector(0, 0, m_dVesselDiameter.Value / 2), 0.5, 1.5 * Math.PI, Math.PI / 2)
                oEndCap = New Revolution3d(ellipticalArc, New Vector(1, 0, 0), New Position(0, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), 2 * Math.PI, True)
                m_PhysicalAspect.Outputs(Output_TankRightEndCap) = oEndCap


                '' calcuate total support height dsupportHeight + 0.25 * dvesselDiameter
                ''Defining the Support origin points
                oPosSupport1Point = New Position((m_dVesselLength.Value * 0.25 - m_dSupportThickness.Value * 0.5), 0.0, (m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.25) * 0.5)
                oPosSupport2Point = New Position((m_dVesselLength.Value * 0.75 - m_dSupportThickness.Value * 0.5), 0.0, (m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.25) * 0.5)

                '' create box for support 1 LeftSupport
                oSymGeomHelper.MoveToPoint(oPosSupport1Point)
                oSymGeomHelper.SetOrientation(oXAxis, oYAxis)
                Dim oSupport As Ingr.SP3D.Common.Middle.BusinessObject
                'We need to switch the dSupportLength and dSupportThickness as we are not modifying the orientaion here
                oSupport = oSymGeomHelper.CreateBox(Nothing, m_dSupportThickness.Value, _
                                                 m_dSupportLength.Value, _
                                                (m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.25))
                m_PhysicalAspect.Outputs(Output_LeftSupport) = oSupport

                'Create the Right Support  
                oSymGeomHelper.MoveToPoint(oPosSupport2Point)
                oSupport = oSymGeomHelper.CreateBox(Nothing, m_dSupportThickness.Value, _
                                                 m_dSupportLength.Value, _
                                                (m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.25))
                m_PhysicalAspect.Outputs(Output_RightSupport) = oSupport


                'Create a Default Surface 1
                oSymGeomHelper.MoveToPoint(New Position(0.0, 0.0, 0.0))
                Dim planeEnds As New Collection(Of Position)
                Dim reference1 As Double = (m_dVesselLength.Value * 0.25 - m_dSupportThickness.Value * 0.5)
                Dim reference2 As Double = (m_dSupportLength.Value * -0.5)

                Dim oPlaneEdges As New Collection(Of Position)
                oPlaneEdges.Add(New Position(reference1, reference2, 0.0))
                oPlaneEdges.Add(New Position(reference1, reference2 + m_dSupportLength.Value, 0.0))
                oPlaneEdges.Add(New Position(reference1 + m_dSupportThickness.Value, reference2 + m_dSupportLength.Value, 0.0))
                oPlaneEdges.Add(New Position(reference1 + m_dSupportThickness.Value, reference2, 0.0))

                Dim defaultSurface1 As New Plane3d(oPlaneEdges)

                m_PhysicalAspect.Outputs(Output_DefaultSurface1) = defaultSurface1

                'Create a Default Surface 2
                oSymGeomHelper.MoveToPoint(New Position(0.0, 0.0, 0.0))
                Dim planeEnds1 As New Collection(Of Position)
                Dim reference3 As Double = (m_dVesselLength.Value * 0.75 - m_dSupportThickness.Value * 0.5)



                Dim oPlaneEdges1 As New Collection(Of Position)
                oPlaneEdges1.Add(New Position(reference3, reference2, 0.0))
                oPlaneEdges1.Add(New Position(reference3, reference2 + m_dSupportLength.Value, 0.0))
                oPlaneEdges1.Add(New Position(reference3 + m_dSupportThickness.Value, reference2 + m_dSupportLength.Value, 0.0))
                oPlaneEdges1.Add(New Position(reference3 + m_dSupportThickness.Value, reference2, 0.0))

                Dim defaultSurface2 As New Plane3d(oPlaneEdges1)

                m_PhysicalAspect.Outputs(Output_DefaultSurface2) = defaultSurface2

                'Create Connect Point 1

                Dim ConnectPoint1 As New Point3d(reference1, reference2, 0.0)
                m_PhysicalAspect.Outputs(Output_ConnectPort1) = ConnectPoint1

                reference3 = (m_dVesselLength.Value * 0.75 + m_dSupportThickness.Value * 0.5)
                'Create Connect Point 2
                Dim ConnectPoint2 As New Point3d(reference3, reference2, 0.0)
                m_PhysicalAspect.Outputs(Output_ConnectPort2) = ConnectPoint2

                'Now adding the geometry for Insulation Aspect
                'output 1(TankBodyInsulation)
                oSymGeomHelper.MoveToPoint(oPosPlacementPoint1)
                oSymGeomHelper.SetOrientation(oXAxis, oYAxis)
                oProjection = oSymGeomHelper.CreateCylinder(Nothing, (m_dVesselDiameter.Value * 0.5 + m_dInsulationThickness.Value), m_dVesselLength.Value)
                m_InsulationAspect.Outputs(Output_TankBodyInsulation) = oProjection

				' output 2 - Right Elliptical Head 
                Dim oInsulationEndCap As Ingr.SP3D.Common.Middle.Revolution3d
                Dim minorMajorRatio As Double = (0.5 * 0.5 * m_dVesselDiameter.Value + m_dInsulationThickness.Value) / (0.5 * m_dVesselDiameter.Value + m_dInsulationThickness.Value)
                ellipticalArc = New EllipticalArc3d(New Position(0, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), New Vector(0, 1, 0), New Vector(0, 0, m_dVesselDiameter.Value / 2 + m_dInsulationThickness.Value), minorMajorRatio, 1.5 * Math.PI, Math.PI / 2)
                oInsulationEndCap = New Revolution3d(ellipticalArc, New Vector(1, 0, 0), New Position(0, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), 2 * Math.PI, True)
                m_InsulationAspect.Outputs(Output_TankRightEndCapInsulation) = oInsulationEndCap

                ' output 3 - Left Elliptical Head - reset position to start of cyliner and set direction in - x
                ellipticalArc = New EllipticalArc3d(New Position(m_dVesselLength.Value, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), New Vector(0, -1, 0), New Vector(0, 0, m_dVesselDiameter.Value / 2 + m_dInsulationThickness.Value), minorMajorRatio, 1.5 * Math.PI, Math.PI / 2)
                oInsulationEndCap = New Revolution3d(ellipticalArc, New Vector(1, 0, 0), New Position(0, 0, m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5), 2 * Math.PI, True)
                m_InsulationAspect.Outputs(Output_TankLeftEndCapInsulation) = oInsulationEndCap

                'Creating the Maintenance Aspect Symbol outputs
                Dim dNzlLength As Double = 0.1 * m_dVesselDiameter.Value
                Dim dLength As Double = 1.2 * m_dVesselLength.Value + 0.575 * m_dVesselDiameter.Value
                Dim dWidth As Double = 1.1 * m_dVesselDiameter.Value
                Dim dHeight As Double = 1.2 * (m_dSupportHeight.Value + 1.1 * m_dVesselDiameter.Value)

                '' move to center position of the start face of the box
                oSymGeomHelper.SetOrientation(oXAxis, oYAxis)
                oSymGeomHelper.MoveToPoint(New Position((-0.1 * m_dVesselLength.Value - 0.3 * m_dVesselDiameter.Value), 0.0, (m_dVesselDiameter.Value * 0.55 + m_dSupportHeight.Value * 0.5)))
                Dim oBox As Ingr.SP3D.Common.Middle.BusinessObject
                oBox = oSymGeomHelper.CreateBox(Nothing, dLength, dWidth, dHeight)
                m_Maitenance.Outputs(Output_MaintenanceGeom) = oBox

                'Now the reference geometry outputs
                Dim oControlpointpos As Position = New Position(0.0, 0.0, 0.0)
                Dim oContlBO As BusinessObject
                Dim oControlPoint As ControlPoint
                oControlPoint = New ControlPoint(oSp3dConnection, oControlpointpos, 0.02)
                oContlBO = DirectCast(oControlPoint, BusinessObject)
                oContlBO.SetPropertyValue(CType(ControlPointType.ControlPoint, Integer), "IJControlPoint", "Type")
                oContlBO.SetPropertyValue(CType(ControlPointSubType.ProcessEquipment, Integer), "IJControlPoint", "SubType")
                m_ReferenceGeometryAspect.Outputs(Output_TankControlPoint) = oControlPoint

                'Check whether they are any warnings. 
                If oWarningCollection.Count > 0 Then
                    Throw oWarningCollection.Item(0)
                End If

            Catch ex As Exception
                Throw
            End Try
        End Sub

#End Region

#Region "Assembly Output Construction"
        Public Overrides Sub EvaluateAssembly()

            Try
                'Initialize the equipment variable
                m_oEquipment = DirectCast(Me.Occurrence, Ingr.SP3D.Equipment.Middle.Equipment)
                'Creation of DatumShape object.
                ''Check whether the output is deleted by the end user. If it is deleted by the user, no need to think about
                ''creation and evalution of the output.
                If m_DatumShape.HasBeenDeletedByUser = False Then
                    'Check whether the output is nothing. If it is nothing, it is not at created. Need to created the output.
                    'Else evaluate the position and orientation of output.
                    If m_DatumShape.Output = Nothing Then
                        m_DatumShape.Output = CreateDP1()
                    ElseIf IsBehaviorControlledByUser(m_DatumShape.Output) = False Then
                        EvaluateDP1(m_DatumShape.Output)
                    End If
                    m_DatumShape.CanDeleteIndependently = IsBehaviorControlledByUser(m_DatumShape.Output)
                End If

                'SuctionNozzle output
                ''Check whether the output is deleted by the end user. If it is deleted by the user, no need to think about
                ''creation and evalution of the output.
                If m_SuctionNozzle.HasBeenDeletedByUser = False Then
                    'Check whether the output is nothing. If it is nothing, it is not at created. Need to created the output.
                    'Else evaluate the position of output.
                    If m_SuctionNozzle.Output = Nothing Then
                        m_SuctionNozzle.Output = CreateSuctionNozzle()
                    ElseIf Not IsBehaviorControlledByUser(m_SuctionNozzle.Output) Then
                        EvaluateSuctionNozzle(m_SuctionNozzle.Output)
                    End If
                    m_SuctionNozzle.CanDeleteIndependently = IsBehaviorControlledByUser(m_SuctionNozzle.Output)
                End If

                'DischargeNozzle output
                ''Check whether the output is deleted by the end user. If it is deleted by the user, no need to think about
                ''creation and evalution of the output.
                If m_DischargeNozzle.HasBeenDeletedByUser = False Then
                    'Check whether the output is nothing. If it is nothing, it is not at created. Need to created the output.
                    'Else evaluate the position and orientation of output.
                    If m_DischargeNozzle.Output = Nothing Then
                        m_DischargeNozzle.Output = CreateDischargeNozzle()
                    ElseIf Not IsBehaviorControlledByUser(m_DischargeNozzle.Output) Then
                        EvaluateDischargeNozzle(m_DischargeNozzle.Output)
                    End If
                    m_DischargeNozzle.CanDeleteIndependently = IsBehaviorControlledByUser(m_DischargeNozzle.Output)
                End If

                'Foundationport1 output
                ''Check whether the output is deleted by the end user. If it is deleted by the user, no need to think about
                ''creation and evalution of the output.
                If m_FoundationPort1.HasBeenDeletedByUser = False Then
                    'Check whether the output is nothing. If it is nothing, it is not at created. Need to created the output.
                    'Else evaluate the position and orientation of output.
                    If m_FoundationPort1.Output = Nothing Then
                        m_FoundationPort1.Output = CreateFoundationPort1()
                    ElseIf Not IsBehaviorControlledByUser(m_FoundationPort1.Output) Then
                        EvaluateFoundationPort1(m_FoundationPort1.Output)
                    End If
                    m_FoundationPort1.CanDeleteIndependently = IsBehaviorControlledByUser(m_FoundationPort1.Output)
                End If

                'Foundationport2 output
                ''Check whether the output is deleted by the end user. If it is deleted by the user, no need to think about
                ''creation and evalution of the output.
                If m_FoundationPort2.HasBeenDeletedByUser = False Then
                    'Check whether the output is nothing. If it is nothing, it is not at created. Need to created the output.
                    'Else evaluate the position and orientation of output.
                    If m_FoundationPort2.Output = Nothing Then
                        m_FoundationPort2.Output = CreateFoundationPort2()
                    ElseIf Not IsBehaviorControlledByUser(m_FoundationPort2.Output) Then
                        EvaluateFoundationPort2(m_FoundationPort2.Output)
                    End If
                    m_FoundationPort2.CanDeleteIndependently = IsBehaviorControlledByUser(m_FoundationPort2.Output)
                End If

                'CablePort output
                ''Check whether the output is deleted by the end user. If it is deleted by the user, no need to think about
                ''creation and evalution of the output.
                If m_CablePort.HasBeenDeletedByUser = False Then
                    'Check whether the output is nothing. If it is nothing, it is not at created. Need to created the output.
                    'Else evaluate the position and orientation of output.
                    If m_CablePort.Output = Nothing Then
                        m_CablePort.Output = CreateCablePort()
                    ElseIf Not IsBehaviorControlledByUser(m_CablePort.Output) Then
                        EvaluateCablePort(m_CablePort.Output)
                    End If
                    m_CablePort.CanDeleteIndependently = IsBehaviorControlledByUser(m_CablePort.Output)
                End If

                'Now set the weight and cg,if this object required
                EvaluateWeightCG(Me.Occurrence)

            Catch ex As Exception
                Throw
            End Try
        End Sub
#End Region

#Region "ICustomWeightCG Members"
        ''' <summary>
        ''' Evaluates the weight and center of gravity of the Equipment and Equipment component objects.
        ''' </summary>
        ''' <param name="businessObject">Equipment or EquipmentComponent business object which aggregates symbol.</param>
        Public Sub EvaluateWeightCG(ByVal businessObject As BusinessObject) Implements ICustomWeightCG.EvaluateWeightCG
            Try
                'Dim dWeight As Double
                'Dim dLocalX As Double, dLocalY As Double, dLocalZ As Double
                'dLocalX = 0.0
                'dLocalY = 0.0
                'dLocalZ = 0.0
                ''Calculate the Weight and COG of the object
                'CalculateWeightAndCOG(dWeight, dLocalX, dLocalY, dLocalZ)
                ''Now set the dry weight and cog values on the object.
                'SetWeightAndCOG(WCGType.DRY, dWeight, dLocalX, dLocalY, dLocalZ)
                ''Now set the wet weight and cog values on the object.
                'SetWeightAndCOG(WCGType.WET, dWeight, dLocalX, dLocalY, dLocalZ)
                'Here we can set the calculated weight and cog values using SetWeightAndCOG(....) method.
            Catch ex As Exception
                Throw
            End Try
        End Sub


#End Region

#Region "Custom Property Managment"
        ''' <summary>
        ''' Indicates whether a property is read-only during the pre-load of the property pages.
        ''' Override this method to modify the read-only state of a property
        ''' which is defined in the  metadata as read/write. 
        ''' 
        ''' Note: Overriding the implementation of the OnPreload method, which provides a context
        ''' of all the properties, results in this method not being invoked.
        ''' </summary>
        ''' <param name="assemblyOutputName"> Name of the assembly output.</param>
        ''' <param name="interfaceName"> Interface name of the property.</param>
        ''' <param name="propertyName"> Name of the property.</param>
        ''' <returns> A Boolean defining whether read-only.</returns>
        Public Overrides Function IsPropertyReadOnly(ByVal assemblyOutputName As String, ByVal interfaceName As String, ByVal propertyName As String) As Boolean
            'Here we can decide whether the property should be read only or not. if yes we need to send the true and vice versa.
            Return MyBase.IsPropertyReadOnly(assemblyOutputName, interfaceName, propertyName)
        End Function
        ''' <summary>
        ''' Indicates whether a property is valid during the modification of the property
        ''' within the property pages. Override this method to validate
        ''' a single property without the context of any other modified properties. If the
        ''' validation requires the context of other changed properties, then the OnPropertyChange
        ''' method must be overridden instead.
        ''' 
        ''' Note: Overriding the implementation of the OnPropertyChange method, which provides 
        ''' a context of all the properties, results in this method not being invoked.
        ''' </summary>
        ''' <param name="assemblyOutputName"> Name of the assembly output or null string for the parent assembly.</param>
        ''' <param name="interfaceName"> Interface name for the property being validated.</param>
        ''' <param name="propertyName"> Name of the property being validated.</param>
        ''' <param name="propertyValue"> New property value being proposed.</param>
        ''' <param name="errorMessage"> Returned error message for the user to indicate why the property is not valid.</param>
        ''' <returns>Boolean indicating whether valid or not.</returns>
        Public Overrides Function IsPropertyValid(ByVal assemblyOutputName As String, ByVal interfaceName As String, ByVal propertyName As String, ByVal propertyValue As Object, ByRef errorMessage As String) As Boolean
            IsPropertyValid = True
            'Here currently we are validating the properties of Equipment Objects only
            If assemblyOutputName.Length = 0 Then
                'It means that these properties belongs to Equipemtnt(occurence object)
                If propertyName.Equals(Input_VesselLength) = True OrElse propertyName.Equals(Input_VesselDiameter) = True OrElse _
                    propertyName.Equals(Input_SupportLength) = True OrElse propertyName.Equals(Input_SupportHeight) = True OrElse _
                    propertyName.Equals(Input_SupportThickness) = True Then
                    ' We do know that for these properties Property Value is double. hence directly cast it to Double                 
                    Dim oNewValue As Double = CDbl(propertyValue)
                    If Not oNewValue > Math3d.DistanceTolerance Then
                        errorMessage = "Dimension value must be grater than zero"
                        IsPropertyValid = False
                    End If
                End If
            End If
        End Function
        ' User can directly overrides the ICustomPropertyManagement functions like below.
        'Public Overrides Sub OnPreLoad(ByVal oBusinessObject As Common.Middle.BusinessObject, ByVal CollAllDisplayedValues As System.Collections.ObjectModel.ReadOnlyCollection(Of Common.Middle.Services.PropertyDescriptor))
        '    MyBase.OnPreLoad(oBusinessObject, CollAllDisplayedValues)
        'End Sub
        'Public Overrides Function OnPropertyChange(ByVal oBusinessObject As Common.Middle.BusinessObject, ByVal CollAllDisplayedValues As System.Collections.ObjectModel.ReadOnlyCollection(Of Common.Middle.Services.PropertyDescriptor), ByVal oPropToChange As Common.Middle.Services.PropertyDescriptor, ByVal oNewPropValue As Common.Middle.PropertyValue, ByRef strErrorMsg As String) As Boolean
        '    Return MyBase.OnPropertyChange(oBusinessObject, CollAllDisplayedValues, oPropToChange, oNewPropValue, strErrorMsg)
        'End Function

#End Region

#Region "Private methods"
        ''' <summary>
        '''   Creates the Datum shape and returns the Business object 
        ''' </summary>
        ''' <returns>Business object reference of Datum shape</returns>
        ''' <remarks></remarks>
        Private Function CreateDP1() As BusinessObject
            Dim oPosition As Position
            Dim oShapeBO As BusinessObject
            Dim oDatumaShape As GenericShape
            oPosition = New Position(0.0, 0.0, (m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5))
            oDatumaShape = CreateDatumShape("DP1", AspectID.ReferenceGeometry, AxisDirection.Primary, oPosition, False)
            oShapeBO = DirectCast(oDatumaShape, BusinessObject)
            EvaluateDP1(oShapeBO)
            Return oShapeBO
        End Function
        ''' <summary>
        '''   Evaluates the Datum Shape position with respect to parent. 
        ''' </summary>
        ''' <param name="oDatumShapeBO">Dataum shape business object</param>
        Private Sub EvaluateDP1(ByVal oDatumShapeBO As BusinessObject)
            Dim oDatumShape As GenericShape
            Dim oPosition As Position
            oDatumShape = DirectCast(oDatumShapeBO, GenericShape)
            oPosition = New Position(0.0, 0.0, (m_dSupportHeight.Value + m_dVesselDiameter.Value * 0.5))
            SetChildPositonRelativeToParent(oDatumShape, oPosition)
        End Sub
        ''' <summary>
        '''   Creates the Suction Nozzle and returns the Business object 
        ''' </summary>
        ''' <returns>Business object reference of Suction Nozzle</returns>
        ''' <remarks></remarks>
        Private Function CreateSuctionNozzle() As BusinessObject
            Dim oPipeNozzle As PipeNozzle
            Dim oPipeNozzleBO As BusinessObject
            oPipeNozzle = New PipeNozzle(m_oEquipment.Part.PartNumber, False, 1, CType(m_oEquipment, ISystem), False)
            oPipeNozzle.Length = 0.175
            'Change the reference of the nozzle
            oPipeNozzle.PortConstraint.ReferenceGeometry = DirectCast(m_DatumShape.Output, BusinessObject)
            oPipeNozzleBO = DirectCast(oPipeNozzle, BusinessObject)
            EvaluateSuctionNozzle(oPipeNozzleBO)
            Return oPipeNozzleBO
        End Function
        ''' <summary>
        '''  Evaluates the Suction nozzle orientation properties.
        ''' </summary>
        ''' <param name="oPipeNozzleBO">Suction nozzle business object</param>

        Private Sub EvaluateSuctionNozzle(ByVal oPipeNozzleBO As BusinessObject)
            Dim oPipeNozzle As PipeNozzle
            oPipeNozzle = DirectCast(oPipeNozzleBO, PipeNozzle)
            oPipeNozzle.PortConstraint.PortPlacementType = PortPlacementType.Radial
            oPipeNozzle.PortConstraint.N1 = m_dVesselLength.Value * 0.125
            oPipeNozzle.PortConstraint.N2 = (m_dVesselDiameter.Value * 0.5) + (oPipeNozzle.Length * 0.9)
            oPipeNozzle.PortConstraint.OR1 = 3 * PI / 2.0
        End Sub
        ''' <summary>
        '''   Creates the Dischage Nozzle and returns the Business object 
        ''' </summary>
        ''' <returns>Business object reference of Dischage Nozzle</returns>
        ''' <remarks></remarks>
        Private Function CreateDischargeNozzle() As BusinessObject
            Dim oPipeNozzle As PipeNozzle
            Dim oPipeNozzleBO As BusinessObject
            oPipeNozzle = New PipeNozzle(m_oEquipment.Part.PartNumber, False, 2, CType(m_oEquipment, ISystem), False)
            oPipeNozzle.Length = 0.175
            'Change the reference of the nozzle
            oPipeNozzle.PortConstraint.ReferenceGeometry = DirectCast(m_DatumShape.Output, BusinessObject)
            oPipeNozzleBO = DirectCast(oPipeNozzle, BusinessObject)
            EvaluateDischargeNozzle(oPipeNozzleBO)
            Return oPipeNozzleBO
        End Function
        ''' <summary>
        '''  Evaluates the Dischage nozzle orientation properties.
        ''' </summary>
        ''' <param name="oPipeNozzleBO">Discharge nozzle Business object</param>
        Private Sub EvaluateDischargeNozzle(ByVal oPipeNozzleBO As BusinessObject)
            Dim oPipeNozzle As PipeNozzle
            oPipeNozzle = DirectCast(oPipeNozzleBO, PipeNozzle)
            oPipeNozzle.PortConstraint.PortPlacementType = PortPlacementType.Radial
            oPipeNozzle.PortConstraint.N1 = m_dVesselLength.Value * 0.85
            oPipeNozzle.PortConstraint.N2 = (m_dVesselDiameter.Value * 0.5) + (oPipeNozzle.Length * 0.9)
            oPipeNozzle.PortConstraint.OR1 = PI / 2.0
        End Sub
        ''' <summary>
        '''   Creates the FoundationPort1 and returns the Business object 
        ''' </summary>
        ''' <returns>Business object reference of Foundation Port</returns>
        ''' <remarks></remarks>
        Private Function CreateFoundationPort1() As BusinessObject
            Dim oFoundationPort As FoundationPort
            Dim oFoundationBO As BusinessObject
            Dim oLCSFP As ILocalCoordinateSystem
            Dim oVecx As Vector
            Dim oVecy As Vector
            oFoundationPort = New FoundationPort(m_oEquipment.Part.PartNumber, 4, m_oEquipment, False)
            oLCSFP = DirectCast(oFoundationPort, ILocalCoordinateSystem)
            oVecx = New Vector(0.0, 0.0, 1.0)
            oVecy = New Vector(1.0, 0.0, 0.0)
            oLCSFP.SetOrientation(oVecx, oVecy)
            'Now se this as reference to datum shape
            oFoundationPort.PortConstraint.PortPlacementType = PortPlacementType.Position_By_Point
            oFoundationPort.PortConstraint.ReferenceGeometry = DirectCast(m_oEquipment, BusinessObject)
            'Set the GeometryPlacementType
            oFoundationBO = DirectCast(oFoundationPort, BusinessObject)
            'Get the codelistitem value "First Bolt Hole Location"
            Dim oCodelistInformation As CodelistInformation = GetCodelistInfo("GeometryPlacementType", "EQUIP")
            Dim oCodeListItem As CodelistItem = oCodelistInformation.GetCodelistItem("First Bolt Hole Location")
            'Set the GeomPlacementType to "First Bolt Hole Location"
            oFoundationBO.SetPropertyValue(oCodeListItem.Value, "IJEqpFndPortGeomPlacementType", "GeomPlacementType")
            EvaluateFoundationPort1(oFoundationBO)
            Return oFoundationBO
        End Function
        ''' <summary>
        '''  Evaluates the FoundationPort1 position with respect to parent. 
        ''' </summary>
        ''' <param name="oFoundationPortBO">Foundation port business object</param>
        Private Sub EvaluateFoundationPort1(ByVal oFoundationPortBO As BusinessObject)
            Dim oFoundationPort As FoundationPort = DirectCast(oFoundationPortBO, FoundationPort)
            Dim oPosition As Position = New Position(m_dVesselLength.Value * 0.25, -0.5 * m_dSupportLength.Value, 0.0)
            TransformFoundationPortToParentLocation(oFoundationPort, oPosition.X, oPosition.Y, oPosition.Z)
        End Sub
        ''' <summary>
        '''   Creates the FoundationPort2 and returns the Business object 
        ''' </summary>
        ''' <returns>Business object reference of Foundation Port</returns>
        ''' <remarks></remarks>
        Private Function CreateFoundationPort2() As BusinessObject
            Dim oFoundationPort As FoundationPort
            Dim oFoundationBO As BusinessObject
            Dim oLCSFP As ILocalCoordinateSystem
            Dim oVecx As Vector
            Dim oVecy As Vector
            oFoundationPort = New FoundationPort(m_oEquipment.Part.PartNumber, 5, m_oEquipment, False)
            oLCSFP = DirectCast(oFoundationPort, ILocalCoordinateSystem)
            oVecx = New Vector(0.0, 0.0, 1.0)
            oVecy = New Vector(1.0, 0.0, 0.0)
            oLCSFP.SetOrientation(oVecx, oVecy)
            'Now set it's reference to datum shape
            oFoundationPort.PortConstraint.PortPlacementType = PortPlacementType.Position_By_Point
            oFoundationPort.PortConstraint.ReferenceGeometry = DirectCast(m_oEquipment, BusinessObject)
            'Set the GeometryPlacementType
            oFoundationBO = DirectCast(oFoundationPort, BusinessObject)
            'Get the codelistitem value "First Bolt Hole Location"
            Dim oCodelistInformation As CodelistInformation = GetCodelistInfo("GeometryPlacementType", "EQUIP")
            Dim oCodeListItem As CodelistItem = oCodelistInformation.GetCodelistItem("First Bolt Hole Location")
            'Set the GeomPlacementType to "First Bolt Hole Location"
            oFoundationBO.SetPropertyValue(oCodeListItem.Value, "IJEqpFndPortGeomPlacementType", "GeomPlacementType")
            EvaluateFoundationPort2(oFoundationBO)
            Return oFoundationBO
        End Function
        ''' <summary>
        '''  Evaluates the FoundationPort2 position with respect to parent. 
        ''' </summary>
        ''' <param name="oFoundationPortBO">Foundation port business object</param>
        Private Sub EvaluateFoundationPort2(ByVal oFoundationPortBO As BusinessObject)
            Dim oFoundationPort As FoundationPort = DirectCast(oFoundationPortBO, FoundationPort)
            Dim oPosition As Position = New Position(m_dVesselLength.Value * 0.75, -0.5 * m_dSupportLength.Value, 0.0)
            TransformFoundationPortToParentLocation(oFoundationPort, oPosition.X, oPosition.Y, oPosition.Z)
        End Sub
        ''' <summary>
        '''   Creates the Cable Port and returns the Business object 
        ''' </summary>
        ''' <returns>Business object reference of CablePort</returns>
        ''' <remarks></remarks>
        Private Function CreateCablePort() As BusinessObject
            Dim oCablePort As CablePort
            Dim oCablePortBO As BusinessObject
            oCablePort = New CablePort(m_oEquipment.Part.PartNumber, 3, CType(m_oEquipment, ISystem), False)
            'Set the Datum shape as reference to this cable port.
            oCablePort.PortConstraint.ReferenceGeometry = DirectCast(m_DatumShape.Output, BusinessObject)
            oCablePortBO = DirectCast(oCablePort, BusinessObject)
            EvaluateCablePort(oCablePortBO)
            Return oCablePortBO
        End Function
        ''' <summary>
        '''  Evaluates the Cable port orientation properties. 
        ''' </summary>
        ''' <param name="oCablePortBO">Cable port business object</param>
        Private Sub EvaluateCablePort(ByVal oCablePortBO As BusinessObject)
            Dim oCablePort As CablePort
            oCablePort = DirectCast(oCablePortBO, CablePort)
            oCablePort.PortConstraint.PortPlacementType = PortPlacementType.Tangential
            oCablePort.PortConstraint.N1 = m_dVesselLength.Value * 0.25
            oCablePort.PortConstraint.N2 = m_dVesselDiameter.Value * 0.5
            oCablePort.PortConstraint.OR1 = 3.0 * PI / 2.0
        End Sub
        ''' <summary>
        '''  Returns the CodelistInformation object. 
        ''' </summary>
        ''' <param name="sCodelistName">Codelist string name</param>
        ''' <param name="sNameSpace">Namespace string name</param>
        Private Function GetCodelistInfo(ByVal sCodelistName As String, ByVal sNameSpace As String) As CodelistInformation
            Dim oCodelistInformation As CodelistInformation
            Dim oMetadataManager As MetadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr
            oCodelistInformation = oMetadataManager.GetCodelistInfo(sCodelistName, sNameSpace)
            Return oCodelistInformation
        End Function

        ''' <summary>
        '''  Transform the passed Foundation Port objects to passed in offset values in 
        ''' in X,Y nad Z direction
        ''' </summary>
        ''' <param name="oFoundationPort">oFoundationPort object which needs to be transformed</param>       
        ''' <param name="dXOffset">Offset value in primary direction with respect to parent position</param>
        ''' <param name="dYOffset">Offset value in Secondary direction with respect to parent position</param>
        ''' <param name="dZOffset">Offset value in Normal direction with respect to parent position</param>
        Private Sub TransformFoundationPortToParentLocation(ByVal oFoundationPort As Ingr.SP3D.Equipment.Middle.FoundationPort, _
                                    ByVal dXOffset As Double, ByVal dYOffset As Double, ByVal dZOffset As Double)
            Try

                If Not oFoundationPort = Nothing Then
                    Dim oEquipMatrix As Matrix4X4 = m_oEquipment.InternalMatrix
                    Dim oEquipPrimary As Vector = New Vector()
                    Dim oEquipSecondary As Vector = New Vector()
                    Dim oEquipNormal As Vector = New Vector()
                    Dim oDirectionVector As Vector = New Vector()

                    'Get the vectors of parent equipment object
                    oEquipPrimary.Set(oEquipMatrix.GetIndexValue(0), oEquipMatrix.GetIndexValue(1), oEquipMatrix.GetIndexValue(2))
                    oEquipSecondary.Set(oEquipMatrix.GetIndexValue(4), oEquipMatrix.GetIndexValue(5), oEquipMatrix.GetIndexValue(6))
                    oEquipNormal.Set(oEquipMatrix.GetIndexValue(8), oEquipMatrix.GetIndexValue(9), oEquipMatrix.GetIndexValue(10))


                    'Now set the position of the Foundation
                    Dim oFoundationPortPosition As Position = m_oEquipment.InternalOrigin

                    If Not dXOffset = 0.0 Then
                        oEquipPrimary.Length = dXOffset
                        oFoundationPortPosition = oFoundationPortPosition.Offset(oEquipPrimary)
                    End If

                    If Not dYOffset = 0.0 Then
                        oEquipSecondary.Length = dYOffset
                        oFoundationPortPosition = oFoundationPortPosition.Offset(oEquipSecondary)
                    End If

                    If Not dZOffset = 0.0 Then
                        oEquipNormal.Length = dZOffset
                        oFoundationPortPosition = oFoundationPortPosition.Offset(oEquipNormal)
                    End If
                    'Set the origin
                    oFoundationPort.Origin = oFoundationPortPosition
                End If

            Catch ex As Exception
                Throw
            End Try

        End Sub
#End Region
    End Class
End Namespace
