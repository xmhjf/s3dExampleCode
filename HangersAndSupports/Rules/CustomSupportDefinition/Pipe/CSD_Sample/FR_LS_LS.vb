'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
'
'   FR_LS_LS.vb
'   CSD_Sample,Ingr.SP3D.Support.Content.Rules.FR_LS_LS
'   Author       :  
'   Creation Date:  
'   Description:

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  23-2-03-2015   Chethan  TR-CP-268570  Namespace inconsistency in .NET content for few H&S project  


Option Explicit On

Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Math
Imports Ingr.SP3D.ReferenceData.Middle
Imports System.Collections.Specialized
Imports Ingr.SP3D.Support.Content.Symbols

Public Class FR_LS_LS
    Inherits CustomSupportDefinition
    Implements ICustomHgrBOMDescription

    Private Const VERT_SECTION = "VERT_SECTION"
    Private Const HOR_SECTION = "HOR_SECTION"
    Private Const PLATE = "PLATE"

    Private m_sBasePlatesOpt As PropertyValueString

    Public Overrides ReadOnly Property Parts() As Collection(Of PartInfo)
        Get

            Dim oSupport As Ingr.SP3D.Support.Middle.Support = SupportHelper.Support

            m_sBasePlatesOpt = oSupport.GetPropertyValue("IJUAHgrAssyPlates", "Plates")

            Dim oPVSPlateSize As PropertyValueString = oSupport.GetPropertyValue("IJUAHgrAssyLS", "PLATE_SIZE")
            Dim oPVSLSize As PropertyValueString = oSupport.GetPropertyValue("IJUAHgrLSize", "LSize")

            Dim oParts As New Collection(Of PartInfo)
            oParts.Add(New PartInfo(VERT_SECTION, oPVSLSize.PropValue))
            oParts.Add(New PartInfo(HOR_SECTION, oPVSLSize.PropValue))

            If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                oParts.Add(New PartInfo(PLATE, "Utility4HolePlate_Sample_" + oPVSPlateSize.PropValue))
            End If

            Return oParts
        End Get
    End Property

    Public Overrides ReadOnly Property ConfigurationCount() As Integer
        Get
            Return 16
        End Get
    End Property

    Public Overrides Sub ConfigureSupport(ByVal oSupCompColl As Collection(Of SupportComponent))



        Dim oSupport As Ingr.SP3D.Support.Middle.Support = SupportHelper.Support
        Dim iConfig As Integer = Configuration()
        Dim oRouteInfo As PipeObjectInfo = SupportedHelper.SupportedObjectInfo(1)
        BoundingBoxHelper.CreateStandardBoundingBoxes(True)

        Dim oBBX As BoundingBox
        If SupportHelper.PlacementType = PlacementType.PlaceByStruct Then
            oBBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting)
        Else
            oBBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported)
        End If

        Dim dShoeH As PropertyValueDouble = oSupport.GetPropertyValue("IJOAHgrAssyShoeH", "SHOE_H")
        Dim dOverlap As PropertyValueDouble = oSupport.GetPropertyValue("IJUAHgrAssyLS", "OVERLAP")
        Dim dOverhang As PropertyValueDouble = oSupport.GetPropertyValue("IJUAHgrAssyLS", "OVERHANG")
        Dim dBPWidth As PropertyValueDouble = oSupport.GetPropertyValue("IJUAHgrAssyLS", "BP_WIDTH")
        Dim dBPHoleSize As PropertyValueDouble = oSupport.GetPropertyValue("IJUAHgrAssyLS", "BP_HOLE_SIZE")
        Dim dBPHoleInset As PropertyValueDouble = oSupport.GetPropertyValue("IJUAHgrAssyLS", "BP_HOLE_INSET")
        Dim dExt As PropertyValueDouble = oSupport.GetPropertyValue("IJOAHgrAssyLS", "EXTENSION")

        Dim ComponentDictionary = SupportHelper.SupportComponentDictionary
        Dim oHorizontalSection As SupportComponent = ComponentDictionary(HOR_SECTION)
        Dim oVerticalSection As SupportComponent = ComponentDictionary(VERT_SECTION)
        Dim oPlate As SupportComponent
        If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
            oPlate = ComponentDictionary(PLATE)
        End If

        Dim propVal As PropertyValue = oHorizontalSection.GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")

        oHorizontalSection.SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength")
        oHorizontalSection.SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength")
        oHorizontalSection.SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation")
        oHorizontalSection.SetPropertyValue(1, "IJUAHgrMiterOcc", "BeginMiter")
        oHorizontalSection.SetPropertyValue(1, "IJUAHgrMiterOcc", "EndMiter")

        oVerticalSection.SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength")
        oVerticalSection.SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength")
        oVerticalSection.SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation")
        oVerticalSection.SetPropertyValue(1, "IJUAHgrMiterOcc", "BeginMiter")
        oVerticalSection.SetPropertyValue(1, "IJUAHgrMiterOcc", "EndMiter")

        Dim oSectionPart As BusinessObject, oCrossSection As CrossSection
        oSectionPart = oHorizontalSection.GetRelationship("madeFrom", "part").TargetObjects(0)
        oCrossSection = oSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects(0)

        Dim dPlateHgrWidth As PropertyValueDouble = oCrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")
        Dim dSteelWidth As Double = dPlateHgrWidth.PropValue
        Dim dPlateHgrDepth As PropertyValueDouble = oCrossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")
        Dim dSteelDepth As Double = dPlateHgrDepth.PropValue
        Dim dLength As Double = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical)
        Dim dPipeOD As Double = oRouteInfo.OutsideDiameter
        Dim dByPointAngle1 As Double = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y)
        Dim dByPointAngle2 As Double = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct)

        Dim ByPointStructPlaneA, ByPointStructPlaneB, ByPointRoutePlaneA, ByPointRoutePlaneB As Plane
        Dim ByPointStructAxisA, ByPointStructAxisB As Axis
        Dim ConfRoutePlaneA, ConfRoutePlaneB, StructRoutePlaneA, StructRoutePlaneB As Plane
        Dim ConfRouteAxisA, ConfRouteAxisB, StructRouteAxisA, StructRouteAxisB As Axis
        Dim dVertOffset, dRouteOffset, ByPointStructOffset, HorOffset, OVERHANG As Double

        OVERHANG = dOverhang.PropValue

        If Abs(dByPointAngle2) > PI / 2 Then    'Structure is Oriented along Route
            If iConfig = 1 Or iConfig = 5 Or iConfig = 9 Or iConfig = 13 Then
                If Abs(dByPointAngle1) < PI / 2 Then

                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 1188
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY

                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 9380
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X

                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY

                End If
            ElseIf iConfig = 2 Or iConfig = 6 Or iConfig = 10 Or iConfig = 14 Then
                If Abs(dByPointAngle1) < PI / 2.0# Then
                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 1188
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 9380
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                End If
            ElseIf iConfig = 3 Or iConfig = 7 Or iConfig = 11 Or iConfig = 15 Then
                If Abs(dByPointAngle1) > PI / 2.0# Then
                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 1252 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY

                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 9444 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X

                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY

                End If
            ElseIf iConfig = 4 Or iConfig = 8 Or iConfig = 12 Or iConfig = 16 Then
                If Abs(dByPointAngle1) > PI / 2.0# Then
                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 1252 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 9444 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                End If
            End If
        Else 'Structure is Oriented in opposite direction
            If iConfig = 1 Or iConfig = 5 Or iConfig = 9 Or iConfig = 13 Then
                If Abs(dByPointAngle1) < PI / 2.0# Then
                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 9380
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 1188
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                End If
            ElseIf iConfig = 2 Or iConfig = 6 Or iConfig = 10 Or iConfig = 14 Then
                If Abs(dByPointAngle1) < PI / 2.0# Then
                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 9380
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X


                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 1188
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                End If
            ElseIf iConfig = 3 Or iConfig = 7 Or iConfig = 11 Or iConfig = 15 Then
                If Abs(dByPointAngle1) > PI / 2.0# Then
                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 9444 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 1252 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                End If
            ElseIf iConfig = 4 Or iConfig = 8 Or iConfig = 12 Or iConfig = 16 Then
                If Abs(dByPointAngle1) > PI / 2.0# Then
                    dVertOffset = -dSteelWidth / 2.0#
                    dRouteOffset = oBBX.Width + OVERHANG
                    ByPointStructOffset = -dSteelDepth / 2.0#

                    'Config Index 9444 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X

                    'Config Index 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                Else
                    dVertOffset = dSteelWidth / 2.0#
                    dRouteOffset = -OVERHANG
                    ByPointStructOffset = dSteelDepth / 2.0#

                    'Config Index 1252 
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneB = Plane.XY
                    ByPointStructAxisA = Axis.NegativeX
                    ByPointStructAxisB = Axis.X

                    'Config Index 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                End If
            End If
        End If

        Dim dByPointLength As Double
        If iConfig = 1 Or iConfig = 3 Or iConfig = 5 Or iConfig = 7 Or iConfig = 9 Or iConfig = 11 Or iConfig = 13 Or iConfig = 15 Then
            dByPointLength = dLength + dPipeOD / 2 + dSteelDepth + dOverlap.PropValue + dShoeH.PropValue
        Else
            dByPointLength = dLength + dPipeOD / 2 + dOverlap.PropValue - dShoeH.PropValue - oBBX.Height
        End If

        Dim dPlateTh As PropertyValueDouble

        If SupportHelper.PlacementType = PlacementType.PlaceByStruct Then

            Dim Beams_Plane_Offset, Beams_Axis_Offset, VertSec_Plane_Offset, Hor_Sec_Length As Double
            Dim strVertPortName1, strVertPortName2 As String

            Hor_Sec_Length = oBBX.Width + dExt.PropValue + dSteelDepth / 2.0# + dOverlap.PropValue + OVERHANG
            strVertPortName1 = ""
            strVertPortName2 = ""

            If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                Dim oPart As BusinessObject = oPlate.GetRelationship("madeFrom", "part").TargetObjects(0)
                dPlateTh = oPart.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")
            End If

            If (iConfig = 1) Then
                'Config Index 1196
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 10404
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = -OVERHANG
                HorOffset = -dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 2) Then
                'Config Index 1196
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 10404
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = -Hor_Sec_Length + oBBX.Width + OVERHANG
                HorOffset = -dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 3) Then
                'Config Index 1260
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 2276
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = oBBX.Width + OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            If (iConfig = 4) Then
                'Config Index 1260
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 2276
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = Hor_Sec_Length - OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            If (iConfig = 5) Then
                'Config Index 1260
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 10404
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = Hor_Sec_Length - OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth - dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 6) Then
                'Config Index 1260
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 10404
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = oBBX.Width + OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 7) Then
                'Config Index 1196
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 2276
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = -Hor_Sec_Length + oBBX.Width + OVERHANG
                HorOffset = -dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth - dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            If (iConfig = 8) Then
                'Config Index 1196
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.NegativeX
                'Config Index 2276
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = -OVERHANG
                HorOffset = -dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            If (iConfig = 9) Then
                'Config Index 9388
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 2212
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = Hor_Sec_Length - OVERHANG
                HorOffset = -dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 10) Then
                'Config Index 9388
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 2212
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = oBBX.Width + OVERHANG
                HorOffset = -dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 11) Then
                'Config Index 9452
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 10468
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = -Hor_Sec_Length + oBBX.Width + OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            If (iConfig = 12) Then
                'Config Index 9452
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 10468
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = -OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dSteelDepth + dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            If (iConfig = 13) Then
                'Config Index 9452
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 2212
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = -OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 14) Then
                'Config Index 9452
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.ZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 2212
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.NegativeXY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.NegativeY

                dRouteOffset = -Hor_Sec_Length + oBBX.Width + OVERHANG
                HorOffset = oBBX.Height + dShoeH.PropValue

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    VertSec_Plane_Offset = dPlateTh.PropValue
                Else
                    VertSec_Plane_Offset = 0
                End If

                Beams_Plane_Offset = -dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth - dOverlap.PropValue

                strVertPortName1 = "EndCap"
                strVertPortName2 = "BeginCap"
            End If

            If (iConfig = 15) Then
                'Config Index 9388
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 10468
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = oBBX.Width + OVERHANG
                HorOffset = -dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            If (iConfig = 16) Then
                'Config Index 9388
                ConfRoutePlaneA = Plane.XY
                ConfRoutePlaneB = Plane.NegativeZX
                ConfRouteAxisA = Axis.X
                ConfRouteAxisB = Axis.X
                'Config Index 10468
                StructRoutePlaneA = Plane.XY
                StructRoutePlaneB = Plane.XY
                StructRouteAxisA = Axis.X
                StructRouteAxisB = Axis.Y

                dRouteOffset = Hor_Sec_Length - OVERHANG
                HorOffset = -dShoeH.PropValue
                VertSec_Plane_Offset = 0

                Beams_Plane_Offset = dSteelDepth + dOverlap.PropValue
                Beams_Axis_Offset = Hor_Sec_Length - dSteelDepth - dOverlap.PropValue

                strVertPortName1 = "BeginCap"
                strVertPortName2 = "EndCap"
            End If

            oHorizontalSection.SetPropertyValue(Hor_Sec_Length, "IJUAHgrOccLength", "Length")

            If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                'Dim oPart As BusinessObject = oPlate.GetRelationship("madeFrom", "part").TargetObjects(0)
                'dPlateTh = oPart.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")

                oPlate.SetPropertyValue(dBPWidth.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH")
                oPlate.SetPropertyValue(dBPWidth.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH")
                oPlate.SetPropertyValue(dBPHoleInset.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C")
                oPlate.SetPropertyValue(dBPHoleSize.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE")

                'Add Joint Between Structure and Vertical Beam
                JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, strVertPortName1, StructRoutePlaneA, StructRoutePlaneB, StructRouteAxisA, StructRouteAxisB, VertSec_Plane_Offset, 0)

                'Add Joint Between the Plate and the Vertical Beam
                JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, strVertPortName1, Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0#, -dSteelDepth / 2.0#, -dSteelWidth / 2.0#)

            Else
                'Add Joint Between Structure and Vertical Beam
                JointHelper.CreatePrismaticJoint("-1", "Structure", VERT_SECTION, strVertPortName1, StructRoutePlaneA, StructRoutePlaneB, StructRouteAxisA, StructRouteAxisB, 0, 0)

            End If

            ' Add joints between Route and Beam
            JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HOR_SECTION, "EndCap", ConfRoutePlaneA, ConfRoutePlaneB, ConfRouteAxisA, ConfRouteAxisB, HorOffset, dRouteOffset)

            'Add Joint Between the Horizontal and Vertical Beams
            If iConfig = 1 Or iConfig = 2 Or iConfig = 3 Or iConfig = 4 Or iConfig = 9 Or iConfig = 10 Or iConfig = 11 Or iConfig = 12 Then
                JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, strVertPortName2, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, Beams_Plane_Offset, Beams_Axis_Offset, 0)
            Else
                JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, strVertPortName2, Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, Beams_Plane_Offset, Beams_Axis_Offset, 0)
            End If

            'Flexible Member
            JointHelper.CreatePrismaticJoint(VERT_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0)

        Else
            If Not (SupportHelper.PlacementType = PlacementType.PlaceByReference) Then
                If SupportingHelper.SupportingObjectInfo(1).SupportingObjectType = SupportingObjectType.Member Then
                    If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                        Dim oPart As BusinessObject = oPlate.GetRelationship("madeFrom", "part").TargetObjects(0)
                        dPlateTh = oPart.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")

                        oPlate.SetPropertyValue(dBPWidth.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH")
                        oPlate.SetPropertyValue(dBPWidth.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH")
                        oPlate.SetPropertyValue(dBPHoleInset.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C")
                        oPlate.SetPropertyValue(dBPHoleSize.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE")

                        oVerticalSection.SetPropertyValue(dByPointLength - dPlateTh.PropValue, "IJUAHgrOccLength", "Length")

                        'Add Joint Between Structure and Vertical Beam
                        If iConfig = 1 Or iConfig = 2 Or iConfig = 5 Or iConfig = 6 Or iConfig = 9 Or iConfig = 10 Or iConfig = 13 Or iConfig = 14 Then
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", ByPointStructPlaneA, ByPointStructPlaneB, ByPointStructAxisA, ByPointStructAxisB, dPlateTh.PropValue, ByPointStructOffset, 0)
                        Else
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "BeginCap", ByPointStructPlaneA, ByPointStructPlaneB, ByPointStructAxisA, ByPointStructAxisB, dPlateTh.PropValue, ByPointStructOffset, 0)
                        End If

                        'Add Joint Between the Plate and the Vertical Beam
                        If iConfig = 1 Or iConfig = 2 Or iConfig = 5 Or iConfig = 6 Or iConfig = 9 Or iConfig = 10 Or iConfig = 13 Or iConfig = 14 Then
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0#, -dSteelDepth / 2.0#, -dSteelWidth / 2.0#)
                        Else
                            JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0#, -dSteelDepth / 2.0#, -dSteelWidth / 2.0#)
                        End If

                    Else
                        oVerticalSection.SetPropertyValue(dByPointLength, "IJUAHgrOccLength", "Length")

                        'Add Joint Between Structure and Vertical Beam
                        If iConfig = 1 Or iConfig = 2 Or iConfig = 5 Or iConfig = 6 Or iConfig = 9 Or iConfig = 10 Or iConfig = 13 Or iConfig = 14 Then
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", ByPointStructPlaneA, ByPointStructPlaneB, ByPointStructAxisA, ByPointStructAxisB, 0.0#, ByPointStructOffset, 0)
                        Else
                            JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "BeginCap", ByPointStructPlaneA, ByPointStructPlaneB, ByPointStructAxisA, ByPointStructAxisB, 0.0#, ByPointStructOffset, 0)
                        End If

                    End If

                    'Add Joint Between the Horizontal and Vertical Beams
                    If iConfig = 1 Or iConfig = 5 Or iConfig = 9 Or iConfig = 13 Then
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, dSteelDepth + dOverlap.PropValue, dOverlap.PropValue + dSteelDepth, 0)
                    ElseIf iConfig = 2 Or iConfig = 6 Or iConfig = 10 Or iConfig = 14 Then
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, -dOverlap.PropValue, -(dOverlap.PropValue + dSteelDepth), 0)
                    ElseIf iConfig = 3 Or iConfig = 7 Or iConfig = 11 Or iConfig = 15 Then
                        JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, dSteelDepth + dOverlap.PropValue, -(dOverlap.PropValue + dSteelDepth), 0)
                    Else
                        JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "EndCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, -dOverlap.PropValue, dOverlap.PropValue + dSteelDepth, 0)
                    End If

                    ' Add joints between Route and Beam
                    If iConfig = 1 Or iConfig = 4 Or iConfig = 5 Or iConfig = 8 Or iConfig = 9 Or iConfig = 12 Or iConfig = 13 Or iConfig = 16 Then
                        JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", ByPointRoutePlaneA, ByPointRoutePlaneB, dRouteOffset)
                    Else
                        JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "BeginCap", ByPointRoutePlaneA, ByPointRoutePlaneB, dRouteOffset)
                    End If

                    'Flexible Member
                    JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0)

                End If
            End If
            Dim bSlabOrPlaceByRef As Boolean

            If SupportHelper.PlacementType = PlacementType.PlaceByReference Then
                bSlabOrPlaceByRef = True
            ElseIf SupportingHelper.SupportingObjectInfo(1).SupportingObjectType = SupportingObjectType.Slab Then
                bSlabOrPlaceByRef = True
            Else
                bSlabOrPlaceByRef = False
            End If

            If bSlabOrPlaceByRef Then
                oHorizontalSection.SetPropertyValue(oBBX.Width + dExt.PropValue + dSteelDepth / 2.0# + dOverlap.PropValue + OVERHANG, "IJUAHgrOccLength", "Length")

                If iConfig = 1 Or iConfig = 9 Then
                    dVertOffset = dExt.PropValue
                    dRouteOffset = oBBX.Width + OVERHANG
                    'Config Index = 1189
                    StructRoutePlaneA = Plane.ZX
                    StructRoutePlaneB = Plane.NegativeXY
                    StructRouteAxisA = Axis.X
                    StructRouteAxisB = Axis.NegativeX
                    'Config Index = 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                    'Config Index = 1188
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.NegativeX
                End If

                If iConfig = 2 Or iConfig = 10 Then
                    dVertOffset = dExt.PropValue
                    dRouteOffset = oBBX.Width + OVERHANG
                    'Config Index = 1253
                    StructRoutePlaneA = Plane.ZX
                    StructRoutePlaneB = Plane.XY
                    StructRouteAxisA = Axis.X
                    StructRouteAxisB = Axis.NegativeX
                    'Config Index = 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                    'Config Index = 1188
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.NegativeX
                End If

                If iConfig = 3 Or iConfig = 11 Then
                    dVertOffset = -dExt.PropValue
                    dRouteOffset = -OVERHANG
                    'Config Index = 2356
                    StructRoutePlaneA = Plane.XY
                    StructRoutePlaneB = Plane.NegativeYZ
                    StructRouteAxisA = Axis.Y
                    StructRouteAxisB = Axis.NegativeY
                    'Config Index = 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                    'Config Index = 2212
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.NegativeY
                End If

                If iConfig = 4 Or iConfig = 12 Then
                    dVertOffset = -dExt.PropValue
                    dRouteOffset = -OVERHANG
                    'Config Index = 2406
                    StructRoutePlaneA = Plane.YZ
                    StructRoutePlaneB = Plane.XY
                    StructRouteAxisA = Axis.Y
                    StructRouteAxisB = Axis.NegativeY
                    'Config Index = 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                    'Config Index = 2212
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.NegativeY
                End If

                If iConfig = 5 Or iConfig = 13 Then
                    dVertOffset = -dExt.PropValue
                    dRouteOffset = -OVERHANG
                    'Config Index = 1189
                    StructRoutePlaneA = Plane.ZX
                    StructRoutePlaneB = Plane.NegativeXY
                    StructRouteAxisA = Axis.X
                    StructRouteAxisB = Axis.NegativeX
                    'Config Index = 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                    'Config Index = 9380
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X
                End If

                If iConfig = 6 Or iConfig = 14 Then
                    dVertOffset = -dExt.PropValue
                    dRouteOffset = -OVERHANG
                    'Config Index = 1253
                    StructRoutePlaneA = Plane.ZX
                    StructRoutePlaneB = Plane.XY
                    StructRouteAxisA = Axis.X
                    StructRouteAxisB = Axis.NegativeX
                    'Config Index = 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                    'Config Index = 9380
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.X
                End If

                If iConfig = 7 Or iConfig = 15 Then
                    dVertOffset = dExt.PropValue
                    dRouteOffset = oBBX.Width + OVERHANG
                    'Config Index = 2356
                    StructRoutePlaneA = Plane.XY
                    StructRoutePlaneB = Plane.NegativeYZ
                    StructRouteAxisA = Axis.Y
                    StructRouteAxisB = Axis.NegativeY
                    'Config Index = 101
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.XY
                    'Config Index = 10404
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.Y
                End If

                If iConfig = 8 Or iConfig = 16 Then
                    dVertOffset = dExt.PropValue
                    dRouteOffset = oBBX.Width + OVERHANG
                    'Config Index = 2406
                    StructRoutePlaneA = Plane.YZ
                    StructRoutePlaneB = Plane.XY
                    StructRouteAxisA = Axis.Y
                    StructRouteAxisB = Axis.NegativeY
                    'Config Index = 37
                    ByPointRoutePlaneA = Plane.ZX
                    ByPointRoutePlaneB = Plane.NegativeXY
                    'Config Index = 10404
                    ByPointStructPlaneA = Plane.XY
                    ByPointStructPlaneA = Plane.NegativeXY
                    ByPointStructAxisA = Axis.X
                    ByPointStructAxisB = Axis.Y
                End If

                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    Dim oPart As BusinessObject = oPlate.GetRelationship("madeFrom", "part").TargetObjects(0)
                    dPlateTh = oPart.GetPropertyValue("IJUAHgrUtility_FOUR_HOLE_PLATE", "THICKNESS")

                    oPlate.SetPropertyValue(dBPWidth.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "WIDTH")
                    oPlate.SetPropertyValue(dBPWidth.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "DEPTH")
                    oPlate.SetPropertyValue(dBPHoleInset.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "C")
                    oPlate.SetPropertyValue(dBPHoleSize.PropValue, "IJOAHgrUtility_FOUR_HOLE_PLATE", "HOLE_SIZE")

                    oVerticalSection.SetPropertyValue(dByPointLength - dPlateTh.PropValue, "IJUAHgrOccLength", "Length")

                    'Add Joint Between Structure and Vertical Beam
                    JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", ByPointStructPlaneA, ByPointStructPlaneB, ByPointStructAxisA, ByPointStructAxisB, dPlateTh.PropValue, dVertOffset, 0)

                    'Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(PLATE, "TopStructure", VERT_SECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0#, -dSteelDepth / 2.0#, -dSteelWidth / 2.0#)
                Else
                    oVerticalSection.SetPropertyValue(dByPointLength, "IJUAHgrOccLength", "Length")

                    'Add Joint Between Structure and Vertical Beam
                    JointHelper.CreateRigidJoint("-1", "Structure", VERT_SECTION, "EndCap", ByPointStructPlaneA, ByPointStructPlaneB, ByPointStructAxisA, ByPointStructAxisB, 0, dVertOffset, 0)
                End If

                'Add Joint Between the Horizontal and Vertical Beams
                If iConfig = 1 Or iConfig = 3 Or iConfig = 5 Or iConfig = 7 Or iConfig = 9 Or iConfig = 11 Or iConfig = 13 Or iConfig = 15 Then
                    JointHelper.CreateRigidJoint(HOR_SECTION, "BeginCap", VERT_SECTION, "BeginCap", StructRoutePlaneA, StructRoutePlaneB, StructRouteAxisA, StructRouteAxisB, dOverlap.PropValue + dSteelDepth, dOverlap.PropValue + dSteelDepth, 0)
                Else
                    JointHelper.CreateRigidJoint(HOR_SECTION, "EndCap", VERT_SECTION, "BeginCap", StructRoutePlaneA, StructRoutePlaneB, StructRouteAxisA, StructRouteAxisB, -dOverlap.PropValue, -(dOverlap.PropValue + dSteelDepth), 0)
                End If

                ' Add joints between Route and Beam
                If iConfig = 1 Or iConfig = 3 Or iConfig = 5 Or iConfig = 7 Or iConfig = 9 Or iConfig = 11 Or iConfig = 13 Or iConfig = 15 Then
                    JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "EndCap", ByPointRoutePlaneA, ByPointRoutePlaneB, dRouteOffset)
                Else
                    JointHelper.CreatePlanarJoint("-1", "BBR_Low", HOR_SECTION, "BeginCap", ByPointRoutePlaneA, ByPointRoutePlaneB, dRouteOffset)
                End If

                'Flexible Member
                JointHelper.CreatePrismaticJoint(HOR_SECTION, "BeginCap", HOR_SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0)

            End If

        End If

    End Sub

    Public Overrides ReadOnly Property SupportedConnections() As Collection(Of ConnectionInfo)
        Get
            Dim oConns As New Collection(Of ConnectionInfo)
            For i As Integer = 1 To SupportHelper.SupportedObjects.Count
                oConns.Add(New ConnectionInfo(HOR_SECTION, i))
            Next

            Return oConns
        End Get
    End Property

    Public Overrides ReadOnly Property SupportingConnections() As Collection(Of ConnectionInfo)
        Get
            Dim oConns As New Collection(Of ConnectionInfo)
            For i As Integer = 1 To SupportHelper.SupportingObjects.Count
                If (m_sBasePlatesOpt.PropValue.ToLower = "with") Then
                    oConns.Add(New ConnectionInfo(PLATE, i))
                Else
                    oConns.Add(New ConnectionInfo(VERT_SECTION, i))
                End If
            Next
            Return oConns
        End Get
    End Property

    Public Function BOMDescription(ByVal oSupportOrComponent As Common.Middle.BusinessObject) As String Implements ICustomHgrBOMDescription.BOMDescription

        Dim oPipeSupp As PipeSupport = oSupportOrComponent

        'Get the Support Definition
        Dim oSupportDefinition As Part = SupportHelper.Support.SupportDefinition
        Dim oBeam1, oBeam2 As SupportComponent
        Dim dLength1, dLength2 As PropertyValueDouble
        Dim dictionary = SupportHelper.SupportComponentDictionary
        oBeam1 = dictionary(VERT_SECTION)
        oBeam2 = dictionary(HOR_SECTION)
        dLength1 = oBeam1.GetPropertyValue("IJUAHgrOccLength", "Length")
        dLength2 = oBeam2.GetPropertyValue("IJUAHgrOccLength", "Length")

        'Get the Vertical Length from Route to Structure
        Dim dVLength As Double
        Dim sLength1, sLength2, sVLength As String
        dVLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical)

        'Format the distances into proper units
        sLength1 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dLength1.PropValue, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)
        sLength2 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dLength2.PropValue, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)
        sVLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, dVLength, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)

        'Get The Primary Pipe Information and Specification
        Dim oPipeInfo As PipeObjectInfo = SupportedHelper.SupportedObjectInfo(1)
        Dim oPipeSpec As SpecificationBase = oPipeInfo.Spec

        'Generate the BOM Description
        Return "From CSD: " & oSupportDefinition.PartDescription & ", " & oPipeInfo.NominalDiameter.Size & " " & oPipeInfo.NominalDiameter.Units & " NPD, " & oPipeSpec.SpecificationName & ", Vertical section length = " & sLength1 & ", Horizontal section length  = " & sLength2 & ", Vertical Length = " & sVLength

    End Function
    
    Public Overrides Function MirroredConfiguration(ByVal CurrentMirrorToggleValue As Integer, ByVal eMirrorPlane As Ingr.SP3D.Support.Middle.MirrorPlane) As Integer
        '    Return MyBase.MirroredConfiguration(CurrentMirrorToggleValue, eMirrorPlane)

        If (SupportHelper.PlacementType = PlacementType.PlaceByPoint) Or (SupportHelper.PlacementType = PlacementType.PlaceByReference) Then

            If (SupportingHelper.SupportingObjectInfo(1) IsNot Nothing) AndAlso (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType = SupportingObjectType.Member) Then

                If CurrentMirrorToggleValue = 1 Or CurrentMirrorToggleValue = 5 Or CurrentMirrorToggleValue = 9 Or CurrentMirrorToggleValue = 13 Then
                    MirroredConfiguration = 3
                ElseIf CurrentMirrorToggleValue = 2 Or CurrentMirrorToggleValue = 6 Or CurrentMirrorToggleValue = 10 Or CurrentMirrorToggleValue = 14 Then
                    MirroredConfiguration = 4
                ElseIf CurrentMirrorToggleValue = 3 Or CurrentMirrorToggleValue = 7 Or CurrentMirrorToggleValue = 11 Or CurrentMirrorToggleValue = 15 Then
                    MirroredConfiguration = 1
                Else
                    MirroredConfiguration = 2
                End If

            Else

                If eMirrorPlane = MirrorPlane.XZPlane Then
                    If CurrentMirrorToggleValue = 1 Or CurrentMirrorToggleValue = 9 Then
                        MirroredConfiguration = 3
                    ElseIf CurrentMirrorToggleValue = 2 Or CurrentMirrorToggleValue = 10 Then
                        MirroredConfiguration = 4
                    ElseIf CurrentMirrorToggleValue = 3 Or CurrentMirrorToggleValue = 11 Then
                        MirroredConfiguration = 1
                    ElseIf CurrentMirrorToggleValue = 4 Or CurrentMirrorToggleValue = 12 Then
                        MirroredConfiguration = 2
                    ElseIf CurrentMirrorToggleValue = 5 Or CurrentMirrorToggleValue = 13 Then
                        MirroredConfiguration = 7
                    ElseIf CurrentMirrorToggleValue = 6 Or CurrentMirrorToggleValue = 14 Then
                        MirroredConfiguration = 8
                    ElseIf CurrentMirrorToggleValue = 7 Or CurrentMirrorToggleValue = 15 Then
                        MirroredConfiguration = 5
                    ElseIf CurrentMirrorToggleValue = 8 Or CurrentMirrorToggleValue = 16 Then
                        MirroredConfiguration = 6
                    End If
                Else
                    If CurrentMirrorToggleValue = 1 Or CurrentMirrorToggleValue = 9 Then
                        MirroredConfiguration = 7
                    ElseIf CurrentMirrorToggleValue = 2 Or CurrentMirrorToggleValue = 10 Then
                        MirroredConfiguration = 8
                    ElseIf CurrentMirrorToggleValue = 3 Or CurrentMirrorToggleValue = 11 Then
                        MirroredConfiguration = 5
                    ElseIf CurrentMirrorToggleValue = 4 Or CurrentMirrorToggleValue = 12 Then
                        MirroredConfiguration = 6
                    ElseIf CurrentMirrorToggleValue = 5 Or CurrentMirrorToggleValue = 13 Then
                        MirroredConfiguration = 3
                    ElseIf CurrentMirrorToggleValue = 6 Or CurrentMirrorToggleValue = 14 Then
                        MirroredConfiguration = 4
                    ElseIf CurrentMirrorToggleValue = 7 Or CurrentMirrorToggleValue = 15 Then
                        MirroredConfiguration = 1
                    ElseIf CurrentMirrorToggleValue = 8 Or CurrentMirrorToggleValue = 16 Then
                        MirroredConfiguration = 2
                    End If
                End If
            End If
        Else
            If eMirrorPlane = MirrorPlane.XZPlane Then
                If CurrentMirrorToggleValue = 1 Then
                    MirroredConfiguration = 7
                ElseIf CurrentMirrorToggleValue = 2 Then
                    MirroredConfiguration = 8
                ElseIf CurrentMirrorToggleValue = 3 Then
                    MirroredConfiguration = 5
                ElseIf CurrentMirrorToggleValue = 4 Then
                    MirroredConfiguration = 6
                ElseIf CurrentMirrorToggleValue = 5 Then
                    MirroredConfiguration = 3
                ElseIf CurrentMirrorToggleValue = 6 Then
                    MirroredConfiguration = 4
                ElseIf CurrentMirrorToggleValue = 7 Then
                    MirroredConfiguration = 1
                ElseIf CurrentMirrorToggleValue = 8 Then
                    MirroredConfiguration = 2
                ElseIf CurrentMirrorToggleValue = 9 Then
                    MirroredConfiguration = 15
                ElseIf CurrentMirrorToggleValue = 10 Then
                    MirroredConfiguration = 16
                ElseIf CurrentMirrorToggleValue = 11 Then
                    MirroredConfiguration = 13
                ElseIf CurrentMirrorToggleValue = 12 Then
                    MirroredConfiguration = 14
                ElseIf CurrentMirrorToggleValue = 13 Then
                    MirroredConfiguration = 11
                ElseIf CurrentMirrorToggleValue = 14 Then
                    MirroredConfiguration = 12
                ElseIf CurrentMirrorToggleValue = 15 Then
                    MirroredConfiguration = 9
                ElseIf CurrentMirrorToggleValue = 16 Then
                    MirroredConfiguration = 10
                End If
            Else
                If CurrentMirrorToggleValue = 1 Then
                    MirroredConfiguration = 16
                ElseIf CurrentMirrorToggleValue = 2 Then
                    MirroredConfiguration = 15
                ElseIf CurrentMirrorToggleValue = 3 Then
                    MirroredConfiguration = 14
                ElseIf CurrentMirrorToggleValue = 4 Then
                    MirroredConfiguration = 13
                ElseIf CurrentMirrorToggleValue = 5 Then
                    MirroredConfiguration = 12
                ElseIf CurrentMirrorToggleValue = 6 Then
                    MirroredConfiguration = 11
                ElseIf CurrentMirrorToggleValue = 7 Then
                    MirroredConfiguration = 10
                ElseIf CurrentMirrorToggleValue = 8 Then
                    MirroredConfiguration = 9
                ElseIf CurrentMirrorToggleValue = 9 Then
                    MirroredConfiguration = 8
                ElseIf CurrentMirrorToggleValue = 10 Then
                    MirroredConfiguration = 7
                ElseIf CurrentMirrorToggleValue = 11 Then
                    MirroredConfiguration = 6
                ElseIf CurrentMirrorToggleValue = 12 Then
                    MirroredConfiguration = 5
                ElseIf CurrentMirrorToggleValue = 13 Then
                    MirroredConfiguration = 4
                ElseIf CurrentMirrorToggleValue = 14 Then
                    MirroredConfiguration = 3
                ElseIf CurrentMirrorToggleValue = 15 Then
                    MirroredConfiguration = 2
                ElseIf CurrentMirrorToggleValue = 16 Then
                    MirroredConfiguration = 1
                End If
            End If
        End If
        Return MirroredConfiguration
    End Function
End Class