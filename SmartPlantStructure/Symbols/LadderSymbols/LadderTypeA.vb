''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright 1992 - 2009 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  LadderTypeA.vb
'
'Abstract
'	This is .NET LadderTypeA symbol. This class subclasses from LadderSymbolDefinition.
'   17Jul2012 Sankar Sutapalli TR214568 Placing Ladder-A results in mirrored matrix - causing drawing issue
'                                       To follow RHS coordinate system changed the top support normal direction
'                                       to opposite direction of the top support.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports System
Imports System.Collections
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports System.Math
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.Structure.Middle.Services
'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Content.Structure
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------

Namespace Ingr.SP3D.Content.Structure

    Public Class LadderTypeA : Inherits LadderSymbolDefinition
        Implements ICustomWeightCG
        '===============================================================================================
        'DefinitionName/ProgID of this symbol is "LadderSymbols,Ingr.SP3D.Content.Structure.LadderTypeA"
        '===============================================================================================

#Region "Definition of Inputs"
        <InputCatalogPart(1)> _
        Public m_oPartInput As InputCatalogPart
        <InputDouble(2, "Width", "Width", 0.035)> _
        Public m_dWidth As InputDouble
        <InputDouble(3, "Angle", "Angle", 90)> _
        Public m_dAngle As InputDouble
        <InputDouble(4, "StepPitch", "Step Pitch", 0.25)> _
        Public m_dStepPitch As InputDouble
        <InputDouble(5, "SupportLegPitch", "Support Leg Pitch", 0.5)> _
        Public m_dSupportLegPitch As InputDouble
        <InputDouble(6, "SupportLegWidth", "Support Leg Width", 0.065)> _
        Public m_dSupportLegWidth As InputDouble
        <InputDouble(7, "SupportLegThickness", "Support Leg Thickness", 0.016)> _
        Public m_dSupportLegThickness As InputDouble
        <InputDouble(8, "SideFrameWidth", "Side Frame Width", 0.065)> _
        Public m_dSideFrameWidth As InputDouble
        <InputDouble(9, "SideFrameThickness", "Side Frame Thickness", 0.016)> _
        Public m_dSideFrameThickness As InputDouble
        <InputDouble(10, "StepDiameter", "Step Diameter", 0.04)> _
        Public m_dStepDiameter As InputDouble
        <InputDouble(11, "VlDim1", "VlDim1", 0.03)> _
        Public m_dVlDim1 As InputDouble
        <InputDouble(12, "VlDim2", "VlDim2", 0.01)> _
        Public m_dVlDim2 As InputDouble
        <InputDouble(13, "VlDim3", "VlDim3", 0.165)> _
        Public m_dVlDim3 As InputDouble
        <InputDouble(14, "WallOffset", "Wall Offset", 0.135)> _
        Public m_dWallOffset As InputDouble
        <InputDouble(15, "Span", "Span", 1.0)> _
        Public m_dSpan As InputDouble
        <InputDouble(16, "Height", "Height", 1.0)> _
        Public m_dHeight As InputDouble
        <InputDouble(17, "Length", "Length", 1.0)> _
        Public m_dLength As InputDouble
        <InputDouble(18, "WithWallSupports", "With Wall Supports", 1.0)> _
        Public m_dWithWallSupports As InputDouble
        <InputDouble(19, "NumSteps", "Number of Steps", 1.0)> _
        Public m_dNumSteps As InputDouble
        <InputDouble(20, "StepProtrusion", "Step Protrusion", 0.005)> _
        Public m_dStepProtrusion As InputDouble
        <InputDouble(21, "WithSafetyHoop", "With Safety Hoop", 0)> _
        Public m_dWithSafetyHoop As InputDouble
        <InputDouble(22, "HoopPitch", "Hoop Pitch", 2)> _
        Public m_dHoopPitch As InputDouble
        <InputDouble(23, "BottomHoopLevel", "Bottom Hoop Level", 2.2)> _
        Public m_dBottomHoopLevel As InputDouble
        <InputDouble(24, "HoopClearance", "Hoop Clearance", 0.75)> _
        Public m_dHoopClearance As InputDouble
        <InputDouble(25, "HoopRadius", "Hoop Radius", 0.7)> _
        Public m_dHoopRadius As InputDouble
        <InputDouble(26, "HoopPlateThickness", "Hoop Plate Thickness", 0.016)> _
        Public m_dHoopPlateThickness As InputDouble
        <InputDouble(27, "HoopPlateWidth", "Hoop Plate Width", 0.065)> _
        Public m_dHoopPlateWidth As InputDouble
        <InputDouble(28, "HoopBendRadius", "Hoop Bend Radius", 0.05)> _
        Public m_dHoopBendRadius As InputDouble
        <InputDouble(29, "HoopOpening", "Hoop Opening", 1)> _
        Public m_dHoopOpening As InputDouble
        <InputDouble(30, "ShDim1", "ShDim1", 0.07)> _
        Public m_dShDim1 As InputDouble
        <InputDouble(31, "ShDim2", "ShDim2", 0.15)> _
        Public m_dShDim2 As InputDouble
        <InputDouble(32, "ShDim3", "ShDim3", 0.3)> _
        Public m_dShDim3 As InputDouble
        <InputDouble(33, "FlareClearance", "Flare Clearance", 1.0)> _
        Public m_dFlareClearance As InputDouble
        <InputDouble(34, "FlareRadius", "Flare Radius", 1.0)> _
        Public m_dFlareRadius As InputDouble
        <InputDouble(35, "HoopFlareBendRadius", "Hoop Flare Bend Radius", 0.05)> _
        Public m_dHoopFlareBendRadius As InputDouble
        <InputDouble(36, "FlareShDim1", "Flare ShDim1", 0.07)> _
        Public m_dFlareShDim1 As InputDouble
        <InputDouble(37, "FlareShDim2", "Flare ShDim2", 0.3)> _
        Public m_dFlareShDim2 As InputDouble
        <InputDouble(38, "FlareShDim3", "Flare ShDim3", 0.3)> _
        Public m_dFlareShDim3 As InputDouble
        <InputDouble(39, "HoopFlareHeight", "Hoop Flare Height", 2.0)> _
        Public m_dHoopFlareHeight As InputDouble
        <InputDouble(40, "HoopFlareMaxHeight", "Hoop Flare Maximum Height", 4.0)> _
        Public m_dHoopFlareMaxHeight As InputDouble
        <InputDouble(41, "VerticalStrapWidth", "Vertical Strap Width", 0.065)> _
        Public m_dVerticalStrapWidth As InputDouble
        <InputDouble(42, "VerticalStrapThickness", "Vertical Strap Thickness", 0.016)> _
        Public m_dVerticalStrapThickness As InputDouble
        <InputDouble(43, "VerticalStrapCount", "Vertical Strap Count", 4.0)> _
        Public m_dVerticalStrapCount As InputDouble
        <InputDouble(44, "TopExtension", "Top Extension", 0.0)> _
        Public m_dTopExtension As InputDouble
        <InputDouble(45, "BottomExtension", "Bottom Extension", 0.0)> _
        Public m_dBottomExtension As InputDouble
        <InputDouble(46, "Justification", "Justification", 1.0)> _
        Public m_dJustification As InputDouble
        <InputDouble(47, "TopSupportSide", "TopSupportSide", 1.0)> _
        Public m_dTopSupportSide As InputDouble
        <InputDouble(48, "IsAssembly", "Is Assembly", 0.0)> _
        Public m_dIsAssembly As InputDouble
        <InputDouble(49, "EnvelopeHeight", "Envelope Height", 0.0)> _
        Public m_dEnvelopeHeight As InputDouble
        <InputString(50, "Primary_SPSMaterial", "Primary Material", "Steel - Carbon")> _
        Public m_sPrimaryMaterial As InputString
        <InputString(51, "Primary_SPSGrade", "Primary Material Grade", "A")> _
        Public m_sPrimaryMaterialGrade As InputString
#End Region

#Region "Definitions of Aspects and their outputs"
        'SimplePhysical Aspect
        <Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)> _
        Public m_oSimplePhysicalAspect As AspectDefinition

        'Operation Aspect
        <Aspect("Operation", "Operation Aspect", AspectID.Operation)> _
        <SymbolOutput("OperationalEnvelope1", "Operational envelope of the vertical ladder")> _
        Public m_oOperationAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

        Protected Overrides Sub ConstructOutputs()
            Try
                '=================================================
                ' Construction of SimplePhysical Aspect 
                '=================================================
                Dim dPosition As Double
                Dim nNumberOfHoops As Integer, nNumberOfHoopsOpen As Integer

                'getting the inputs
                Dim dLength As Double = m_dLength.Value
                Dim dHeight As Double = m_dHeight.Value
                Dim dAngle As Double = m_dAngle.Value
                Dim dTopExtension As Double = m_dTopExtension.Value
                Dim dSpan As Double = m_dSpan.Value
                Dim dWidth As Double = m_dWidth.Value
                Dim dWallOffset As Double = m_dWallOffset.Value
                Dim dSideFrameWidth As Double = m_dSideFrameWidth.Value
                Dim dSideFrameThickness As Double = m_dSideFrameThickness.Value
                Dim dSupportLegPitch As Double = m_dSupportLegPitch.Value
                Dim dStepDiameter As Double = m_dStepDiameter.Value
                Dim dHoopPlateWidth As Double = m_dHoopPlateWidth.Value
                Dim dBottomHoopLevel As Double = m_dBottomHoopLevel.Value
                Dim dHoopPitch As Double = m_dHoopPitch.Value
                Dim dEnvelopeHeight As Double = m_dEnvelopeHeight.Value
                Dim nJustification As Integer = CInt(m_dJustification.Value)
                Dim nHoopOpening As Integer = CInt(m_dHoopOpening.Value)
                Dim bWithWallSupports As Boolean = CBool(m_dWithWallSupports.Value)
                Dim bWithSafetyHoop As Boolean = CBool(m_dWithSafetyHoop.Value)

                'Get the connection where outputs will be created.
                Dim oConnection As SP3DConnection = OccurrenceConnection

                Dim dVerticalOffset As Double = (dHeight - dTopExtension - dSpan)
                Dim dTopSupportLocation As Double = -(dVerticalOffset + dTopExtension)

                'Set ladder position based on justification
                If nJustification = SPSSymbolConstants.ALIGNMENT_LEFT Then
                    dPosition = ((dWidth / 2) + dSideFrameThickness)
                ElseIf nJustification = SPSSymbolConstants.ALIGNMENT_RIGHT Then
                    dPosition = -((dWidth / 2) + dSideFrameThickness)
                End If

                'Place left side frame
                PlaceLeftSideFrame(dWidth, dSideFrameThickness, dPosition, dWallOffset, dSideFrameWidth, _
                                   dLength, dAngle)

                'Place right side frame
                PlaceRightSideFrame(dWidth, dSideFrameThickness, dPosition, dSideFrameWidth, dLength, _
                                    dAngle, dWallOffset)

                'Place rungs
                PlaceRungs(dLength, dTopExtension, dVerticalOffset, dSupportLegPitch, dSideFrameThickness, _
                           dPosition, dStepDiameter, dWidth, dWallOffset, dSideFrameWidth, dAngle)

                'Place wall supports
                If (bWithWallSupports) And dWallOffset <> 0 Then
                    PlaceWallSupport(dWallOffset, dSideFrameWidth, dSideFrameThickness, dTopExtension, _
                                     dLength, dVerticalOffset, dAngle, dSupportLegPitch, dWidth, dPosition)
                End If

                'Place hoops
                Dim hoopCurves As ComplexString3d = Nothing, flareCurves As ComplexString3d = Nothing
                If (bWithSafetyHoop) Then
                    PlaceHoop(dStepDiameter, dSideFrameThickness, dAngle, dHoopPlateWidth, dWidth, _
                              dLength, dBottomHoopLevel, dHoopPitch, dTopSupportLocation, nHoopOpening, dPosition, _
                               dWallOffset, dSideFrameWidth, nNumberOfHoops, nNumberOfHoopsOpen, hoopCurves, flareCurves)
                End If

                'Place vertical strap
                If (bWithSafetyHoop) And nNumberOfHoops > 1 Then
                    PlaceVerticalStrap(dStepDiameter, dHoopPitch, dLength, dAngle, dHoopPlateWidth, _
                                       nHoopOpening, nNumberOfHoops, nNumberOfHoopsOpen, hoopCurves, flareCurves)
                End If

                '===========================================================
                ' Construction of Net Volume and Center of Gravity
                '===========================================================
                Dim outputsOnAspect As OutputDictionary = m_oSimplePhysicalAspect.Outputs
                Dim outputGeometries As Collection(Of Projection3d) = New Collection(Of Projection3d)
                For Each output As DictionaryEntry In outputsOnAspect
                    Dim outputGeometry As Projection3d = TryCast(output.Value, Projection3d)
                    If Not outputGeometry Is Nothing Then
                        outputGeometries.Add(outputGeometry)
                    End If
                Next
                'Get the material and pass it as an argument to the AddWeigthCG() method as the method needs material to get the density and calculate the weight.
                Dim catalogStructHelper As CatalogStructHelper = New CatalogStructHelper()
                Dim materialName$ = m_sPrimaryMaterial.Value
                Dim materialGrade$ = m_sPrimaryMaterialGrade.Value
                Dim material As Material = catalogStructHelper.GetMaterial(materialName, materialGrade)
                If outputGeometries.Count > 0 Then
                    'As all surface are Projectoin3D, we can accumulate their weight through one collection.
                    MyBase.AddWeightCG(outputGeometries, material)

                    ' Add the weightCG as an output which can be retrieved later in EvaluateWeightCG
                    MyBase.SetWeightCOG(m_oSimplePhysicalAspect)
                End If

                '=================================================
                ' Construction of Operation Aspect 
                '=================================================
                'Place Operational envelope1
                PlaceOperationalEnvelope1(bWithSafetyHoop, dWidth, dLength, _
                                          dPosition, dSideFrameThickness, dEnvelopeHeight)

                'Place Operational envelope2
                If bWithSafetyHoop Then
                    PlaceOperationalEnvelope2(dLength, dPosition, dEnvelopeHeight)
                End If

            Catch Ex As Exception ' General Unhandled exception 
                If MyBase.ToDoListMessage Is Nothing Then 'Check ToDoListMessgae created already or not
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureErrorsMessageCodelist, _
                                                          SPSSymbolConstants.TDL_INVALID_LADDER_GEOMETRY, String.Format(LadderSymbolsLocalizer.GetString(LadderSymbolsResourceIDs.ErrConstructOutputs, "Error in constructing outputs for Ladder in {0}. Check custom code or contact S3D support."), Me.ToString()))
                End If
            End Try
        End Sub

#End Region

#Region "Private Functions and Methods"

#Region "Left Side Frame Creation"

        Private Sub PlaceLeftSideFrame(ByVal dWidth As Double, _
                                       ByVal dSideFrameThickness As Double, ByVal dPosition As Double, _
                                       ByVal dWallOffset As Double, ByVal dSideFrameWidth As Double, _
                                       ByVal dLength As Double, ByVal dAngle As Double)

            Dim aUpperLeftCorner(2) As Double
            Dim oTempPosition As New Position()
            Dim oProjectionVector As New Vector, oUpperVector As New Vector

            aUpperLeftCorner(0) = 0 - (dWidth / 2.0) - dSideFrameThickness / 2.0 - dPosition
            aUpperLeftCorner(1) = dWallOffset + dSideFrameWidth / 2.0
            aUpperLeftCorner(2) = 0

            oTempPosition.X = aUpperLeftCorner(0)
            oTempPosition.Y = aUpperLeftCorner(1)
            oTempPosition.Z = aUpperLeftCorner(2)

            oProjectionVector.X = 0
            oProjectionVector.Y = (dLength) * Math.Tan(Math.PI / 2 - dAngle)
            oProjectionVector.Z = -dLength

            oUpperVector.X = 0
            oUpperVector.Y = 1.0
            oUpperVector.Z = 1 * Tan(PI / 2 - dAngle)

            Dim oProjection3d As Projection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oUpperVector, dSideFrameThickness, dSideFrameWidth)
            m_oSimplePhysicalAspect.Outputs.Add("LeftSideFrame1", oProjection3d)

        End Sub

#End Region

#Region "Right Side Frame Creation"

        Private Sub PlaceRightSideFrame(ByVal dWidth As Double, _
                                        ByVal dSideFrameThickness As Double, ByVal dPosition As Double, _
                                        ByVal dSideFrameWidth As Double, ByVal dLength As Double, _
                                        ByVal dAngle As Double, ByVal dWallOffset As Double)

            Dim aUpperRightCorner(2) As Double
            Dim oTempPosition As New Position()
            Dim oProjectionVector As New Vector, oUpperVector As New Vector

            aUpperRightCorner(0) = 0 + dWidth / 2.0 + dSideFrameThickness / 2.0 - dPosition
            aUpperRightCorner(1) = dWallOffset + dSideFrameWidth / 2.0
            aUpperRightCorner(2) = 0

            oTempPosition.X = aUpperRightCorner(0)
            oTempPosition.Y = aUpperRightCorner(1)
            oTempPosition.Z = aUpperRightCorner(2)

            oProjectionVector.X = 0
            oProjectionVector.Y = (dLength) * Tan(PI / 2 - dAngle)
            oProjectionVector.Z = -dLength

            oUpperVector.X = 0
            oUpperVector.Y = 1.0
            oUpperVector.Z = 1 * Math.Tan(Math.PI / 2 - dAngle)

            Dim oProjection3d As Projection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oUpperVector, dSideFrameThickness, dSideFrameWidth)
            m_oSimplePhysicalAspect.Outputs.Add("RightSideFrame1", oProjection3d)
        End Sub

#End Region

#Region "Rungs Creation"

        Private Sub PlaceRungs(ByVal dLength As Double, _
                               ByVal dTopExtension As Double, ByVal dVerticalOffset As Double, _
                               ByVal dSupportLegPitch As Double, ByVal dSideFrameThickness As Double, _
                               ByVal dPosition As Double, ByVal dStepDiameter As Double, ByVal dWidth As Double, _
                               ByVal dWallOffset As Double, ByVal dSideFrameWidth As Double, ByVal dAngle As Double)

            Dim oTempPosition As New Position()
            Dim oProjectionVector As New Vector, oUpperVector As New Vector
            Dim oProjection3d As Projection3d
            Dim aUpperLeftCorner(2) As Double

            Dim nNumberOfSteps As Integer = CInt(m_dNumSteps.Value)
            Dim dStepProtrusion As Double = m_dStepProtrusion.Value
            Dim dStepPitch As Double = m_dStepPitch.Value

            aUpperLeftCorner(0) = 0 - (dWidth / 2.0) - dSideFrameThickness / 2.0 - dPosition
            aUpperLeftCorner(1) = dWallOffset + dSideFrameWidth / 2.0
            aUpperLeftCorner(2) = 0

            'Account for vertical offset. This will allow the wall supports to be place correctly.
            Dim nSupports As Integer = CInt((dLength - (dTopExtension + dVerticalOffset)) / dSupportLegPitch)
            If nSupports <= 0 Then
                nSupports = 1
            End If

            'Set start pos for steps
            oTempPosition.X = aUpperLeftCorner(0) - dSideFrameThickness / 2.0 - dStepProtrusion

            'Align the top of the rung with the top reference surface.
            oTempPosition.Y = aUpperLeftCorner(1)
            oTempPosition.Z = aUpperLeftCorner(2) - dTopExtension - dStepDiameter / 2

            'Vector length determines projection length and it runs across the width of the ladder
            oProjectionVector.X = dWidth + 2 * (dSideFrameThickness + dStepProtrusion)
            oProjectionVector.Y = 0.0
            oProjectionVector.Z = 0.0

            'Set UpperVector to rotate the step bar on edge to 45 degrees from Z
            oUpperVector.X = 0.0
            oUpperVector.Y = 1.0
            oUpperVector.Z = 1.0

            For iStepNumber = 1 To nNumberOfSteps
                oProjection3d = CreateCircularProjection(oTempPosition, oProjectionVector, oUpperVector, dStepDiameter)
                m_oSimplePhysicalAspect.Outputs.Add("Step" & iStepNumber, oProjection3d)
                oTempPosition.Y = oTempPosition.Y + dStepPitch * Cos(dAngle)
                oTempPosition.Z = oTempPosition.Z - dStepPitch * Sin(dAngle)
            Next iStepNumber

        End Sub

#End Region

#Region "Wall Support Creation"

        Private Sub PlaceWallSupport(ByVal dWallOffset As Double, _
                                     ByVal dSideFrameWidth As Double, ByVal dSideFrameThickness As Double, _
                                     ByVal dTopExtension As Double, ByVal dLength As Double, _
                                     ByVal dVerticalOffset As Double, ByVal dAngle As Double, _
                                     ByVal dSupportLegPitch As Double, ByVal dWidth As Double, ByVal dPosition As Double)

            Dim aUpperLeftCorner(2) As Double, aUpperRightCorner(2) As Double
            Dim oUpperVector As New Vector, oProjectionVector As New Vector
            Dim oTempPosition As New Position()
            Dim oProjection3d As Projection3d
            Dim dSupportLegThickness As Double = m_dSupportLegThickness.Value
            Dim dSupportLegWidth As Double = m_dSupportLegWidth.Value

            aUpperLeftCorner(0) = 0 - (dWidth / 2.0) - dSideFrameThickness / 2.0 - dPosition
            aUpperLeftCorner(1) = dWallOffset + dSideFrameWidth / 2.0
            aUpperLeftCorner(2) = 0

            aUpperRightCorner(0) = 0 + dWidth / 2.0 + dSideFrameThickness / 2.0 - dPosition
            aUpperRightCorner(1) = dWallOffset + dSideFrameWidth / 2.0
            aUpperRightCorner(2) = 0

            Dim dSupportLegLength As Double = dWallOffset + (dSideFrameWidth / 2.0)

            'Account for vertical offset. This will allow the wall supports to be place correctly.
            Dim nSupports As Integer = CInt((dLength - (dTopExtension + dVerticalOffset)) / dSupportLegPitch)
            If nSupports <= 0 Then
                nSupports = 1
            End If

            oUpperVector.X = 0.0
            oUpperVector.Y = 0.0
            oUpperVector.Z = 1.0

            oTempPosition.X = aUpperLeftCorner(0) - dSideFrameThickness / 2.0 - dSupportLegThickness / 2.0
            oTempPosition.Y = 0
            oTempPosition.Z = aUpperLeftCorner(2) - 0.01016 - dSupportLegWidth / 2.0 - (dTopExtension + dVerticalOffset)

            oProjectionVector.X = 0
            oProjectionVector.Y = dSupportLegLength
            oProjectionVector.Z = 0.0

            'Left side support leg
            For i = 1 To nSupports
                oProjectionVector.Y = dSupportLegLength - oTempPosition.Z / Math.Tan(dAngle)
                oProjection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oUpperVector, dSupportLegThickness, dSupportLegWidth)
                m_oSimplePhysicalAspect.Outputs.Add("LeftSupportLeg" & i, oProjection3d)
                oTempPosition.Z = oTempPosition.Z - dSupportLegPitch
            Next

            'Right side support leg
            oTempPosition.X = aUpperRightCorner(0) + dSideFrameThickness / 2.0 + dSupportLegThickness / 2.0
            oTempPosition.Y = 0
            oTempPosition.Z = aUpperRightCorner(2) - 0.01016 - dSupportLegWidth / 2.0 - (dTopExtension + dVerticalOffset)

            For i = 1 To nSupports
                oProjectionVector.Y = dSupportLegLength - oTempPosition.Z / Math.Tan(dAngle)
                oProjection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oUpperVector, dSupportLegThickness, dSupportLegWidth)
                m_oSimplePhysicalAspect.Outputs.Add("RightSupportLeg" & i, oProjection3d)
                oTempPosition.Z = oTempPosition.Z - dSupportLegPitch
            Next

        End Sub

#End Region

#Region "Hoop Creation"

        Private Sub PlaceHoop(ByVal dStepDiameter As Double, _
                              ByVal dSideFrameThickness As Double, ByVal dAngle As Double, _
                              ByVal dHoopPlateWidth As Double, ByVal dWidth As Double, _
                              ByVal dLength As Double, ByVal dBottomHoopLevel As Double, _
                              ByVal dHoopPitch As Double, ByVal dTopSupportLocation As Double, _
                              ByVal nHoopOpening As Integer, ByVal dPosition As Double, _
                              ByVal dWallOffset As Double, ByVal dSideFrameWidth As Double, _
                              ByRef nNumberOfHoops As Integer, ByRef nNumberOfHoopsOpen As Integer, _
                              ByRef oHoopComplex As ComplexString3d, ByRef oFlareComplex As ComplexString3d)

            Dim aShDim(2) As Double, aFlareShDim(2) As Double, aUpperLeftCorner(2) As Double
            Dim oProjectionVector As New Vector, oUpperVector As New Vector
            Dim oTempPosition As New Position()
            Dim oProjection3d As Projection3d
            Dim dShDim1 As Double = m_dShDim1.Value
            Dim dShDim2 As Double = m_dShDim2.Value
            Dim dShDim3 As Double = m_dShDim3.Value
            Dim dFlareShDim1 As Double = m_dFlareShDim1.Value
            Dim dFlareShDim2 As Double = m_dFlareShDim2.Value
            Dim dFlareShDim3 As Double = m_dFlareShDim3.Value
            Dim dHoopPlateThickness As Double = m_dHoopPlateThickness.Value
            Dim dHoopBendRadius As Double = m_dHoopBendRadius.Value

            aUpperLeftCorner(0) = 0 - (dWidth / 2.0) - dSideFrameThickness / 2.0 - dPosition
            aUpperLeftCorner(1) = dWallOffset + dSideFrameWidth / 2.0
            aUpperLeftCorner(2) = 0

            Dim dWeldClearance As Double = 3 * dStepDiameter / 2

            oTempPosition.X = aUpperLeftCorner(0) - dSideFrameThickness / 2.0
            oTempPosition.Y = aUpperLeftCorner(1) + (0.127 + dWeldClearance) * Math.Cos(dAngle)
            oTempPosition.Z = aUpperLeftCorner(2) - (0.127 + dWeldClearance) * Math.Sin(dAngle)

            oProjectionVector.X = 0
            oProjectionVector.Y = dHoopPlateWidth * Math.Cos(dAngle)
            oProjectionVector.Z = -(dHoopPlateWidth) * Math.Sin(dAngle)

            oUpperVector.X = 0
            oUpperVector.Y = 1.0
            oUpperVector.Z = 1 * Math.Tan(Math.PI / 2 - dAngle)

            'Calculate the HoopRadius and HoopClearance based on Ladder Width
            Dim dHoopRadius As Double = dShDim2 + (dWidth / 2) + (2 * dHoopBendRadius) + dSideFrameThickness
            Dim dHoopClearance As Double = dHoopRadius + (dShDim3 - dShDim1)

            aShDim(0) = dShDim1
            aShDim(1) = dHoopRadius - (dWidth / 2) - 2 * dHoopBendRadius - dSideFrameThickness
            aShDim(2) = dShDim3

            nNumberOfHoops = CInt((dLength - dBottomHoopLevel) / dHoopPitch)

            'Start placing hoops from top to bottom except the bottom hoop.
            If nNumberOfHoops > 1 Then
                For i = 1 To nNumberOfHoops
                    'Change for HoopOpening
                    If oTempPosition.Z >= dTopSupportLocation Then
                        'Hoops above the top support. These may need side opening.
                        oProjection3d = CreateProjectionOfCurve(oTempPosition, oProjectionVector, oUpperVector, _
                                                                dHoopBendRadius, aShDim, dHoopRadius, dHoopRadius, dHoopPlateThickness, _
                                                                "SafetyHoop", i, oHoopComplex, oFlareComplex, nHoopOpening, dWidth + 2 * dSideFrameThickness)
                        m_oSimplePhysicalAspect.Outputs.Add("SafetyHoop" & i, oProjection3d)
                        nNumberOfHoopsOpen = nNumberOfHoopsOpen + 1
                    Else
                        'Hoops below the top support, treat them as ordinary hoops.
                        oProjection3d = CreateProjectionOfCurve(oTempPosition, oProjectionVector, oUpperVector, _
                                                                dHoopBendRadius, aShDim, dHoopRadius, dHoopRadius, dHoopPlateThickness, _
                                                                "SafetyHoop", i, oHoopComplex, oFlareComplex)
                        m_oSimplePhysicalAspect.Outputs.Add("SafetyHoop" & i, oProjection3d)
                    End If
                    oTempPosition.Z = oTempPosition.Z - dHoopPitch * Math.Sin(dAngle)
                    oTempPosition.Y = oTempPosition.Y + dHoopPitch * Math.Cos(dAngle)
                Next

                'Now place the Bottom Hoop
                'Calculate the FlareRadius based on Ladder Width
                Dim dFlareRadius As Double = dFlareShDim2 + (dWidth / 2) + 2 * dHoopBendRadius + dSideFrameThickness

                aFlareShDim(0) = dFlareShDim1
                aFlareShDim(1) = dFlareRadius - (dWidth / 2) - 2 * dHoopBendRadius - dSideFrameThickness
                aFlareShDim(2) = aShDim(2)
                oProjection3d = CreateProjectionOfCurve(oTempPosition, oProjectionVector, oUpperVector, dHoopBendRadius, _
                                                        aFlareShDim, dFlareRadius, dHoopClearance, dHoopPlateThickness, _
                                                        "FlareSafetyHoop", nNumberOfHoops + 1, oHoopComplex, oFlareComplex)
                m_oSimplePhysicalAspect.Outputs.Add("FlareSafetyHoop" & nNumberOfHoops + 1, oProjection3d)
            End If

        End Sub

#End Region

#Region "Vertical Strap Creation"

        Private Sub PlaceVerticalStrap(ByVal dStepDiameter As Double, _
                                       ByVal dHoopPitch As Double, ByVal dLength As Double, _
                                       ByVal dAngle As Double, ByVal dHoopPlateWidth As Double, _
                                       ByVal nHoopOpening As Integer, ByVal nNumberOfHoops As Integer, _
                                       ByVal nNumberOfHoopsOpen As Integer, ByVal oHoopComplex As ComplexString3d, _
                                       ByVal oFlareComplex As ComplexString3d)

            Dim oTempVector1 As New Vector, oTempProjectionVector As New Vector, oTempVector As New Vector
            Dim dHoopRotAngle As Double, dTheta As Double
            Dim dFlareRotAngle As Double
            Dim oProjectionVector As New Vector, oUpperVector As New Vector
            Dim oHoopArc As Arc3d
            Dim oTempPosition As New Position()
            Dim oProjection3d As Projection3d
            Dim oTmpVector As New Vector

            Dim nVerticalStrapCount As Integer = CInt(m_dVerticalStrapCount.Value)
            Dim dVerticalStrapWidth As Double = m_dVerticalStrapWidth.Value
            Dim dVerticalStrapThickness As Double = m_dVerticalStrapThickness.Value

            Dim aStrapPos(nVerticalStrapCount) As Position
            Dim aFlarePos(nVerticalStrapCount) As Position

            For i = 1 To nVerticalStrapCount
                aStrapPos(i) = New Position
                aFlarePos(i) = New Position
            Next i

            Dim dWeldClearance As Double = 3 * dStepDiameter / 2
            Dim dTempPosZ As Double = 0.0 - (0.127 + dWeldClearance) * Math.Sin(dAngle)

            Dim oCurveColl As Collection(Of ICurve) = Nothing
            oHoopComplex.GetCurves(oCurveColl)
            oHoopArc = DirectCast(oCurveColl(0), Arc3d)
            oCurveColl = Nothing

            Dim dHoopAngle As Double = oHoopArc.SweepAngle
            Dim dHoopRadius As Double = oHoopArc.Radius
            Dim oCenterPosition As Position = oHoopArc.Center

            If nVerticalStrapCount > 1 Then
                dHoopRotAngle = dHoopAngle / (nVerticalStrapCount - 1)
            End If

            For i = 1 To nVerticalStrapCount
                If i = 1 Then
                    If nVerticalStrapCount = 1 Then
                        'we have only one strap. Place this strap at mid position
                        oTmpVector.X = Math.Cos(dHoopAngle / 2) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                        oTmpVector.Y = Math.Sin(dHoopAngle / 2) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                        aStrapPos(i).X = oTmpVector.X + oCenterPosition.X
                        aStrapPos(i).Y = oTmpVector.Y + oCenterPosition.Y
                    Else
                        'first strap
                        dTheta = (dVerticalStrapWidth / 2) / dHoopRadius
                        oTmpVector.X = Math.Cos(dTheta + (dHoopRotAngle * (i - 1))) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                        oTmpVector.Y = Math.Sin(dTheta + (dHoopRotAngle * (i - 1))) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                        aStrapPos(i).X = oTmpVector.X + oCenterPosition.X
                        aStrapPos(i).Y = oTmpVector.Y + oCenterPosition.Y
                    End If

                ElseIf i = nVerticalStrapCount Then
                    'last strap
                    dTheta = (dVerticalStrapWidth / 2) / dHoopRadius
                    oTmpVector.X = Math.Cos(dTheta + (dHoopRotAngle * (i - 1))) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                    oTmpVector.Y = Math.Sin(dTheta + (dHoopRotAngle * (i - 1))) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                    aStrapPos(i).X = oTmpVector.X + oCenterPosition.X
                    aStrapPos(i).Y = oTmpVector.Y + oCenterPosition.Y
                Else
                    'any strap in between
                    oTmpVector.X = Math.Cos(dHoopRotAngle * (i - 1)) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                    oTmpVector.Y = Math.Sin(dHoopRotAngle * (i - 1)) * (dHoopRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                    aStrapPos(i).X = oTmpVector.X + oCenterPosition.X
                    aStrapPos(i).Y = oTmpVector.Y + oCenterPosition.Y
                End If
            Next i

            Dim dSideOpeningStrapZPosition As Double = dTempPosZ - nNumberOfHoopsOpen * dHoopPitch
            oTempPosition.Z = dTempPosZ 'this point on first hoop

            oProjectionVector.X = 0
            oProjectionVector.Y = dLength * Math.Tan((Math.PI / 2 - dAngle))
            oProjectionVector.Z = -((nNumberOfHoops - 1) * dHoopPitch) - (dHoopPlateWidth)

            Dim dSideOpeningStrapLength As Double = oProjectionVector.Z - nNumberOfHoopsOpen * (-dHoopPitch)

            oUpperVector.X = 0
            oUpperVector.Y = 1.0
            oUpperVector.Z = 1 * Math.Tan(Math.PI / 2 - dAngle)

            'Change for HoopOpening
            For i = 1 To nVerticalStrapCount
                oTempPosition.X = aStrapPos(i).X
                oTempPosition.Y = aStrapPos(i).Y

                If nHoopOpening <> SPSSymbolConstants.ALIGNMENT_CENTER Then
                    If nHoopOpening = SPSSymbolConstants.ALIGNMENT_LEFT And i = 1 And nVerticalStrapCount > 1 Then
                        oTempPosition.Z = dSideOpeningStrapZPosition
                        oProjectionVector.Z = dSideOpeningStrapLength
                    ElseIf nHoopOpening = SPSSymbolConstants.ALIGNMENT_RIGHT And i = nVerticalStrapCount And nVerticalStrapCount > 1 Then
                        oTempPosition.Z = dSideOpeningStrapZPosition
                        oProjectionVector.Z = dSideOpeningStrapLength
                    Else
                        oTempPosition.Z = dTempPosZ
                        oProjectionVector.Z = -((nNumberOfHoops - 1) * dHoopPitch) - (dHoopPlateWidth)
                    End If
                End If

                If nVerticalStrapCount = 1 Then
                    oUpperVector.X = 1 * Math.Cos(dHoopAngle / 2)
                    oUpperVector.Y = 1 * Math.Sin(dHoopAngle / 2)
                Else
                    oUpperVector.X = 1 * Math.Cos((dHoopRotAngle * (i - 1)))
                    oUpperVector.Y = 1 * Math.Sin((dHoopRotAngle * (i - 1)))
                End If
                oProjection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oUpperVector, dVerticalStrapWidth, dVerticalStrapThickness)
                m_oSimplePhysicalAspect.Outputs.Add("VerticalStap" & i, oProjection3d)
            Next i

            'reset for flare straps
            oProjectionVector.Z = -((nNumberOfHoops - 1) * dHoopPitch) - (dHoopPlateWidth)

            'Now place Vertical Straps upto the Flare Hoop                  
            oFlareComplex.GetCurves(oCurveColl)
            Dim oFlareArc As Arc3d = DirectCast(oCurveColl(0), Arc3d)
            Dim dFlareAngle As Double = oFlareArc.SweepAngle
            Dim dFlareRadius As Double = oFlareArc.Radius

            If nVerticalStrapCount > 1 Then
                dFlareRotAngle = dFlareAngle / (nVerticalStrapCount - 1)
            End If

            oTempPosition.Z = dTempPosZ + oProjectionVector.Z 'this point on last hoop

            For i = 1 To nVerticalStrapCount
                If nVerticalStrapCount = 1 Then
                    oUpperVector.X = 1 * Math.Cos((dHoopAngle / 2))
                    oUpperVector.Y = 1 * Math.Sin((dHoopAngle / 2))
                    aFlarePos(i).X = oCenterPosition.X
                    aFlarePos(i).Y = oCenterPosition.Y + dFlareRadius
                Else
                    oUpperVector.X = 1 * Math.Cos((dHoopRotAngle * (i - 1)))
                    oUpperVector.Y = 1 * Math.Sin((dHoopRotAngle * (i - 1)))

                    oTmpVector.X = Math.Cos((dFlareRotAngle * (i - 1))) * (dFlareRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                    oTmpVector.Y = Math.Sin((dFlareRotAngle * (i - 1))) * (dFlareRadius - dVerticalStrapThickness - dVerticalStrapThickness / 2)
                    aFlarePos(i).X = oTmpVector.X + oCenterPosition.X

                    If CBool(CInt(i = 1) Or nVerticalStrapCount) Then
                        aFlarePos(i).Y = oTmpVector.Y + oCenterPosition.Y + (dVerticalStrapThickness + dVerticalStrapThickness / 2)
                    Else
                        aFlarePos(i).Y = oTmpVector.Y + oCenterPosition.Y
                    End If
                End If
                oUpperVector.Z = 1 * Math.Tan(Math.PI / 2 - dAngle)

                oTempPosition.X = aStrapPos(i).X
                oTempPosition.Y = aStrapPos(i).Y

                aFlarePos(i).Z = oTempPosition.Z - dHoopPitch

                oTempVector1 = oProjectionVector.Cross(oUpperVector)

                oProjectionVector.Set(aFlarePos(i).X - aStrapPos(i).X, aFlarePos(i).Y - aStrapPos(i).Y, aFlarePos(i).Z - oTempPosition.Z)

                oTempProjectionVector.Set(oProjectionVector.X, oProjectionVector.Y, oProjectionVector.Z)

                oTempVector = oTempVector1.Cross(oTempProjectionVector)

                oProjection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oTempVector, dVerticalStrapWidth, dVerticalStrapThickness)
                m_oSimplePhysicalAspect.Outputs.Add("VerticalStap" & (nVerticalStrapCount + i), oProjection3d)
            Next i
        End Sub

#End Region

#Region "Operational Envelope Creation"

        Private Sub PlaceOperationalEnvelope1(ByVal bWithSafetyHoop As Boolean, _
                                              ByVal dWidth As Double, ByVal dLength As Double, _
                                              ByVal dPosition As Double, ByVal dSideFrameThickness As Double, _
                                              ByVal dEnvelopeHeight As Double)

            Dim dTempHeight As Double, dTotalWidth As Double, dTotalHeight As Double
            Dim oProjectionVector As New Vector(0.0, 0.0, 0.0), oUpperVector As New Vector(0.0, 1.0, 0.0)
            Dim oTempPosition As New Position()
            Dim dFlareClearance As Double = m_dFlareClearance.Value
            Dim dFlareRadius As Double = m_dFlareRadius.Value
            Dim dHoopPlateThickness As Double = m_dHoopPlateThickness.Value
            Dim dSupportLegThickness As Double = m_dSupportLegThickness.Value
            Dim dHoopPlateWidth As Double = m_dHoopPlateWidth.Value
            Dim dBottomHoopLevel As Double = m_dBottomHoopLevel.Value
            Dim dHoopPitch As Double = m_dHoopPitch.Value
            Dim dWallOffset As Double = m_dWallOffset.Value

            If bWithSafetyHoop Then
                dTempHeight = dBottomHoopLevel + dHoopPitch + (dHoopPlateWidth * 2)
                oTempPosition.Set(-dPosition, ((dFlareClearance + dHoopPlateThickness) / 2) + (dWallOffset / 2), -(dLength))
                oProjectionVector.Z = dTempHeight
                dTotalWidth = dFlareRadius * 2
                dTotalHeight = dFlareClearance + (dHoopPlateThickness / 2) + dWallOffset
            Else
                oTempPosition.Set(-dPosition, (0.762 / 2) + (dWallOffset / 2), -(dLength))
                oProjectionVector.Z = dLength + dEnvelopeHeight
                dTotalWidth = dWidth + dSideFrameThickness * 2 + dSupportLegThickness * 2
                dTotalHeight = 0.762 + dWallOffset
            End If
            Dim oProjection3d As Projection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oUpperVector, dTotalWidth, dTotalHeight)
            m_oOperationAspect.Outputs.Add("OperationalEnvelope1", oProjection3d)

        End Sub

        Private Sub PlaceOperationalEnvelope2(ByVal dLength As Double, _
                                              ByVal dPosition As Double, ByVal dEnvelopeHeight As Double)

            Dim oProjectionVector As New Vector(0.0, 0.0, 0.0), oUpperVector As New Vector(0.0, 1.0, 0.0)
            Dim dFlareClearance As Double = m_dFlareClearance.Value
            Dim dHoopRadius As Double = m_dHoopRadius.Value
            Dim dHoopClearance As Double = m_dHoopClearance.Value
            Dim dBottomHoopLevel As Double = m_dBottomHoopLevel.Value
            Dim dHoopPlateThickness As Double = m_dHoopPlateThickness.Value
            Dim dHoopPitch As Double = m_dHoopPitch.Value
            Dim dHoopPlateWidth As Double = m_dHoopPlateWidth.Value
            Dim dWallOffset As Double = m_dWallOffset.Value

            Dim dTempHeight As Double = dBottomHoopLevel + dHoopPitch + (dHoopPlateWidth * 2)
            Dim oTempPosition As New Position(-dPosition, ((dHoopClearance + dHoopPlateThickness) / 2) + (dWallOffset / 2), -dLength + dTempHeight)
            oProjectionVector.Z = dLength + dEnvelopeHeight - dTempHeight
            Dim dTotalWidth As Double = dHoopRadius * 2
            Dim dTotalHeight As Double = dHoopClearance + (dHoopPlateThickness / 2) + dWallOffset
            Dim oProjection3d As Projection3d = CreateRectangularProjection(oTempPosition, oProjectionVector, oUpperVector, dTotalWidth, dTotalHeight)
            m_oOperationAspect.Outputs.Add("OperationalEnvelope2", oProjection3d)

        End Sub

#End Region

#Region "Different Type of Projection Creation"

        Private Function CreateRectangularProjection(ByVal oCenterPos As Position, _
                                                     ByVal oProjectionVector As Vector, ByVal oUpperVector As Vector, _
                                                     ByVal dRectWidth As Double, ByVal dRectHeight As Double) As Projection3d

            Dim oTempVect As New Vector
            Dim oTempProjVect As New Vector
            Dim oProjection3D As Projection3d
            Dim oMatrix As Matrix4X4
            Dim aMatrix(15) As Double

            oProjection3D = CreateUnitRectangularProjection(dRectWidth, dRectHeight, oProjectionVector.Length)

            oTempProjVect.Set(oProjectionVector.X, oProjectionVector.Y, oProjectionVector.Z)
            oTempProjVect.Length = 1.0

            oUpperVector.Length = 1.0

            oTempVect = oTempProjVect.Cross(oUpperVector)
            oTempVect.Length = 1.0

            aMatrix(0) = oTempVect.X
            aMatrix(1) = oTempVect.Y
            aMatrix(2) = oTempVect.Z

            aMatrix(4) = oUpperVector.X
            aMatrix(5) = oUpperVector.Y
            aMatrix(6) = oUpperVector.Z

            aMatrix(8) = oTempProjVect.X
            aMatrix(9) = oTempProjVect.Y
            aMatrix(10) = oTempProjVect.Z

            aMatrix(12) = oCenterPos.X
            aMatrix(13) = oCenterPos.Y
            aMatrix(14) = oCenterPos.Z
            aMatrix(15) = 1.0

            oMatrix = New Matrix4X4(aMatrix, True)
            oProjection3D.Transform(oMatrix)

            CreateRectangularProjection = oProjection3D

        End Function

        Private Function CreateCircularProjection(ByVal oCenterPos As Position, _
                                                  ByVal oProjectionVector As Vector, ByVal oUpperVector As Vector, _
                                                  ByVal dDiameter As Double) As Projection3d

            Dim oTempVect As Vector
            Dim oTempProjVect As Vector
            Dim oProjection3D As Projection3d
            Dim aMatrix(15) As Double
            Dim oMatrix As Matrix4X4

            oProjection3D = CreateUnitCircularProjection(dDiameter, oProjectionVector.Length)

            oTempProjVect = New Vector(oProjectionVector.X, oProjectionVector.Y, oProjectionVector.Z)
            oTempProjVect.Length = 1.0

            oUpperVector.Length = 1.0
            oTempVect = oTempProjVect.Cross(oUpperVector)
            oTempVect.Length = 1.0

            aMatrix(0) = oTempVect.X
            aMatrix(1) = oTempVect.Y
            aMatrix(2) = oTempVect.Z

            aMatrix(4) = oUpperVector.X
            aMatrix(5) = oUpperVector.Y
            aMatrix(6) = oUpperVector.Z

            aMatrix(8) = oTempProjVect.X
            aMatrix(9) = oTempProjVect.Y
            aMatrix(10) = oTempProjVect.Z

            aMatrix(12) = oCenterPos.X
            aMatrix(13) = oCenterPos.Y
            aMatrix(14) = oCenterPos.Z
            aMatrix(15) = 1.0

            oMatrix = New Matrix4X4(aMatrix, True)
            oProjection3D.Transform(oMatrix)

            CreateCircularProjection = oProjection3D

        End Function

        Private Function CreateProjectionOfCurve(ByVal oCenterPos As Position, _
                                                 ByVal oProjectionVector As Vector, ByVal oUpperVector As Vector, _
                                                 ByVal dHoopBendRadius As Double, ByVal aShDim() As Double, _
                                                 ByVal dDimX As Double, ByVal dDimY As Double, _
                                                 ByVal dHoopPlateThickness As Double, ByVal sHoop As String, _
                                                 ByVal nHoopNumber As Integer, ByRef oHoopComplex As ComplexString3d, _
                                                 ByRef oFlareComplex As ComplexString3d, Optional ByVal nHoopOpening As Integer = SPSSymbolConstants.ALIGNMENT_CENTER, _
                                                 Optional ByVal dLadderWidth As Double = 0.0) As Projection3d

            Dim oTempVect As Vector
            Dim oTempProjVect As New Vector()
            Dim oProjection3D As Projection3d
            Dim oMatrix As Matrix4X4
            Dim aMatrix(15) As Double
            Dim oComplex As ComplexString3d = Nothing

            oProjection3D = CreateUnitProjection(dHoopBendRadius, aShDim, dDimX, dDimY, _
                        dHoopPlateThickness, oProjectionVector.Length, sHoop, nHoopNumber, oComplex, _
                        oHoopComplex, oFlareComplex, nHoopOpening, dLadderWidth)

            oTempProjVect.Set(oProjectionVector.X, oProjectionVector.Y, oProjectionVector.Z)
            oTempProjVect.Length = 1.0

            oUpperVector.Length = 1.0
            oTempVect = oTempProjVect.Cross(oUpperVector)
            oTempVect.Length = 1.0

            aMatrix(0) = oTempVect.X
            aMatrix(1) = oTempVect.Y
            aMatrix(2) = oTempVect.Z

            aMatrix(4) = -oUpperVector.X
            aMatrix(5) = -oUpperVector.Y
            aMatrix(6) = -oUpperVector.Z

            aMatrix(8) = oTempProjVect.X
            aMatrix(9) = oTempProjVect.Y
            aMatrix(10) = oTempProjVect.Z

            aMatrix(12) = oCenterPos.X
            aMatrix(13) = oCenterPos.Y
            aMatrix(14) = oCenterPos.Z
            aMatrix(15) = 1.0

            oMatrix = New Matrix4X4(aMatrix, True)
            oComplex.Transform(oMatrix)

            If nHoopNumber = 1 Then
                oHoopComplex.Transform(oMatrix)
            ElseIf sHoop = "FlareSafetyHoop" Then
                oFlareComplex.Transform(oMatrix)
            End If

            oProjection3D.Transform(oMatrix)

            CreateProjectionOfCurve = oProjection3D

        End Function

        Private Function CreateUnitRectangularProjection(ByVal dRectWidth As Double, _
                                                  ByVal dRectHeight As Double, ByVal dProjLength As Double) As Projection3d

            Dim oLineString As ILineString
            Dim oPosColl As Collection(Of Position)
            Dim oVectorProjDir As New Vector(0.0, 0.0, 1.0)
            Dim oProjection3d As Projection3d

            oPosColl = New Collection(Of Position)
            CreateRectangularCurvePoints(dRectWidth, dRectHeight, oPosColl)
            oLineString = New LineString3d(oPosColl)

            oProjection3d = New Projection3d(DirectCast(oLineString, ICurve), oVectorProjDir, dProjLength, True)

            CreateUnitRectangularProjection = oProjection3d

        End Function

        Private Function CreateUnitCircularProjection(ByVal dDiameter As Double, _
                                                  ByVal dProjLength As Double) As Projection3d

            Dim oCircle As Circle3d
            Dim oCenterPos As Position
            Dim oVectorNormal As Vector
            Dim oProjection3d As Projection3d

            oCenterPos = New Position(0, 0, 0)
            oVectorNormal = New Vector(0.0, 0.0, 1.0)
            oCircle = New Circle3d(oCenterPos, oVectorNormal, dDiameter / 2)

            oProjection3d = New Projection3d(DirectCast(oCircle, ICurve), oVectorNormal, dProjLength, True)

            CreateUnitCircularProjection = oProjection3d

        End Function

        Private Function CreateUnitProjection(ByVal dHoopBendRadius As Double, _
                                              ByVal aShDim() As Double, ByVal dDimX As Double, ByVal dDimY As Double, _
                                              ByVal dHoopPlateThickness As Double, ByVal dProjLength As Double, _
                                              ByVal sHoop As String, ByVal nHoopNumber As Integer, _
                                              ByRef oComplex As ComplexString3d, ByRef oHoopComplex As ComplexString3d, _
                                              ByRef oFlareComplex As ComplexString3d, Optional ByVal nHoopOpening As Integer = SPSSymbolConstants.ALIGNMENT_CENTER, _
                                              Optional ByVal dLadderWidth As Double = 0.0) As Projection3d

            Dim oProjection3d As Projection3d
            Dim oVectorProjDir As New Vector(0.0, 0.0, 1.0)

            If "FlareSafetyHoop" = sHoop Or nHoopOpening = SPSSymbolConstants.ALIGNMENT_CENTER Then
                CreateHoop(aShDim(0), aShDim(1), aShDim(2), dHoopBendRadius, dDimX, dDimY, dHoopPlateThickness, sHoop, nHoopNumber, oComplex, oHoopComplex, oFlareComplex)
            ElseIf nHoopOpening = SPSSymbolConstants.ALIGNMENT_LEFT Then
                'left opening
                CreateLeftOpenHoop(aShDim(0), aShDim(1), aShDim(2), dHoopBendRadius, dDimX, dDimY, dHoopPlateThickness, sHoop, nHoopNumber, dLadderWidth, oComplex, oHoopComplex, oFlareComplex)
            ElseIf nHoopOpening = SPSSymbolConstants.ALIGNMENT_RIGHT Then
                'right opening
                CreateRightOpenHoop(aShDim(0), aShDim(1), aShDim(2), dHoopBendRadius, dDimX, dDimY, dHoopPlateThickness, sHoop, nHoopNumber, dLadderWidth, oComplex, oHoopComplex, oFlareComplex)
            End If

            oProjection3d = New Projection3d(DirectCast(oComplex, ICurve), oVectorProjDir, dProjLength, True)

            CreateUnitProjection = oProjection3d

        End Function

        Private Sub CreateRectangularCurvePoints(ByVal dWidth As Double, ByVal dHeight As Double, _
                                        ByRef oPosColl As Collection(Of Position))

            Dim oTempPos As New Position

            'Build points in local XY plane at the centroid of the rectangle
            oTempPos.X = -(dWidth / 2.0)
            oTempPos.Y = -(dHeight / 2.0)
            oTempPos.Z = 0
            oPosColl.Add(oTempPos)

            oTempPos = New Position
            oTempPos.X = dWidth / 2.0
            oTempPos.Y = -(dHeight / 2.0)
            oTempPos.Z = 0
            oPosColl.Add(oTempPos)

            oTempPos = New Position
            oTempPos.X = dWidth / 2.0
            oTempPos.Y = dHeight / 2.0
            oTempPos.Z = 0
            oPosColl.Add(oTempPos)

            oTempPos = New Position
            oTempPos.X = -(dWidth / 2.0)
            oTempPos.Y = dHeight / 2.0
            oTempPos.Z = 0
            oPosColl.Add(oTempPos)

            oTempPos = New Position
            oTempPos.X = -(dWidth / 2.0) 'Same as first point for closed shape
            oTempPos.Y = -(dHeight / 2.0)
            oTempPos.Z = 0
            oPosColl.Add(oTempPos)
        End Sub

#End Region

#Region "Different Type of Hoop Creation"

        Private Sub CreateHoop(ByVal dShDim1 As Double, ByVal dShDim2 As Double, _
                               ByVal dShDim3 As Double, ByVal dHoopBendRad As Double, _
                               ByVal dDimX As Double, ByVal dDimY As Double, _
                               ByVal dHoopPlateThickness As Double, ByVal sHoop As String, _
                               ByVal nHoopNumber As Integer, ByRef oComplex As ComplexString3d, _
                               ByRef oHoopComplex As ComplexString3d, ByRef oFlareComplex As ComplexString3d)

            Dim oStPos As New Position(), oEndPos As New Position(), oMidPos As New Position, oCenterPos As New Position()
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, z As Double
            Dim dHoopBendInnerRad, dDimX_Inner, dDimY_Inner As Double
            Dim oCurveColl As New Collection(Of ICurve)
            Dim oArc As Arc3d
            Dim oLine As Line3d

            x1 = 0.0        '                        /\+Y
            y1 = 0          '  2 |                   |
            x2 = 0          '    |                   |
            y2 = dShDim1    '  1 |.(0,0)             .------> +X

            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oCurveColl.Add(New Line3d(oStPos, oEndPos))
            oComplex = New ComplexString3d(oCurveColl)
            oCurveColl = Nothing

            x1 = x2                     'Start Point           3__
            y1 = y2                     'Start Point             . \
            x2 = x1 - dHoopBendRad      'Center Point            2 | 1
            y2 = y2                     'Center Point
            x3 = x2                     'End Point
            y3 = y2 + dHoopBendRad      'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc = New Arc3d(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            x1 = x3                     '               2______1
            y1 = y3
            x2 = x3 - dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine = New Line3d(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                     'Start Point        1________
            y1 = y2                     'Start Point        / .2    . \             ^ +Y
            x2 = x1                     'Center Point     3|           |            |
            y2 = y2 - dHoopBendRad      'Center Point      |           |            |
            x3 = x1 - dHoopBendRad      'End Point         |           |            .--------> +X
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            x1 = x3                     '                  1|
            y1 = y3                     '                   |
            x2 = x1                     '                   |
            y2 = y1 - dShDim3           '                  2|
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                   'Start Point                  |--dDimX--|
            y1 = y2                   'Start Point     1 |                    | 3        -
            x2 = x1 + dDimX           'Center Point       -                  -           | dDimY
            y2 = y1 - dDimX           'Center Point         -              -             |
            x3 = x1 + 2 * dDimX       'End Point                - __2__ -                -
            y3 = y1                   'End Point
            oStPos.Set(x1, y1, z)
            oMidPos.Set(x2, y2, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            If sHoop = "SafetyHoop" And nHoopNumber = 1 Then
                oCurveColl = New Collection(Of ICurve)
                oCurveColl.Add(oArc)
                oHoopComplex = New ComplexString3d(oCurveColl)
            End If

            x1 = x3                                              '                  2|
            y1 = y3                                              '                   |
            x2 = x1                                              '                   |
            y2 = y1 + dShDim3                                    '                  1|
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                     'Start Point                         ______3__
            y1 = y2                     'Start Point                        /      2. \ 1
            x2 = x1 - dHoopBendRad      'Center Point                      |           |
            y2 = y1                     'Center Point                                  |
            x3 = x1 - dHoopBendRad      'End Point                                     |
            y3 = y1 + dHoopBendRad      'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            x1 = x3                     '               2______1
            y1 = y3
            x2 = x1 - dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                     'Start Point        __1
            y1 = y2                     'Start Point       / .2
            x2 = x1                     'Center Point     3|
            y2 = y1 - dHoopBendRad      'Center Point
            x3 = x1 - dHoopBendRad      'End Point
            y3 = y1 - dHoopBendRad      'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            x1 = x3                               '  1 |
            y1 = y3                               '    |
            x2 = x1                               '  2 |
            y2 = y3 - dShDim1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'Now return along the same lines with an offset of dHoopPlateThickness
            dHoopBendInnerRad = dHoopBendRad - dHoopPlateThickness
            dDimX_Inner = dDimX - dHoopPlateThickness
            dDimY_Inner = dDimY - dHoopPlateThickness

            x1 = x2                              '  1--2
            y1 = y2
            x2 = x1 + dHoopPlateThickness
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                               '  2 |
            y1 = y2                               '    |
            x2 = x1                               '  1 |
            y2 = y1 + dShDim1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                                 'Start Point        __3
            y1 = y2                                 'Start Point       / .2
            x2 = x1 + dHoopBendInnerRad             'Center Point     1|
            y2 = y1                                 'Center Point
            x3 = x2                                 'End Point
            y3 = y2 + dHoopBendInnerRad             'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            x1 = x3                     '               1______2
            y1 = y3
            x2 = x1 + dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                     'Start Point         ______1__
            y1 = y2                     'Start Point        /      2. \ 3
            x2 = x1                     'Center Point      |           |
            y2 = y1 - dHoopBendInnerRad 'Center Point      |
            x3 = x2 + dHoopBendInnerRad 'End Point         |
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            x1 = x3                '                  1|
            y1 = y3                '                   |
            x2 = x1                '                   |
            y2 = y1 - dShDim3      '                  2|
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)
            '''''''''''''''''''''''''''                            |-DimX_inner-|
            x1 = x2                         'Start Point
            y1 = y2                   'Start Point      3 |                     | 1      -
            x2 = x1 - dDimX_Inner            'Center       -                  -           |dDimY_Inner
            y2 = y1 - dDimX_Inner            'Center         -              -             |
            x3 = x1 - 2 * dDimX_Inner        'End Point         - __2__ -                -
            y3 = y1                                 'End Point
            oStPos.Set(x1, y1, z)
            oMidPos.Set(x2, y2, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            oComplex.AddCurve(oArc, True)
            If sHoop = "FlareSafetyHoop" Then
                oCurveColl = New Collection(Of ICurve)
                oCurveColl.Add(oArc)
                oFlareComplex = New ComplexString3d(oCurveColl)
            End If

            x1 = x3                     '                  2|
            y1 = y3                     '                   |
            x2 = x1                     '                   |
            y2 = y1 + dShDim3           '                  1|
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                     'Start Point        3________
            y1 = y2                     'Start Point        / .2    . \             ^ +Y
            x2 = x1 + dHoopBendInnerRad 'Center Point     1|           |            |
            '''''''''''''               'Center Point      |           |            |
            x3 = x2                     'End Point         |           |            .--------> +X
            y3 = y2 + dHoopBendInnerRad 'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            x1 = x3                     '               1______2
            y1 = y3
            x2 = x3 + dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                     'Start Point           1__
            y1 = y2                     'Start Point             . \
            ''''''''''                  'Center Point            2 | 3
            y2 = y2 - dHoopBendInnerRad 'Center Point
            x3 = x2 + dHoopBendInnerRad 'End Point
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)
            ''''''''''''''''''''''''''''''''''''''''                       /\+Y
            x1 = x3                                '  1 |                   |
            y1 = y3                                '    |                   |
            x2 = x1                                '  2 |                   .------> +X
            y2 = y1 - dShDim1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                                '   1-------2 (0,0)
            y1 = y2
            x2 = 0
            y2 = 0
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

        End Sub

        Private Sub CreateRightOpenHoop(ByVal dShDim1 As Double, ByVal dShDim2 As Double, _
                                       ByVal dShDim3 As Double, ByVal dHoopBendRad As Double, _
                                       ByVal dDimX As Double, ByVal dDimY As Double, _
                                       ByVal dHoopPlateThickness As Double, ByVal sHoop As String, _
                                       ByVal nHoopNumber As Integer, ByVal dLadderWidth As Double, _
                                       ByRef oComplex As ComplexString3d, ByRef oHoopComplex As ComplexString3d, _
                                       ByRef oFlareComplex As ComplexString3d)

            Dim oStPos As New Position(), oEndPos As New Position(), oMidPos As New Position, oCenterPos As New Position()
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, z As Double
            Dim dHoopBendInnerRad, dHoopBendOuterRad, dDimX_Inner, dDimY_Inner As Double
            Dim dHoopBotX As Double, dHoopBotY As Double
            Dim oCurveColl As New Collection(Of ICurve)
            Dim oArc As Arc3d
            Dim oLine As Line3d

            'TR214568
            'Right side opening. The order of the segments for the hoop is shown below.

            'Construct the segments for the hoop
            '
            '                       C|----|D
            '          ^y           |      |
            '          |          B|        |
            '          |           |         |E
            '          --->x       |A        |
            '                                |
            '   I--H                         |
            '      -                         |
            '       -|G                      |F
            '         |                     |
            '          |                  |
            '            |             |
            '               |      |
            '                 ____           
            '
            'Though hoop in the figure is in -ve y direction, to get
            'hoop in +ve y direction, the local geometry is flipped about x axis as shown code in the below. 
            'Finally the geometry is transformed using the occurrence matrix to flip in both X and Y directions
            'Then the final hoop looks like
            '
            '      D|----|C
            '      |      |           
            '    E|       B|           
            '     |        |           
            '     |        A     x<----
            '     |                   |
            '     |                   |           H|---|I
            '     |                 y v          |
            '    F|                             |G
            '       |                        |
            '          |                  |
            '            |             |
            '               |      |
            '                 ____
            '   ^GY
            '   |
            '   |
            '   - - ->GX



            'segment AB
            x1 = dLadderWidth
            y1 = 0
            x2 = x1
            y2 = dShDim1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oCurveColl.Add(New Line3d(oStPos, oEndPos))
            oComplex = New ComplexString3d(oCurveColl)
            oCurveColl = Nothing

            'segment BC
            x1 = x2                     'Start Point
            y1 = y2                     'Start Point
            x2 = x1 + dHoopBendRad      'Center Point
            y2 = y2                     'Center Point
            x3 = x2                     'End Point
            y3 = y2 + dHoopBendRad      'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc = New Arc3d(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment CD
            x1 = x3
            y1 = y3
            x2 = x3 + dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine = New Line3d(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment DE
            x1 = x2                     'Start Point
            y1 = y2                     'Start Point
            x2 = x1                     'Center Point
            y2 = y2 - dHoopBendRad      'Center Point
            x3 = x1 + dHoopBendRad      'End Point
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment EF
            x1 = x3                     '
            y1 = y3                     '
            x2 = x1                     '
            y2 = y1 - dShDim3           '
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment FG
            ' To leave enough space for user to slide on the side, we need the cage to stop short of full semi-circle
            ' we will leave short at 45deg angle
            x1 = x2                   'Start Point
            y1 = y2                   'Start Point
            x2 = x1 - dDimX                 'Center
            y2 = y1 - dDimX                 'Center
            dHoopBotX = x2
            dHoopBotY = y2
            '''''''''''''''''''                   'End Point
            ''''''''                               'End Point
            x3 = x2 - dDimX * Sin(PI / 4.0)          'End Point
            y3 = y2 + dDimX * (1.0 - Cos(PI / 4.0))    'End Point
            oStPos.Set(x1, y1, z)
            oMidPos.Set(x2, y2, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            If sHoop = "SafetyHoop" And nHoopNumber = 1 Then
                oCurveColl = New Collection(Of ICurve)
                'redefine x3, y3 for the flare hoop
                oStPos.Set(x1, y1, z)
                oMidPos.Set(x2, y2, z)
                oEndPos.Set(x2 - dDimX, y1, z)
                oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
                oCurveColl.Add(oArc)
                oHoopComplex = New ComplexString3d(oCurveColl)
            End If

            'segment GH
            x1 = x3                     'Start Point
            y1 = y3                     'Start Point
            x2 = x1 - dHoopBendRad      'Center Point
            y2 = y1                     'Center Point
            x3 = x2                     'End Point
            y3 = y1 + dHoopBendRad      'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment HI
            x1 = x3
            y1 = y3
            x2 = x1 - dShDim2 / 2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'Now return along the same lines with an offset of dHoopPlateThickness
            dHoopBendInnerRad = dHoopBendRad - dHoopPlateThickness
            dHoopBendOuterRad = dHoopBendRad + dHoopPlateThickness
            dDimX_Inner = dDimX - dHoopPlateThickness
            dDimY_Inner = dDimY - dHoopPlateThickness

            'Adjust coordinates for plate thickness
            x1 = x2                              '  1--2
            y1 = y2
            x2 = x1
            y2 = y1 + dHoopPlateThickness
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment IH
            x1 = x2
            y1 = y2
            x2 = x1 + dShDim2 / 2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment HG
            x1 = x2                         'Start Point
            y1 = y2                         'Start Point
            x2 = x1                         'Center Point
            y2 = y1 - dHoopBendOuterRad     'Center Point
            x3 = x2 + dHoopBendOuterRad     'End Point
            y3 = y2                         'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment GF
            x1 = x3
            y1 = y3
            x2 = dHoopBotX
            y2 = dHoopBotY + dHoopPlateThickness
            x3 = x2 + dDimX_Inner
            y3 = y2 + dDimX_Inner                   'End Point
            oStPos.Set(x1, y1, z)
            oMidPos.Set(x2, y2, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            oComplex.AddCurve(oArc, True)
            'If sHoop = "FlareSafetyHoop" Then
            '    oCurveColl = New Collection(Of ICurve)
            '    'redefine x1, y1 for the flare hoop
            '    x1 = x2 + dDimX_Inner
            '    y1 = y2 + dDimX_Inner
            '    oStPos.Set(x1, y1, z)
            '    oMidPos.Set(x2, y2, z)
            '    oEndPos.Set(x3, y3, z)
            '    oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            '    oCurveColl.Add(oArc)
            '    oFlareComplex = New ComplexString3d(oCurveColl)
            'End If

            'segment FE
            x1 = x3                     '                  2|
            y1 = y3                     '                   |
            x2 = x1                     '                   |
            y2 = y1 + dShDim3           '                  1|
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment ED
            x1 = x2                     'Start Point        3________
            y1 = y2                     'Start Point        / .2    . \             ^ +Y
            x2 = x1 - dHoopBendInnerRad 'Center Point     1|           |            |
            '''''''                     'Center Point      |           |            |
            x3 = x2                     'End Point         |           |            .--------> +X
            y3 = y2 + dHoopBendInnerRad 'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment DC
            x1 = x3                     '               1______2
            y1 = y3
            x2 = x3 - dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment CB
            x1 = x2                     'Start Point           1__
            y1 = y2                     'Start Point             . \
            ''''''''''''''''            'Center Point            2 | 3
            y2 = y2 - dHoopBendInnerRad 'Center Point
            x3 = x2 - dHoopBendInnerRad 'End Point
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            '''''''''''''''''''''''''''''''''''''''
            'segment BA                                                    /\+Y
            x1 = x3                                '  1 |                   |
            y1 = y3                                '    |                   |
            x2 = x1                                '  2 |                   .------> +X
            y2 = y1 - dShDim1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2                                '   1-------2 (0,0)
            y1 = y2
            x2 = dLadderWidth
            y2 = 0
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)
        End Sub

        Private Sub CreateLeftOpenHoop(ByVal dShDim1 As Double, ByVal dShDim2 As Double, _
                                       ByVal dShDim3 As Double, ByVal dHoopBendRad As Double, _
                                       ByVal dDimX As Double, ByVal dDimY As Double, _
                                       ByVal dHoopPlateThickness As Double, ByVal sHoop As String, _
                                       ByVal nHoopNumber As Integer, ByVal dLadderWidth As Double, _
                                       ByRef oComplex As ComplexString3d, ByRef oHoopComplex As ComplexString3d, _
                                       ByRef oFlareComplex As ComplexString3d)

            Dim oStPos As New Position(), oEndPos As New Position(), oMidPos As New Position, oCenterPos As New Position()
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, z As Double
            Dim dHoopBendInnerRad, dHoopBendOuterRad, dDimX_Inner, dDimY_Inner As Double
            Dim dHoopBotX As Double, dHoopBotY As Double
            Dim oCurveColl As New Collection(Of ICurve)
            Dim oLine As Line3d
            Dim oArc As Arc3d

            'TR214568
            'Left side opening. The order of the segments for the hoop is shown below.

            'Construct the segments for the hoop
            '      D|----|C
            '      |      |^y
            '    E|       B| 
            '     |        | 
            '     |        A----->x
            '     |
            '     |                               H|---|I
            '     |                              |
            '    F|                             |G
            '       |                        |
            '          |                  |
            '            |             |
            '               |      |
            '                 ____

            'Though hoop in the figure is in -ve y direction, to get
            'hoop in +ve y direction, the local geometry is flipped about x axis as shown code in the below. 
            'Finally the geometry is transformed using the occurrence matrix to flip in both X and Y directions
            'Then the final hoop looks like
            '
            '                     
            '                       C|----|D
            '                      |      |
            '                     B|        |
            '                      |         |E
            '                      |A        |
            '                 x<----         |
            '   I--H               |         |
            '      -               |         |
            '       -|G          y v        |F
            '         |                     |
            '          |                  |
            '            |             |
            '               |      |
            '                 ____           
            '   ^GY
            '   |
            '   |
            '   - - ->GX

            'segment AB
            x1 = 0.0
            y1 = 0
            x2 = x1
            y2 = dShDim1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oCurveColl.Add(New Line3d(oStPos, oEndPos))
            oComplex = New ComplexString3d(oCurveColl)
            oCurveColl = Nothing

            'segment BC
            x1 = x2                     'Start Point
            y1 = y2                     'Start Point
            x2 = x1 - dHoopBendRad      'Center Point
            y2 = y2                     'Center Point
            x3 = x2                     'End Point
            y3 = y2 + dHoopBendRad      'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc = New Arc3d(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment CD
            x1 = x3                     '
            y1 = y3
            x2 = x3 - dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine = New Line3d(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment DE
            x1 = x2                     'Start Point
            y1 = y2                     'Start Point
            x2 = x1                     'Center Point
            y2 = y2 - dHoopBendRad      'Center Point
            x3 = x1 - dHoopBendRad      'End Point
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment EF
            x1 = x3
            y1 = y3
            x2 = x1
            y2 = y1 - dShDim3
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment FG - the actual hoop arc
            x1 = x2                          'Start Point
            y1 = y2
            x2 = x1 + dDimX                  'Second Point
            y2 = y1 - dDimX
            x3 = x2 + dDimX * Sin(PI / 4.0)  'End Point
            y3 = y2 + dDimX * (1 - Cos(PI / 4.0))
            oStPos.Set(x1, y1, z)
            oMidPos.Set(x2, y2, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'save x2,y2 for later use
            dHoopBotX = x2
            dHoopBotY = y2

            If sHoop = "SafetyHoop" And nHoopNumber = 1 Then
                oCurveColl = New Collection(Of ICurve)
                'redefine arc for the flare hoop
                oStPos.Set(x1, y1, z)
                oMidPos.Set(x2, y2, z)
                oEndPos.Set(x1 + 2 * dDimX, y1, z)
                oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
                oCurveColl.Add(oArc)
                oHoopComplex = New ComplexString3d(oCurveColl)
            End If

            'segment GH
            x1 = x3                     'Start Point
            y1 = y3                     'Start Point
            x2 = x1 + dHoopBendRad      'Center Point
            y2 = y1                     'Center Point
            x3 = x2                     'End Point
            y3 = y1 + dHoopBendRad      'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment HI
            x1 = x3
            y1 = y3
            x2 = x1 + dShDim2 / 2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            ' Now return along the same lines with an offset of dHoopPlateThickness
            dHoopBendInnerRad = dHoopBendRad - dHoopPlateThickness
            dHoopBendOuterRad = dHoopBendRad + dHoopPlateThickness
            dDimX_Inner = dDimX - dHoopPlateThickness
            dDimY_Inner = dDimY - dHoopPlateThickness

            x1 = x2
            y1 = y2
            x2 = x1
            y2 = y1 + dHoopPlateThickness
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment IH
            x1 = x2
            y1 = y2
            x2 = x1 - dShDim2 / 2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment HG
            x1 = x2                     'Start Point
            y1 = y2                     'Start Point
            x2 = x1                     'Center Point
            y2 = y1 - dHoopBendOuterRad 'Center Point
            x3 = x2 - dHoopBendOuterRad 'End Point
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment GF
            x1 = x3                              'Start Point
            y1 = y3                              'Start Point
            x2 = dHoopBotX                       'Center
            y2 = dHoopBotY + dHoopPlateThickness 'Center
            x3 = x2 - dDimX_Inner                'End Point
            y3 = y2 + dDimX_Inner                'End Point
            oStPos.Set(x1, y1, z)
            oMidPos.Set(x2, y2, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'If sHoop = "FlareSafetyHoop" Then
            '    oCurveColl = New Collection(Of ICurve)
            '    'redefine arc for flare hoop
            '    x1 = x2 + dDimX_Inner             'Center
            '    y1 = y2 + dDimX_Inner             'Center
            '    oStPos.Set(x1, y1, z)
            '    oMidPos.Set(x2, y2, z)
            '    oEndPos.Set(x3, y3, z)
            '    oArc.DefineBy3Points(oStPos, oMidPos, oEndPos)
            '    oCurveColl.Add(oArc)
            '    oFlareComplex = New ComplexString3d(oCurveColl)
            'End If

            'segment FE
            x1 = x3
            y1 = y3
            x2 = x1
            y2 = y1 + dShDim3
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment ED
            x1 = x2                     'Start Point
            y1 = y2                     'Start Point
            x2 = x1 + dHoopBendInnerRad 'Center Point
            y2 = y1                     'Center Point
            x3 = x2                     'End Point
            y3 = y2 + dHoopBendInnerRad 'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment DC
            x1 = x3
            y1 = y3
            x2 = x3 + dShDim2
            y2 = y1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            'segment CB
            x1 = x2                     'Start Point
            y1 = y2                     'Start Point

            y2 = y2 - dHoopBendInnerRad 'Center Point
            x3 = x2 + dHoopBendInnerRad 'End Point
            y3 = y2                     'End Point
            oCenterPos.Set(x2, y2, z)
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x3, y3, z)
            oArc.DefineByCtrNormStartEnd(oCenterPos, Nothing, oStPos, oEndPos)
            oComplex.AddCurve(oArc, True)

            'segment BA
            x1 = x3
            y1 = y3
            x2 = x1
            y2 = y1 - dShDim1
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)

            x1 = x2
            y1 = y2
            x2 = 0.0
            y2 = 0
            oStPos.Set(x1, y1, z)
            oEndPos.Set(x2, y2, z)
            oLine.DefineBy2Points(oStPos, oEndPos)
            oComplex.AddCurve(oLine, True)
        End Sub

#End Region

#End Region

#Region "ICustomWeightCG Methods"

        ''' <summary>
        ''' Evaluates the weight and center of gravity of the ladder part and sets it on the ladder business object.
        ''' </summary>
        ''' <param name="businessObject">Ladder business object which aggregates symbol.</param>
        Public Sub EvaluateWeightCG(ByVal businessObject As BusinessObject) Implements ICustomWeightCG.EvaluateWeightCG
            'If position is not there, means symbol is not computed yet, better to skip weight and COG calculation now,
            'we will be called again after the symbol computed.
            Dim accumulatedCG As Position = MyBase.AccumulatedCG
            If Not accumulatedCG Is Nothing Then
                Dim cogX As Double = accumulatedCG.X
                Dim cogY As Double = accumulatedCG.Y
                Dim cogZ As Double = accumulatedCG.Z

                'Getting weight
                Dim weight As Double? = MyBase.AccumulatedWeight

                'Set the net weight and COG on the ladder business object using helper method provided in WeightCOGServices
                Dim weightCOGServices As New WeightCOGServices()
                Try
                    If Not weight Is Nothing Then
                        weightCOGServices.SetWeightAndCOG(businessObject, weight.Value, cogX, cogY, cogZ)
                    End If

                Catch ex As Exception
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, LadderSymbolsLocalizer.GetString(LadderSymbolsResourceIDs.ErrSetWeightAndCOG, _
                                            "Cannot set weight and center of gravity on the Ladder in LadderTypeA Symbol. Please check custom code or contact S3D."))
                End Try
            End If
        End Sub

#End Region

#Region "Overrides Functions And Methods"

        ''' <summary>
        ''' Computes the length of the ladder. Typical implementation of this method should include checking if any dimensional property of
        ''' ladder needs to be updated and calculating length.
        ''' </summary>
        ''' <param name="ladder">Ladder business object which aggregates symbol.</param>
        ''' <param name="ladderPart">Ladder part.</param>
        ''' <param name="height">Height of the ladder.</param>
        ''' <param name="ladderWidth">Width of the ladder.</param>
        ''' <param name="stepDiameter">The step diameter .</param>
        ''' <returns>The ladder length.</returns>
        Public Overrides Function ComputeLadderLength(ByVal ladder As Ladder, ByVal ladderPart As Part, ByVal height As Double, ByVal ladderWidth As Double, ByRef stepDiameter As Double) As Double

            Dim sideFrameThickness#, bottomExtension#, shDimension1#, shDimension2#, shDimension3#, flareShDimension1#
            Dim flareShDimension2#, hoopBendRadius#, oldHoopRadius#

            Try
                sideFrameThickness = StructHelper.GetDoubleProperty(ladder, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.SideFrameThickness)
                bottomExtension = StructHelper.GetDoubleProperty(ladder, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.BottomExtension)
                shDimension1 = StructHelper.GetDoubleProperty(ladderPart, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.ShDimension1)
                shDimension2 = StructHelper.GetDoubleProperty(ladderPart, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.ShDimension2)
                shDimension3 = StructHelper.GetDoubleProperty(ladderPart, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.ShDimension3)
                flareShDimension1 = StructHelper.GetDoubleProperty(ladderPart, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.FlareShDimension1)
                flareShDimension2 = StructHelper.GetDoubleProperty(ladderPart, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.FlareShDimension2)
                hoopBendRadius = StructHelper.GetDoubleProperty(ladderPart, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.HoopBendRadius)
                oldHoopRadius = StructHelper.GetDoubleProperty(ladder, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.HoopRadius)
                stepDiameter = StructHelper.GetDoubleProperty(ladder, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.StepDiameter)
            Catch ex As Exception
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, ladder.ToString + " " + LadderSymbolsLocalizer.GetString(LadderSymbolsResourceIDs.ErrUserAttributeDataNotFound, _
                                        "Error while computing length of ladder in LadderTypeA symbol, as some of the user attribute values cannot be obtained from the catalog part. Check the error log and catalog data."))
                Return 0 'stop evaluating
            End Try

            Dim hoopRadius# = shDimension2 + (ladderWidth / 2) + (2 * hoopBendRadius) + sideFrameThickness

            'Check whether the old HoopRadius is same or different to update the other depending parameters
            If Not StructHelper.AreEqual(oldHoopRadius, hoopRadius) Then
                oldHoopRadius = hoopRadius
                Dim hoopClearance# = hoopRadius + (shDimension3 - shDimension1)
                Dim flareRadius# = flareShDimension2 + (ladderWidth / 2) + (2 * hoopBendRadius) + sideFrameThickness
                Dim flareClearance# = flareRadius + (shDimension3 - flareShDimension1)

                Try
                    ladder.SetPropertyValue(oldHoopRadius, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.HoopRadius)
                    ladder.SetPropertyValue(hoopClearance, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.HoopClearance)
                    ladder.SetPropertyValue(flareRadius, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.FlareRadius)
                    ladder.SetPropertyValue(flareClearance, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.FlareClearance)
                Catch ex As Exception
                    'attributes might be missing, create todo record
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, ladder.ToString + " " + LadderSymbolsLocalizer.GetString(LadderSymbolsResourceIDs.ErrSetUserAttributeData, _
                                        "Error while computing length of ladder in LadderTypeA symbol, as some of the user attribute values can not be obtained from the catalog part. Check the error log and catalog data."))
                    Return 0 'stop evaluating
                End Try
            End If

            Return height - bottomExtension

        End Function

        ''' <summary>
        ''' Sets the step properties - number of steps and first step pitch on the ladder.
        ''' </summary>
        ''' <param name="ladder">Ladder business object which aggregates symbol.</param>
        ''' <param name="numberOfSteps">The number of steps on the ladder.</param>
        ''' <param name="firstStepPitch">The first step pitch.</param>
        Public Overrides Sub SetStepProperties(ByVal ladder As Ladder, ByVal numSteps As Integer, ByVal firstStepPitch As Double)
            ladder.SetPropertyValue(numSteps, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.NumSteps)
            ladder.SetPropertyValue(firstStepPitch, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.FirstStepPitch)
        End Sub

        ''' <summary>
        ''' Get the top extension length above the top support for ladder symbol.
        ''' </summary>
        ''' <param name="ladder">Ladder business object which aggregates symbol.</param>
        ''' <returns>Top extension length above the top support.</returns>
        Public Overrides Function GetTopExtension(ByVal ladder As Ladder) As Double
            Try
                Return StructHelper.GetDoubleProperty(ladder, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.TopExtension)
            Catch ex As Exception
                'unable to get the topextension,create a warning todo record and exit
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SPSSymbolConstants.TopExtension + " " + _
                                            LadderSymbolsLocalizer.GetString(LadderSymbolsResourceIDs.ErrPropertyNotFound, _
                                        "Property does not exist on the ladder in LadderTypeA symbol. Please check the error log and catalog part."))
                Return 0
            End Try

        End Function

        ''' <summary>
        ''' Makes the property read-only. 
        ''' </summary>
        ''' <param name="interfaceName">Interface name of the property.</param>
        ''' <param name="propertyName">The name of the property.</param>
        Public Overrides Function IsPropertyReadOnly(ByVal interfaceName$, ByVal propertyName$) As Boolean

            Select Case propertyName
                'making following property values as read-only. 
                Case SPSSymbolConstants.FirstStepPitch, SPSSymbolConstants.NumSteps, SPSSymbolConstants.VlDimension1, _
                SPSSymbolConstants.VlDimension2, SPSSymbolConstants.VlDimension3, SPSSymbolConstants.HoopPlateThickness, _
                 SPSSymbolConstants.HoopPlateWidth, SPSSymbolConstants.HoopClearance, SPSSymbolConstants.HoopRadius, _
                 SPSSymbolConstants.FlareClearance, SPSSymbolConstants.FlareRadius, SPSSymbolConstants.HoopFlareBendRadius, _
                 SPSSymbolConstants.HoopFlareHeight, SPSSymbolConstants.Span, SPSSymbolConstants.Height, SPSSymbolConstants.Length, _
                 SPSSymbolConstants.Angle
                    Return True
            End Select

        End Function

        ''' <summary>
        ''' Validates the given property.
        ''' </summary>
        ''' <param name="interfaceName">Interface name of the property.</param>
        ''' <param name="propertyName">The name of the property.</param>
        ''' <param name="propertyValue">The value of the property.</param>
        ''' <param name="errorMessage">The error message if validation fails.</param>
        ''' <returns>True if property value validation succeeds.</returns>
        Public Overrides Function IsPropertyValid(ByVal interfaceName$, ByVal propertyName$, ByVal propertyValue As Object, ByRef errorMessage$) As Boolean

            'by default set the property value as valid. Override the value later for known checks
            Dim isValidPropertyValue As Boolean = True

            If propertyValue IsNot Nothing Then
                Select Case propertyName
                    'following property values need to be in between 75-90 degrees 
                    Case SPSSymbolConstants.Angle
                        isValidPropertyValue = ValidationHelper.IsBetween90And75(CDbl(propertyValue), errorMessage)
                        Exit Select
                        'following property values must be lower than 100
                    Case SPSSymbolConstants.Span
                        isValidPropertyValue = ValidationHelper.IsLowerThan100(CDbl(propertyValue), errorMessage)
                        Exit Select
                        'following property values must be positive
                    Case SPSSymbolConstants.TopExtension, SPSSymbolConstants.BottomExtension
                        isValidPropertyValue = Not ValidationHelper.IsNegative(CDbl(propertyValue), errorMessage)
                        Exit Select
                        'following property values must be greater than 0
                    Case SPSSymbolConstants.VerticalStrapCount
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(CInt(propertyValue), errorMessage)
                        Exit Select
                        'following property values must be greater than 0
                    Case SPSSymbolConstants.Width, SPSSymbolConstants.StepPitch, SPSSymbolConstants.EnvelopeHeight, SPSSymbolConstants.WallOffset, _
                    SPSSymbolConstants.BottomHoopLevel, SPSSymbolConstants.VerticalStrapWidth, SPSSymbolConstants.VerticalStrapThickness, _
                    SPSSymbolConstants.FlareRadius, SPSSymbolConstants.FlareClearance, SPSSymbolConstants.HoopRadius, _
                    SPSSymbolConstants.HoopClearance, SPSSymbolConstants.HoopPitch, SPSSymbolConstants.StepProtrusion, _
                    SPSSymbolConstants.StepDiameter, SPSSymbolConstants.SideFrameWidth, SPSSymbolConstants.SideFrameThickness, _
                    SPSSymbolConstants.SupportLegWidth, SPSSymbolConstants.SupportLegThickness, SPSSymbolConstants.SupportLegPitch
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(CDbl(propertyValue), errorMessage)
                        Exit Select
                End Select
            End If

            Return isValidPropertyValue

        End Function

        ''' <summary>
        ''' Sets the properties on the ladder part on mirror.
        ''' </summary>
        ''' <param name="ladder">Ladder business object which aggregates symbol.</param>
        ''' <param name="ladderPart">The ladder part.</param>
        Public Overrides Sub SetPropertiesOnMirror(ByVal ladder As Ladder, ByVal ladderPart As IPart)

            'Default behavior for replacement part mirror is to set the properties on the target object
            ' to those of source object but we want the ladder to retain the hoop opening property of part.

            'Get mirror behavior specified on the catalog part
            Dim mirrorBehaviorOption As Integer = ladderPart.MirrorBehaviorOption

            If mirrorBehaviorOption = SPSSymbolConstants.REPLACEMENT_PART_VALUE Then

                'Now get the mirror hoop opening value on the occurrence and definition
                Dim hoopOpeningPart As Integer = StructHelper.GetIntProperty(DirectCast(ladderPart, BusinessObject), SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.HoopOpening)
                Dim hoopOpeningBO As Integer = StructHelper.GetIntProperty(ladder, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.HoopOpening)

                If hoopOpeningPart <> hoopOpeningBO Then
                    ladder.SetPropertyValue(hoopOpeningPart, SPSSymbolConstants.IJUALadderTypeA, SPSSymbolConstants.HoopOpening)
                End If

            End If
        End Sub

        ''' <summary>
        ''' Sets the properties in ladder.
        ''' Provides a way to override the part default propeties.
        ''' </summary>
        ''' <param name="ladder">Ladder business object which aggregates symbol.</param>
        Public Overrides Sub SetProperties(ByVal ladder As Ladder)
            'for the vertical ladder here we setting the angle value as 90 degree or Math.PI / 2
            ladder.SetPropertyValue(Math.PI / 2, SPSSymbolConstants.IJSPSCommonStairLadderProps, SPSSymbolConstants.Angle)
        End Sub

        ''' <summary>
        ''' Checks for undefined value and raise error.
        ''' </summary>
        ''' <param name="ladder">Ladder business object which aggregates symbol.</param>
        Public Overrides Sub ValidateUndefinedCodelistValue(ByVal ladder As Ladder)
            'validate justification
            ValidateUndefinedCodelistValue(ladder, SPSSymbolConstants.Justification, SPSSymbolConstants.TDL_INVALID_JUSTIFICATION)

            'validate hoop opening
            ValidateUndefinedCodelistValue(ladder, SPSSymbolConstants.HoopOpening, SPSSymbolConstants.TDL_INVALID_HOOPOPENING)
        End Sub

        ''' <summary>
        ''' Validates the undefined codelist value.
        ''' </summary>
        ''' <param name="ladder">The ladder.</param>
        ''' <param name="propertyName">Name of the property.</param>
        ''' <param name="errNumber">The error number.</param>
        Private Overloads Sub ValidateUndefinedCodelistValue(ByVal ladder As Ladder, ByVal propertyName$, ByVal errNumber As Integer)
            Dim propertyValue As Integer
            Try
                propertyValue = StructHelper.GetIntProperty(ladder, SPSSymbolConstants.IJUALadderTypeA, propertyName)
            Catch ex As Exception
                'unable to get the property value,create a warning todo record and exit
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, propertyName + " " + _
                                         LadderSymbolsLocalizer.GetString(LadderSymbolsResourceIDs.ErrPropertyNotFound, _
                                                            "Property does not exist on the ladder in LadderTypeA symbol. Please check the error log and catalog part."))
                Return
            End Try

            'now, check the value valid or not
            If Not StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructAlignment, SPSSymbolConstants.REFDAT, propertyValue) Then
                'create a todo record with appropriate message
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, errNumber, String.Empty, ladder)
            End If

        End Sub
#End Region

    End Class

End Namespace



