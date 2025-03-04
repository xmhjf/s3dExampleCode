''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright 1992 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  StairTypeA.vb
'
'Abstract
'	This is .NET StairTypeA symbol. This class subclasses from StairSymbolDefinition.
'
'   July 09, 2012 SS - CR184504 GetCrossSectionDimensions should be moved from SymbolHelper to CrossSectionService
'
'   July 17, 2012 SS - DM219538 Placing Ladder-A results in mirrored matrix - causing drawing issue
'                                       To follow RHS coordinate system changed the top support normal direction
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports System
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Exceptions
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.Structure.Middle.Services
Imports Ingr.Common.Middle.Services.Internal

'===========================================================================================
'Namespace of this class is Ingr.SP3D.Content.Structure
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'===========================================================================================

Namespace Ingr.SP3D.Content.Structure

    ''' <summary>
    ''' StairTypeA symbol definition constructs geometry of a Stair with optional Top Landing. 3 aspects- SimplePhysical representing actual physical geometry,
    ''' Operational- representing operational volume of the Stair , and Centerline- representing a simplified geometry for Drawings are created.
    ''' </summary>
    Public Class StairTypeA : Inherits StairSymbolDefinition
        Implements ICustomWeightCG
        '=============================================================================================
        'DefinitionName/ProgID of this symbol is "StairSymbols,Ingr.SP3D.Content.Structure.StairTypeA"
        '=============================================================================================

#Region "Definition of Inputs"

        <InputCatalogPart(1)> _
            Public partInput As InputCatalogPart
        <InputDouble(2, "Width", "Width", 0.6)> _
            Public widthInput As InputDouble
        <InputDouble(3, "Angle", "Angle", 55)> _
        Public angleInput As InputDouble
        <InputDouble(4, "StepPitch", "Step Pitch", 0.25)> _
        Public stepPitchInput As InputDouble
        <InputDouble(5, "Height", "Height", 1.0)> _
        Public heightInput As InputDouble
        <InputDouble(6, "NumSteps", "Number of Steps", 0.0)> _
        Public numberOfStepsInput As InputDouble
        <InputDouble(7, "Span", "Span", 1)> _
        Public spanInput As InputDouble
        <InputDouble(8, "Length", "Length", 1)> _
        Public lengthInput As InputDouble
        <InputDouble(9, "Justification", "Justification", 1)> _
        Public justificationInput As InputDouble
        <InputDouble(10, "TopSupportSide", "Top Support Side", 1)> _
        Public topSupportSideInput As InputDouble
        <InputDouble(11, "SideFrameSectionCP", "Side Frame Section Cardinal Point", 4)> _
        Public sideFrameSectionCPInput As InputDouble
        <InputDouble(12, "SideFrameSectionAngle", "Side Frame Section Angle", 0)> _
        Public sideFrameSectionAngleInput As InputDouble
        <InputDouble(13, "HandRailSectionCP", "HandRail Section Cardinal Point", 0)> _
        Public handRailSectionCPInput As InputDouble
        <InputDouble(14, "HandRailSectionAngle", "HandRail Section Angle", 0)> _
        Public handRailSectionAngleInput As InputDouble
        <InputDouble(15, "StepSectionCP", "Step Section Cardinal Point", 5)> _
        Public stepSectionCPInput As InputDouble
        <InputDouble(16, "StepSectionAngle", "Step Section Angle", 0)> _
        Public stepSectionAngleInput As InputDouble
        <InputDouble(17, "PlatformThickness", "Platform Thickness", 0.02)> _
        Public platformThicknessInput As InputDouble
        <InputDouble(18, "WithTopLanding", "With Top Landing", 1)> _
        Public withTopLandingInput As InputDouble
        <InputDouble(19, "TopLandingLength", "Top Landing Length", 0.06)> _
        Public topLandingLengthInput As InputDouble
        <InputDouble(20, "PostHeight", "Post Height", 1.2)> _
        Public postHeightInput As InputDouble
        <InputDouble(21, "HandRailPostPitch", "HandRail Post Pitch", 1)> _
        Public handRailPostPitchInput As InputDouble
        <InputDouble(22, "NumMidRails", "Number of MidRails", 3)> _
        Public numberOfMidRailsInput As InputDouble
        <InputDouble(23, "IsAssembly", "Is Assembly", 0)> _
        Public isAssemblyInput As InputDouble
        <InputDouble(24, "IsSystem", "Is System", 0)> _
        Public isSystemInput As InputDouble
        <InputDouble(25, "EnvelopeHeight", "Envelope Height", 0)> _
        Public envelopeHeightInput As InputDouble
        <InputString(26, "SideFrame_SPSSectionName", "Side Frame Section Name", "C12x25")> _
        Public sideFrameSectionNameInput As InputString
        <InputString(27, "SideFrame_SPSSectionRefStandard", "Side Frame Section Reference Standard", "AISC-LRFD-3.1")> _
        Public sideFrameSectionRefStandardInput As InputString
        <InputString(28, "HandRail_SPSSectionName", "HandRail Section Name", "PIPE2sch40")> _
        Public handRailSectionNameInput As InputString
        <InputString(29, "HandRail_SPSSectionRefStandard", "HandRail Section Reference Standard", "AISC-LRFD-3.1")> _
        Public handRailSectionRefStandardInput As InputString
        <InputString(30, "Step_SPSSectionName", "Step Section Name", "C12x25")> _
        Public stepSectionNameInput As InputString
        <InputString(31, "Step_SPSSectionRefStandard", "Step Section Reference Standard", "AISC-LRFD-3.1")> _
        Public stepSectionRefStandardInput As InputString
        <InputString(32, "Primary_SPSMaterial", "Primary Material", "Steel - Carbon")> _
        Public primaryMaterialInput As InputString
        <InputString(33, "Primary_SPSGrade", "Primary Material Grade", "A")> _
        Public primaryMaterialGradeInput As InputString

#End Region

#Region "Definitions of Aspects and their outputs"

        'SimplePhysical Aspect
        <Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)> _
        Public simplePhysicalAspect As AspectDefinition

        'Operation Aspect
        <Aspect("Operation", "Operation Aspect", AspectID.Operation)> _
        <SymbolOutput("OperationalEnvelope1", "Operational envelope of the stair")> _
        Public operationAspect As AspectDefinition

        'Centerline Aspect
        <Aspect("Centerline", "Centerline Aspect", AspectID.Centerline)> _
        Public centerlineAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

        ''' <summary>
        ''' Implemented by the final inheriting concrete class that constructs
        ''' the symbol outputs for all aspects.
        ''' </summary>
        Protected Overrides Sub ConstructOutputs()
            Try
                '==============================================================
                ' Construction of Simple Physical Aspect and Centerline Aspect 
                '==============================================================
                Dim sectionWidth As Double, tempDepth As Double, sideWidthY As Double, sideWidthZ As Double
                Dim totalWidth As Double, depth As Double
                Dim stepDepth As Double, stepSpan As Double
                Dim handRailDepth As Double, handRailWidth As Double
                Dim upperLeftCorner As New Position, upperRightCorner As New Position
                Dim lowerLeftCorner As New Position, lowerRightCorner As New Position
                Dim position As Double
                Dim sweepOptions As SweepOptions = sweepOptions.CreateCaps
                Dim currentAspectID As AspectID

                'Get the connection where outputs will be created.
                Dim connection As SP3DConnection = MyBase.OccurrenceConnection
                'Get Input values              
                Dim length As Double = lengthInput.Value
                Dim width As Double = widthInput.Value
                Dim angle As Double = angleInput.Value
                Dim height As Double = heightInput.Value
                Dim stepPitch As Double = stepPitchInput.Value
                Dim sideSectionAngle As Double = sideFrameSectionAngleInput.Value
                Dim platformThickness As Double = platformThicknessInput.Value
                Dim topLandingLength As Double = topLandingLengthInput.Value
                Dim envelopeHeight As Double = envelopeHeightInput.Value
                Dim handRailSectionAngle As Double = handRailSectionAngleInput.Value
                Dim handRailPostPitch As Double = handRailPostPitchInput.Value
                Dim postHeight As Double = postHeightInput.Value
                Dim stepSectionRefStandard As String = CStr(stepSectionRefStandardInput.Value)
                Dim stepSectionName As String = CStr(stepSectionNameInput.Value)
                Dim handRailSectionReferenceStandard As String = CStr(handRailSectionRefStandardInput.Value)
                Dim handRailSectionName As String = CStr(handRailSectionNameInput.Value)
                Dim sideFrameSectionName As String = CStr(sideFrameSectionNameInput.Value)
                Dim sideFrameSectionReferenceStandard As String = CStr(sideFrameSectionRefStandardInput.Value)
                Dim withTopLanding As Boolean = CBool(withTopLandingInput.Value)
                Dim justification As Integer = CInt(justificationInput.Value)
                Dim numberOfMidRails As Integer = CInt(numberOfMidRailsInput.Value)
                Dim numberOfSteps As Integer = CInt(numberOfStepsInput.Value)
                Dim sideFrameSectionCP As Integer = CInt(sideFrameSectionCPInput.Value)
                Dim handRailSectionCP As Integer = CInt(handRailSectionCPInput.Value)

                If String.IsNullOrEmpty(stepSectionRefStandard) Or String.IsNullOrEmpty(stepSectionName) Then
                    'With empty or null section properties, symbol cannot be created.
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, _
                                              StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrMissingStepSectionProperties,
                                            "Cannot construct Stair outputs in StairTypeA Symbol, as either or both of StepSectionName and StepSectionRefStandard are missing for Step section properties. Please check catalog or contact S3D support."))
                    Return
                End If

                If String.IsNullOrEmpty(handRailSectionReferenceStandard) Or String.IsNullOrEmpty(handRailSectionName) Then
                    'With empty or null section properties, symbol cannot be created.
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, _
                                              StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrMissingHandRailSectionProperties,
                                            "Cannot construct Stair outputs in StairTypeA Symbol, as either or both of HandRailSectionName and HandRailSectionReferenceStandard are missing for Handrail section properties. Please check catalog or contact S3D support."))
                    Return
                End If
                If String.IsNullOrEmpty(sideFrameSectionName) Or String.IsNullOrEmpty(sideFrameSectionReferenceStandard) Then
                    'With empty or null section properties, symbol cannot be created.
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, _
                                              StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrMissingSideFrameSectionProperties,
                                            "Cannot construct Stair outputs in StairTypeA Symbol, as either or both of SideFrameSectionName and SideFrameSectionReferenceStandard are missing for SideFrame section properties. Please check catalog or contact S3D support."))
                    Return
                End If

                angle = (Math.PI / 2 - angle)

                If justification = SPSSymbolConstants.ALIGNMENT_LEFT Then
                    position = -(width / 2)
                ElseIf justification = SPSSymbolConstants.ALIGNMENT_RIGHT Then
                    position = (width / 2)
                End If

                'setting corner positions
                upperLeftCorner.X = -width / 2 - position
                upperLeftCorner.Y = 0.0
                upperLeftCorner.Z = 0.0

                upperRightCorner.X = width / 2 - position
                upperRightCorner.Y = 0.0
                upperRightCorner.Z = 0.0

                lowerLeftCorner.X = -width / 2 - position
                lowerLeftCorner.Y = -Math.Tan(angle) * height
                lowerLeftCorner.Z = -height

                lowerRightCorner.X = width / 2 - position
                lowerRightCorner.Y = -Math.Tan(angle) * height
                lowerRightCorner.Z = -height

                Dim sideFrameCrosssection As CrossSection = MyBase.GetCrossSection(sideFrameSectionReferenceStandard, sideFrameSectionName)
                'might be created a todo record if section is missing
                If MyBase.ToDoListMessage IsNot Nothing Then
                    If MyBase.ToDoListMessage.Type = ToDoMessageTypes.ToDoMessageError Then
                        Return
                    End If
                End If

                Dim stepCrosssection As CrossSection = MyBase.GetCrossSection(stepSectionRefStandard, stepSectionName)
                'might be created a todo record if section is missing
                If MyBase.ToDoListMessage IsNot Nothing Then
                    If MyBase.ToDoListMessage.Type = ToDoMessageTypes.ToDoMessageError Then
                        Return
                    End If
                End If
                Dim handRailCrosssection As CrossSection = MyBase.GetCrossSection(handRailSectionReferenceStandard, handRailSectionName)
                'might be created a todo record if section is missing
                If MyBase.ToDoListMessage IsNot Nothing Then
                    If MyBase.ToDoListMessage.Type = ToDoMessageTypes.ToDoMessageError Then
                        Return
                    End If
                End If

                'The code within the loop, should be executed twice for both SimplePhysical and Centerline Aspect 
                For nLoop = 1 To 2
                    If nLoop = 1 Then
                        currentAspectID = AspectID.SimplePhysical
                    Else
                        currentAspectID = AspectID.Centerline
                    End If

                    If currentAspectID = AspectID.SimplePhysical Then
                        depth = sideFrameCrosssection.Depth 'CR184504
                        sectionWidth = sideFrameCrosssection.Width
                        tempDepth = (depth / 2) / Math.Cos(angle) ' to get to the center axis of the member and place steps on that axis
                    End If

                    sideWidthY = 2 * tempDepth
                    sideWidthZ = 2 * tempDepth / Math.Tan(angle)

                    If withTopLanding Then
                        '=======================
                        'Stair with TopLanding
                        '=======================
                        'using same section for sideframes and toplanding

                        'Place left side frame 
                        PlaceLeftSideFrameWithTopLanding(upperLeftCorner, lowerLeftCorner, _
                                                          lowerRightCorner, angle, height, topLandingLength, _
                                                          sideFrameCrosssection, currentAspectID, sideWidthY, sideWidthZ, _
                                                          sideFrameSectionCP, sideSectionAngle, sweepOptions)

                        'Place right side frame 
                        PlaceRightSideFrameWithTopLanding(upperRightCorner, topLandingLength, _
                                                          lowerRightCorner, lowerLeftCorner, sideFrameCrosssection, currentAspectID, _
                                                          sideWidthY, sideWidthZ, sideFrameSectionCP, sideSectionAngle, sweepOptions)

                        'Place front toplanding
                        PlaceFrontTopLanding(upperLeftCorner, upperRightCorner, sideFrameCrosssection, _
                                             currentAspectID, sideFrameSectionCP, sideSectionAngle, topLandingLength, sweepOptions)

                        'Place back toplanding
                        PlaceBackTopLanding(upperRightCorner, upperLeftCorner, sideFrameCrosssection, _
                                            currentAspectID, sideFrameSectionCP, sideSectionAngle, sweepOptions)

                        'Place middle frame
                        If topLandingLength > 1.5 Then
                            PlaceMiddleFrame(upperRightCorner, upperLeftCorner, currentAspectID, topLandingLength, _
                                             sideFrameCrosssection, sideFrameSectionCP, sideSectionAngle, sweepOptions)
                        End If

                        'Place platform
                        PlacePlatform(position, topLandingLength, currentAspectID, _
                                      upperRightCorner, upperLeftCorner, sideWidthZ, width, platformThickness)

                    Else
                        '==========================
                        'Stair without TopLanding
                        '==========================
                        'Place left side frame 
                        PlaceLeftSideFrame(upperLeftCorner, lowerLeftCorner, _
                                           sideFrameCrosssection, currentAspectID, sideWidthY, sideWidthZ, _
                                           sideFrameSectionCP, sideSectionAngle, sweepOptions)

                        'Place right side frame 
                        PlaceRightSideFrame(upperRightCorner, lowerRightCorner, _
                                            lowerLeftCorner, sideFrameCrosssection, currentAspectID, sideWidthY, _
                                            sideWidthZ, sideFrameSectionCP, sideSectionAngle, sweepOptions)

                    End If

                    'Place steps
                    'using different section for steps
                    If currentAspectID = AspectID.SimplePhysical Then
                        tempDepth = (depth / 2) / Math.Sin(angle)

                        'here the StepDepth is really the width of the step section
                        stepSpan = stepCrosssection.Depth 'CR184504
                        stepDepth = stepCrosssection.Width
                        'only support step section CP5 as default CP
                        stepDepth = stepDepth / 2
                        stepSpan = stepSpan / 2
                    End If

                    Dim actualPitch As Double = height / (numberOfSteps + 1)

                    PlaceSteps(upperLeftCorner, upperRightCorner, numberOfSteps, angle, _
                                   actualPitch, tempDepth, stepDepth, stepSpan, withTopLanding, _
                                   topLandingLength, stepCrosssection, currentAspectID)

                    'Place handrails and posts
                    upperLeftCorner.X = upperLeftCorner.X - (sectionWidth / 2)
                    upperRightCorner.X = upperRightCorner.X + (sectionWidth / 2)
                    lowerLeftCorner.X = lowerLeftCorner.X - (sectionWidth / 2)
                    lowerRightCorner.X = lowerRightCorner.X + (sectionWidth / 2)

                    'using same section for posts and handrails
                    If currentAspectID = AspectID.SimplePhysical Then
                        handRailDepth = handRailCrosssection.Depth 'CR184504
                        handRailWidth = handRailCrosssection.Width
                    End If

                    'Place posts
                    PlacePosts(handRailCrosssection, angle, upperLeftCorner, lowerLeftCorner, _
                               upperRightCorner, lowerRightCorner, length, handRailSectionCP, handRailSectionAngle, _
                                postHeight, withTopLanding, topLandingLength, height, handRailDepth, handRailPostPitch, currentAspectID)

                    'Place handrail
                    If withTopLanding Then
                        upperLeftCorner.Y = 0.0
                        upperRightCorner.Y = 0.0
                        PlaceHandRailsWithToplanding(handRailCrosssection, upperLeftCorner, lowerLeftCorner, upperRightCorner, _
                                                     lowerRightCorner, handRailSectionCP, handRailSectionAngle, _
                                                     postHeight, topLandingLength, numberOfMidRails, handRailDepth, currentAspectID)
                    Else
                        PlaceHandrails(handRailCrosssection, upperLeftCorner, lowerLeftCorner, _
                                           upperRightCorner, lowerRightCorner, handRailSectionCP, _
                                           handRailSectionAngle, postHeight, topLandingLength, _
                                           numberOfMidRails, handRailDepth, currentAspectID)
                    End If
                Next

                '======================================
                'Construction of Operational Aspect
                '======================================
                'we will create the side of the stair and then project it along the width to create the envelope
                'all the points for the side will be in local xy, assuming width to be along the z-axis.
                'Total width is 2 times the handrail section width plus the stair width.
                totalWidth = width + 2 * sectionWidth

                'Place operational envelope
                PlaceEnvelope(withTopLanding, totalWidth, position)

            Catch Ex As Exception ' General Unhandled exception 
                If MyBase.ToDoListMessage Is Nothing Then 'Check ToDoListMessgae created already or not
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureErrorsMessageCodelist, SPSSymbolConstants.TDL_INVALID_STAIR_GEOMETRY, _
                                                          StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrConstructOutputs, "Error in constructing Stair outputs in StairTypeA Symbol due to invalid inputs. Check custom code or contact S3D support"))
                End If
            End Try

        End Sub

#End Region

#Region "Private Functions and Methods"

#Region "Left Side Frame Creation"

        ''' <summary>
        ''' Places the left side frame with top landing.
        ''' </summary>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="lowerLeftCorner">The lower left corner.</param>
        ''' <param name="lowerRightCorner">The lower right corner.</param>
        ''' <param name="angle">The angle.</param>
        ''' <param name="height">The height.</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        ''' <param name="sideWidthY">The side width y.</param>
        ''' <param name="sideWidthZ">The side width z.</param>
        ''' <param name="sideFrameSectionCP">The side frame section cp.</param>
        ''' <param name="sideSectionAngle">The side section angle.</param>
        ''' <param name="sweepOptions">The sweep options.</param>
        Private Sub PlaceLeftSideFrameWithTopLanding(ByVal upperLeftCorner As Position, _
                                                     ByVal lowerLeftCorner As Position, ByVal lowerRightCorner As Position, _
                                                     ByVal angle As Double, ByVal height As Double, ByVal topLandingLength As Double, _
                                                     ByVal crosssection As CrossSection, ByVal currentAspectID As AspectID, _
                                                     ByVal sideWidthY As Double, ByVal sideWidthZ As Double, _
                                                     ByVal sideFrameSectionCP As Integer, ByVal sideSectionAngle As Double, _
                                                     ByVal sweepOptions As SweepOptions)

            Dim startPosition As New Position(), endPosition As New Position()
            Dim line As Line3d

            lowerLeftCorner.Y = (-Math.Tan(angle) * height) - topLandingLength
            lowerRightCorner.Y = (-Math.Tan(angle) * height) - topLandingLength

            If currentAspectID = AspectID.Centerline Then
                startPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y - topLandingLength), upperLeftCorner.Z)
                endPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y - topLandingLength), upperLeftCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_1", line)

                startPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y - topLandingLength), upperLeftCorner.Z - sideWidthZ)
                endPosition.Set(lowerLeftCorner.X, lowerLeftCorner.Y + sideWidthY, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_2", line)

                startPosition.Set(lowerLeftCorner.X, lowerLeftCorner.Y + sideWidthY, lowerLeftCorner.Z)
                line = New Line3d(startPosition, lowerLeftCorner)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_3", line)

                endPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y - topLandingLength), upperLeftCorner.Z)
                line = New Line3d(lowerLeftCorner, endPosition)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_4", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                startPosition.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
                endPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y - topLandingLength), upperLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)

                Dim curves As New Collection(Of ICurve)
                curves.Add(line)
                Dim complexString As New ComplexString3d(curves)

                startPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y - topLandingLength), upperLeftCorner.Z)
                endPosition.Set(lowerLeftCorner.X, lowerLeftCorner.Y, lowerLeftCorner.Z)
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                Dim startNormal As New Vector(0.0, 0.0, 1.0)
                Dim crossSectionServices As New CrossSectionServices()
                Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, sideFrameSectionCP, False, sideSectionAngle, Nothing, Nothing, lowerLeftCorner, startNormal, sweepOptions)
                If Not surfaces Is Nothing Then
                    For j = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("LeftSideFrame" & j, surfaces(j))
                    Next j
                End If

            End If
        End Sub

        ''' <summary>
        ''' Places the left side frame.
        ''' </summary>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="lowerLeftCorner">The lower left corner.</param>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        ''' <param name="sideWidthY">The side width y.</param>
        ''' <param name="sideWidthZ">The side width z.</param>
        ''' <param name="sideFrameSectionCP">The side frame section cp.</param>
        ''' <param name="sideSectionAngle">The side section angle.</param>
        ''' <param name="sweepOptions">The sweep options.</param>
        Private Sub PlaceLeftSideFrame(ByVal upperLeftCorner As Position, _
                                       ByVal lowerLeftCorner As Position, ByVal crosssection As CrossSection, _
                                       ByVal currentAspectID As AspectID, ByVal sideWidthY As Double, _
                                       ByVal sideWidthZ As Double, ByVal sideFrameSectionCP As Integer, _
                                       ByVal sideSectionAngle As Double, ByVal sweepOptions As SweepOptions)

            Dim startPosition As New Position(), endPosition As New Position()
            Dim line As Line3d

            If currentAspectID = AspectID.Centerline Then
                startPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y), upperLeftCorner.Z)
                endPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y), upperLeftCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_1", line)

                startPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y), upperLeftCorner.Z - sideWidthZ)
                endPosition.Set(lowerLeftCorner.X, lowerLeftCorner.Y + sideWidthY, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_2", line)

                startPosition.Set(lowerLeftCorner.X, lowerLeftCorner.Y + sideWidthY, lowerLeftCorner.Z)
                line = New Line3d(startPosition, lowerLeftCorner)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_3", line)

                endPosition.Set(upperLeftCorner.X, (upperLeftCorner.Y), upperLeftCorner.Z)
                line = New Line3d(lowerLeftCorner, endPosition)
                centerlineAspect.Outputs.Add("LeftSideFrame" & "_4", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                startPosition.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
                endPosition.Set(lowerLeftCorner.X, lowerLeftCorner.Y, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)

                Dim startNormal As New Vector(0.0, 1.0, 0.0)
                Dim endNormal As New Vector(0.0, 0.0, 1.0)
                Dim crossSectionServices As New CrossSectionServices()
                Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, sideFrameSectionCP, False, sideSectionAngle, upperLeftCorner, startNormal, lowerLeftCorner, endNormal, sweepOptions)
                If Not surfaces Is Nothing Then
                    For i = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("LeftSideFrame" & CStr(i), surfaces(i))
                    Next i
                End If
            End If
        End Sub

#End Region

#Region "Right Side Frame Creation"

        ''' <summary>
        ''' Places the right side frame with top landing.
        ''' </summary>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="lowerRightCorner">The lower right corner.</param>
        ''' <param name="lowerLeftCorner">The lower left corner.</param>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        ''' <param name="sideWidthY">The side width y.</param>
        ''' <param name="sideWidthZ">The side width z.</param>
        ''' <param name="sideFrameSectionCP">The side frame section cp.</param>
        ''' <param name="sideSectionAngle">The side section angle.</param>
        ''' <param name="sweepOptions">The sweep options.</param>
        Private Sub PlaceRightSideFrameWithTopLanding(ByVal upperRightCorner As Position, _
                                                      ByVal topLandingLength As Double, ByVal lowerRightCorner As Position, _
                                                      ByVal lowerLeftCorner As Position, ByVal crosssection As CrossSection, _
                                                      ByVal currentAspectID As AspectID, ByVal sideWidthY As Double, _
                                                      ByVal sideWidthZ As Double, ByVal sideFrameSectionCP As Integer, _
                                                      ByVal sideSectionAngle As Double, ByVal sweepOptions As SweepOptions)

            Dim startPosition As New Position(), endPosition As New Position()
            Dim line As Line3d

            If currentAspectID = AspectID.Centerline Then
                startPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                endPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_1", line)

                startPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z - sideWidthZ)
                endPosition.Set(lowerRightCorner.X, lowerRightCorner.Y + sideWidthY, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_2", line)

                startPosition.Set(lowerRightCorner.X, lowerRightCorner.Y + sideWidthY, lowerLeftCorner.Z)
                endPosition.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_3", line)

                startPosition.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerLeftCorner.Z)
                endPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_4", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                startPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                endPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                Dim curves As New Collection(Of ICurve)
                curves.Add(line)
                Dim complexString As New ComplexString3d(curves)

                startPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                endPosition.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerLeftCorner.Z)
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                Dim endNormal As New Vector(0.0, 0.0, 1.0)
                endPosition.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerLeftCorner.Z)
                Dim crossSectionServices As New CrossSectionServices()
                Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, sideFrameSectionCP, True, sideSectionAngle, Nothing, Nothing, endPosition, endNormal, sweepOptions)
                If Not surfaces Is Nothing Then
                    For j = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("RightSideFrame" & j, surfaces(j))
                    Next j
                End If
            End If
        End Sub

        ''' <summary>
        ''' Places the right side frame.
        ''' </summary>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="lowerRightCorner">The lower right corner.</param>
        ''' <param name="lowerLeftCorner">The lower left corner.</param>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        ''' <param name="sideWidthY">The side width y.</param>
        ''' <param name="sideWidthZ">The side width z.</param>
        ''' <param name="sideFrameSectionCP">The side frame section cp.</param>
        ''' <param name="sideSectionAngle">The side section angle.</param>
        ''' <param name="sweepOptions">The sweep options.</param>
        Private Sub PlaceRightSideFrame(ByVal upperRightCorner As Position, _
                                        ByVal lowerRightCorner As Position, ByVal lowerLeftCorner As Position, _
                                        ByVal crosssection As CrossSection, ByVal currentAspectID As AspectID, _
                                        ByVal sideWidthY As Double, ByVal sideWidthZ As Double, _
                                        ByVal sideFrameSectionCP As Integer, ByVal sideSectionAngle As Double, _
                                        ByVal sweepOptions As SweepOptions)

            Dim startPosition As New Position(), endPosition As New Position()
            Dim line As Line3d

            If currentAspectID = AspectID.Centerline Then
                startPosition.Set(upperRightCorner.X, (upperRightCorner.Y), upperRightCorner.Z)
                endPosition.Set(upperRightCorner.X, (upperRightCorner.Y), upperRightCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_1", line)

                startPosition.Set(upperRightCorner.X, (upperRightCorner.Y), upperRightCorner.Z - sideWidthZ)
                endPosition.Set(lowerRightCorner.X, lowerRightCorner.Y + sideWidthY, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_2", line)

                startPosition.Set(lowerRightCorner.X, lowerRightCorner.Y + sideWidthY, lowerLeftCorner.Z)
                endPosition.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_3", line)

                startPosition.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerLeftCorner.Z)
                endPosition.Set(upperRightCorner.X, (upperRightCorner.Y), upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("RightSideFrame" & "_4", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                startPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                endPosition.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerLeftCorner.Z)
                line = New Line3d(startPosition, endPosition)

                Dim startNormal As New Vector(0.0, 1.0, 0.0)
                Dim endNormal As New Vector(0.0, 0.0, 1.0)
                Dim crossSectionServices As New CrossSectionServices()
                Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, sideFrameSectionCP, True, sideSectionAngle, upperRightCorner, startNormal, lowerRightCorner, endNormal, sweepOptions)
                If Not surfaces Is Nothing Then
                    For i = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("RightSideFrame" & CStr(i), surfaces(i))
                    Next i
                End If
            End If
        End Sub

#End Region

#Region "Front Top Landing Creation"

        ''' <summary>
        ''' Places the front top landing.
        ''' </summary>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        ''' <param name="sideFrameSectionCP">The side frame section cp.</param>
        ''' <param name="sideSectionAngle">The side section angle.</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="sweepOptions">The sweep options.</param>
        Private Sub PlaceFrontTopLanding(ByVal upperLeftCorner As Position, _
                                         ByVal upperRightCorner As Position, ByVal crosssection As CrossSection, _
                                         ByVal currentAspectID As AspectID, ByVal sideFrameSectionCP As Integer, _
                                         ByVal sideSectionAngle As Double, ByVal topLandingLength As Double, _
                                         ByVal sweepOptions As SweepOptions)

            Dim startPosition As New Position(upperLeftCorner.X, (upperLeftCorner.Y - topLandingLength), upperLeftCorner.Z)
            Dim endPosition As New Position(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
            Dim line As Line3d

            If currentAspectID = AspectID.Centerline Then
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("FrontTopLanding", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                line = New Line3d(startPosition, endPosition)
                Dim crossSectionServices As New CrossSectionServices()
                Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, sideFrameSectionCP, True, sideSectionAngle, sweepOptions)
                If Not surfaces Is Nothing Then
                    For j = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("FrontTopLanding" & CStr(j), surfaces(j))
                    Next j
                End If
            End If

        End Sub

#End Region

#Region "Back Top Landing Creation"

        ''' <summary>
        ''' Places the back top landing.
        ''' </summary>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        ''' <param name="sideFrameSectionCP">The side frame section cp.</param>
        ''' <param name="sideSectionAngle">The side section angle.</param>
        ''' <param name="sweepOptions">The sweep options.</param>
        Private Sub PlaceBackTopLanding(ByVal upperLeftCorner As Position, _
                                        ByVal upperRightCorner As Position, ByVal crosssection As CrossSection, _
                                        ByVal currentAspectID As AspectID, ByVal sideFrameSectionCP As Integer, _
                                        ByVal sideSectionAngle As Double, ByVal sweepOptions As SweepOptions)

            Dim startPosition As New Position(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            Dim endPosition As New Position(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
            Dim line As Line3d

            If currentAspectID = AspectID.Centerline Then
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("BackTopLanding", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                line = New Line3d(startPosition, endPosition)
                Dim crossSectionServices As New CrossSectionServices
                Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, sideFrameSectionCP, True, sideSectionAngle, sweepOptions)
                If Not surfaces Is Nothing Then
                    For j = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("BackTopLanding" & j, surfaces(j))
                    Next j
                End If
            End If
        End Sub

#End Region

#Region "Middle Frame Creation"

        Private Sub PlaceMiddleFrame(ByVal upperLeftCorner As Position, _
                                     ByVal upperRightCorner As Position, ByVal currentAspectID As AspectID, _
                                     ByVal topLandingLength As Double, ByVal crosssection As CrossSection, _
                                     ByVal sideFrameSectionCP As Integer, ByVal sideSectionAngle As Double, _
                                     ByVal sweepOptions As SweepOptions)

            Dim startPosition As New Position(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            Dim endPosition As New Position(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
            Dim line As Line3d

            If currentAspectID = AspectID.Centerline Then
                line = New Line3d(upperLeftCorner, endPosition)
                centerlineAspect.Outputs.Add("MiddleFrame", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                line = New Line3d(startPosition, endPosition)
                Dim crossSectionServices As New CrossSectionServices()
                Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, sideFrameSectionCP, False, sideSectionAngle, sweepOptions)
                If Not surfaces Is Nothing Then
                    For j = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("MiddleFrame" & j, surfaces(j))
                    Next j
                End If
            End If

        End Sub

#End Region

#Region "Platform Creation"

        ''' <summary>
        ''' Places the platform.
        ''' </summary>
        ''' <param name="position">The position.</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="sideWidthZ">The side width z.</param>
        ''' <param name="width">The width.</param>
        ''' <param name="platformThickness">The platform thickness.</param>
        Private Sub PlacePlatform(ByVal position As Double, _
                                  ByVal topLandingLength As Double, ByVal currentAspectID As AspectID, _
                                  ByVal upperRightCorner As Position, ByVal upperLeftCorner As Position, _
                                  ByVal sideWidthZ As Double, ByVal width As Double, _
                                  ByVal platformThickness As Double)

            Dim projectionVector As New Vector, upDirection As New Vector
            Dim startPosition As New Position(), endPosition As New Position()
            Dim projection As Projection3d
            Dim center As New Position
            Dim line As Line3d

            center.X = 0 - position
            center.Y = 0.0
            center.Z = 0.0

            projectionVector.X = 0.0
            projectionVector.Y = -topLandingLength
            projectionVector.Z = 0.0

            upDirection.X = 0.0
            upDirection.Y = 0.0
            upDirection.Z = 1.0

            If currentAspectID = AspectID.Centerline Then
                endPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                line = New Line3d(upperRightCorner, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_1", line)

                startPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                endPosition.Set(upperLeftCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_2", line)

                startPosition.Set(upperLeftCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z)
                endPosition.Set(upperLeftCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_3", line)

                startPosition.Set(upperLeftCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                endPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_4", line)

                startPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z - sideWidthZ)
                endPosition.Set(upperRightCorner.X, upperRightCorner.Y - topLandingLength, upperRightCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_5", line)

                startPosition.Set(upperRightCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z - sideWidthZ)
                endPosition.Set(upperLeftCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_6", line)

                startPosition.Set(upperLeftCorner.X, (upperRightCorner.Y - topLandingLength), upperRightCorner.Z - sideWidthZ)
                endPosition.Set(upperLeftCorner.X, upperRightCorner.Y, upperRightCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_7", line)

                startPosition.Set(upperLeftCorner.X, upperRightCorner.Y, upperRightCorner.Z - sideWidthZ)
                endPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z - sideWidthZ)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_8", line)

                startPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z - sideWidthZ)
                endPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_9", line)

                startPosition.Set(upperLeftCorner.X, upperRightCorner.Y, upperRightCorner.Z - sideWidthZ)
                endPosition.Set(upperLeftCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                line = New Line3d(startPosition, endPosition)
                centerlineAspect.Outputs.Add("Platform" & "_10", line)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                projection = CreateRectangularProjection(center, projectionVector, upDirection, width, platformThickness)
                simplePhysicalAspect.Outputs.Add("Platform", projection)
            End If
        End Sub

#End Region

#Region "Steps Creation"

        ''' <summary>
        ''' Places the steps.
        ''' </summary>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="nNumSteps">The number of steps.</param>
        ''' <param name="angle">The angle.</param>
        ''' <param name="dPitch">The step pitch.</param>
        ''' <param name="inclinedDepth">The inclined depth.</param>
        ''' <param name="stepDepth">The step depth.</param>
        ''' <param name="stepSpan">The step span.</param>
        ''' <param name="withTopLanding">if set to <c>true</c>, Stair is with top landing.</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        Private Sub PlaceSteps(ByVal upperLeftCorner As Position, _
                               ByVal upperRightCorner As Position, ByVal numberOfSteps As Integer, _
                               ByVal angle As Double, ByVal pitch As Double, _
                               ByVal inclinedDepth As Double, ByVal stepDepth As Double, ByVal stepSpan As Double, _
                               ByVal withTopLanding As Boolean, ByVal topLandingLength As Double, _
                               ByVal crosssection As CrossSection, ByVal currentAspectID As AspectID)

            Dim startPosition As New Position, endPosition As New Position
            Dim startPoint As New Position, endPoint As New Position
            Dim line As Line3d

            Dim stepSectionCP As Integer = CInt(stepSectionCPInput.Value)
            Dim stepSectionAngle As Double = stepSectionAngleInput.Value

            startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            endPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)

            For i = 1 To numberOfSteps

                startPoint.Y = -Math.Tan(angle) * ((pitch * i) - inclinedDepth + stepDepth)
                endPoint.Y = startPoint.Y

                If withTopLanding Then
                    startPoint.Y = startPoint.Y - topLandingLength
                    endPoint.Y = endPoint.Y - topLandingLength
                End If

                startPoint.Z = -(i * pitch + stepDepth)
                endPoint.Z = startPoint.Z

                startPosition.Set(startPoint.X, startPoint.Y, startPoint.Z)
                endPosition.Set(endPoint.X, endPoint.Y, endPoint.Z)
                line = New Line3d(startPosition, endPosition)

                If currentAspectID = AspectID.Centerline Then
                    startPosition.Set(startPoint.X, startPoint.Y - stepSpan, startPoint.Z)
                    endPosition.Set(endPoint.X, endPoint.Y - stepSpan, endPoint.Z)
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("Step" & i & "_1", line)

                    startPosition.Set(startPoint.X, startPoint.Y + stepSpan, startPoint.Z)
                    endPosition.Set(endPoint.X, endPoint.Y + stepSpan, endPoint.Z)
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("Step" & i & "_2", line)

                    startPosition.Set(startPoint.X, startPoint.Y + stepSpan, startPoint.Z)
                    endPosition.Set(startPoint.X, startPoint.Y - stepSpan, startPoint.Z)
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("Step" & i & "_3", line)

                    startPosition.Set(endPoint.X, endPoint.Y - stepSpan, endPoint.Z)
                    endPosition.Set(endPoint.X, endPoint.Y + stepSpan, endPoint.Z)
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("Step" & i & "_4", line)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    Dim crossSectionServices As New CrossSectionServices
                    Dim sweepOptions As SweepOptions = sweepOptions.CreateCaps
                    Dim surfaces As Collection(Of ISurface) = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, stepSectionCP, False, stepSectionAngle, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("Step" & i & j, surfaces(j))
                        Next j
                    End If
                End If
            Next i

        End Sub

#End Region

#Region "Posts Creation"

        ''' <summary>
        ''' Places the posts.
        ''' </summary>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="angle">The angle.</param>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="lowerLeftCorner">The lower left corner.</param>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="lowerRightCorner">The lower right corner.</param>
        ''' <param name="length">The length.</param>
        ''' <param name="handRailSectionCP">The hand rail section cp.</param>
        ''' <param name="handRailSectionAngle">The hand rail section angle.</param>
        ''' <param name="postHeight">Height of the post.</param>
        ''' <param name="withTopLanding">if set to <c>true</c> [with top landing].</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="height">The height.</param>
        ''' <param name="handRailDepth">The hand rail depth.</param>
        ''' <param name="handrailPostPitch">The handrail post pitch.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        Private Sub PlacePosts(ByVal crosssection As CrossSection, _
                               ByVal angle As Double, ByVal upperLeftCorner As Position, ByVal lowerLeftCorner As Position, _
                               ByVal upperRightCorner As Position, ByVal lowerRightCorner As Position, _
                               ByVal length As Double, ByVal handRailSectionCP As Integer, _
                               ByVal handRailSectionAngle As Double, ByVal postHeight As Double, _
                               ByVal withTopLanding As Boolean, ByVal topLandingLength As Double, _
                               ByVal height As Double, ByVal handRailDepth As Double, _
                               ByVal handrailPostPitch As Double, ByVal currentAspectID As AspectID)

            Dim startPosition As New Position, endPosition As New Position
            Dim crossSectionServices As CrossSectionServices = Nothing
            Dim startPoint As New Position, endPoint As New Position
            Dim depthCorrection As Double
            Dim surfaces As Collection(Of ISurface)
            Dim line As Line3d

            Dim nNumPosts As Integer = CInt(Math.Floor(((length - 0.6) / handrailPostPitch) + 1))
            Dim newSpacing As Double = length / nNumPosts
            Dim sweepOptions As SweepOptions = SweepOptions.CreateCaps

            If currentAspectID = AspectID.SimplePhysical Then
                crossSectionServices = New CrossSectionServices
            End If

            If withTopLanding Then
                startPosition.Set(upperLeftCorner.X, upperLeftCorner.Y - 0.1, upperLeftCorner.Z)
                endPosition.Set(upperLeftCorner.X, upperLeftCorner.Y - 0.1, upperLeftCorner.Z + postHeight)
                If currentAspectID = AspectID.Centerline Then
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("LPost", line)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    line = New Line3d(startPosition, endPosition)
                    If crossSectionServices IsNot Nothing Then
                        surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                        If Not surfaces Is Nothing Then
                            For j = 0 To surfaces.Count - 1
                                simplePhysicalAspect.Outputs.Add("LPost" & j, surfaces(j))
                            Next j
                        End If
                    End If
                End If
                startPosition.Set(upperLeftCorner.X, upperLeftCorner.Y - topLandingLength + 0.1, upperLeftCorner.Z)
                endPosition.Set(upperLeftCorner.X, upperLeftCorner.Y - topLandingLength + 0.1, upperLeftCorner.Z + postHeight)
                If currentAspectID = AspectID.Centerline Then
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("LPostX", line)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    line = New Line3d(startPosition, endPosition)

                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("LPostX" & j, surfaces(j))
                        Next j
                    End If

                End If

                Select Case handRailSectionCP
                    Case 7, 8, 9, 14  'top edge
                        depthCorrection = handRailDepth
                    Case 4, 5, 6  'along half depth
                    Case 1, 2, 3, 11 'bottom  edge
                        depthCorrection = -(handRailDepth)
                    Case 10, 12, 13  'along centroid in depth direction
                    Case 15  'shear center
                End Select

                startPosition.Set(upperRightCorner.X, upperRightCorner.Y - 0.1, upperRightCorner.Z)
                endPosition.Set(upperRightCorner.X, upperRightCorner.Y - 0.1, upperRightCorner.Z + postHeight)
                If currentAspectID = AspectID.Centerline Then
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("RPost", line)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    line = New Line3d(startPosition, endPosition)
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("RPost" & j, surfaces(j))
                        Next j
                    End If
                End If
                startPosition.Set(upperRightCorner.X, upperRightCorner.Y - topLandingLength + 0.1, upperRightCorner.Z)
                endPosition.Set(upperRightCorner.X, upperRightCorner.Y - topLandingLength + 0.1, upperRightCorner.Z + postHeight)
                If currentAspectID = AspectID.Centerline Then
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("RPostX", line)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    line = New Line3d(startPosition, endPosition)
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("RPostX" & j, surfaces(j))
                        Next j
                    End If
                End If

                upperLeftCorner.Y = upperLeftCorner.Y - topLandingLength
                upperRightCorner.Y = upperRightCorner.Y - topLandingLength
                lowerLeftCorner.Y = (-Math.Tan(angle) * height) - topLandingLength
                lowerRightCorner.Y = (-Math.Tan(angle) * height) - topLandingLength
            End If

            startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)

            startPoint.Z = startPoint.Z - Math.Cos(angle) * 0.3
            startPoint.Y = startPoint.Y - Math.Sin(angle) * 0.3
            endPoint.Set(startPoint.X, startPoint.Y, startPoint.Z)
            endPoint.Z = endPoint.Z + postHeight

            Dim startNormal As New Vector(0.0, (-1 / Math.Sin(angle)), (1 / Math.Cos(angle)))

            'Intermediate Posts
            For i = 1 To nNumPosts
                If i <> 1 Then
                    startPoint.Y = startPoint.Y - (Math.Sin(angle) * newSpacing)
                    startPoint.Z = startPoint.Z - (Math.Cos(angle) * newSpacing)
                    endPoint.Set(startPoint.X, startPoint.Y, startPoint.Z)
                    endPoint.Z = endPoint.Z + postHeight
                End If

                startPosition.Set(startPoint.X, startPoint.Y, startPoint.Z)
                endPosition.Set(endPoint.X, endPoint.Y, endPoint.Z)
                If currentAspectID = AspectID.Centerline Then
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("LPost" & i, line)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    line = New Line3d(startPosition, endPosition)
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, startPoint, startNormal, Nothing, Nothing, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("LPost" & i & j, surfaces(j))
                        Next j
                    End If
                End If
            Next i

            startPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
            startPoint.Z = startPoint.Z - Math.Cos(angle) * 0.3
            startPoint.Y = startPoint.Y - Math.Sin(angle) * 0.3

            'now get depth correction
            Select Case handRailSectionCP
                Case 7, 8, 9, 14  'top edge
                    startPoint.X = startPoint.X + handRailDepth
                Case 4, 5, 6  'along half depth
                Case 1, 2, 3, 11 'bottom  edge
                    startPoint.X = startPoint.X - handRailDepth
                Case 10, 12, 13  'along centroid in depth direction
                Case 15  'shear center
            End Select

            endPoint.Set(startPoint.X, startPoint.Y, startPoint.Z)
            endPoint.Z = endPoint.Z + postHeight

            For i = 1 To nNumPosts
                If i <> 1 Then
                    startPoint.Y = startPoint.Y - (Math.Sin(angle) * newSpacing)
                    startPoint.Z = startPoint.Z - (Math.Cos(angle) * newSpacing)
                    endPoint.Set(startPoint.X, startPoint.Y, startPoint.Z)
                    endPoint.Z = endPoint.Z + postHeight
                End If

                startPosition.Set(startPoint.X, startPoint.Y, startPoint.Z)
                endPosition.Set(endPoint.X, endPoint.Y, endPoint.Z)
                If currentAspectID = AspectID.Centerline Then
                    line = New Line3d(startPosition, endPosition)
                    centerlineAspect.Outputs.Add("RPost" & i, line)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    line = New Line3d(startPosition, endPosition)
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, startPoint, startNormal, Nothing, Nothing, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("RPost" & i & j, surfaces(j))
                        Next j
                    End If
                End If
            Next i

        End Sub
#End Region

#Region "Handrails Creation"

        ''' <summary>
        ''' Places the hand rails with toplanding.
        ''' </summary>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="lowerLeftCorner">The lower left corner.</param>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="lowerRightCorner">The lower right corner.</param>
        ''' <param name="handRailSectionCP">The hand rail section cp.</param>
        ''' <param name="handRailSectionAngle">The hand rail section angle.</param>
        ''' <param name="postHeight">Height of the post.</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="numberOfMidRails">The number of mid rails.</param>
        ''' <param name="handRailDepth">The hand rail depth.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        Private Sub PlaceHandRailsWithToplanding(ByVal crosssection As CrossSection, _
                                                 ByVal upperLeftCorner As Position, ByVal lowerLeftCorner As Position, _
                                                 ByVal upperRightCorner As Position, ByVal lowerRightCorner As Position, _
                                                 ByVal handRailSectionCP As Integer, ByVal handRailSectionAngle As Double, _
                                                 ByVal postHeight As Double, ByVal topLandingLength As Double, _
                                                 ByVal numberOfMidRails As Integer, ByVal handRailDepth As Double, _
                                                 ByVal currentAspectID As AspectID)

            Dim startPosition As New Position, endPosition As New Position
            Dim startPoint As New Position, endPoint As New Position
            Dim surfaces As Collection(Of ISurface)
            Dim complexStringCurves As Collection(Of ICurve) = Nothing
            Dim curves As Collection(Of ICurve)
            Dim depthCorrection As Double
            Dim outCurve As ComplexString3d, complexString As ComplexString3d
            Dim crossSectionServices As CrossSectionServices = Nothing
            Dim line As Line3d

            'creation of left handrail
            Select Case handRailSectionCP
                Case 6, 8, 13, 14
                    depthCorrection = -handRailDepth / 2
                Case 1
                    depthCorrection = handRailDepth
                Case 2, 4, 11, 12
                    depthCorrection = handRailDepth / 2
                Case 3, 5, 7, 10, 15
                Case 9
                    depthCorrection = -handRailDepth
            End Select

            If currentAspectID = AspectID.SimplePhysical Then
                crossSectionServices = New CrossSectionServices()
            End If

            Dim handRailSpacing As Double = postHeight / (numberOfMidRails + 1)
            Dim sweepOptions As SweepOptions = SweepOptions.CreateCaps
            startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            endPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            endPoint.Y = endPoint.Y - topLandingLength

            '1
            startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
            endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
            line = New Line3d(startPosition, endPosition)
            curves = New Collection(Of ICurve)
            curves.Add(line)
            complexString = New ComplexString3d(curves)
            curves = Nothing

            '2
            startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            startPoint.Y = startPoint.Y - topLandingLength
            endPoint.Set(lowerLeftCorner.X, lowerLeftCorner.Y, lowerLeftCorner.Z)
            startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
            endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
            line.DefineBy2Points(startPosition, endPosition)
            complexString.AddCurve(line, True)

            If currentAspectID = AspectID.Centerline Then
                complexString.GetCurves(complexStringCurves)
                outCurve = New ComplexString3d(complexStringCurves)
                centerlineAspect.Outputs.Add("LeftHandRail", outCurve)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                If Not surfaces Is Nothing Then
                    For j = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("LeftHandRail" & j, surfaces(j))
                    Next j
                End If
            End If

            'creation of midrails
            If numberOfMidRails >= 1 Then
                depthCorrection = 0

                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
                endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (numberOfMidRails)))
                line.DefineBy2Points(startPosition, endPosition)
                curves = New Collection(Of ICurve)
                curves.Add(line)
                complexString = New ComplexString3d(curves)
                curves = Nothing

                '4
                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (numberOfMidRails)))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (numberOfMidRails)))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                '5
                startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (numberOfMidRails)))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y + topLandingLength, startPoint.Z + postHeight - (handRailSpacing * (numberOfMidRails)))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                '6
                startPosition.Set(startPoint.X + depthCorrection, startPoint.Y + topLandingLength, startPoint.Z + postHeight - (handRailSpacing * (numberOfMidRails)))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y + topLandingLength, startPoint.Z + postHeight)
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                If currentAspectID = AspectID.Centerline Then
                    complexString.GetCurves(complexStringCurves)
                    outCurve = New ComplexString3d(complexStringCurves)
                    centerlineAspect.Outputs.Add("MidRail", outCurve)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("MidRail" & j, surfaces(j))
                        Next j
                    End If
                End If
            End If

            'creation of left midrails
            depthCorrection = 0
            Select Case handRailSectionCP
                Case 6, 8, 13, 14
                    depthCorrection = -handRailDepth / 2
                Case 1
                    depthCorrection = handRailDepth
                Case 2, 4, 11, 12
                    depthCorrection = handRailDepth / 2
                Case 3, 5, 7, 10, 15
                Case 9
                    depthCorrection = -handRailDepth
            End Select

            If numberOfMidRails >= 2 Then
                For i = 2 To numberOfMidRails

                    startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
                    endPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
                    endPoint.Y = endPoint.Y - topLandingLength

                    startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    line.DefineBy2Points(startPosition, endPosition)
                    curves = New Collection(Of ICurve)
                    curves.Add(line)
                    complexString = New ComplexString3d(curves)

                    startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
                    endPoint.Set(lowerLeftCorner.X, lowerLeftCorner.Y, lowerLeftCorner.Z)
                    startPoint.Y = startPoint.Y - topLandingLength

                    startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))

                    line.DefineBy2Points(startPosition, endPosition)
                    complexString.AddCurve(line, True)

                    If currentAspectID = AspectID.Centerline Then
                        complexString.GetCurves(complexStringCurves)
                        outCurve = New ComplexString3d(complexStringCurves)
                        centerlineAspect.Outputs.Add("LMidRail" & i, outCurve)
                    ElseIf currentAspectID = AspectID.SimplePhysical Then
                        surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                        If Not surfaces Is Nothing Then
                            For j = 0 To surfaces.Count - 1
                                simplePhysicalAspect.Outputs.Add("LMidRail" & i & j, surfaces(j))
                            Next j
                        End If
                    End If
                Next i
            End If

            startPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
            endPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
            endPoint.Y = endPoint.Y - topLandingLength
            depthCorrection = 0

            'creation of toprail
            Select Case handRailSectionCP
                Case 1, 5, 9, 10, 15
                    'Nothing
                Case 2, 6, 11, 13
                    depthCorrection = -handRailDepth / 2
                Case 3
                    depthCorrection = -handRailDepth
                Case 4, 8, 12, 14
                    depthCorrection = handRailDepth / 2
                Case 7
                    depthCorrection = handRailDepth
            End Select

            '1
            startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
            endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
            line.DefineBy2Points(startPosition, endPosition)
            curves = New Collection(Of ICurve)
            curves.Add(line)
            complexString = New ComplexString3d(curves)
            curves = Nothing

            '2
            startPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
            startPoint.Y = startPoint.Y - topLandingLength
            endPoint.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerRightCorner.Z)
            startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
            endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
            line.DefineBy2Points(startPosition, endPosition)
            complexString.AddCurve(line, True)

            If currentAspectID = AspectID.Centerline Then
                complexString.GetCurves(complexStringCurves)
                outCurve = New ComplexString3d(complexStringCurves)
                centerlineAspect.Outputs.Add("TopRail", outCurve)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                If Not surfaces Is Nothing Then
                    For j = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("TopRail" & j, surfaces(j))
                    Next j
                End If
            End If

            'creation of bottomrail
            If numberOfMidRails >= 1 Then
                depthCorrection = 0
                Select Case handRailSectionCP
                    Case 4, 5, 6, 10, 13, 15
                        'Nothing
                    Case 1, 2, 3, 11
                        depthCorrection = -handRailDepth
                    Case 7, 8, 9, 14
                        depthCorrection = handRailDepth
                End Select

                '3
                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
                endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                line.DefineBy2Points(startPosition, endPosition)
                curves = New Collection(Of ICurve)
                curves.Add(line)
                complexString = New ComplexString3d(curves)
                curves = Nothing

                '4
                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                '5
                startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y + topLandingLength, startPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                '6
                startPosition.Set(startPoint.X + depthCorrection, startPoint.Y + topLandingLength, startPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y + topLandingLength, startPoint.Z + postHeight)
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                If currentAspectID = AspectID.Centerline Then
                    complexString.GetCurves(complexStringCurves)
                    outCurve = New ComplexString3d(complexStringCurves)
                    centerlineAspect.Outputs.Add("BottomRail", outCurve)
                ElseIf currentAspectID = AspectID.SimplePhysical Then
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                    If Not surfaces Is Nothing Then
                        For j = 0 To surfaces.Count - 1
                            simplePhysicalAspect.Outputs.Add("BottomRail" & j, surfaces(j))
                        Next j
                    End If
                End If
            End If

            'creation of right midrails
            depthCorrection = 0
            Select Case handRailSectionCP
                Case 1, 5, 9, 10, 15
                Case 2, 6, 11, 13
                    depthCorrection = -handRailDepth / 2
                Case 3
                    depthCorrection = -handRailDepth
                Case 4, 8, 12, 14
                    depthCorrection = handRailDepth / 2
                Case 7
                    depthCorrection = handRailDepth
            End Select

            If numberOfMidRails >= 2 Then
                For i = 2 To numberOfMidRails

                    startPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                    endPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                    endPoint.Y = endPoint.Y - topLandingLength

                    startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    line.DefineBy2Points(startPosition, endPosition)

                    curves = New Collection(Of ICurve)
                    curves.Add(line)
                    complexString = New ComplexString3d(curves)
                    curves = Nothing

                    startPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
                    endPoint.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerRightCorner.Z)
                    startPoint.Y = startPoint.Y - topLandingLength

                    startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    line.DefineBy2Points(startPosition, endPosition)
                    complexString.AddCurve(line, True)

                    If currentAspectID = AspectID.Centerline Then
                        complexString.GetCurves(complexStringCurves)
                        outCurve = New ComplexString3d(complexStringCurves)
                        centerlineAspect.Outputs.Add("RMidRail" & i, outCurve)
                    ElseIf currentAspectID = AspectID.SimplePhysical Then
                        surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                        If Not surfaces Is Nothing Then
                            For j = 0 To surfaces.Count - 1
                                simplePhysicalAspect.Outputs.Add("RMidRail" & i & j, surfaces(j))
                            Next j
                        End If
                    End If
                Next i
            End If

        End Sub

        ''' <summary>
        ''' Places the handrails.
        ''' </summary>
        ''' <param name="crosssection">The crosssection.</param>
        ''' <param name="upperLeftCorner">The upper left corner.</param>
        ''' <param name="lowerLeftCorner">The lower left corner.</param>
        ''' <param name="upperRightCorner">The upper right corner.</param>
        ''' <param name="lowerRightCorner">The lower right corner.</param>
        ''' <param name="handRailSectionCP">The hand rail section cp.</param>
        ''' <param name="handRailSectionAngle">The hand rail section angle.</param>
        ''' <param name="postHeight">Height of the post.</param>
        ''' <param name="topLandingLength">Length of the top landing.</param>
        ''' <param name="numberOfMidRails">The number of mid rails.</param>
        ''' <param name="handRailDepth">The hand rail depth.</param>
        ''' <param name="currentAspectID">The current aspect identifier.</param>
        Private Sub PlaceHandrails(ByVal crosssection As CrossSection, _
                                   ByVal upperLeftCorner As Position, ByVal lowerLeftCorner As Position, _
                                   ByVal upperRightCorner As Position, ByVal lowerRightCorner As Position, _
                                   ByVal handRailSectionCP As Integer, ByVal handRailSectionAngle As Double, _
                                   ByVal postHeight As Double, ByVal topLandingLength As Double, _
                                   ByVal numberOfMidRails As Integer, ByVal handRailDepth As Double, _
                                   ByVal currentAspectID As AspectID)

            Dim startPosition As New Position, endPosition As New Position
            Dim startPoint As New Position, endPoint As New Position
            Dim surfaces As Collection(Of ISurface)
            Dim complexStringCurves As Collection(Of ICurve) = Nothing
            Dim curves As New Collection(Of ICurve)
            Dim depthCorrection As Double
            Dim outCurve As ComplexString3d, complexString As ComplexString3d
            Dim crossSectionServices As CrossSectionServices = Nothing
            Dim line As Line3d
            Dim handRailSpacing As Double = postHeight / (numberOfMidRails + 1)

            'creation if left handrails
            startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            endPoint.Set(lowerLeftCorner.X, lowerLeftCorner.Y, lowerLeftCorner.Z)

            Dim sweepOptions As SweepOptions = SweepOptions.CreateCaps

            Select Case handRailSectionCP
                Case 6, 8, 13, 14
                    depthCorrection = -handRailDepth / 2
                Case 1
                    depthCorrection = handRailDepth
                Case 2, 4, 11, 12
                    depthCorrection = handRailDepth / 2
                Case 3, 5, 7, 10, 15
                Case 9
                    depthCorrection = -handRailDepth
            End Select

            If currentAspectID = AspectID.SimplePhysical Then
                crossSectionServices = New CrossSectionServices()
            End If

            '1
            startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
            endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)

            line = New Line3d(startPosition, endPosition)
            curves.Add(line)
            complexString = New ComplexString3d(curves)
            curves = Nothing

            '2
            If numberOfMidRails >= 1 Then
                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
                endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)
                '3
                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (numberOfMidRails)))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)
                '4
                startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)
            End If

            If currentAspectID = AspectID.Centerline Then
                complexString.GetCurves(complexStringCurves)
                outCurve = New ComplexString3d(complexStringCurves)
                centerlineAspect.Outputs.Add("LHandRail1", outCurve)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                If Not surfaces Is Nothing Then
                    For i = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("LHandRail1" & i, surfaces(i))
                    Next i
                End If
            End If

            'creation of left midrails
            startPoint.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            endPoint.Set(lowerLeftCorner.X, lowerLeftCorner.Y, lowerLeftCorner.Z)

            If numberOfMidRails >= 2 Then
                For i = 2 To numberOfMidRails
                    startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    line.DefineBy2Points(startPosition, endPosition)
                    If currentAspectID = AspectID.Centerline Then
                        startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                        endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                        line = New Line3d(startPosition, endPosition)
                        centerlineAspect.Outputs.Add("LHandRail1" & i, line)
                    ElseIf currentAspectID = AspectID.SimplePhysical Then
                        surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                        If Not surfaces Is Nothing Then
                            For j = 0 To surfaces.Count - 1
                                simplePhysicalAspect.Outputs.Add("LHandRail1" & i & j, surfaces(j))
                            Next j
                        End If
                    End If
                Next i
            End If

            'creation of right handrails
            startPoint.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)
            endPoint.Set(lowerRightCorner.X, lowerRightCorner.Y, lowerRightCorner.Z)
            depthCorrection = 0
            Select Case handRailSectionCP
                Case 4, 8, 12, 14
                    depthCorrection = handRailDepth / 2
                Case 1, 5, 9, 10, 15
                Case 2, 6, 11, 13
                    depthCorrection = -handRailDepth / 2
                Case 3
                    depthCorrection = -handRailDepth
                Case 7
                    depthCorrection = handRailDepth
            End Select

            startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
            endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
            line = New Line3d(startPosition, endPosition)
            curves = New Collection(Of ICurve)
            curves.Add(line)
            complexString = New ComplexString3d(curves)
            curves = Nothing

            If numberOfMidRails >= 1 Then
                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight)
                endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                startPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)

                startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * numberOfMidRails))
                endPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight)
                line.DefineBy2Points(startPosition, endPosition)
                complexString.AddCurve(line, True)
            End If

            If currentAspectID = AspectID.Centerline Then
                complexString.GetCurves(complexStringCurves)
                outCurve = New ComplexString3d(complexStringCurves)
                centerlineAspect.Outputs.Add("RHandRail1", outCurve)
            ElseIf currentAspectID = AspectID.SimplePhysical Then
                surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, complexString, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                If Not surfaces Is Nothing Then
                    For i = 0 To surfaces.Count - 1
                        simplePhysicalAspect.Outputs.Add("RHandRail1" & i, surfaces(i))
                    Next i
                End If
            End If

            'creation of right midrails
            If numberOfMidRails >= 2 Then
                For i = 2 To numberOfMidRails
                    startPosition.Set(startPoint.X + depthCorrection, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    endPosition.Set(endPoint.X + depthCorrection, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                    line.DefineBy2Points(startPosition, endPosition)
                    If currentAspectID = AspectID.Centerline Then
                        startPosition.Set(startPoint.X, startPoint.Y, startPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                        endPosition.Set(endPoint.X, endPoint.Y, endPoint.Z + postHeight - (handRailSpacing * (i - 1)))
                        line = New Line3d(startPosition, endPosition)
                        centerlineAspect.Outputs.Add("RHandRail1" & i, line)
                    ElseIf currentAspectID = AspectID.SimplePhysical Then
                        surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(crosssection, line, handRailSectionCP, False, handRailSectionAngle, sweepOptions)
                        If Not surfaces Is Nothing Then
                            For j = 0 To surfaces.Count - 1
                                simplePhysicalAspect.Outputs.Add("RHandRail1" & i & j, surfaces(j))
                            Next j
                        End If
                    End If
                Next i
            End If

        End Sub
#End Region

#Region "Operational Envelope Creation"
        ''' <summary>
        ''' Places the envelope.
        ''' </summary>
        ''' <param name="withTopLanding">if set to <c>true</c> [with top landing].</param>
        ''' <param name="totalWidth">The total width.</param>
        ''' <param name="position">The position.</param>
        Private Sub PlaceEnvelope(ByVal withTopLanding As Boolean, _
                                  ByVal totalWidth As Double, ByVal position As Double)

            Dim xAxis As New Vector, yAxis As New Vector, zAxis As New Vector
            Dim position1 As New Position, position2 As New Position
            Dim positions As New Collection(Of Position)
            Dim matrixArray(15) As Double

            Dim envelopHeight As Double = envelopeHeightInput.Value
            Dim topLandingLength As Double = topLandingLengthInput.Value
            Dim angle As Double = angleInput.Value
            Dim length As Double = lengthInput.Value

            'We start at the top of the stair and go down along the stair. 
            'Origin at the top right- both w and w/o landing
            '                       ______
            '                   -         |
            '               -       ______|
            '           -       -
            '       -       -
            '   |       -
            '   |   -

            position1.Set(0.0, 0.0, 0.0)
            position2.Set(-length * Math.Cos(angle), -length * Math.Sin(angle), 0.0)
            positions.Add(position1)
            positions.Add(position2)
            Dim lineStrings3d As New LineString3d(positions)

            position1.Set(-length * Math.Cos(angle), -length * Math.Sin(angle) + envelopHeight, 0.0)
            lineStrings3d.AddPoint(position1)

            position1.Set(0.0, envelopHeight, 0.0)
            lineStrings3d.AddPoint(position1)

            If withTopLanding Then
                position1.Set(topLandingLength, envelopHeight, 0.0)
                lineStrings3d.AddPoint(position1)
                position1.Set(topLandingLength, 0.0, 0.0)
                lineStrings3d.AddPoint(position1)
            End If

            position1.Set(0.0, 0.0, 0.0)
            lineStrings3d.AddPoint(position1) 'Close out the line string

            'Now project it along z
            Dim projectionDirection As New Vector(0.0, 0.0, 1.0)
            Dim projection As New Projection3d(lineStrings3d, projectionDirection, totalWidth, True)

            ' x-axis is the y-axis on the projection
            xAxis.X = 0.0
            xAxis.Y = 1.0
            xAxis.Z = 0.0
            xAxis.Length = 1.0

            'Global Y axis is the z axis on the projection
            yAxis.X = 0.0
            yAxis.Y = 0.0
            yAxis.Z = 1.0
            yAxis.Length = 1.0

            'Global z axis is the x axis on the projection
            zAxis.X = 1.0
            zAxis.Y = 0.0
            zAxis.Z = 0.0
            zAxis.Length = 1.0

            'Now apply the xform to correctly orient the envelope
            matrixArray(0) = xAxis.X
            matrixArray(1) = xAxis.Y
            matrixArray(2) = xAxis.Z

            matrixArray(4) = yAxis.X
            matrixArray(5) = yAxis.Y
            matrixArray(6) = yAxis.Z

            matrixArray(8) = zAxis.X
            matrixArray(9) = zAxis.Y
            matrixArray(10) = zAxis.Z

            matrixArray(12) = -totalWidth / 2.0 - position
            If withTopLanding Then
                matrixArray(13) = 0.0 - topLandingLength
            Else
                matrixArray(13) = 0.0
            End If
            matrixArray(14) = 0.0
            matrixArray(15) = 1.0

            Dim matrix As New Matrix4X4(matrixArray, False)
            projection.Transform(matrix)

            operationAspect.Outputs.Add("OperationalEnvelope1", projection)

        End Sub
#End Region

#Region "Different Type of Projection Creation"

        ''' <summary>
        ''' Creates the rectangular projection.
        ''' </summary>
        ''' <param name="center">The center.</param>
        ''' <param name="projectionVector">The projection vector.</param>
        ''' <param name="upperVector">The upper vector.</param>
        ''' <param name="rectangleWidth">Width of the rectangle.</param>
        ''' <param name="rectangleHeight">Height of the rectangle.</param>
        ''' <returns></returns>
        Private Function CreateRectangularProjection(ByVal center As Position, _
                                                     ByVal projectionVector As Vector, ByVal upperVector As Vector, _
                                                     ByVal rectangleWidth As Double, ByVal rectangleHeight As Double) As Projection3d

            Dim tempVector As Vector
            Dim tempProjectionVector As Vector
            Dim projection As Projection3d
            Dim matrix As Matrix4X4
            Dim matrixArray(15) As Double

            tempProjectionVector = New Vector(projectionVector.X, projectionVector.Y, projectionVector.Z)
            tempProjectionVector.Length = 1.0

            projection = CreateUnitRectangulaProjection(rectangleWidth, rectangleHeight, projectionVector.Length)
            upperVector.Length = 1.0
            tempVector = tempProjectionVector.Cross(upperVector)
            tempVector.Length = 1.0

            matrixArray(0) = tempVector.X
            matrixArray(1) = tempVector.Y
            matrixArray(2) = tempVector.Z

            matrixArray(4) = upperVector.X
            matrixArray(5) = upperVector.Y
            matrixArray(6) = upperVector.Z

            matrixArray(8) = tempProjectionVector.X
            matrixArray(9) = tempProjectionVector.Y
            matrixArray(10) = tempProjectionVector.Z

            matrixArray(12) = center.X
            matrixArray(13) = center.Y
            matrixArray(14) = center.Z
            matrixArray(15) = 1.0

            matrix = New Matrix4X4(matrixArray, True)
            projection.Transform(matrix)

            CreateRectangularProjection = projection

        End Function

        ''' <summary>
        ''' Creates the unit rectangular projection in the positive z-direction.
        ''' </summary>
        ''' <param name="rectangleWidth">Width of the rectangle.</param>
        ''' <param name="rectangleHeight">Height of the rectangle.</param>
        ''' <param name="projectionLength">Length of the projection.</param>
        ''' <returns></returns>
        Private Function CreateUnitRectangulaProjection(ByVal rectangleWidth As Double, _
                                            ByVal rectangleHeight As Double, ByVal projectionLength As Double) As Projection3d
            Dim projectionDirection As New Vector(0.0, 0.0, 1.0)

            Dim rectangle As LineString3d = SymbolHelper.CreateRectangle(rectangleWidth, rectangleHeight, 0)
            Dim projection As Projection3d = New Projection3d(DirectCast(rectangle, ICurve), projectionDirection, projectionLength, True)

            CreateUnitRectangulaProjection = projection

        End Function

#End Region

#Region "CalculateVolumeCOG"

        ''' <summary>
        ''' Calculates the volume COG.
        ''' </summary>
        ''' <param name="businessObject">The business object.</param>
        ''' <param name="accumulatedVolume">The accumulated volume.</param>
        ''' <param name="cogX">The cog x.</param>
        ''' <param name="cogY">The cog y.</param>
        ''' <param name="cogZ">The cog z.</param>
        Private Sub CalculateVolumeCOG(ByVal businessObject As BusinessObject, ByRef accumulatedVolume As Double, _
                                       ByRef cogX As Double, ByRef cogY As Double, ByRef cogZ As Double)

            'Getting required values after constructing symbol for volume and COG evaluate
            Dim length As Double, angle As Double
            Dim height As Double
            Dim width As Double
            Dim stepPitch As Double
            Dim handrailPostPitch As Double
            Dim postHeight As Double, topLandingLength As Double
            Dim platformThickness As Double
            Dim sideFrameSectionName As String = ""
            Dim sideFrameSectionReferenceStandard As String = ""
            Dim stepSectionName As String = ""
            Dim stepSectionReferenceStandard As String = ""
            Dim handRailSectionName As String = ""
            Dim handRailSectionReferenceStandard As String = ""
            Dim justificationOption As Integer
            Dim numberOfMidRails As Integer
            Dim numberOfSteps As Integer
            Dim withTopLanding As Boolean
            Try
                length = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJSPSCommonStairLadderProps, SPSSymbolConstants.Length)
                angle = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJSPSCommonStairLadderProps, SPSSymbolConstants.Angle)
                height = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJSPSCommonStairLadderProps, SPSSymbolConstants.Height)
                width = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJSPSCommonStairLadderProps, SPSSymbolConstants.Width)
                stepPitch = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJSPSCommonStairLadderProps, SPSSymbolConstants.StepPitch)
                handrailPostPitch = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.HandRailPostPitch)
                postHeight = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.PostHeight)
                topLandingLength = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.TopLandingLength)
                platformThickness = StructHelper.GetDoubleProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.PlatformThickness)
                sideFrameSectionName = StructHelper.GetStringProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.SideFrameSectionName)
                sideFrameSectionReferenceStandard = StructHelper.GetStringProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.SideFrameSectionRefStandard)
                stepSectionName = StructHelper.GetStringProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.StepSectionName)
                stepSectionReferenceStandard = StructHelper.GetStringProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.StepSectionRefStandard)
                handRailSectionName = StructHelper.GetStringProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.HandRailSectionName)
                handRailSectionReferenceStandard = StructHelper.GetStringProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.HandRailSectionRefStandard)
                justificationOption = StructHelper.GetIntProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.Justification)
                numberOfMidRails = StructHelper.GetIntProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.NumberOfMidRails)
                numberOfSteps = StructHelper.GetIntProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.NumSteps)
                withTopLanding = StructHelper.GetBoolProperty(businessObject, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.WithTopLanding)
            Catch ex As Exception
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, businessObject.ToString + StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrWCOGMissingSystemAttributeData, _
                                                            "Weight and COG failed to evaluate in CalculateVolumeCOG in StairTypeA Symbol, as some of the required system attribute values cannot be obtained. Check the error log and catalog data."))
                Return 'stop evaluating
            End Try

            Dim catalogHelper As CatalogStructHelper = New CatalogStructHelper()
            'Getting the cross-section for side frame, step and handrail
            Dim sideFrameCrossSection As CrossSection = catalogHelper.GetCrossSection(sideFrameSectionReferenceStandard, sideFrameSectionName)
            'might be created a todo record if section is missing
            If MyBase.ToDoListMessage IsNot Nothing Then
                If MyBase.ToDoListMessage.Type = ToDoMessageTypes.ToDoMessageError Then
                    Return
                End If
            End If
            Dim stepCrossSection As CrossSection = catalogHelper.GetCrossSection(stepSectionReferenceStandard, stepSectionName)
            'might be created a todo record if section is missing
            If MyBase.ToDoListMessage IsNot Nothing Then
                If MyBase.ToDoListMessage.Type = ToDoMessageTypes.ToDoMessageError Then
                    Return
                End If
            End If
            Dim handrailCrossSection As CrossSection = catalogHelper.GetCrossSection(handRailSectionReferenceStandard, handRailSectionName)
            'might be created a todo record if section is missing
            If MyBase.ToDoListMessage IsNot Nothing Then
                If MyBase.ToDoListMessage.Type = ToDoMessageTypes.ToDoMessageError Then
                    Return
                End If
            End If

            Dim sideFrameWidth As Double, sideFrameDepth As Double
            Dim stepDepth As Double, stepWidth As Double


            sideFrameWidth = sideFrameCrossSection.Depth 'CR184504
            sideFrameDepth = sideFrameCrossSection.Width

            stepWidth = stepCrossSection.Depth 'CR184504
            stepDepth = stepCrossSection.Width

            'only support step section CP5 as default CP
            stepDepth = stepDepth / 2

            'Right now, only default CP (CP7) is supported for SideFrameSectionCP. 
            Dim sideFrameCSArea As Double
            Dim stepCSArea As Double
            Dim handRailCSArea As Double
            Try
                sideFrameCSArea = StructHelper.GetDoubleProperty(sideFrameCrossSection, SPSSymbolConstants.IStructCrossSectionDimensions, SPSSymbolConstants.Area)
                stepCSArea = StructHelper.GetDoubleProperty(stepCrossSection, SPSSymbolConstants.IStructCrossSectionDimensions, SPSSymbolConstants.Area)
                handRailCSArea = StructHelper.GetDoubleProperty(handrailCrossSection, SPSSymbolConstants.IStructCrossSectionDimensions, SPSSymbolConstants.Area)
            Catch ex As Exception
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, businessObject.ToString + StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrWCOGMissingSystemAttributeData, _
                                                            "Weight and COG failed to evaluate in CalculateVolumeCOG in StairTypeA Symbol, as some of the required system attribute values can not obtained. Check the error log and catalog data."))
                Return 'stop evaluating
            End Try

            Dim numberOfPosts As Integer = CInt(Math.Floor(((length - 0.6) / handrailPostPitch) + 1))

            Dim handRailSpacing As Double = postHeight / (numberOfMidRails + 1)

            'Calculation of COG
            Dim projectionVolume As Double, position As Double
            Dim accumulatedCOGPosition As New Position, cogPosition As New Position
            Dim upperLeftCorner As New Position, upperRightCorner As New Position
            Dim lowerLeftCorner As New Position, lowerRightCorner As New Position

            angle = (Math.PI / 2 - angle)

            Dim sideWidthY As Double = sideFrameWidth / Math.Cos(angle)
            Dim sideWidthZ As Double = sideFrameWidth / Math.Sin(angle)

            If justificationOption = SPSSymbolConstants.ALIGNMENT_LEFT Then
                position = -(width / 2)
            ElseIf justificationOption = SPSSymbolConstants.ALIGNMENT_RIGHT Then
                position = (width / 2)
            End If

            upperLeftCorner.X = -width / 2
            upperLeftCorner.Y = 0.0
            upperLeftCorner.Z = 0.0

            upperRightCorner.X = width / 2
            upperRightCorner.Y = 0.0
            upperRightCorner.Z = 0.0

            lowerLeftCorner.X = -width / 2
            lowerLeftCorner.Y = -Math.Tan(angle) * height
            lowerLeftCorner.Z = -height

            lowerRightCorner.X = width / 2
            lowerRightCorner.Y = -Math.Tan(angle) * height
            lowerRightCorner.Z = -height

            'End Treatment (Left + Right)
            If numberOfMidRails >= 1 Then
                cogPosition.X = 0
                cogPosition.Y = 0
                cogPosition.Z = upperLeftCorner.Z + (postHeight + handRailSpacing) / 2

                projectionVolume = (postHeight - handRailSpacing) * handRailCSArea * 2
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If
            End If

            If withTopLanding Then
                'PlatForm
                cogPosition.X = 0
                cogPosition.Y = -topLandingLength / 2
                cogPosition.Z = upperLeftCorner.Z - platformThickness / 2

                projectionVolume = topLandingLength * width * platformThickness
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If

                'Middle Top landing
                If topLandingLength > 1.5 Then
                    cogPosition.X = 0
                    cogPosition.Y = -topLandingLength / 2
                    cogPosition.Z = upperLeftCorner.Z - sideFrameWidth / 2 'CP7

                    projectionVolume = Math.Sqrt(Math.Pow(topLandingLength, 2) + Math.Pow(width, 2)) * sideFrameCSArea
                    accumulatedVolume = accumulatedVolume + projectionVolume

                    If accumulatedVolume > 0 Then
                        accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                        accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                        accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                    End If
                End If

                'Side Top landing (Left + Right)
                cogPosition.X = 0
                cogPosition.Y = -topLandingLength / 2
                cogPosition.Z = upperLeftCorner.Z - sideFrameWidth / 2 'CP7

                projectionVolume = topLandingLength * sideFrameCSArea * 2
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If

                'Front and Back Landing
                cogPosition.X = 0
                cogPosition.Y = -topLandingLength / 2
                cogPosition.Z = -sideFrameWidth / 2  'CP7

                projectionVolume = width * sideFrameCSArea * 2
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If

                '4 Posts
                cogPosition.X = 0
                cogPosition.Y = -topLandingLength / 2
                cogPosition.Z = postHeight / 2
                '
                projectionVolume = postHeight * handRailCSArea * 4
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If

                'Top Handrails on Toplanding (Left + Right)
                cogPosition.X = 0
                cogPosition.Y = -topLandingLength / 2
                cogPosition.Z = upperLeftCorner.Z + postHeight

                projectionVolume = topLandingLength * handRailCSArea * 2
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If

                'Bottom Handrail on Toplanding (Left + Right)
                If numberOfMidRails >= 1 Then
                    cogPosition.X = 0
                    cogPosition.Y = -topLandingLength / 2
                    cogPosition.Z = upperLeftCorner.Z + postHeight - (handRailSpacing * (numberOfMidRails))

                    projectionVolume = topLandingLength * handRailCSArea * 2
                    accumulatedVolume = accumulatedVolume + projectionVolume

                    If accumulatedVolume > 0 Then
                        accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                        accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                        accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                    End If
                End If

                'MiddleRails on Toplanding (Left + Right)
                If numberOfMidRails >= 2 Then
                    For i = 2 To numberOfMidRails
                        cogPosition.X = 0
                        cogPosition.Y = -topLandingLength / 2
                        cogPosition.Z = upperLeftCorner.Z + postHeight - (handRailSpacing * (i - 1))

                        projectionVolume = topLandingLength * handRailCSArea * 2
                        accumulatedVolume = accumulatedVolume + projectionVolume

                        If accumulatedVolume > 0 Then
                            accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                            accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                            accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                        End If
                    Next i
                End If

                lowerLeftCorner.Y = lowerLeftCorner.Y - topLandingLength
                lowerRightCorner.Y = lowerRightCorner.Y - topLandingLength
                upperLeftCorner.Y = upperLeftCorner.Y - topLandingLength
                upperRightCorner.Y = upperRightCorner.Y - topLandingLength
            End If

            'COG of Left + Right stringer
            cogPosition.X = 0
            cogPosition.Y = (lowerLeftCorner.Y + upperLeftCorner.Y) / 2 + sideWidthY / 2 'CP7
            cogPosition.Z = (lowerLeftCorner.Z + upperLeftCorner.Z) / 2 - sideWidthZ / 2 'CP7

            projectionVolume = length * sideFrameCSArea * 2
            accumulatedVolume = accumulatedVolume + projectionVolume

            If accumulatedVolume > 0 Then
                accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
            End If

            'COG of Steps
            Dim startPosition As New Position
            Dim endPosition As New Position
            startPosition.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            endPosition.Set(upperRightCorner.X, upperRightCorner.Y, upperRightCorner.Z)

            For i = 1 To numberOfSteps

                cogPosition.X = 0
                cogPosition.Y = upperRightCorner.Y - (Math.Tan(angle)) * (((stepPitch * i) - sideWidthY / 2 + stepDepth))
                cogPosition.Z = upperRightCorner.Z - (i * stepPitch + stepDepth)

                projectionVolume = width * stepCSArea
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If

            Next i

            'COG for posts
            startPosition.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            startPosition.Y = startPosition.Y - Math.Sin(angle) * 0.3
            startPosition.Z = startPosition.Z - Math.Cos(angle) * 0.3

            Dim newSpacing As Double = length / numberOfPosts

            For i = 1 To numberOfPosts
                If i <> 1 Then
                    startPosition.Y = startPosition.Y - (Math.Sin(angle) * newSpacing)
                    startPosition.Z = startPosition.Z - (Math.Cos(angle) * newSpacing)
                End If

                cogPosition.X = 0
                cogPosition.Y = startPosition.Y
                cogPosition.Z = startPosition.Z + postHeight / 2

                projectionVolume = postHeight * handRailCSArea * 2
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If

            Next i

            'COG of Handrails
            startPosition.Set(upperLeftCorner.X, upperLeftCorner.Y, upperLeftCorner.Z)
            endPosition.Set(lowerLeftCorner.X, lowerLeftCorner.Y, lowerLeftCorner.Z)

            'Top rail (Left + Right)
            cogPosition.X = 0
            cogPosition.Y = (endPosition.Y + startPosition.Y) / 2
            cogPosition.Z = (endPosition.Z + startPosition.Z) / 2 + postHeight

            projectionVolume = length * handRailCSArea * 2
            accumulatedVolume = accumulatedVolume + projectionVolume

            If accumulatedVolume > 0 Then
                accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
            End If

            'Bottom rails (Left + Right)
            If numberOfMidRails >= 1 Then
                cogPosition.X = 0
                cogPosition.Y = (endPosition.Y + startPosition.Y) / 2
                cogPosition.Z = (endPosition.Z + startPosition.Z) / 2 + postHeight - handRailSpacing * numberOfMidRails

                projectionVolume = length * handRailCSArea * 2
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If
            End If

            'MiddleRails (Left + Right)
            If numberOfMidRails >= 2 Then
                For i = 2 To numberOfMidRails
                    cogPosition.X = 0
                    cogPosition.Y = (endPosition.Y + startPosition.Y) / 2
                    cogPosition.Z = (endPosition.Z + startPosition.Z) / 2 + postHeight - (handRailSpacing * (i - 1))

                    projectionVolume = length * handRailCSArea * 2
                    accumulatedVolume = accumulatedVolume + projectionVolume

                    If accumulatedVolume > 0 Then
                        accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                        accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                        accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                    End If
                Next i
            End If

            'End treatment (Left + Right)
            If numberOfMidRails >= 1 Then
                cogPosition.X = 0
                cogPosition.Y = lowerLeftCorner.Y
                cogPosition.Z = lowerLeftCorner.Z + (postHeight + handRailSpacing) / 2

                projectionVolume = (postHeight - handRailSpacing) * handRailCSArea * 2
                accumulatedVolume = accumulatedVolume + projectionVolume

                If accumulatedVolume > 0 Then
                    accumulatedCOGPosition.X = accumulatedCOGPosition.X + (cogPosition.X - accumulatedCOGPosition.X) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Y = accumulatedCOGPosition.Y + (cogPosition.Y - accumulatedCOGPosition.Y) * projectionVolume / accumulatedVolume
                    accumulatedCOGPosition.Z = accumulatedCOGPosition.Z + (cogPosition.Z - accumulatedCOGPosition.Z) * projectionVolume / accumulatedVolume
                End If
            End If

            cogX = accumulatedCOGPosition.X - position
            cogY = accumulatedCOGPosition.Y
            cogZ = accumulatedCOGPosition.Z

        End Sub

#End Region

#End Region

#Region "ICustomWeightCG Members"

        ''' <summary>
        ''' Evaluates the weight and center of gravity of the stair part and sets it on the stair business object.
        ''' </summary>
        ''' <param name="businessObject">Stair business object which aggregates symbol.</param>
        Public Sub EvaluateWeightCG(ByVal businessObject As BusinessObject) Implements ICustomWeightCG.EvaluateWeightCG
            'If LeftSideFrame0 output is not there, means symbol is not computed yet, better to skip weight and COG calculation now,
            'we will be called again after the symbol computed.
            Dim aspectName$ = "SimplePhysical"
            Dim outputName$ = "LeftSideFrame0"

            If MyBase.DoesOutputExist(businessObject, aspectName, outputName) Then
                'Getting Weight and COG origin
                Dim weightCOGOrigin As Integer
                Try
                    weightCOGOrigin = StructHelper.GetIntProperty(businessObject, SPSSymbolConstants.IJWCGValueOrigin, SPSSymbolConstants.DryWCGOrigin)
                Catch ex As Exception
                    'attributes might be missing, create todo record
                    MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrWeightCGAttributeData, _
                                "Cannot calculate weight and center of gravity as required system attribute dry weight and center of gravity origin value cannot be obtained. Check the error log and catalog data."))
                    Return 'stop evaluating
                End Try

                'we need to calculate weight and COG as DryWCGOrigin is Computed 
                If weightCOGOrigin = SPSSymbolConstants.DRY_WCOG_ORIGIN_COMPUTED Then

                    'Getting volume and COG values
                    Dim volume As Double = 0.0
                    Dim cogX As Double = 0.0, cogY As Double = 0.0, cogZ As Double = 0.0
                    CalculateVolumeCOG(businessObject, volume, cogX, cogY, cogZ)
                    'Getting weight from volume
                    Dim weight#
                    Try
                        weight = SymbolHelper.EvaluateWeightFromVolume(businessObject, volume, SPSSymbolConstants.IJUAStairTypeA)

                    Catch ex As RefDataMaterialNotFoundException
                        'attributes might be missing, create todo record
                        MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrMaterialNotFound, _
                                                "Error while calculating weight and center of gravity in StairTypeA Symbol while evaluating Weight and CG, as the required material is not found in catalog. Check the error log and catalog data."))
                        Return 'stop evaluating
                    Catch ex As Exception
                        'attributes might be missing, create todo record
                        MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrMaterialAttributeData, _
                                                "Error while Calculating weight and center of gravity in StairTypeA Symbol, as required user attribute material name and grade value cannot be obtained from the catalog. Check the error log and catalog data."))
                        Return 'stop evaluating
                    End Try

                    'Set the net weight and COG on the stair business object using helper method provided in WeightCOGServices
                    Dim weightCOGServices As New WeightCOGServices()
                    Try
                        weightCOGServices.SetWeightAndCOG(businessObject, weight, cogX, cogY, cogZ)

                    Catch ex As Exception
                        MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrSetWeightAndCOG, _
                                                "Cannot set weight and center of gravity on the Stair in WeightCOGServices in StairTypeA Symbol. Please check custom code or contact S3D support."))
                    End Try
                End If
            End If

        End Sub

#End Region

#Region "Overrides Functions And Methods"

        ''' <summary>
        ''' Sets the step properties.
        ''' </summary>
        ''' <param name="stair">The stair.</param>
        ''' <param name="numberOfSteps">The number of steps.</param>
        ''' <param name="actualStepPitch">The actual step pitch.</param>
        Public Overrides Sub SetStepProperties(ByVal stair As Stair, ByVal numberOfSteps As Integer, ByVal actualStepPitch As Double)
            stair.SetPropertyValue(numberOfSteps, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.NumSteps)
            stair.SetPropertyValue(actualStepPitch, SPSSymbolConstants.IUAStairPitchParams, SPSSymbolConstants.ActualPitch)
        End Sub

        ''' <summary>
        ''' Makes the property read-only. 
        ''' </summary>
        ''' <param name="interfaceName">Interface name of the property.</param>
        ''' <param name="propertyName">The name of the property.</param>
        Public Overrides Function IsPropertyReadOnly(ByVal interfaceName$, ByVal propertyName$) As Boolean

            Select Case propertyName
                'making following property values as read-only. 
                Case SPSSymbolConstants.SideFrameSectionRefStandard, SPSSymbolConstants.HandRailSectionRefStandard, _
                SPSSymbolConstants.StepSectionRefStandard, SPSSymbolConstants.NumSteps, SPSSymbolConstants.Span, _
                SPSSymbolConstants.Height, SPSSymbolConstants.Length
                    Return True
            End Select

        End Function

        ''' <summary>
        ''' Stair is with top landing or not. 
        ''' </summary>
        ''' <param name="stair">Stair business object which aggregates symbol.</param>
        ''' <returns>True if stair is with top landing otherwise false.</returns>
        Public Overrides Function WithTopLanding(ByVal stair As Stair) As Boolean
            Return StructHelper.GetBoolProperty(stair, SPSSymbolConstants.IJUAStairTypeA, SPSSymbolConstants.WithTopLanding)
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
                    'following property values need to be in between 25-65 degrees 
                    Case SPSSymbolConstants.Angle
                        isValidPropertyValue = ValidationHelper.IsBetween65And25(CDbl(propertyValue), errorMessage)
                        Exit Select
                        'following property value has combo-box to select the proper option, so set the property value as valid.
                    Case SPSSymbolConstants.SideFrameSectionName, SPSSymbolConstants.HandRailSectionName, _
                    SPSSymbolConstants.StepSectionName, SPSSymbolConstants.PrimaryMaterial
                        isValidPropertyValue = True
                        Exit Select
                        'following property values need to be in between 0-360 degrees 
                    Case SPSSymbolConstants.SideFrameSectionAngle, SPSSymbolConstants.HandrailSectionAngle, SPSSymbolConstants.StepSectionAngle
                        isValidPropertyValue = ValidationHelper.IsBetween0And360(CDbl(propertyValue), errorMessage)
                        Exit Select
                        'following property values must be greater than 0 and lower than 100
                    Case SPSSymbolConstants.Width, SPSSymbolConstants.StepPitch
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(CDbl(propertyValue), errorMessage)
                        If isValidPropertyValue Then
                            isValidPropertyValue = ValidationHelper.IsLowerThan100(CDbl(propertyValue), errorMessage)
                        End If
                        Exit Select
                        'following property values must be lower than 100
                    Case SPSSymbolConstants.Span
                        isValidPropertyValue = ValidationHelper.IsLowerThan100(CDbl(propertyValue), errorMessage)
                        Exit Select
                        'following property values must be greater than 0
                    Case SPSSymbolConstants.NumberOfMidRails
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(CInt(propertyValue), errorMessage)
                        Exit Select
                        'following property values must be greater than 0
                    Case SPSSymbolConstants.PlatformThickness, SPSSymbolConstants.EnvelopeHeight, SPSSymbolConstants.Height, SPSSymbolConstants.Length, _
                    SPSSymbolConstants.TopLandingLength, SPSSymbolConstants.PostHeight, SPSSymbolConstants.HandRailPostPitch
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(CDbl(propertyValue), errorMessage)
                        Exit Select
                End Select
            End If

            Return isValidPropertyValue

        End Function

        ''' <summary>
        ''' Checks for undefined value and raise error.
        ''' </summary>
        ''' <param name="stair">Stair business object which aggregates symbol.</param>
        Public Overrides Sub ValidateUndefinedCodelistValue(ByVal stair As Stair)

            'validate justification
            ValidateUndefinedCodelistValue(stair, SPSSymbolConstants.Justification, SPSSymbolConstants.StructAlignment, SPSSymbolConstants.REFDAT, SPSSymbolConstants.TDL_INVALID_JUSTIFICATION)

            'validate side frame cross-section cardinal point
            ValidateUndefinedCodelistValue(stair, SPSSymbolConstants.SideFrameSectionCP, SPSSymbolConstants.CrossSectionCardinalPoints, SPSSymbolConstants.CMNSCH, SPSSymbolConstants.TDL_INVALID_SIDEFRAME_SECTIONCP)

            'validate handrail cross-section cardinal point
            ValidateUndefinedCodelistValue(stair, SPSSymbolConstants.HandRailSectionCP, SPSSymbolConstants.CrossSectionCardinalPoints, SPSSymbolConstants.CMNSCH, SPSSymbolConstants.TDL_INVALID_HANDRAIL_SECTIONCP)

            'validate step cross-section cardinal point
            ValidateUndefinedCodelistValue(stair, SPSSymbolConstants.StepSectionCP, SPSSymbolConstants.CrossSectionCardinalPoints, SPSSymbolConstants.CMNSCH, SPSSymbolConstants.TDL_INVALID_STEP_SECTIONCP)

        End Sub

        Private Overloads Sub ValidateUndefinedCodelistValue(ByVal stair As Stair, ByVal propertyName$, _
                                                             ByVal codelistTableName$, ByVal codelistTableNamespace$, _
                                                             ByVal errNumber As Integer)

            Dim propertyValue As Integer
            Try
                propertyValue = StructHelper.GetIntProperty(stair, SPSSymbolConstants.IJUAStairTypeA, propertyName)
            Catch ex As Exception
                'unable to get the property value, create a warning todo record and exit
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, propertyName + " " + StairSymbolsLocalizer.GetString(StairSymbolsResourceIDs.ErrPropertyNotFound, _
                                                            "property does not exist on the Stair in StariTypeA Symbol. Check the error log and catalog data."))
                Return
            End Try

            'now, check the value is valid or not
            If Not StructHelper.IsValidCodeListValue(codelistTableName, codelistTableNamespace, propertyValue) Then
                'create a todo record with appropriate message
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, errNumber,
                                                             String.Format("Error while validating stair PropertyValue as propertyValue: {0} does not exist in {1} table of StairTypeA Symbol", propertyValue.ToString, codelistTableName), stair)
            End If

        End Sub

#End Region

    End Class

End Namespace

