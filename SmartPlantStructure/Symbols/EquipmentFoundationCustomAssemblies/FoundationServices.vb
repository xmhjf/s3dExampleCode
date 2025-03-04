'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright 1992 - 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  FoundationServices.vb
'
'Abstract
'	FoundationServices is a module designed for common functionality for equipment foundation components and assemblies.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.Content.Structure
Imports Ingr.SP3D.ReferenceData.Middle
Imports System.Collections.ObjectModel
Imports System
Imports System.Collections
Imports Ingr.SP3D.Structure.Middle.Services
Imports Ingr.SP3D.Common.Exceptions

Module FoundationServices
    ''' <summary>
    ''' Foundation component types
    ''' </summary>
    Public Enum FoundationComponentType
        Block = 0
        Slab = 1
    End Enum

    Public CatalogStructHelper As New CatalogStructHelper()

    ''' <summary>
    ''' Creates the foundation component of given type as a child of the given foundation.
    ''' </summary>
    ''' <param name="foundation">The parent foundation.</param>
    ''' <param name="foundationType">Type of the foundation component.</param>
    ''' <returns>Foundation component</returns>
    Public Function CreateComponent(ByVal foundation As Foundation, ByVal foundationType As FoundationComponentType) As FoundationComponent
        Dim componentPartName As String = String.Empty

        If foundation Is Nothing Then
            Throw New CmnArgumentNullException("foundation")
        End If

        'depending upon the foundation component type we get the component part name from the catalog properties
        Select Case foundationType
            Case FoundationComponentType.Block
                componentPartName = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSBlockAndSlabFndnAsm, SPSSymbolConstants.BlockComponent)
            Case FoundationComponentType.Slab
                componentPartName = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSBlockAndSlabFndnAsm, SPSSymbolConstants.SlabComponent)
        End Select

        'Get catalog part with the help of component part name
        Dim catalogStructHelper As New CatalogStructHelper
        Dim part As Part = DirectCast(catalogStructHelper.GetPart(componentPartName), Part)

        'Construct equipment foundation component 
        Return New FoundationComponent(foundation, part)

    End Function

    ''' <summary>
    ''' Places the rectangular foundation with the given dimensions. Creates a single projection3D surface with end caps, if no supporting surface or
    ''' no need to slip by supporting surface; else creates clipped surfaces for the side along with top and bottom cap surfaces.
    ''' </summary>
    ''' <param name="connection">The connection.</param>
    ''' <param name="length">The length.</param>
    ''' <param name="width">The width.</param>
    ''' <param name="height">The height.</param>
    ''' <param name="clearance">The clearance.</param>
    ''' <param name="foundationType">Type of the foundation.</param>
    ''' <param name="clippingPlane">The clipping plane in local coordinate system of the foundation.</param>
    ''' <param name="baseZOffset">The base Z offset.</param>
    ''' <param name="aspect">The aspect.</param>
    ''' <exception cref="CmnArgumentNullException">
    ''' connection
    ''' or
    ''' foundation
    ''' or
    ''' aspect
    ''' </exception>
    ''' <exception cref="InvalidOperationException">Raised when given dimensions are invalid for creating rectangular geometry</exception>
    Public Sub PlaceRectangularFoundation(ByVal connection As SP3DConnection, length#, width#, height#, clearance#, ByVal foundationType As FoundationComponentType, _
    ByVal clippingPlane As Plane3d, ByVal baseZOffset#, ByVal aspect As AspectDefinition)

        If connection Is Nothing Then
            Throw New CmnArgumentNullException("connection")
        End If

        If aspect Is Nothing Then
            Throw New CmnArgumentNullException("aspect")
        End If

        'Construct the rectangular block geometry. 
        'We are creating a rectangular block. Need to add clearance to Width and length.
        Dim effectiveLength = length + 2 * clearance
        Dim effectiveWidth = width + 2 * clearance

        'Get the rectangular solid geometry with given dimensions
        Dim surfaces As Collection(Of ISurface) = SymbolHelper.CreateRectangularSolid(connection, effectiveLength, effectiveWidth, height, baseZOffset, clippingPlane)

        'Decide name for the output based on component Type
        Dim outputNamePrefix$ = String.Empty
        If foundationType = FoundationComponentType.Block Then
            outputNamePrefix$ = SPSSymbolConstants.Block
        ElseIf foundationType = FoundationComponentType.Slab Then
            outputNamePrefix$ = SPSSymbolConstants.Slab
        End If

        If surfaces.Count = 0 Then
            Throw New InvalidOperationException(EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrFoundationIputs, "No outputs can be created for the foundation based on given inputs."))
        End If

        If surfaces.Count = 1 Then
            'Single output. Should be Projection3D
            aspect.Outputs.Add(outputNamePrefix, surfaces.Item(0))
        Else
            'Multiple outputs. Includes caps and clipped surfaces separately.
            aspect.Outputs.Add("TopPlane", surfaces.Item(0))
            aspect.Outputs.Add("BottomPlane", surfaces.Item(1))

            For i As Integer = 2 To surfaces.Count - 1
                aspect.Outputs.Add(outputNamePrefix & i, surfaces.Item(i))
            Next
        End If

    End Sub

    ''' <summary>
    ''' Gets the block foundation component from the system children of the equipment foundation.
    ''' </summary>
    ''' <param name="foundation">The foundation - slab component or equipment foundation.</param>
    ''' <returns>Block foundation component</returns>
    Public Function GetBlockFoundationComponent(ByVal foundation As FoundationBase) As FoundationBase
        If foundation = Nothing Then
            Throw New CmnArgumentNullException("foundation")
        End If
        Dim blockFoundationComponent As FoundationComponent = Nothing
        Dim equipmentFoundation As EquipmentFoundation = Nothing

        If TypeOf foundation Is FoundationComponent Then
            'I must be Slab. Get the block from parent
            equipmentFoundation = DirectCast(SymbolHelper.GetCustomAssemblyParent(foundation), EquipmentFoundation)
        ElseIf TypeOf foundation Is EquipmentFoundation Then
            If foundation.SupportsInterface(SPSSymbolConstants.IJUASPSBlockFndn) Then
                'Non-assembly. EF itself supports the I/F.
                Return foundation
            Else
                'Parent Equip Fnd. 
                equipmentFoundation = DirectCast(foundation, EquipmentFoundation)
            End If
        End If

        If equipmentFoundation = Nothing Then
            Throw New CmnException(EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrGetEquipmentFoundation, "EquipmentFoundation cannot be obtained from the given foundation."))
        End If

        'Loop through all the components till we find one which supports the block I/F
        Dim foundationComponents As ReadOnlyCollection(Of ISystemChild) = equipmentFoundation.SystemChildren
        For Each component As ISystemChild In foundationComponents
            Dim businessObject As BusinessObject = DirectCast(component, BusinessObject)
            If (businessObject.SupportsInterface(SPSSymbolConstants.IJUASPSBlockFndn)) Then
                blockFoundationComponent = DirectCast(businessObject, FoundationComponent)
                Exit For
            End If
        Next

        Return blockFoundationComponent
    End Function

    ''' <summary>
    ''' Calculates and sets the physical properties on the foundation- Surface Area,  Volume and Center of Gravity
    ''' VolumeCG is added an output on the given aspect and surface area property is set on the foundation.
    ''' </summary>
    ''' <param name="connection">The connection.</param>
    ''' <param name="foundation">The foundation - FoundationComponent or EquipmentFoundation.</param>
    ''' <param name="length">The length.</param>
    ''' <param name="width">The width.</param>
    ''' <param name="height">The height.</param>
    ''' <param name="clearance">The clearance.</param>
    ''' <param name="componentType">Type of the component.</param>
    ''' <param name="clipByPlane">if set to true, the footing is clipped by plane.</param>
    ''' <param name="totalSurfaceArea">The total surface area.</param>
    ''' <param name="baseZOffset">The base Z offset.</param>
    ''' <param name="aspect">The aspect.</param>
    ''' <exception cref="CmnArgumentNullException">
    ''' connection
    ''' or
    ''' foundation
    ''' </exception>
    Public Sub SetPhysicalProperties(ByVal connection As SP3DConnection, ByVal foundation As FoundationBase, length#, width#, height#, clearance#, ByVal componentType As FoundationComponentType, ByVal clipByPlane As Boolean, ByVal totalSurfaceArea As Double, _
    ByVal baseZOffset#, ByRef aspect As AspectDefinition)
        If connection Is Nothing Then
            Throw New CmnArgumentNullException("connection")
        End If

        If foundation Is Nothing Then
            Throw New CmnArgumentNullException("foundation")
        End If

        Dim surfaceArea#
        'Get dimensions for the foundation
        Dim effectiveLength# = length + 2 * clearance
        Dim effectiveWidth# = width + 2 * clearance
        'Volume is L*B*H but need to include clearance
        Dim totalVolume As Double = effectiveLength * effectiveWidth * height

        Dim volumeCG As VolumeCG
        'Get all outputs from the simple physical aspect
        Dim outputs As OutputDictionary = aspect.Outputs

        If clipByPlane = True And outputs.Count > 1 Then
            'FOundation is clipped by the supporting surface. Need to get area/volume from clipped surfaces
            'We need to get only the clipped surfaces from the outputs list. So we will exclude Top and bottom planes
            Dim planes As Collection(Of IPlane) = New Collection(Of IPlane)()
            For Each dictionaryEntry As DictionaryEntry In outputs
                If dictionaryEntry.Key.Equals("TopPlane") Or dictionaryEntry.Key.Equals("BottomPlane") Then
                    Continue For
                End If
                Dim plane As IPlane = DirectCast(dictionaryEntry.Value, IPlane)
                planes.Add(plane)
            Next

            Dim clippedVerticalSurfaceArea As Double
            'Get the total vertical surface area 
            Dim totalVerticalSurfaceArea# = GetVerticalSurfaceArea(length, width, height, clearance)
            'Get the remaining surface area for top face.
            Dim topFaceSurfaceArea# = totalSurfaceArea - totalVerticalSurfaceArea
            'Get the clipped volume, COG and clipped surface area using clipped surfaces.
            volumeCG = SymbolHelper.GetVolumeCOGFromClippedSurfaces(connection, planes, totalVerticalSurfaceArea, totalVolume, clippedVerticalSurfaceArea)
            ' Surface area needs top face addition back.
            surfaceArea = topFaceSurfaceArea + clippedVerticalSurfaceArea
        Else
            'No clipping. In this case simple calculation based on length, width, height is enough.
            volumeCG = SymbolHelper.GetVolumeCGForRectangularSolid(connection, effectiveLength, effectiveWidth, height)
            surfaceArea = totalSurfaceArea

            'For Slab we will need to adjust the COG.Z by block height if not using clipped surfaces for calculation. Volume from clipped surfaces
            'is already wrt the occurrence matrix origin.
            If componentType = FoundationComponentType.Slab Then
                volumeCG.COGZ = volumeCG.COGZ - baseZOffset
            End If
        End If

        outputs.Add("VolumeCOG", volumeCG)

        'Set the volume and surface area on the given foundation
        foundation.SetPropertyValue(volumeCG.Volume, SPSSymbolConstants.IJGenericVolume, SPSSymbolConstants.Volume)
        foundation.SetPropertyValue(surfaceArea, SPSSymbolConstants.IJSurfaceArea, SPSSymbolConstants.SurfaceArea)

    End Sub

    ''' <summary>
    ''' Gets the height of the block.
    ''' </summary>
    ''' <param name="foundation">The foundation- Equipment Foundation or Foundation component.</param>
    ''' <returns>Height of the block component</returns>
    Public Function GetBlockHeight(ByVal foundation As FoundationBase) As Double
        If foundation = Nothing Then
            Throw New CmnArgumentNullException("equipmentFoundation")
        End If

        Dim block As FoundationBase = GetBlockFoundationComponent(foundation)

        If block Is Nothing Then
            Throw New CmnException("Block can not be obtained from the given foundation.")
        End If
        Return StructHelper.GetDoubleProperty(block, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockHeight)
    End Function

    ''' <summary>
    ''' Sets the foundation custom dimensional properties to given values.
    ''' </summary>
    ''' <param name="foundation">The foundation - FoundationComponent or EquipmentFoundation.</param>
    ''' <param name="componentType">Type of the component.</param>
    ''' <param name="length">The length.</param>
    ''' <param name="width">The width.</param>
    ''' <param name="height">The height.</param>
    ''' <param name="placedByPoint">if set to <c>true</c> fundation is placed by point.</param>
    ''' <param name="boundedByPlane">if set to <c>true</c> foundation is bounded by plane.</param>
    ''' <param name="hasSlab">if set to <c>true</c> [has slab].</param>
    Public Sub SetDimensionalProperties(ByVal foundation As FoundationBase, ByVal componentType As FoundationComponentType, ByVal length As Double, ByVal width As Double, ByVal height As Double, ByVal placedByPoint As Boolean, ByVal boundedByPlane As Boolean, Optional ByVal hasSlab As Boolean = True)

        If foundation = Nothing Then
            Throw New CmnArgumentNullException("foundation")
        End If

        ' Length,Width and Height is to be calculated only for the components. 
        'Calculate the width and length after considering the offsets.
        Dim xOffset As Double = 0.0, yOffset As Double = 0.0

        'Set interface and property names related to sizing
        Dim componentInterface$ = SPSSymbolConstants.IJUASPSBlockFndn
        Dim sizeDrivenByRuleProperty$ = SPSSymbolConstants.IsBlockSizeDrivenByRule
        Dim lengthProperty$ = SPSSymbolConstants.BlockLength
        Dim widthProperty$ = SPSSymbolConstants.BlockWidth
        Dim heightProperty$ = SPSSymbolConstants.BlockHeight
        If componentType = FoundationComponentType.Slab Then
            componentInterface = SPSSymbolConstants.IJUASPSSlabFndn
            sizeDrivenByRuleProperty = SPSSymbolConstants.IsSlabSizeDrivenByRule '"IsSlabSizeDrivenByRule"
            lengthProperty = SPSSymbolConstants.SlabLength '"SlabLength"
            widthProperty = SPSSymbolConstants.SlabWidth '"SlabWidth"
            heightProperty = SPSSymbolConstants.SlabHeight '"SlabHeight"
        End If

        'Set the Sizing rule to read-only if placed by point
        If placedByPoint Then
            foundation.SetPropertyValue(False, componentInterface, sizeDrivenByRuleProperty)
        End If

        'Need the Sizing rule property to figure out if we need to set the sizes or not
        'Sizing Rule may be set by the user manually in property page. So get it's current value on the foundation
        Dim isSizeDrivenByRule As Boolean = StructHelper.GetBoolProperty(foundation, componentInterface, sizeDrivenByRuleProperty)
        If isSizeDrivenByRule Then
            'Set sizes on the foundation
            foundation.SetPropertyValue(width, componentInterface, widthProperty)
            foundation.SetPropertyValue(length, componentInterface, lengthProperty)
        End If

        'Set height only if bounded by plane
        If boundedByPlane Then
            'When we have assembly with slab, in that case reset the block height to half the height so that 
            ' block and slab extend equally. Slab, however may be clipped by the supporting plane later.
            If hasSlab Then
                height = height / 2
                'Dim blockComponent As FoundationBase = GetBlockFoundationComponent(foundation)
                'blockComponent.SetPropertyValue(height, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockHeight)
            End If

            foundation.SetPropertyValue(height, componentInterface, heightProperty)
        End If
    End Sub

    ''' <summary>
    ''' Gets the foundation dimensions.
    ''' </summary>
    ''' <param name="foundation">The foundation - FoundationComponent or EquipmentFoundation.</param>
    ''' <param name="componentType">Type of the foundation.</param>
    ''' <param name="length">The length.</param>
    ''' <param name="width">The width.</param>
    ''' <param name="height">The height.</param>
    ''' <param name="clearance">The clearance.</param>
    Public Sub GetDimensions(ByVal foundation As FoundationBase, ByVal componentType As FoundationComponentType, ByRef length As Double, ByRef width As Double, ByRef height As Double, ByRef clearance As Double)

        If foundation = Nothing Then
            Throw New CmnArgumentNullException("foundation")
        End If

        'Set interface and property names related to sizing
        Dim componentInterface$ = "IJUASPSBlockFndn"
        Dim sizeDrivenByRuleProperty$ = "IsBlockSizeDrivenByRule"
        Dim lengthProperty$ = "BlockLength"
        Dim widthProperty$ = "BlockWidth"
        Dim heightProperty$ = "BlockHeight"
        Dim clearanceProperty$ = "BlockEdgeClearance"
        If componentType = FoundationComponentType.Slab Then
            componentInterface = "IJUASPSSlabFndn"
            sizeDrivenByRuleProperty = "IsSlabSizeDrivenByRule"
            lengthProperty = "SlabLength"
            widthProperty = "SlabWidth"
            heightProperty = "SlabHeight"
            clearanceProperty = "SlabEdgeClearance"
        End If
        length = StructHelper.GetDoubleProperty(foundation, componentInterface, lengthProperty)
        width = StructHelper.GetDoubleProperty(foundation, componentInterface, widthProperty)
        height = StructHelper.GetDoubleProperty(foundation, componentInterface, heightProperty)
        clearance = StructHelper.GetDoubleProperty(foundation, componentInterface, clearanceProperty)
    End Sub

    ''' <summary>
    ''' Gets the vertical surface area (only) for given rectangular solid.
    ''' </summary>
    ''' <param name="length">Length.</param>
    ''' <param name="width">Width.</param>
    ''' <param name="height">Height.</param>
    ''' <param name="clearance">clearance.</param>
    ''' <returns>Surface area of the vertical sides</returns>
    Private Function GetVerticalSurfaceArea(ByVal length As Double, ByVal width As Double, ByVal height As Double, ByVal clearance As Double) As Double
        'if the foundation is placed via a surface as a supporting object then we must pass the vertical surface area to the 
        'clipping function so the correct volume is calculated.  
        Dim effectiveLength# = length + 2 * clearance
        Dim effectiveWidth# = width + 2 * clearance
        Return (2 * effectiveLength * height) + (2 * effectiveWidth * height)
    End Function

    ''' <summary>
    ''' Gets the exposed surface area for the foundation with given dimensions.
    ''' </summary>
    ''' <param name="length">The length.</param>
    ''' <param name="width">The width.</param>
    ''' <param name="height">The height.</param>
    ''' <param name="clearance">The clearance.</param>
    ''' <returns></returns>
    Public Function GetExposedSurfaceArea(length#, width#, height#, clearance#) As Double
        Dim adjustedLength# = length + (2 * clearance)
        Dim adjustedWidth# = width + (2 * clearance)

        'Exclude Bottom face as the component should be placed on top of some surface and that face is not exposed
        Dim topFaceArea# = adjustedLength * adjustedWidth
        Dim totalArea# = topFaceArea + 2 * adjustedLength * height + 2 * adjustedWidth * height

        Return totalArea

    End Function

    ''' <summary>
    ''' Gets the exposed surface area for a Slab with given dimensions and dimensions of the block above it.
    ''' </summary>
    ''' <param name="length">The length.</param>
    ''' <param name="width">The width.</param>
    ''' <param name="height">The height.</param>
    ''' <param name="clearance">The clearance.</param>
    ''' <param name="blockLength">Length of the block.</param>
    ''' <param name="blockWidth">Width of the block.</param>
    ''' <param name="blockHeight">Height of the block.</param>
    ''' <param name="blockClearance">The block clearance.</param>
    ''' <returns></returns>
    Public Function GetExposedSurfaceArea(length#, width#, height#, clearance#, blockLength#, blockWidth#, blockHeight#, blockClearance#) As Double
        'First get the foundation dimensions
        Dim adjustedLength# = length + (2 * clearance)
        Dim adjustedWidth# = width + (2 * clearance)

        'Exclude Bottom face as the component should be placed on top of some surface and that face is not exposed
        Dim topFaceArea# = adjustedLength * adjustedWidth
        Dim totalArea# = topFaceArea + 2 * adjustedLength * height + 2 * adjustedWidth * height


        'If the slab is part of an assembly with a block then we must subtract the block bottom face from the slab
        'get the slab foundation component parent
        adjustedLength# = blockLength + (2 * blockClearance)
        adjustedWidth# = blockWidth + (2 * blockClearance)
        Dim blockTopFaceArea# = adjustedLength * adjustedWidth
        If topFaceArea > blockTopFaceArea Then
            Return totalArea - blockTopFaceArea
        End If
        Return totalArea - topFaceArea
    End Function

    ''' <summary>
    ''' Evaluates weight and center of gravity of the equipment foundation and sets it on the equipment foundation business object.
    ''' </summary>
    Public Sub EvaluateWeightCG(ByVal foundation As FoundationBase, ByVal foundationType As FoundationComponentType)

        If foundation = Nothing Then
            Throw New CmnArgumentNullException("foundation")
        End If

        ' Get the output named VolumeCOG from Simple physical aspect
        Dim symbolOutput As BusinessObject = SymbolHelper.GetSymbolOutput(foundation, "SimplePhysical", "VolumeCOG")
        'If VolumeCOG output is not there, it means the symbol is not computed yet, better to skip weight and COG calculation now.
        If Not symbolOutput Is Nothing Then
            'Getting Weight and COG origin
            Dim weightCOGOrigin As Integer
            Try
                weightCOGOrigin = StructHelper.GetIntProperty(foundation, SPSSymbolConstants.IJWCGValueOrigin, SPSSymbolConstants.DryWCGOrigin)
            Catch ex As Exception
                'attributes might be missing, need to create todo record, stop evaluating
                Throw ex
            End Try

            'we need to calculate weight and COG as DryWCGOrigin is Computed 
            If weightCOGOrigin = SPSSymbolConstants.DRY_WCOG_ORIGIN_COMPUTED Then
                Dim blockEquipmentFoundationVCG As VolumeCG = DirectCast(symbolOutput, VolumeCG)
                Dim volume As Double = blockEquipmentFoundationVCG.Volume
                Dim cogX As Double = blockEquipmentFoundationVCG.COGX
                Dim cogY As Double = blockEquipmentFoundationVCG.COGY
                Dim cogZ As Double = blockEquipmentFoundationVCG.COGZ

                'Transform the center of gravity (COG) from local to global
                Dim localCOG As New Position(cogX, cogY, cogZ)
                Dim matrix = foundation.Matrix
                Dim globalCOG As Position = matrix.Transform(localCOG)

                'get correct interface and property name based on foundation type. Default to Block
                Dim interfaceName$ = SPSSymbolConstants.IJUASPSBlockFndn
                Dim materialProperty$ = SPSSymbolConstants.BlockMaterial
                Dim materialGradeProperty$ = SPSSymbolConstants.BlockMaterialGrade
                If foundationType = FoundationComponentType.Slab Then
                    interfaceName = SPSSymbolConstants.IJUASPSSlabFndn
                    materialProperty = SPSSymbolConstants.SlabMaterial
                    materialGradeProperty = SPSSymbolConstants.SlabMaterialGrade
                End If

                'Getting weight from volume
                Dim weight#
                Try
                    weight = SymbolHelper.GetWeightFromVolume(foundation, interfaceName, materialProperty, materialGradeProperty, volume)
                Catch ex As Exception
                    'attributes might be missing, need to create todo record, stop evaluating
                    Throw ex
                End Try

                'Set the net weight and COG on the slab foundation componet business object using helper method provided in WeightCOGServices
                Dim weightCOGServices As New WeightCOGServices()
                weightCOGServices.SetWeightAndCOG(foundation, weight, globalCOG.X, globalCOG.Y, globalCOG.Z)

            End If
        End If
    End Sub

End Module

