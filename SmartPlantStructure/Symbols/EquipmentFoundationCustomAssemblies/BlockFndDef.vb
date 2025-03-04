'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright 1992 - 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  BlockFndDef.vb
'
'Abstract
'	BlockFndDef is a .NET custom assembly definition which creates graphic outputs for representing a block foundation in the model.
'   This class subclasses from EquipmentFoundationCustomAssemblyDefinition.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Imports System
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.Structure.Middle
Imports Ingr.SP3D.ReferenceData.Exceptions
Imports Ingr.SP3D.Structure.Exceptions

'===========================================================================================
'Namespace of this class is Ingr.SP3D.Content.Structure
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'===========================================================================================
Namespace Ingr.SP3D.Content.Structure

    'All the graphic outputs are not known at symbol design time. 
    'Based on inputs, the symbol will create outputs dynamically at runtime, this .NET custom assembly definition has variable outputs.
    <SymbolVersion("1.0.0.0")> _
    <CacheOption(CacheOptionType.NonCached)> _
    <VariableOutputs()> _
    <OutputNotification(SPSSymbolConstants.IID_IJUASPSBlockFndn)> _
    <OutputNotification(SPSSymbolConstants.IID_IJStructMaterial)> _
    Public Class BlockFndDef : Inherits EquipmentFoundationCustomAssemblyDefinition
        Private isOnPreLoad As Boolean = False

        '=====================================================================================================================
        'DefinitionName/ProgID of this symbol is "EquipmentFoundationCustomAssemblies,Ingr.SP3D.Content.Structure.BlockFndDef"
        '=====================================================================================================================

#Region "Definition of Inputs"

        'Following inputs are needed to create the graphic outputs, modification of these inputs will trigger the re-computation of the graphic outputs.
        <InputCatalogPart(1)> _
        Public catalogPart As InputCatalogPart
        <InputDouble(2, "BlockLength", "Length of the equipment containing bolt-holes position", 1.05)> _
        Public blockLength As InputDouble
        <InputDouble(3, "BlockWidth", "Mounting width which comes from the equipment", 0.55)> _
        Public blockWidth As InputDouble
        <InputDouble(4, "BlockHeight", "Height of the block foundation", 0.5)> _
        Public blockHeight As InputDouble
        <InputDouble(5, "IsBlockSizeDrivenByRule", "Block size driven by rule (bolt hole locations)", 1)> _
        Public isBlockSizeDrivenByRule As InputDouble
        <InputDouble(6, "BlockEdgeClearance", "Block edge clearance", 0.0001)> _
        Public blockEdgeClearance As InputDouble
        <InputString(7, "BlockSPSMaterial", "Block material", "Concrete")> _
        Public blockMaterial As InputString
        <InputString(8, "BlockSPSGrade", "Block material grade", "Fc 4000")> _
        Public blockMaterialGrade As InputString

#End Region

#Region "Definitions of Aspects and their outputs"

        'SimplePhysical Aspect
        <SymbolOutput(SPSSymbolConstants.Block, "Block Geometry")> _
        <Aspect("SimplePhysical", "Simple Physical Aspect", AspectID.SimplePhysical)> _
        Public simplePhysicalAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"
        ''' <summary>
        ''' Method to be overridden by the inheriting class when the
        ''' need arises to perform actions prior to calling the
        ''' construction of the symbol outputs.
        ''' We will access the occurrence and evaluate the dimensions here so that they can
        ''' be accessed later in ConstructOutputs
        ''' </summary>
        Public Overrides Sub PreConstructOutputs()
            '========================================
            ' Construction of Simple Physical Aspect 
            '========================================
            Dim length#, width#, height#
            Dim foundation As EquipmentFoundation = DirectCast(Occurrence, EquipmentFoundation)

            '1. First evaluate the foundation to set the origin and orientation based on its supported/supporting objects
            'Based on these dimensions, appropriate foundation geometry needs to be constructed later.
            MyBase.EvaluateFoundation(foundation, width, length, height)

            '2. Now set the custom properties on the foundation
            Dim placedByPoint As Boolean = MyBase.IsPlacedByPoint(foundation)
            Dim boundedByPlane As Boolean = MyBase.IsBoundedBySurface(foundation)
            Dim hasSlab As Boolean = False ' No Slab in this case
            FoundationServices.SetDimensionalProperties(foundation, FoundationComponentType.Block, length, width, height, placedByPoint, boundedByPlane, hasSlab)
        End Sub
        ''' <summary>
        ''' Evaluates the foundation orientation and size and constructs its symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Try
                Dim foundation As EquipmentFoundation = DirectCast(Occurrence, EquipmentFoundation)

                Dim length# = Me.blockLength.Value
                Dim width# = Me.blockWidth.Value
                Dim height# = Me.blockHeight.Value
                Dim clearance = Me.blockEdgeClearance.Value

                '1. Now, construct the block geometry. 
                'For single block always need to clipped by supporting surface if there is one.
                Dim clipBySupportingSurface As Boolean = True
                'Get the clipping plane in local coordinate system of the foundation
                Dim clippingPlane As Plane3d = MyBase.GetLocalClippingPlane(clipBySupportingSurface, foundation, foundation.Matrix)
                'Block geometry base offset is zero as the base or top of the bock is at the local origin. 
                Dim baseZOffset# = 0
                'Place rectangular block geometry
                Try
                    FoundationServices.PlaceRectangularFoundation(Me.OccurrenceConnection, length, width, height, clearance, FoundationComponentType.Block, clippingPlane, baseZOffset, Me.simplePhysicalAspect)
                Catch ex As InvalidOperationException
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrBlockFoundationConstruction,
                                    "Error in placing rectangular foundation in Block Foundation Definition. Please check the input dimensions or symbol definition code."))
                End Try

                '2. Finally get and set the physical properties on the foundation
                'Get physical properties - exposed Surface Area
                Dim surfaceArea As Double = FoundationServices.GetExposedSurfaceArea(length, width, height, clearance)
                'Set the physical properties on the foundation- Surface Area,  Volume and Center of Gravity
                'VolumeCG is added an output and surface area property is set on the foundation.
                FoundationServices.SetPhysicalProperties(OccurrenceConnection, foundation, length, width, height, clearance, FoundationComponentType.Block, clipBySupportingSurface, surfaceArea, baseZOffset, Me.simplePhysicalAspect)
            Catch Ex As Exception ' General Unhandled exception 
                If MyBase.ToDoListMessage Is Nothing Then 'Check ToDoListMessgae created already or not
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrBlockFoundationConstructOutputs,
                                    "Error in constructing outputs for foundation component in Block Foundation Definition. Please check your custom code or contact S3D support."))
                End If
            End Try

        End Sub

#End Region

#Region "Public Override Functions and Methods"

        ''' <summary>
        ''' This method is expected to be overridden by the inheriting class to construct and re-evaluate the custom assembly outputs.
        ''' Set the material and evaluate the Weight and COG of the component
        ''' </summary>
        Public Overrides Sub EvaluateAssembly()
            Dim foundation As EquipmentFoundation = DirectCast(Occurrence, EquipmentFoundation)
            Try
                'Set the material 
                'Read the material information from the inputs and get the material object.
                Dim materialName$ = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockMaterial)
                Dim materialGrade$ = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockMaterialGrade)
                Dim catalogStructHelper = FoundationServices.CatalogStructHelper
                Dim material As Material = catalogStructHelper.GetMaterial(materialName, materialGrade)
                foundation.SetMaterial(SPSSymbolConstants.Block, material)
            Catch ex As RefDataMaterialNotFoundException
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrBlockFoundationMaterialNotFound, _
                                        "Error in setting the material on the foundation while evaluating Weight and CG in Block Foundation Definition, as the required material is not found in catalog. Check the error log and catalog data."))
                Return
            End Try

            'Evaluate weight and COG for the foundation here as evaluate is called after construction of outputs.
            'Also this allows COG to be updated during copy-paste/mirror operations when actual symbol geoemtry is not re-constructed again.
            Try
                FoundationServices.EvaluateWeightCG(DirectCast(Occurrence, FoundationBase), FoundationComponentType.Block)
            Catch ex As Exception
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Occurrence.ToString + " " + EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrBlockFoundationWCOGMissingSystemAttributeData, _
                                        "Error in calculating weight and center of gravity in Block Foundation Definition, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data."))
            End Try

        End Sub

        ''' <summary>
        ''' Returns  the name, material, geometry and attributes of the components of the equipment foundation.
        ''' Implement this method if the equipment foundation doesn't have children business objects to represent
        ''' its components but information about the components is needed to be provided during data exchange
        ''' with other applications.
        ''' </summary>
        ''' <param name="businessObject">Equipment foundation business object.</param>
        ''' <returns>Returns a collection of custom output.</returns>
        Public Overrides Function GetComponents(ByVal businessObject As BusinessObject) As Collection(Of CustomOutput)

            If businessObject Is Nothing Then
                Throw New ArgumentNullException("businessObject")
            End If

            Dim equipmentFoundation As EquipmentFoundation = DirectCast(businessObject, EquipmentFoundation)

            Dim components As New Collection(Of CustomOutput)
            Dim blockFaces As New Collection(Of Surface3d)

            'get output geometry of each component from the physical representation and add to the respective collection
            Dim blockSurface As Surface3d = DirectCast(SymbolHelper.GetSymbolOutput(equipmentFoundation, "SimplePhysical", SPSSymbolConstants.Block), Surface3d)
            If Not blockSurface Is Nothing Then
                blockFaces.Add(blockSurface)
            End If

            'create output to represent block
            If blockFaces.Count > 0 Then
                'get the block material type and material grade
                Dim blockMaterialType As String = StructHelper.GetStringProperty(equipmentFoundation, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockMaterial)
                Dim blockMaterialGrade As String = StructHelper.GetStringProperty(equipmentFoundation, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockMaterialGrade)

                'get properties specific to the block component from the equipment foundation and to the respective properties collection
                Dim blockProperties As New Collection(Of PropertyValue)
                For Each propertyValue As PropertyValue In equipmentFoundation.GetAllProperties()
                    Dim interfaceName As String = propertyValue.PropertyInfo.InterfaceInfo.Name
                    If interfaceName = SPSSymbolConstants.IJUASPSBlockFndn Then
                        blockProperties.Add(propertyValue)
                    End If
                Next

                Dim blockSurfaceOutput As New CustomSurfaceOutput(SPSSymbolConstants.Block, blockMaterialType, blockMaterialGrade, blockFaces, blockProperties)
                components.Add(blockSurfaceOutput)
            End If

            Return components

        End Function

#End Region

#Region "Property Management Methods"

        ''' <summary>
        ''' OnPreLoad gets called immediately before the properties are loaded in the property page control. 
        ''' Any change to the display status of properties is to be done here.
        ''' </summary>
        ''' <param name="businessObject">EquipmentFoundation business object which aggregates symbol.</param>
        ''' <param name="colAllDisplayedValues">Read-only collection of all properties displayed in the property pages control.</param>
        Public Overrides Sub OnPreLoad(ByVal businessObject As BusinessObject, ByVal colAllDisplayedValues As ReadOnlyCollection(Of PropertyDescriptor))

            'check if any of the arguments are null and throw null argument exception.  
            'businessObject can be null.
            If colAllDisplayedValues Is Nothing Then
                Throw New ArgumentNullException("colAllDisplayedValues")
            End If

            Dim equipmentFoundation As EquipmentFoundation = Nothing
            If Not businessObject Is Nothing Then
                equipmentFoundation = DirectCast(businessObject, EquipmentFoundation)
            End If

            'optimization to avoid value validation in OnPropertyChange
            Me.isOnPreLoad = True

            Dim isOnPropertyChange As Boolean = False
            Dim errorMessage As String = String.Empty
            Dim propertyDescriptor As PropertyDescriptor
            Dim propertyValue As PropertyValue
            Dim propertyName$

            For i As Integer = 0 To colAllDisplayedValues.Count - 1
                propertyDescriptor = colAllDisplayedValues(i)
                propertyValue = propertyDescriptor.Property
                propertyName = propertyValue.PropertyInfo.Name

                'make all these properties read-only
                Select Case propertyName
                    'need to gray out the BlockHeight if it has constraining plane
                    Case SPSSymbolConstants.BlockHeight
                        If Not equipmentFoundation Is Nothing Then
                            If MyBase.IsBoundedBySurface(equipmentFoundation) Then
                                propertyDescriptor.ReadOnly = True
                            End If
                        End If
                        'need to gray out the IsBlockSizeDrivenByRule if it is created By Point option
                    Case SPSSymbolConstants.IsBlockSizeDrivenByRule
                        If Not equipmentFoundation Is Nothing Then
                            If MyBase.IsPlacedByPoint(equipmentFoundation) Then
                                'single point case
                                propertyDescriptor.ReadOnly = True
                            End If
                        End If
                        'need to gray out the SurfaceArea and Volume
                    Case SPSSymbolConstants.SurfaceArea, SPSSymbolConstants.Volume
                        propertyDescriptor.ReadOnly = True
                End Select

                isOnPropertyChange = OnPropertyChange(businessObject, colAllDisplayedValues, propertyDescriptor, propertyValue, errorMessage)

                If errorMessage.Length > 0 Then
                    Me.isOnPreLoad = False
                    Exit For
                End If
            Next

            Me.isOnPreLoad = False

        End Sub


        ''' <summary>
        ''' OnPropertyChange is called each time a property is modified. Any custom validation to be done here.
        ''' </summary>
        ''' <param name="businessObject">EquipmentFoundation business object which aggregates symbol.</param>
        ''' <param name="colAllDisplayedValues">Read-only collection of all properties displayed in the property pages control.</param>
        ''' <param name="propertyToChange">Property being modified.</param>
        ''' <param name="newPropertyValue">New value of the property.</param>
        ''' <param name="errorMessage">Custom error message returned by validation.</param>
        ''' <returns>Returns false if change in property value is not valid.</returns>
        Public Overrides Function OnPropertyChange(ByVal businessObject As BusinessObject, ByVal colAllDisplayedValues As ReadOnlyCollection(Of PropertyDescriptor), _
                                         ByVal propertyToChange As PropertyDescriptor, ByVal newPropertyValue As PropertyValue, _
                                         ByRef errorMessage As String) As Boolean

            'check if any of the arguments are null and throw null argument exception.  
            ' businessObject and colAllDisplayedValues can be null.
            If propertyToChange Is Nothing Then
                Throw New ArgumentNullException("propertyToChange")
            End If
            If newPropertyValue Is Nothing Then
                Throw New ArgumentNullException("newPropertyValue")
            End If

            Dim isOnPropertyChange As Boolean = False
            Dim propertyInformation As PropertyInformation = propertyToChange.Property.PropertyInfo
            Dim propertyName$ = propertyInformation.Name
            Dim interfaceName$ = propertyInformation.InterfaceInfo.Name
            Dim propertyDescriptor As PropertyDescriptor
            Dim propertyValue As PropertyValue
            Dim propertyNameToReadOnly$

            'If property name is IsBlockSizeDrivenByRule and the value is True we need to gray out BlockEdgeClearance, BlockLength and BlockWidth   
            'If property name is ReportingRequirements and no value defined we need to gray out the ReportingType
            If propertyName = SPSSymbolConstants.IsBlockSizeDrivenByRule Or propertyName = SPSSymbolConstants.ReportingRequirements Then
                For i As Integer = 0 To colAllDisplayedValues.Count - 1
                    propertyDescriptor = colAllDisplayedValues(i)
                    propertyValue = propertyDescriptor.Property
                    propertyNameToReadOnly = propertyValue.PropertyInfo.Name

                    Select Case propertyNameToReadOnly
                        Case SPSSymbolConstants.BlockEdgeClearance, SPSSymbolConstants.BlockLength, SPSSymbolConstants.BlockWidth
                            If propertyName = SPSSymbolConstants.IsBlockSizeDrivenByRule Then
                                If DirectCast(newPropertyValue, PropertyValueBoolean).PropValue Then
                                    propertyDescriptor.ReadOnly = True
                                Else
                                    propertyDescriptor.ReadOnly = False
                                End If
                            End If
                        Case SPSSymbolConstants.ReportingType
                            If propertyName = SPSSymbolConstants.ReportingRequirements Then
                                If DirectCast(newPropertyValue, PropertyValueCodelist).PropValue = -1 Then
                                    propertyDescriptor.ReadOnly = True
                                Else
                                    propertyDescriptor.ReadOnly = False
                                End If
                            End If
                    End Select
                Next
            End If

            If Not Me.isOnPreLoad Then
                'If the OnPropertyChange event is triggered explicitly,not through OnPreLoad event, we need to go for the following check.
                isOnPropertyChange = CustomPropertyValidate(propertyName, newPropertyValue, errorMessage)
                If isOnPropertyChange Then
                    errorMessage = String.Empty ' empty string indicates a success
                ElseIf errorMessage.Length = 0 Then
                    'Expected an error message to be provided to describe the reason for the failure
                    'this is required for posting the failure to the user.                
                    Throw New CustomSymbolErrorValidationMustProvideMessage(EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrCustomContextMessageMissing, _
                    "A context message must be provided for the user when a property is tagged as invalid, InterfaceName:PropertyName -:" + interfaceName + ":" + propertyName + ".")) 'to track it in the error log
                End If
            End If

            Return isOnPropertyChange

        End Function

        ''' <summary>
        ''' This method is used to validate the block equipment foundation property value.
        ''' </summary>
        ''' <param name="propertyName">The name of the property.</param>
        ''' <param name="newPropertyValue">The value of the property to validate.</param>
        ''' <param name="errorMessage">The error message if validation fails.</param>
        ''' <returns>True if property value validation succeeds.</returns>
        Private Function CustomPropertyValidate(ByVal propertyName As String, ByVal newPropertyValue As PropertyValue, ByRef errorMessage As String) As Boolean

            'by default set the property value as valid. Override the value later for known checks
            Dim isValidCustomProperty As Boolean = True
            Dim propertyValue As Double = Nothing

            Dim propertyType As SP3DPropType = newPropertyValue.PropertyInfo.PropertyType

            If propertyType = SP3DPropType.PTDouble Then
                Dim propertyValueDbl As PropertyValueDouble = DirectCast(newPropertyValue, PropertyValueDouble)
                If propertyValueDbl.PropValue IsNot Nothing Then
                    propertyValue = CDbl(propertyValueDbl.PropValue)
                Else
                    Return isValidCustomProperty
                End If
            End If

            'for following property the property value should be greater than zero
            Select Case propertyName
                Case SPSSymbolConstants.BlockLength, SPSSymbolConstants.BlockWidth, _
                    SPSSymbolConstants.BlockHeight
                    isValidCustomProperty = ValidationHelper.IsGreaterThanZero(propertyValue, errorMessage)
                Case SPSSymbolConstants.BlockEdgeClearance
                    isValidCustomProperty = Not ValidationHelper.IsNegative(propertyValue, errorMessage)
            End Select

            Return isValidCustomProperty

        End Function

#End Region

    End Class

End Namespace

