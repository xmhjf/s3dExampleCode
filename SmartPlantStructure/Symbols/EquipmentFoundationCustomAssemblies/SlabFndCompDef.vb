﻿'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright 1992 - 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  SlabFndCompDef.vb
'
'Abstract
'	SlabFndCompDef is a .NET custom assembly definition which creates graphic outputs for representing a slab foundation component in the model.
'   This class subclasses from EquipmentFoundationCustomAssemblyDefinition.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Public Class SlabFndCompDef : Inherits EquipmentFoundationCustomAssemblyDefinition
        Private isOnPreLoad As Boolean = False

        '========================================================================================================================
        'DefinitionName/ProgID of this symbol is "EquipmentFoundationCustomAssemblies,Ingr.SP3D.Content.Structure.SlabFndCompDef"
        '========================================================================================================================

#Region "Definition of Inputs"

        'Following inputs are needed to create the graphic outputs, modification of these inputs will trigger the re-computation of the graphic outputs.
        <InputCatalogPart(1)> _
        Public catalogPart As InputCatalogPart
        <InputDouble(2, "SlabLength", "Length of the equipment containing bolt-holes position", 1.05)> _
        Public slabLength As InputDouble
        <InputDouble(3, "SlabWidth", "Mounting width which comes from the equipment", 0.55)> _
        Public slabWidth As InputDouble
        <InputDouble(4, "SlabHeight", "Height of the slab foundation", 0.5)> _
        Public slabHeight As InputDouble
        <InputDouble(5, "IsSlabSizeDrivenByRule", "Is slab size driven by rule (bolt hole locations)", 1)> _
        Public isSlabSizeDrivenByRule As InputDouble
        <InputDouble(6, "SlabEdgeClearance", "Slab edge clearance", 0.0001)> _
        Public slabEdgeClearance As InputDouble
        <InputString(7, "SlabSPSMaterial", "Slab material", "Concrete")> _
        Public slabMaterial As InputString
        <InputString(8, "SlabSPSGrade", "Slab material grade", "Fc 4000")> _
        Public slabMaterialGrade As InputString

#End Region

#Region "Definitions of Aspects and their outputs"

        'SimplePhysical Aspect
        <SymbolOutput(SPSSymbolConstants.Slab, "Slab Geometry")> _
        <Aspect("SimplePhysical", "Simple Physical Aspect", AspectID.SimplePhysical)> _
        Public simplePhysicalAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects and evaluate the assembly"

        ''' <summary>
        ''' Construct the foundation and calculate its physical properties
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Try
                '========================================
                ' Construction of Simple Physical Aspect 
                '========================================
                'Can get occurrence as it is non-cached symbol. Need it for getting bounding surfaces for
                ' creation of outputs
                Dim foundation As FoundationComponent = DirectCast(Occurrence, FoundationComponent)

                'Get all dimensions from symbol inputs
                Dim length# = Me.slabLength.Value
                Dim width# = Me.slabWidth.Value
                Dim height# = Me.slabHeight.Value
                Dim clearance = Me.slabEdgeClearance.Value

                'Get Is bounded by surface
                Dim parent As EquipmentFoundation = DirectCast(SymbolHelper.GetCustomAssemblyParent(foundation), EquipmentFoundation)
                If parent IsNot Nothing Then
                    Dim isSupportedBySurface As Boolean = MyBase.IsBoundedBySurface(parent)
                    'Get the clipping plane in local coordinate system of the foundation
                    Dim clippingPlane As Plane3d = MyBase.GetLocalClippingPlane(isSupportedBySurface, parent, foundation.Matrix)

                    '1. Construct the slab geometry. 
                    'Slab geometry base offset is same as block height as the base or top of the slab is below the block and local origin is at top of the block. 
                    'Get block and its dimensions
                    Dim blockComponent As FoundationBase = FoundationServices.GetBlockFoundationComponent(parent)
                    Dim blockHeight#, blockWidth#, blockLength#, blockClearance#
                    FoundationServices.GetDimensions(blockComponent, FoundationComponentType.Block, blockLength, blockWidth, blockHeight, blockClearance)
                    Dim baseZOffset# = blockHeight# ' FoundationServices.GetBlockHeight(parent)
                    'Place rectangular slab geometry
                    Try
                        FoundationServices.PlaceRectangularFoundation(MyBase.OccurrenceConnection, length, width, height, clearance, FoundationComponentType.Slab, clippingPlane, baseZOffset, simplePhysicalAspect)
                    Catch ex As InvalidOperationException
                        ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrSlabConstruction,
                                                        "Error in placing rectangular foundation in Slab Foundation Component Definition. Please check the input dimensions or symbol definition code."))
                    End Try

                    '2. Finally get and set the physical properties on the foundation
                    'Get physical properties - exposed Surface Area
                    Dim surfaceArea As Double = FoundationServices.GetExposedSurfaceArea(length, width, height, clearance, blockLength, blockWidth, blockHeight, blockClearance)
                    'Set the physical properties on the foundation- Surface Area,  Volume and Center of Gravity
                    'VolumeCG is added an output and surface area property is set on the foundation.
                    FoundationServices.SetPhysicalProperties(MyBase.OccurrenceConnection, foundation, length, width, height, clearance, FoundationComponentType.Slab, isSupportedBySurface, surfaceArea, baseZOffset, simplePhysicalAspect)
                Else
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrAssemblyParent,
                                                        "Error in getting custom assembly parent for output in Slab Foundation Component Definition. Check custom code or contact S3D support."))
                End If
            Catch Ex As Exception ' General Unhandled exception 
                If MyBase.ToDoListMessage Is Nothing Then 'Check ToDoListMessgae created already or not
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrSlabFoundationConstructOutputs,
                                                        "Error in constructing outputs for foundation component in Slab Foundation Component Definition. Please check your custom code or contact S3D support."))
                End If
            End Try

        End Sub

#End Region

#Region "Evaluation of Assembly"
        ''' <summary>
        ''' This method is expected to be overridden by the inheriting class to construct and re-evaluate
        ''' the custom assembly outputs.
        ''' Set the material and evaluate the Weight and COG of the component
        ''' </summary>
        Public Overrides Sub EvaluateAssembly()

            Try
                Dim foundation As FoundationBase = DirectCast(Me.Occurrence, FoundationBase)
                'Set the material 
                'Read the material information from the foundation component and get the material object.
                Dim materialName$ = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSSlabFndn, SPSSymbolConstants.SlabMaterial)
                Dim materialGrade$ = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSSlabFndn, SPSSymbolConstants.SlabMaterialGrade)
                Dim catalogStructHelper = FoundationServices.CatalogStructHelper
                Dim material As Material = catalogStructHelper.GetMaterial(materialName, materialGrade)
                DirectCast(foundation, FoundationComponent).Material = material
            Catch ex As RefDataMaterialNotFoundException
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrSlabFoundationMaterialNotFound, _
                                        "Error in setting the material on the foundation while evaluating Weight and CG in Slab Foundation Component Definition, as the required material is not found in catalog. Check the error log and catalog data."))
                Return
            End Try

            'Evaluate weight and COG for the foundation here as evaluate is called after construction of outputs.
            'Also this allows COG to be updated during copy-paste/mirror operations when actual symbol geoemtry is not re-constructed again.
            Try
                FoundationServices.EvaluateWeightCG(DirectCast(Occurrence, FoundationBase), FoundationComponentType.Slab)
            Catch ex As Exception
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Occurrence.ToString + " " + EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrSlabFoundationWCOGMissingSystemAttributeData, _
                                        "Error in calculating weight and center of gravity in Slab Foundation Component Definition, as some of the required user attribute values cannot be obtained from the catalog. Check the error log and catalog data."))
            End Try
        End Sub

#End Region

#Region "Property Management Methods"

        ''' <summary>
        ''' OnPreLoad gets called immediately before the properties are loaded in the property page control. 
        ''' Any change to the display status of properties is to be done here.
        ''' </summary>
        ''' <param name="businessObject">Slab equipment foundation component business object which aggregates symbol.</param>
        ''' <param name="colAllDisplayedValues">Read-only collection of all properties displayed in the property pages control.</param>
        Public Overrides Sub OnPreLoad(ByVal businessObject As BusinessObject, ByVal colAllDisplayedValues As ReadOnlyCollection(Of PropertyDescriptor))

            'check if any of the arguments are null and throw null argument exception.  
            'businessObject can be null.
            If colAllDisplayedValues Is Nothing Then
                Throw New ArgumentNullException("colAllDisplayedValues")
            End If

            'if foundation component is there then we need to get the parent equipment foundation to get supported and supporting objects
            Dim equipmentFoundation As EquipmentFoundation = Nothing
            If Not businessObject Is Nothing Then
                Dim foundationComponent As FoundationComponent = DirectCast(businessObject, FoundationComponent)
                Dim systemParent As ISystem = foundationComponent.SystemParent
                equipmentFoundation = DirectCast(systemParent, EquipmentFoundation)
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
                propertyValue = propertyDescriptor.[Property]
                propertyName = propertyValue.PropertyInfo.Name

                'make all these properties read-only
                Select Case propertyName
                    'need to gray out the SlabHeight if it has a constraining plane
                    Case SPSSymbolConstants.SlabHeight
                        If Not equipmentFoundation Is Nothing Then
                            If MyBase.IsBoundedBySurface(equipmentFoundation) Then
                                propertyDescriptor.ReadOnly = True
                            End If
                        End If
                        'need to gray out the IsSlabSizeDrivenByRule if it is created By Point option
                    Case SPSSymbolConstants.IsSlabSizeDrivenByRule
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
        ''' <param name="businessObject">Slab equipment foundation component business object which aggregates symbol.</param>
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

            'If property name is IsSlabSizeDrivenByRule and the value is True we need to gray out SlabEdgeClearance, SlabLength and SlabWidth   
            'If property name is ReportingRequirements and no value defined we need to gray out the ReportingType
            If propertyName = SPSSymbolConstants.IsSlabSizeDrivenByRule Or propertyName = SPSSymbolConstants.ReportingRequirements Then
                For i As Integer = 0 To colAllDisplayedValues.Count - 1
                    propertyDescriptor = colAllDisplayedValues(i)
                    propertyValue = propertyDescriptor.Property
                    propertyNameToReadOnly = propertyValue.PropertyInfo.Name

                    Select Case propertyNameToReadOnly
                        Case SPSSymbolConstants.SlabEdgeClearance, SPSSymbolConstants.SlabLength, SPSSymbolConstants.SlabWidth, SPSSymbolConstants.SlabHeight
                            If propertyName = SPSSymbolConstants.IsSlabSizeDrivenByRule Then
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
        ''' This method is used to validate the slab equipment foundation component property value.
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
                Case SPSSymbolConstants.SlabLength, SPSSymbolConstants.SlabWidth, _
                    SPSSymbolConstants.SlabHeight
                    isValidCustomProperty = ValidationHelper.IsGreaterThanZero(propertyValue, errorMessage)
                Case SPSSymbolConstants.SlabEdgeClearance
                    isValidCustomProperty = Not ValidationHelper.IsNegative(propertyValue, errorMessage)
            End Select

            Return isValidCustomProperty

        End Function

#End Region

    End Class

End Namespace
