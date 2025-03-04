'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright 1992 - 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  BlockFndCompDef.vb
'
'Abstract
'	BlockFndCompDef is a .NET custom assembly definition which creates graphic outputs for representing a block foundation component in the model.
'   This class subclasses from EquipmentFoundationCustomAssemblyDefinition.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Public Class BlockFndCompDef : Inherits EquipmentFoundationCustomAssemblyDefinition
        Private isOnPreLoad As Boolean = False

        '=========================================================================================================================
        'DefinitionName/ProgID of this symbol is "EquipmentFoundationCustomAssemblies,Ingr.SP3D.Content.Structure.BlockFndCompDef"
        '=========================================================================================================================

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
        <InputDouble(5, "IsBlockSizeDrivenByRule", "Is block size driven by rule (bolt hole locations)", 1)> _
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
                Dim foundation As FoundationBase = DirectCast(Occurrence, FoundationBase)

                'Get all dimensions from symbol inputs
                Dim length# = Me.blockLength.Value
                Dim width# = Me.blockWidth.Value
                Dim height# = Me.blockHeight.Value
                Dim clearance = Me.blockEdgeClearance.Value

                'Get WithSlab property from the custom assembly parent equipment foundation
                Dim parent As Foundation = DirectCast(SymbolHelper.GetCustomAssemblyParent(foundation), Foundation)
                Dim hasSlab As Boolean = StructHelper.GetBoolProperty(parent, SPSSymbolConstants.IJUASPSBlockAndSlabFndnAsm, SPSSymbolConstants.WithSlab)
                Dim isSupportedBySurface As Boolean = MyBase.IsBoundedBySurface(parent)
                'For block, need to clip by supporting surface if this block is a custom assembly output and has no slab under it.
                Dim clipBySupportingSurface As Boolean = isSupportedBySurface And (Not hasSlab)
                'Get the clipping plane in local coordinate system of the foundation
                Dim clippingPlane As Plane3d = MyBase.GetLocalClippingPlane(clipBySupportingSurface, parent, foundation.Matrix)

                '1. Construct the block geometry. 
                'Block geometry base offset is zero as the base or top of the bock is at the local origin. 
                Dim baseZOffset# = 0
                'Place rectangular block geometry
                Try
                    FoundationServices.PlaceRectangularFoundation(MyBase.OccurrenceConnection, length, width, height, clearance, FoundationComponentType.Block, clippingPlane, baseZOffset, Me.simplePhysicalAspect)
                Catch ex As InvalidOperationException
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrBlockConstruction,
                                     "Error in placing rectangular foundation in Block Foundation Component Definition. Please check the input dimensions or symbol definition code."))
                End Try

                '2. Get and set the physical properties on the foundation
                'Get physical properties - exposed Surface Area
                Dim surfaceArea As Double = FoundationServices.GetExposedSurfaceArea(length, width, height, clearance)
                'Set the physical properties on the foundation- Surface Area,  Volume and Center of Gravity
                'VolumeCG is added an output and surface area property is set on the foundation.
                FoundationServices.SetPhysicalProperties(MyBase.OccurrenceConnection, foundation, length, width, height, clearance, FoundationComponentType.Block, clipBySupportingSurface, surfaceArea, baseZOffset, Me.simplePhysicalAspect)
            Catch Ex As Exception ' General Unhandled exception 
                If MyBase.ToDoListMessage Is Nothing Then 'Check ToDoListMessgae created already or not
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrConstructOutputs,
                                        "Error in constructing outputs for foundation component in Block Foundation Component Definition. Please check your custom code or contact S3D support."))
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
            Dim foundation As FoundationBase = DirectCast(Me.Occurrence, FoundationBase)
            Try
                'Set the material 
                'Read the material information from the foundation component and get the material object.
                Dim materialName$ = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockMaterial)
                Dim materialGrade$ = StructHelper.GetStringProperty(foundation, SPSSymbolConstants.IJUASPSBlockFndn, SPSSymbolConstants.BlockMaterialGrade)
                Dim catalogStructHelper = FoundationServices.CatalogStructHelper
                Dim material As Material = catalogStructHelper.GetMaterial(materialName, materialGrade)
                DirectCast(foundation, FoundationComponent).Material = material
            Catch ex As RefDataMaterialNotFoundException
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrMaterialNotFound, _
                                      "Error in setting the material on the foundation while evaluating Weight and CG in Block Foundation Component Definition, as the required material is not found in catalog. Check the error log and catalog data."))
                Return
            End Try
            'Evaluate weight and COG for the foundation here as evaluate is called after construction of outputs.
            'Also this allows COG to be updated during copy-paste/mirror operations when actual symbol geometry is not re-constructed again.
            Try
                FoundationServices.EvaluateWeightCG(foundation, FoundationComponentType.Block)
            Catch ex As Exception
                'attributes might be missing, create todo record
                MyBase.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, EquipmentFoundationLocalizer.GetString(EquipmentFoundationResourceIDs.ErrWCOGMissingSystemAttributeData, _
                                        "Error in calculating weight and center of gravity in Block Foundation Component Definition, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data."))
            End Try

        End Sub

#End Region

#Region "Property Management Methods"

        ''' <summary>
        ''' OnPreLoad gets called immediately before the properties are loaded in the property page control. 
        ''' Any change to the display status of properties is to be done here.
        ''' </summary>
        ''' <param name="businessObject">Block equipment foundation component business object which aggregates symbol.</param>
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
                propertyValue = propertyDescriptor.Property
                propertyName = propertyValue.PropertyInfo.Name

                'make all these properties read-only
                Select Case propertyName
                    'need to gray out the BlockHeight if it has a constraining plane
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
        ''' <param name="businessObject">Block equipment foundation component business object which aggregates symbol.</param>
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
                        Case SPSSymbolConstants.BlockEdgeClearance, SPSSymbolConstants.BlockLength, SPSSymbolConstants.BlockWidth, SPSSymbolConstants.BlockHeight
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
        ''' This method is used to validate the block equipment foundation component property value.
        ''' </summary>
        ''' <param name="propertyName">The name of the property.</param>
        ''' <param name="newPropertyValue">The value of the property to validate.</param>
        ''' <param name="errorMessage">The error message if validation fails.</param>
        ''' <returns>True if property value validation succeeds.</returns>
        Private Function CustomPropertyValidate(ByVal propertyName As String, ByVal newPropertyValue As PropertyValue, ByRef errorMessage As String) As Boolean

            'by default set the property value as valid. Override the value later for known checks
            Dim isValidCustomProperty As Boolean = True
            Dim propertyValue As Double

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
