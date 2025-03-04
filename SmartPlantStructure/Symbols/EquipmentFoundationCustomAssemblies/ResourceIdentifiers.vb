'=====================================================================================================
'
'Copyright 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'Equipment foundation custom assemblies resource identifiers.
'
'File
'  ResourceIdentifiers.vb
'
'=====================================================================================================
Namespace Ingr.SP3D.Content.Structure

    NotInheritable Class EquipmentFoundationResourceIDs

        ''' <summary>
        ''' Equipment foundation custom assemblies creation localizer resource.
        ''' </summary>
        Public Const DEFAULT_RESOURCE As String = "EquipmentFoundationCustomAssemblies"

        ''' <summary>
        ''' Equipment foundation custom assemblies localizer assembly.
        ''' </summary>
        Public Const DEFAULT_ASSEMBLY As String = "EquipmentFoundationCustomAssemblies"

        ''' <summary>
        ''' Error in calculating weight and center of gravity, as the required material is not found in catalog. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrMaterialNotFound As Integer = 1

        ''' <summary>
        ''' Error in  calculating weight and center of gravity, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrWCOGMissingSystemAttributeData As Integer = 2

        ''' <summary>
        ''' When the custom context message is missing.
        ''' </summary>
        Public Const ErrCustomContextMessageMissing As Integer = 3

        ''' <summary>
        ''' Error in constructing outputs for foundation component.
        ''' </summary>
        Public Const ErrConstructOutputs As Integer = 4

        ''' <summary>
        ''' Error in evaluate of equipment foundation assembly.
        ''' </summary>
        Public Const ErrEvaluateAssembly As Integer = 5

        ''' <summary>
        ''' EquipmentFoundation or FoundationComponent expected.
        ''' </summary>
        Public Const ErrFoundationType As Integer = 6

        ''' <summary>
        ''' No outputs can be created for the foundation based on given inputs.
        ''' </summary>
        Public Const ErrFoundationIputs As Integer = 7

        ''' <summary>
        ''' Error in constructing block foundation component.
        ''' </summary>
        Public Const ErrBlockConstruction As Integer = 8

        ''' <summary>
        ''' Error in constructing slab foundation component.
        ''' </summary>
        Public Const ErrSlabConstruction As Integer = 9

        ''' <summary>
        ''' EquipmentFoundation can not be obtained from the given foundation.
        ''' </summary>
        ''' <remarks></remarks>
        Public Const ErrGetEquipmentFoundation As Integer = 10

        ''' <summary>
        ''' Cannot get custom assembly parent for given assembly output.
        ''' </summary>
        Public Const ErrAssemblyParent As Integer = 11

        ''' <summary>
        ''' Error in constructing block foundation component.
        ''' </summary>
        Public Const ErrBlockFoundationConstruction As Integer = 12

        ''' <summary>
        ''' Error in constructing outputs for foundation component.
        ''' </summary>
        Public Const ErrBlockFoundationConstructOutputs As Integer = 13

        ''' <summary>
        ''' Error in  calculating weight and center of gravity, as the required material is not found in catalog. 
        ''' </summary>
        Public Const ErrBlockFoundationMaterialNotFound As Integer = 14

        ''' <summary>
        ''' Error in  calculating weight and center of gravity, as some of the required user attribute value cannot be obtained from the catalog. 
        ''' </summary>
        Public Const ErrBlockFoundationWCOGMissingSystemAttributeData As Integer = 15

        ''' <summary>
        ''' Error in  calculating weight and center of gravity, as some of the required user attribute value cannot be obtained from the catalog. 
        ''' </summary>
        Public Const ErrBlockSlabFoundationWCOGMissingSystemAttributeData As Integer = 16

        ''' <summary>
        ''' Error in constructing outputs for foundation component.
        ''' </summary>
        Public Const ErrSlabFoundationConstructOutputs As Integer = 17

        ''' <summary>
        ''' Error in calculating weight and center of gravity, as the required material is not found in catalog. 
        ''' </summary>
        Public Const ErrSlabFoundationMaterialNotFound As Integer = 18

        ''' <summary>
        ''' Error in  calculating weight and center of gravity, as some of the required user attribute value cannot be obtained from the catalog. 
        ''' </summary>
        Public Const ErrSlabFoundationWCOGMissingSystemAttributeData As Integer = 19
    End Class

End Namespace