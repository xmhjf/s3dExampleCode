'=====================================================================================================
'
'Copyright 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'StairSymbols resource identifiers.
'
'File
'  ResourceIdentifiers.vb
'
'=====================================================================================================
Namespace Ingr.SP3D.Content.Structure

    NotInheritable Class StairSymbolsResourceIDs

        ''' <summary>
        ''' Stair symbols creation localizer resource.
        ''' </summary>
        Public Const DEFAULT_RESOURCE As String = "StairSymbols"

        ''' <summary>
        ''' Stair symbols creation localizer assembly.
        ''' </summary>
        Public Const DEFAULT_ASSEMBLY As String = "StairSymbols"

        ''' <summary>
        ''' Cannot get symbol output to calculate weight and center of gravity.
        ''' </summary>
        Public Const ErrNoVolumeCOGSymbolOutput As Integer = 1

        ''' <summary>
        ''' Property does not exist on the stair. Please check the catalog part.
        ''' </summary>
        Public Const ErrPropertyNotFound As Integer = 2

        ''' <summary>
        ''' Section not found in catalog.
        ''' </summary>
        Public Const ErrSectionNotFound As Integer = 3

        ''' <summary>
        ''' Cannot calculate weight and center of gravity as required system attribute dry weight and center of gravity origin value can not obtained. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrWeightCGAttributeData As Integer = 5

        ''' <summary>
        ''' Cannot calculate weight and center of gravity, as required user attribute material name and grade value cannot be obtained from the catalog. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrMaterialAttributeData As Integer = 6

        ''' <summary>
        ''' Cannot calculate weight and center of gravity, as the required material is not found in catalog. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrMaterialNotFound As Integer = 7

        ''' <summary>
        ''' Weight and COG failed to evaluate, as some of the required system attribute values can not obtained. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrWCOGMissingSystemAttributeData As Integer = 8

        ''' <summary>
        ''' Cannot set weight and center of gravity on the stair.
        ''' </summary>
        Public Const ErrSetWeightAndCOG As Integer = 9

        ''' <summary>
        ''' Step section properties are missing, please check catalog.
        ''' </summary>
        Public Const ErrMissingStepSectionProperties As Integer = 10

        ''' <summary>
        ''' Handrail section properties are missing, please check catalog.
        ''' </summary>
        Public Const ErrMissingHandRailSectionProperties As Integer = 11

        ''' <summary>
        ''' Side frame section properties are missing, please check catalog.
        ''' </summary>
        Public Const ErrMissingSideFrameSectionProperties As Integer = 12

        ''' <summary>
        ''' Error in constructing outputs for stair.
        ''' </summary>
        Public Const ErrConstructOutputs As Integer = 13
    End Class

End Namespace

