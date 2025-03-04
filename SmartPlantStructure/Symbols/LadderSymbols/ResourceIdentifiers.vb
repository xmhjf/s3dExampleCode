'=====================================================================================================
'
'Copyright 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'LadderSymbols resource identifiers.
'
'File
'  ResourceIdentifiers.vb
'
'=====================================================================================================
Namespace Ingr.SP3D.Content.Structure

    NotInheritable Class LadderSymbolsResourceIDs

        ''' <summary>
        ''' Ladder symbols creation localizer resource.
        ''' </summary>
        Public Const DEFAULT_RESOURCE As String = "LadderSymbols"

        ''' <summary>
        ''' Ladder symbols creation localizer assembly.
        ''' </summary>
        Public Const DEFAULT_ASSEMBLY As String = "LadderSymbols"

        ''' <summary>
        ''' Cannot get symbol output to calculate weight and center of gravity.
        ''' </summary>
        Public Const ErrNoVolumeCOGSymbolOutput As Integer = 1

        ''' <summary>
        ''' Property does not exist on the ladder. Please check the catalog part.
        ''' </summary>
        Public Const ErrPropertyNotFound As Integer = 2

        ''' <summary>
        ''' Geometry failed to evaluate, as some user attribute values can not be obtained from the catalog part. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrUserAttributeDataNotFound As Integer = 3

        ''' <summary>
        ''' Geometry failed to evaluate, error during setting of some of the user attribute value. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrSetUserAttributeData As Integer = 4

        ''' <summary>
        ''' Cannot calculate weight and center of gravity, as required user attribute material name and grade value cannot be obtained from the catalog. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrMaterialAttributeData As Integer = 6

        ''' <summary>
        ''' Weight and COG failed to evaluate, as the required material is not found in catalog. Check the error log and catalog data.
        ''' </summary>
        Public Const ErrMaterialNotFound As Integer = 7

        ''' <summary>
        ''' Cannot set weight and center of gravity on the ladder.
        ''' </summary>
        Public Const ErrSetWeightAndCOG As Integer = 8

        ''' <summary>
        ''' Error in constructing outputs for ladder.
        ''' </summary>
        Public Const ErrConstructOutputs As Integer = 9
    End Class

End Namespace

