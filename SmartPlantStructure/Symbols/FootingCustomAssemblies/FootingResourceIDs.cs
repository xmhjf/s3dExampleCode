//-------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FootingResourceIDs.cs
//
//Abstract
//	FootingResourceIDs is Footings custom assemblies resource identifiers.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Footings custom assemblies resource identifiers.
    /// </summary>
    internal static class FootingResourceIDs
    {
        /// <summary>
        /// Footing custom assemblies creation localizer resource.
        /// </summary>
        internal const string DEFAULT_RESOURCE = "FootingCustomAssemblies";

        /// <summary>
        /// Footing custom assemblies creation localizer assembly.
        /// </summary>
        internal const string DEFAULT_ASSEMBLY = "FootingCustomAssemblies";

        /// <summary>
        /// Error in constructing outputs for slab.
        /// </summary>
        internal const int ErrSlabConstructOutputs = 1;

        /// <summary>
        /// Unexpected error when shape code list value is not found in code list table.
        /// </summary>
        internal const int ErrGroutShapeCodeListValue = 2;

        /// <summary>
        /// Unexpected error when sizing rule code list value is not found in code list table.
        /// </summary>
        internal const int ErrGroutSizingRuleCodeListValue = 3;

        /// <summary>
        /// Unexpected error when orientation code list value is not found in code list table.
        /// </summary>
        internal const int ErrGroutOrientationCodeListValue = 4;

        /// <summary>
        /// Unexpected error while evaluating Footing Grout Pad definition.
        /// </summary>
        internal const int ErrEvaluateFootingGroutPadDef = 5;

        /// <summary>
        /// Error in constructing outputs for grout pad.
        /// </summary>
        internal const int ErrEvaluatePierAndSlabFootingAssembly = 6;

        /// <summary>
        /// Error in constructing outputs for pier.
        /// </summary>
        internal const int ErrPierConstructOutputs = 7;

        /// <summary>
        /// cannot calculate weight and centre of gravity, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data.
        /// </summary>
        internal const int ErrGroutWCOGMissingSystemAttributeData = 8;

        /// <summary>
        /// Error in constructing outputs for pier and slab.
        /// </summary>
        internal const int ErrPierAndSlabConstructOutputs = 9;

        /// <summary>
        /// Footing cannot be placed using selected inputs. Delete and replace.
        /// </summary>
        internal const int ErrFootingInsufficientInputs = 10;

        /// <summary>
        /// Specify an elevation that is below the top of the Footing
        /// </summary>
        internal const int ErrFootingInvalidElevation = 11;

        /// <summary>
        /// Octagonal Grout is not supported
        /// </summary>
        internal const int ErrOctagonalGroutNotSupported = 12;

        /// <summary>
        /// Footing support plane is not valid. Software cannot compute the pier height. Delete and replace footing.
        /// </summary>
        internal const int ErrInvalidSupportingPlane = 13;

        /// <summary>
        /// Footing support plane is not valid for some supported members. Footing pier height is an invalid negative value.
        /// </summary>
        internal const int ErrInvalidPierHeightNegativeValue = 14;

        /// <summary>
        /// Unexpected error while evaluating of CombinedPierSlab Footing custom assembly definition.
        /// </summary>
        internal const int ErrEvaluateCombinedPierSlabFootingAssembly = 15;

        /// <summary>
        /// Combined footings requires at least one supported object.
        /// </summary>
        internal const int ErrCombinedFtgInsufficientNumberOfSupportedObject = 16;

        /// <summary>
        /// Pier height is invalid. Check the bottom plane and the defined grout thickness.
        /// </summary>
        internal const int ErrInvalidPierHeight = 17;

        /// <summary>
        /// Unable to get slab component.
        /// </summary>
        internal const int ErrMissingSlabComponent = 18;

        /// <summary>
        /// No outputs can be created for the footing on given inputs.
        /// </summary>
        internal const int ErrNoOutput = 19;

        /// <summary>
        /// Error when material is not found in catalog.
        /// </summary>
        internal const int ErrMaterialNotFound = 20;

        /// <summary>
        /// Error in PreConstructOutouts.
        /// </summary>
        internal const int ErrPreConstructOutputs = 21;

        /// <summary>
        /// Error in calculating weight and centre of gravity for pier, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data.
        /// </summary>
        internal const int ErrCalculatingWCOG = 22;

        /// <summary>
        /// Error while accesing shape code list value which is not found in code list table.
        /// </summary>
        internal const int ErrPierShapeCodeListValue = 23;

        /// <summary>
        /// Error while accessing  sizing rule code list value which is not found in code list table.
        /// </summary>
        internal const int ErrPierSizingRuleCodeListValue = 24;

        /// <summary>
        /// Error while accesing orientation code list value which is not found in code list table.
        /// </summary>
        internal const int ErrPierOrientationCodeListValue = 25;

        /// <summary>
        /// Error while accessing shape code list value which is not found in code list table.
        /// </summary>
        internal const int ErrSlabShapeCodeListValue = 26;

        /// <summary>
        /// Error  while accessing sizing rule code list value which is not found in code list table.
        /// </summary>
        internal const int ErrSlabSizingRuleCodeListValue = 27;

        /// <summary>
        /// Error while accesing orientation code list value which is not found in code list table.
        /// </summary>
        internal const int ErrSlabOrientationCodeListValue = 28;

        /// <summary>
        /// Error in calculating weight and centre of gravity for pier, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data.
        /// </summary>
        internal const int ErrPierAndSlabWCOGMissingSystemAttributeData = 29;

        /// <summary>
        /// Error in calculating weight and centre of gravity for pier, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data.
        /// </summary>
        internal const int ErrSlabWCOGMissingSystemAttributeData = 30;


    }
}