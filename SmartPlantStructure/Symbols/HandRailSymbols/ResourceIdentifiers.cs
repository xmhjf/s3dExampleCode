
//=====================================================================================================
//
//Copyright 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//HandRailSymbols resource identifiers.
//
//File
//  ResourceIdentifiers.cs
//
//History:
//  Feb 16, 2015    Ninad   DI-CP-267808  Implement content changes for support of drop of Handrail  
// 
//=====================================================================================================
namespace Ingr.SP3D.Content.Structure
{

    sealed class HandRailSymbolsResourceIDs
    {
        /// <summary>
        /// HandRail symbols creation localizer resource.
        /// </summary>
        public const string DEFAULT_RESOURCE = "HandRailSymbols";

        /// <summary>
        /// HandRail symbols creation localizer assembly.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "HandRailSymbols";

        /// <summary>
        /// Section not found in catalog.
        /// </summary>
        public const int ErrSectionNotFound = 1;

        /// <summary>
        /// Some of the required user attribute values can not be obtained from the catalog part. Check the error log and catalog data.
        /// </summary>
        public const int ErrUserAttributesMissing = 3;

        /// <summary>
        /// Cannot set weight and center of gravity on the handrail.
        /// </summary>
        public const int ErrSetWeightAndCOG = 5;
       
        /// <summary>
        /// Weight and COG failed to evaluate, as required user attribute material name and grade value can not be obtained from the catalog. Check the error log and catalog data.
        /// </summary>
        public const int ErrMaterialAttributeData = 7;

        /// <summary>
        /// Weight and COG failed to evaluate, as the required material is not found in catalog. Check the error log and catalog data.
        /// </summary>
        public const int ErrMaterialNotFound = 8;

        /// <summary>
        /// Toprail section properties are missing, please check catalog.
        /// </summary>
        public const int ErrMissingToprailSectionProperty = 9;

        /// <summary>
        /// Unexpected error while constructing outputs.
        /// </summary>
        public const int ErrConstructOutputs = 10;

        /// <summary>
        /// Invalid orientation code list value.
        /// </summary>
        public const int ErrOrientationCodeListValue = 11;

        /// <summary>
        /// Invalid cross section cardinal points code list value.
        /// </summary>
        public const int ErrCrossSectionCPCodeListValue = 12;

        /// <summary>
        /// Invalid handrail offset type code list value.
        /// </summary>
        public const int ErrHandrailOffsetCodeListValue = 13;

        /// <summary>
        /// Invalid handrail treatment type code list value.
        /// </summary>
        public const int ErrHandrailTreatmentCodeListValue = 14;

        /// <summary>
        /// Invalid handrail connection type code list value.
        /// </summary>
        public const int ErrHandrailConnTypeCodeListValue = 15;

        /// <summary>
        /// Midrail section properties are missing, please check catalog.
        /// </summary>
        public const int ErrMissingMidrailSectionProperty = 16;
        
        /// <summary>
        /// Post section properties are missing, please check catalog.
        /// </summary>
        public const int ErrMissingPostSectionProperty = 17;
        
        /// <summary>
        /// ToePlate section properties are missing, please check catalog.
        /// </summary>
        public const int ErrMissingToePlateSectionProperty = 18;
        
        /// <summary>
        /// The given handrail orientation is not supported: 
        /// </summary>
        public const int ErrInvalidOrientationType = 19;
        
        /// <summary>
        /// Handrail path segment is not a line or arc. Edit handrail path and redefine all path segments to be either a line or an arc
        /// </summary>
        public const int ErrInvalidPath = 20;

        /// <summary>
        /// Invalid handrail path. The path should contain atleast one curve.
        /// </summary>
        /// <remarks></remarks>
        public const int ErrNoInputCurves = 21;

        /// <summary>
        /// Error occurred while creating members for given curve
        /// </summary>
        public const int ErrCreateMembersForGivenCurve = 22;

        /// <summary>
        /// Invalid End Treatment Type.
        /// </summary>
        public const int ErrInvalidEndTreatmentType = 23;

        /// <summary>
        /// Error occurred while adding outputs for given curve
        /// </summary>
        public const int ErrAddHandrailOutputForGivenCurve = 24;

        /// <summary>
        /// Error occurred while connecting components
        /// </summary>
        public const int ErrConnectComponents = 25;

        /// <summary>
        /// Invalid midrail cross section cardinal points code list value.
        /// </summary>
        public const int ErrMidrailCrossSectionCPCodeListValue = 26;

        /// <summary>
        /// Invalid Toe plate cross section cardinal points code list value.
        /// </summary>
        public const int ErrToePlateCrossSectionCPCodeListValue = 27;

        /// <summary>
        /// Invalid Toe Post cross section cardinal points code list value.
        /// </summary>
        public const int ErrToePostSectionCPCodeListValue = 28;

      
    }
}


