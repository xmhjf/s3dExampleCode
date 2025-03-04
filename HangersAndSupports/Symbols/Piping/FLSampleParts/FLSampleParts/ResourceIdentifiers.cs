//-----------------------------------------------------------------------------
// Copyright 1992 - 2012 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      ResourceIdentifiers.cs
// Author:     
//      Vijay    
//
// Abstract:
//     FLSampleParts resource identifiers
//-----------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// Smart Parts Symbol resource identifiers.
    /// </summary>
    public static class FLSampleSymbolResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Support.Symbols.Resources.FLSampleParts";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "FLSample";
        /// <summary>
        /// Error while constructing outputs
        /// </summary>
        public const int ErrConstructOutputs = 1;
        /// <summary>
        /// Argument Radius should not be less then Zero or equal to Zero
        /// </summary>
        public const int ErrRadiusArguments = 2;
        /// <summary>
        /// Argument L should not be less then Zero
        /// </summary>
        public const int ErrLArguments = 3;
        /// <summary>
        /// Argument C should not be less then Zero or equal to Zero
        /// </summary>
        public const int ErrCArguments = 4;
        /// <summary>
        /// Argument Hole Size should not be less then Zero or equal to Zero
        /// </summary>
        public const int ErrHoleSizeArguments = 5;
        /// <summary>
        /// Argument T should not be less then Zero or equal to Zero
        /// </summary>
        public const int ErrTArguments = 6;
        /// <summary>
        /// Argument Width should not be less then Zero or equal to Zero
        /// </summary>
        public const int ErrWidthArguments = 7;
        /// <summary>
        /// Argument Depth should not be less then Zero or equal to Zero
        /// </summary>
        public const int ErrDepthArguments = 8;
        /// <summary>
        /// Argument Thickness should not be less then Zero
        /// </summary>
        public const int ErrThicknessArguments = 9;
        /// <summary>
        /// Argument Thickness should not be Zero
        /// </summary>
        public const int ErrThicknessArgumentsNEZ = 10;
        /// <summary>
        /// Error while BOM Description
        /// </summary>
        public const int ErrBOMDescription = 11;
        /// <summary>
        /// Error while WeightCG
        /// </summary>
        public const int ErrWeightCG = 12;
        /// <summary>
        /// Argument Width cannot be Zero
        /// </summary>
        public const int ErrInvalidWidthNEZ = 13;
        /// <summary>
        /// Argument Height cannot be Zero
        /// </summary>
        public const int ErrInvalidHeightNEZ = 14;
    }
}
