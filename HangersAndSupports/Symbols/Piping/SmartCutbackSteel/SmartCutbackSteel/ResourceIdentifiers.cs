//-----------------------------------------------------------------------------
// Copyright 1992 - 2012 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      ResourceIdentifiers.cs
// Author:     
//      Vijay    
//
// Abstract:
//     Smart Parts resource identifiers
//-----------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// Smart Parts Symbol resource identifiers.
    /// </summary>
    public static class SmartCutbackSteelSymbolResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Support.Symbols.Resources.SmartCutbackSteel";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "SmartCutbackSteel";
        /// <summary>
        /// Error while constructing outputs
        /// </summary>
        public const int ErrConstructOutputs = 1;
        /// <summary>
        /// Error while BOM Description
        /// </summary>
        public const int ErrBOMDescription = 2;
        /// <summary>
        /// Error while WeightCG
        /// </summary>
        public const int ErrWeightCG = 3;
        /// <summary>
        /// Error while Cardinal Points range exceeds between 1 to 15
        /// </summary>
        public const int ErrCardinalPnts = 4;
        /// <summary>
        /// FlangeThickness Should not be Zero
        /// </summary>
        public const int ErrFlangeThicknessArgument = 5;
        /// <summary>
        /// WebThickness Should not be Zero
        /// </summary>
        public const int ErrWebThicknessArgument = 6;
        /// <summary>
        /// Width Should not be Zero
        /// </summary>
        public const int ErrWidthArgument = 7;
        /// <summary>
        /// NomThickness Should not be Zero
        /// </summary>
        public const int ErrCNomThickArgument = 8;
        /// <summary>
        /// CrossSection not found in Catalog
        /// </summary>
        public const int ErrCrossSectionNotFound = 9;
    }
}
