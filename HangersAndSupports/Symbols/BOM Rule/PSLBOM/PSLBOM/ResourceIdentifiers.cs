//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      ResourceIdentifiers.cs
// Author:     PVK
//     
//
// Abstract:
//   PSLBOM resource identifiers
//-----------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// Smart Parts Symbol resource identifiers.
    /// </summary>
    public static class PSLBOMResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "PSLBOM";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "PSLBOM";
        /// <summary>
        /// Error while Length is not in the Range
        /// </summary>
        public const int ErrInvalidMinMaxLength = 1;
        /// <summary>
        /// Error while defining BOMDescription        
        /// </summary>
        public const int ErrBOMDescription = 2;
    }
}