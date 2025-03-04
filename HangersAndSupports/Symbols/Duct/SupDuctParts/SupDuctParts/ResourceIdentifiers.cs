//-----------------------------------------------------------------------------
// Copyright 1992 - 2012 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      ResourceIdentifiers.cs
// Author:     
//      Sridhar Bathina     
//
// Abstract:
//     SupDuctParts resource identifiers
//-----------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// SupDuctParts Symbol resource identifiers.
    /// </summary>
    public static class SupDuctPartSymbolResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Support.Symbols.Resources.SupDuctParts";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "SupDuctParts";
        /// <summary>
        /// Invalid input arguments.
        /// </summary>
        public const int ErrInvalidArguments = 1;
        /// <summary>
        /// Error while constructing outputs
        /// </summary>
        public const int ErrConstructOutputs = 2;
        /// <summary>
        /// Error while creating the Ports
        /// </summary>
        public const int ErrCreatePorts = 3;
        /// <summary>
        /// Error while setting the Port Orientation
        /// </summary>
        public const int ErrPortOrientation = 4;
        /// <summary>
        /// Invalid width.
        /// </summary>
        public const int ErrInvalidWidthNZero = 5;
        /// <summary>
        /// Invalid thickness.
        /// </summary>
        public const int ErrInvalidThicknessNZero = 6;
        
    }
}
