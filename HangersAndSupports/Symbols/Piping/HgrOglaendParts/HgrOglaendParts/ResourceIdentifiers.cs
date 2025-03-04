//-----------------------------------------------------------------------------
// Copyright 1992 - 2012 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      ResourceIdentifiers.cs
// Author:     
//      Ramya Pandala     
//
// Abstract:
//     Oglaend Parts resource identifiers
//-----------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// Oglaend Parts Symbol resource identifiers.
    /// </summary>
    public static class OglaendPartSymbolResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Support.Symbols.Resources.HgrOglaendParts";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "HgrOglaendParts";
        /// <summary>
        /// Invalid input arguments.
        /// </summary>
        public const int ErrInvalidArguments = 1;
        /// <summary>
        /// Error while constructing outputs
        /// </summary>
        public const int ErrConstructOutputs = 2;
        /// <summary>
        /// Error in getting Cross Section
        /// </summary>
        public const int ErrSectionNotFound = 3;
        /// <summary>
        /// Error in getting the Material
        /// </summary>
        public const int ErrGettingMaterial = 4;
        /// <summary>
        /// Error in Connection Component Ports creation
        /// </summary>
        public const int ErrCreateConnectionComponentPorts = 5;
        /// <summary>
        /// Error in Orientation of Port
        /// </summary>
        public const int ErrPortOrientation = 6;
        /// <summary>
        /// Error while Ceating Ports
        /// </summary>
        public const int ErrCreatePorts = 7;
    }
}