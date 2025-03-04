//-----------------------------------------------------------------------------
// Copyright 1992 - 2012 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      ResourceIdentifiers.cs
// Author:     
//      Ramya Pandala     
//
// Abstract:
//     RichHgrBeam Symbol resource identifiers
//-----------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// RichHgrBeam Symbol resource identifiers.
    /// </summary>
    public static class HangerBeamSymbolResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Support.Symbols.Resources.HgrStructural";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "HgrStructural";
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
        /// Invalid input arguments.
        /// </summary>
        public const int ErrInvalidArguments = 4;
        /// <summary>
        /// Invalid HgrBeam input arguments.
        /// </summary>
        public const int ErrInvalidHgrBeamArguments = 5;
        /// <summary>
        /// Invalid CutBackSteel input arguments.
        /// </summary>
        public const int ErrInvalidCutBackSteelArguments = 6;
        /// <summary>
        /// Invalid Snip Steel input arguments.
        /// </summary>
        public const int ErrInvalidSnipSteelArguments = 7;
        /// <summary>
        /// Error while getting Material Type or Material Grade
        /// </summary>
        public const int ErrGettingMaterial = 8;
        /// <summary>
        /// CrossSection not found in Catalog
        /// </summary>
        public const int ErrCrossSectionNotFound = 9;
        /// <summary>
        /// Section not found in Catalog
        /// </summary>
        public const int ErrSectionNotFound = 10;
        /// <summary>
        /// Error in Create Connection Component Ports method for HgrBeamInputs
        /// </summary>
        public const int ErrCreateConnectionComponentPortsHgrBeam = 11;
        /// <summary>
        /// Error in Create Connection Component Ports method for CutBackSteelInputs
        /// </summary>
        public const int ErrCreateConnectionComponentPortsCutBackSteel = 12;
        /// <summary>
        /// Error in Create Connection Component Ports method for SnipSteelInputs
        /// </summary>
        public const int ErrCreateConnectionComponentPortsSnipSteel = 13;
        /// <summary>
        /// Error while setting the Orientation for BeginCap Port
        /// </summary>
        public const int ErrBeginCapPortOrientation = 14;
        /// <summary>
        /// Error while setting the Orientation for EndCap Port
        /// </summary>
        public const int ErrEndCapPortOrientation = 15;
        /// <summary>
        /// Error while setting the Orientation for BeginFace Port
        /// </summary>
        public const int ErrBeginFacePortOrientation = 16;
        /// <summary>
        /// Error while setting the Orientation for EndFace Port
        /// </summary>
        public const int ErrEndFacePortOrientation = 17;
        /// <summary>
        /// Error while setting the Orientation for Neutral Port
        /// </summary>
        public const int ErrNeutralPortOrientation = 18;
        /// <summary>
        /// Error while setting the Orientation for BeginFlex Port
        /// </summary>
        public const int ErrBeginFlexPortOrientation = 19;
        /// <summary>
        /// Error while setting the Orientation for EndFlex Port
        /// </summary>
        public const int ErrEndFlexPortOrientation = 20;
        /// <summary>
        /// Error while creating BeginCap Port
        /// </summary>
        public const int ErrCreateBeginCapPort = 21;
        /// <summary>
        /// Error while creating EndCap Port
        /// </summary>
        public const int ErrCreateEndCapPort = 22;
        /// <summary>
        /// Error while creating BeginFace Port
        /// </summary>
        public const int ErrCreateBeginFacePort = 23;
        /// <summary>
        /// Error while creating EndFace Port
        /// </summary>
        public const int ErrCreateEndFacePort = 24;
        /// <summary>
        /// Error while creating Neutral Port
        /// </summary>
        public const int ErrCreateNeutralPort = 25;
        /// <summary>
        /// Error while creating BeginFlex Port
        /// </summary>
        public const int ErrCreateBeginFlexPort = 26;
        /// <summary>
        /// Error while creating EndFlex Port
        /// </summary>
        public const int ErrCreateEndFlexPort = 27;
        /// <summary>
        /// Error while Snip functionality is supported only for L section
        /// </summary>
        public const int ErrSnipLSectionMaterial = 28;
        /// <summary>
        /// Error while Cardinal Points range exceeds between 1 to 15
        /// </summary>
        public const int ErrCardinalPnts = 29;
    }
}
