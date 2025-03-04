using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ingr.SP3D.Content.Support.Symbols
{
    /// <summary>
    /// RichHgrBeam Symbol resource identifiers.
    /// </summary>
    public static class SmartStellResourceIDs
    {
        /// <summary>
        /// Default Resource Name.
        /// </summary>   
        public const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Support.Symbols.Resources.SmartSteel";
        /// <summary>
        /// Default Assembly Name.
        /// </summary>
        public const string DEFAULT_ASSEMBLY = "SmartSteel";
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
        /// CrossSection not found in Catalog
        /// </summary>
        public const int ErrCrossSectionNotFound = 4;
    }
}
