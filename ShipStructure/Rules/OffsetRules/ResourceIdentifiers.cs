//-------------------------------------------------------------------------------------------------------
//Copyright 2014 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ResourceIdentifiers.cs
//
//Abstract
//	ResourceIdentifiers is .NET tripping offset rules resource identifiers.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Tripping offset rules resource identifiers.
    /// </summary>
    internal static class ResourceIdentifiers
    {
        /// <summary>
        /// Tripping offset rules localizer resource name.
        /// </summary>
        internal const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Structure.Resources.OffsetRules";

        /// <summary>
        /// Tripping offset rules localizer assembly name.
        /// </summary>
        internal const string DEFAULT_ASSEMBLY = "OffsetRules";

        /// <summary>
        /// Invalid support index. 
        /// </summary>
        internal static int ErrInvalidSupportIndex = 101;
    }
}