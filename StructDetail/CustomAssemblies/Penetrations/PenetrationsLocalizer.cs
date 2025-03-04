//-------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  PenetrationsLocalizer.cs
//
//Abstract
//	PenetrationsLocalizer is Penetrations custom assemblies resource localizer.
//-------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Penetrations custom assemblies localizer.
    /// </summary>
    internal static class PenetrationsLocalizer
    {
        /// <summary>
        /// Gets the localized string.
        /// </summary>
        /// <param name="resourceId">The resourceID.</param>
        /// <param name="defaultString">The default string.</param>
        /// <returns>The localized string.</returns>
        internal static string GetString(int resourceId, string defaultString)
        {
            return CmnLocalizer.GetString(resourceId, defaultString, PenetrationsResourceIds.DEFAULT_RESOURCE, PenetrationsResourceIds.DEFAULT_ASSEMBLY);
        }
    }
}
