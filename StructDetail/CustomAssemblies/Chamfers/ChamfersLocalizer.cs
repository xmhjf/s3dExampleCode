//-------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ChamferLocalizer.cs
//
//Abstract
//	ChamfersLocalizer is Chamfer custom assemblies resource localizer.
//-------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Chamfers custom assemblies localizer.
    /// </summary>
    internal static class ChamfersLocalizer
    {
        /// <summary>
        /// Gets the localized string.
        /// </summary>
        /// <param name="resourceId">The resourceID.</param>
        /// <param name="defaultString">The default string.</param>
        /// <returns>The localized string.</returns>
        internal static string GetString(int resourceId, string defaultString)
        {
            return CmnLocalizer.GetString(resourceId, defaultString, ChamfersResourceIds.DEFAULT_RESOURCE, ChamfersResourceIds.DEFAULT_ASSEMBLY);
        }
    }
}