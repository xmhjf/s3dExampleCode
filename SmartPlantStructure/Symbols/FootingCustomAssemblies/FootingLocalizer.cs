//-------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FootingLocalizer.cs
//
//Abstract
//	FootingLocalizer is Footings custom assemblies resource localizer.
//-------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Structure
{
    internal static class FootingLocalizer
    {
        /// <summary>
        /// Gets the localized string.
        /// If the string matches with the string in resource file, appropriate message from resource file is fetched else the default message.
        /// </summary>
        /// <param name="resourceId">The resourceID.</param>
        /// <param name="defaultString">The default string.</param>
        /// <returns>The localized string.</returns>
        internal static string GetString(int resourceId, string defaultString)
        {
            return CmnLocalizer.GetString(resourceId, defaultString, FootingResourceIDs.DEFAULT_RESOURCE, FootingResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}