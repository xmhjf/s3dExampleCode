//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      PSLBOMLocalizer.cs
// Author:     PVK
//
// Abstract:
//     PSLBOM Localizer
//-----------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class PSLBOMLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, PSLBOMResourceIDs.DEFAULT_ASSEMBLY,PSLBOMResourceIDs.DEFAULT_RESOURCE);
        }
    }
}
