//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      JIMCBOMLocalizer.cs
// Author:     
//        Hema 
//
// Abstract:
//     JIMCBOM Localizer
//-----------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class JIMCBOMLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, JIMCBOMResourceIDs.DEFAULT_ASSEMBLY, JIMCBOMResourceIDs.DEFAULT_RESOURCE);
        }
    }
}
