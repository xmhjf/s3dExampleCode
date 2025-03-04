//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      PSLLocalizer.cs
// Author:     
//        Hema 
//
// Abstract:
//     PTPBOM Localizer
//-----------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class PTPBOMLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, PTPBOMResourceIDs.DEFAULT_ASSEMBLY, PTPBOMResourceIDs.DEFAULT_RESOURCE);
        }
    }
}
