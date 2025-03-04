//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      PSLLocalizer.cs
// Author:     
//         
//
// Abstract:
//     PSL Parts Localizer
//-----------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class PSLLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, PSLSymbolResourceIDs.DEFAULT_RESOURCE, PSLSymbolResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
