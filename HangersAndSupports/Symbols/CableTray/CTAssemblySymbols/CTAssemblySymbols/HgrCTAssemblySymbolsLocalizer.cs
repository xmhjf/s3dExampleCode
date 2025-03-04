//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      HgrCTAssemblySymbolsLocalizer.cs
// Author:     
//         Vijay
//
// Abstract:
//     PSL Parts Localizer
//-----------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class HgrCTAssemblySymbolsLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, HgrCTAssemblySymbolsResourceIDs.DEFAULT_RESOURCE, HgrCTAssemblySymbolsResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
