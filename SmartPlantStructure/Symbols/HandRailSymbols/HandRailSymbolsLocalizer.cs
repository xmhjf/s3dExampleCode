//===========================================================================
//
//Copyright 2010 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//HandRail symbols localizer
//
//File
//  HandRailSymbolsLocalizer.cs
//
//History:
//  Feb 16, 2015    Ninad   DI-CP-267808  Implement content changes for support of drop of Handrail  
// 
//===========================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Structure
{

    sealed class HandRailSymbolsLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, HandRailSymbolsResourceIDs.DEFAULT_RESOURCE, HandRailSymbolsResourceIDs.DEFAULT_ASSEMBLY);
        }
    }

}
