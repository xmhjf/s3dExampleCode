using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    class Util_MetricLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, Util_MetricSymbolsResourceIDs.DEFAULT_RESOURCE, Util_MetricSymbolsResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
