using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class FINL_PartsLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, FINL_PartsSymbolResourceIDs.DEFAULT_RESOURCE, FINL_PartsSymbolResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
