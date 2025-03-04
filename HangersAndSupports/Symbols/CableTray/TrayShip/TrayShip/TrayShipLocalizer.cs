using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    class TrayShipLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, TrayShiplResourceIDs.DEFAULT_RESOURCE, TrayShiplResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
