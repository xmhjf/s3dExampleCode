﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class GenericLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, GenericResourceIdentifiers.DEFAULT_RESOURCE, GenericResourceIdentifiers.DEFAULT_ASSEMBLY);
        }
    }
}