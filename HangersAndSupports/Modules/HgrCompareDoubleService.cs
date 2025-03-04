//-----------------------------------------------------------------------------
// Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
// File
//      HgrCompareDoubleService.cs
// Author:     
//      Vinay     
//  11.Dec.2014   PVK  TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//
// Abstract:
//     

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public static class HgrCompareDoubleService
    {
        /// <summary>
        /// Compare the double vaules
        /// </summary>
        /// <param name="leftvaule">The vaule to be compared.</param>
        /// <param name="right value">The vaule to be compared against.</param>
        public static bool cmpdbl(double leftvaule, double rightvalue)
        {
            bool result = true;
            if (leftvaule > rightvalue - 0.00001 && leftvaule < rightvalue + 0.00001)
                result = true;
            else
                result = false;
            return result;
        }
    }
}