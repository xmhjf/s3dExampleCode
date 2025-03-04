//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: PinJig Remarking Rule. It provides the following :
//               - Remarking Surface
//               - Entities and Geometries that create a Remarking Line of the PinJig of given type.
//               - Remarking types that satisfy a particular purpose.
//               - Filter criteria for the PinJig Remarking step based on the remarking type.
//
//      Author:  Suma Mallena
//
//      History:
//      March 1st, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// PinJig Remarking Rule.
    /// </summary>
    public class RemarkRule : PinJigRemarkRuleBase
    {
        // User can override the methods of the base class to have different implementation.  
    }
}
