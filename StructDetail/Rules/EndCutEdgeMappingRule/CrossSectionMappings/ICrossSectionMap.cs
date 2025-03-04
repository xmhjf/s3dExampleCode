//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  ICrossSectionMap interface to be implemented by all classes
//                  that are returning a cross section map.
//
//      Author:  3XCalibur
//
//      History:
//      November 03, 2010       WR                  Created
//
//      October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) 'GetCrossSectionMap' method is added with new parameter 'pointMap'.
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure.CrossSectionMappings
{
    /// <summary>
    /// Implemented by the cross section map classes.
    /// </summary>
    interface ICrossSectionMap
    {
        /// <summary>
        /// Gets the cross section map.
        /// </summary>
        /// <param name="flipLeftAndRight">if set to <c>true</c> [is mirrored].</param>
        /// <param name="quadrant">The quadrant.</param>
        /// <returns></returns>
        Dictionary<SectionFaceType, SectionFaceType> GetCrossSectionMap(bool flipLeftAndRight, int quadrant, Dictionary<int, int> pointMap);
    }
}
