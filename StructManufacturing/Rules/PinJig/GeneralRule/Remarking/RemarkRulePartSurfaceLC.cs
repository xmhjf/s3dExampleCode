//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Part Surface with Logical Connection based PinJig Remarking Rule. It provides the following :
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

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Grids.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Part Surface with Logical Connection based PinJig Remarking Rule.
    /// </summary>
    public class RemarkRulePartSurfaceLC : PinJigRemarkRuleBase
    {
        /// <summary>
        /// Returns the remarking surface of the PinJig.
        /// </summary>
        /// <param name="pinJigInformation">The pinjig info.</param>
        public override ISurface GetRemarkingSurface(PinJigInformation pinJigInformation)
        {
            ISurface remarkingSurface = null;
            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                PinJig mfgPinJig = null;
                if (pinJigInformation.ManufacturingPart != null)
                {
                    mfgPinJig = (PinJig)pinJigInformation.ManufacturingPart;
                }

                if (mfgPinJig == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing
                
                // Get Remarking Surface.
                //It is the surface opposite to base plane which is on the remarking sides of the supported plate parts.

                int offsetType = Convert.ToInt32(pinJigInformation.GetArguments("OffsetType", "RemarkingSurface").FirstOrDefault().Value);
                int referenceSurfaceType = Convert.ToInt32(pinJigInformation.GetArguments("ReferenceSurfaceType", "RemarkingSurface").FirstOrDefault().Value);

                if (base.DoPlatePartRemarkingSurfacesHaveGaps(mfgPinJig) == true)
                {
                    // If there are gaps between the remarking surfaces of the pinjig supported plate parts, then get remarking surface as below:
                    // 1. Get offset from PlateSystem based on average of the weighted thickness.
                    // 2. Get surface based on this offset from PlateSystem.

                    offsetType = Convert.ToInt32(OffsetType.WeightedThicknessAverage);
                    referenceSurfaceType = Convert.ToInt32(ReferenceSurfaceType.PlatePartSupportedSurface);
                }

                remarkingSurface = base.GetOffsetSurface(mfgPinJig, (OffsetType)offsetType, 0.0, (ReferenceSurfaceType) referenceSurfaceType);
  
                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e,"SMCustomWarningMessages", 5022, "Call to PinJig Remarking Surface Rule failed with the error" + e.Message);
            }

            return remarkingSurface;

        }

        // User can override the other methods of the base class to have different implementation.  
    }
}
