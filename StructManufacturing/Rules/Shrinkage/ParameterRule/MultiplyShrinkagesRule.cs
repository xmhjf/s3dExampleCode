using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections;

using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;


namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// 
    /// </summary>
    public class MultiplyShrinkagesRule : ShrinkageAccumulationRuleBase
    {
        /// <summary>
        /// Gets the acumulated shrinkage factors for an assembly part.
        /// </summary>
        /// <param name="acumulatedShrinkageInformation">Input acumulatedShrinkageInformation object.</param>
        public override void GetFactors(ShrinkageAccumulationInformation acumulatedShrinkageInformation, out double parimaryfactor, out double secondaryfactor)
        {
            parimaryfactor = 0.0;
            secondaryfactor = 0.0;

            parimaryfactor = 0.0;
            secondaryfactor = 0.0;

            if (acumulatedShrinkageInformation != null)
            {
                double primaryfactor1 = acumulatedShrinkageInformation.PartAssemblyShrinkage.PrimaryFactor;
                double primaryfactor2 = acumulatedShrinkageInformation.PartPrivateShrinkage.PrimaryFactor;
                parimaryfactor = (primaryfactor1 + primaryfactor2);


                double secondaryfactor1 = acumulatedShrinkageInformation.PartAssemblyShrinkage.SecondaryFactor;
                double secondaaryfactor2 = acumulatedShrinkageInformation.PartPrivateShrinkage.SecondaryFactor;
                secondaryfactor = (secondaryfactor1 + secondaaryfactor2);
            }

        }

    }
}

