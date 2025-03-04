using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections;

using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// 
    /// </summary>
    public class BuiltUpProfileStandardRule : ShrinkageParameterRuleBase
    {

        /// <summary>
        /// Gets the shrinkage parameters for a plate depending on the plate orientation.
        /// </summary>
        /// <param name="shrinkageInformation">shrinkage information acts as input output carrier.</param>
        public override ShrinkageParameters GetParameters(ShrinkageEntityInformation shrinkageInformation)
        {
            return null;

        }


        /// <summary>
        /// Gets the dependent stiffener's shrinkage parameters from its parent plate part.
        /// </summary>
        /// <param name="plateShrinkageInformation">Input plate shrinkage information.</param>              
        public override ShrinkageParameters GetParameters(ShrinkageConnectedEntityInformation plateShrinkageInformation)
        {
            return null;

        }

        /// <summary>
        /// Gets the collection of stiffner systems on the plate part.
        /// </summary>
        /// <param name="shrinkageInformation">shrinkage information acts as input output carrier.</param>
        public override ReadOnlyCollection<BusinessObject> GetConnectedEntities(ShrinkageEntityInformation shrinkageInformation)
        {
            return null;
        }          
       
    }
}
