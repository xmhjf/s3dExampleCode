using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections.ObjectModel;
//using Ingr.SP3D.Structure.Middle;
using System.Collections;

using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Grids.Middle;
using Ingr.SP3D.Grids.Middle.Services;
using Ingr.SP3D.Common.Exceptions;

namespace Ingr.SP3D.Content.Manufacturing
{
    public class PartRule : ShrinkageParameterRuleBase
    {

        /// <summary>
        /// Gets the shrinkage parameters for a plate depending on the plate orientation.
        /// </summary>
        /// <param name="shrinkageInformation">shrinkage information acts as input output carrier.</param>
        public override ShrinkageParameters GetParameters(ShrinkageEntityInformation shrinkageInformation)
        {
            if (shrinkageInformation == null)
            {
                throw new CmnNullArgumentException("shrinkageInformation");
            }
            return new ShrinkageParameters(ShrinkageType.ByAxis,null,0.2,null,0.4,null,0);
        }


        /// <summary>
        /// Gets the dependent stiffener's shrinkage parameters from its parent part shrinkage parameters.
        /// </summary>
        /// <param name="plateShrinkageInformation">Input plate shrinkage information.</param>              
        public override ShrinkageParameters GetParameters(ShrinkageConnectedEntityInformation plateShrinkageInformation)
        {
            if (plateShrinkageInformation == null)
            {
                throw new CmnNullArgumentException("plateShrinkageInformation");
            }

            ShrinkageParameters outParams = new ShrinkageParameters(ShrinkageType.ByAxis,null,0.2,null,0.4,null,0);
            return outParams;
        }

        /// <summary>
        /// Gets the collection of stiffner systems on the plate part.
        /// </summary>
        /// <param name="shrinkageInformation">shrinkage information acts as input output carrier.</param>
        public override ReadOnlyCollection<BusinessObject> GetConnectedEntities(ShrinkageEntityInformation shrinkageInformation)
        {
            PlatePartBase plate = shrinkageInformation.ManufacturingParent as PlatePartBase;
            ReadOnlyCollection<BusinessObject> connectedObj = null;
            if (plate != null)
            {
                connectedObj = plate.GetConnectedObjects();
            }            

            Collection<BusinessObject> stiffenerSystemsColl = new Collection<BusinessObject>();
            if (connectedObj != null && connectedObj.Count  > 0)
            {
                for( int i = 0; i < connectedObj.Count; i++)
                {
                    
                    if( connectedObj[i] is StiffenerPartBase)
                    {
                        StiffenerPartBase stiffBase = connectedObj[i] as StiffenerPartBase;
                        if (stiffBase != null)
                        {
                            StiffenerSystemBase rootStiffnerSystem = stiffBase.RootStiffenerSystem;
                            stiffenerSystemsColl.Add(rootStiffnerSystem);
                        }
                    }
                }

            }
            return new ReadOnlyCollection<BusinessObject>(stiffenerSystemsColl);
        }       

    }
}
