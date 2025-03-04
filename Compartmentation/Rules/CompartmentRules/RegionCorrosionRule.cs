using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System;

namespace Ingr.SP3D.Content.Compartmentation
{
    public class RegionCorrosionRule : CompartmentComputeRuleBase
    {
        public override void Evaluate(SP3D.Compartmentation.Middle.Compartment compartment)
        {
            #region Input Error Handling
            if (compartment == null)
            {
                throw new CmnArgumentNullException("input argument is null");
            }
            #endregion
            try
            {
                compartment.SetPropertyValue(Convert.ToInt32(1234), "IJUARegionCorrosionRule", "CorrosionRule");
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }
    }
}
