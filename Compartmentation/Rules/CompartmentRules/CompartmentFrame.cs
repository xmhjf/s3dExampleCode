using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Compartmentation.Middle;
using Ingr.SP3D.Grids.Middle;
using System;

namespace Ingr.SP3D.Content.Compartmentation
{
    public class CompartmentFrame : CompartmentComputeRuleBase
    {
        public override void Evaluate(Compartment compartment)
        {
            #region Input Error Handling
            if (compartment == null)
            {
                throw new CmnArgumentNullException("input argument is null");
            }
            #endregion

            try
            {
                string propertyValue = GetLowBoundingFrameName(compartment, AxisType.Y);
                compartment.SetPropertyValue(propertyValue, "IJUACompartFrame", "LongitudinalMin");

                propertyValue = GetHighBoundingFrameName(compartment, AxisType.Y);
                compartment.SetPropertyValue(propertyValue, "IJUACompartFrame", "LongitudinalMax");

                propertyValue = GetLowBoundingFrameName(compartment, AxisType.Z);
                compartment.SetPropertyValue(propertyValue, "IJUACompartFrame", "DeckMin");

                propertyValue = GetHighBoundingFrameName(compartment, AxisType.Z);
                compartment.SetPropertyValue(propertyValue, "IJUACompartFrame", "DeckMax");

                propertyValue = GetLowBoundingFrameName(compartment, AxisType.X);
                compartment.SetPropertyValue(propertyValue, "IJUACompartFrame", "TransversalMin");

                propertyValue = GetHighBoundingFrameName(compartment, AxisType.X);
                compartment.SetPropertyValue(propertyValue, "IJUACompartFrame", "TransversalMax");
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }
    }
}
