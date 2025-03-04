using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Compartmentation.Middle;
using System;

namespace Ingr.SP3D.Content.Compartmentation
{
    public class CompartmentRefer : CompartmentComputeRuleBase
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
                ReferencePosition referencePosition = GetReferencePosition(compartment);                
                compartment.SetPropertyValue((int)referencePosition, "IJCompartReference", "ReferencePosition");
            }
            catch (CmnException e)
            {

                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        /// <summary>
        /// Gets the Reference position of the given compartment
        /// </summary>
        /// <param name="compartment"></param>
        /// <returns></returns>
        private ReferencePosition GetReferencePosition(Compartment compartment)
        {
            ReferencePosition RefPos = ReferencePosition.Undefined;

            try
            {
                RangeBox rangeBox = compartment.Range;

                if (rangeBox.Low.Y >= 0 && rangeBox.High.Y >= 0)
                {
                    RefPos = ReferencePosition.PortSide;
                }
                else if (rangeBox.Low.Y <= 0 && rangeBox.High.Y <= 0)
                {
                    RefPos = ReferencePosition.StarBoard;
                }
                else if (rangeBox.Low.X <= 0 && rangeBox.High.X <= 0)
                {
                    RefPos = ReferencePosition.Aft;
                }
                else if (rangeBox.Low.X >= 0 && rangeBox.High.X >= 0)
                {
                    RefPos = ReferencePosition.Fore;
                }
                else if (rangeBox.Low.Z <= 0 && rangeBox.High.Z <= 0)
                {
                    RefPos = ReferencePosition.Below;
                }
                else if (rangeBox.Low.Z >= 0 && rangeBox.High.Z >= 0)
                {
                    RefPos = ReferencePosition.Above;
                }
            }
            catch (CmnException e)
            {

                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }

            return RefPos;
        }
    }
}
