using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Compartmentation.Middle;
using System;

namespace Ingr.SP3D.Content.Compartmentation
{
    public class CompartmentStructure : CompartmentComputeRuleBase
    {
        public override void Evaluate(Compartment compartEntity)
        {
            #region Input Error Handling
            if (compartEntity == null)
            {
                throw new CmnArgumentNullException("input argument is null");
            }
            #endregion

            try
            {
                double propertyValue = GetVolume(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "VolumeMoulded");

                propertyValue = GetWallArea(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "WallArea");

                propertyValue = GetWallLength(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "WallLength");

                propertyValue = GetDeckHeight(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "DeckHeight");

                propertyValue = GetSideWallArea(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "SideWallArea");

                propertyValue = GetBottomArea(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "BottomArea");

                propertyValue = GetBottomCGX(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "BottomCGX");

                propertyValue = GetBottomCGY(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "BottomXGY");

                propertyValue = GetBottomCGZ(compartEntity);
                compartEntity.SetPropertyValue(propertyValue, "IJUACompartStructure", "BottomXGZ");
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        private double GetVolume(Compartment compartEntity)
        {
            return 100;
        }

        /// <summary>
        /// returns the wall area related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetWallArea(Compartment compartEntity)
        {
            double wallArea = 0.0;
            
            try
            {
                PropertyValueDouble propValueDouble = compartEntity.GetPropertyValue("IJCompartAttributes", "SurfaceArea") as PropertyValueDouble;

                if (propValueDouble != null && propValueDouble.PropValue != null)
                {
                    wallArea = (double)propValueDouble.PropValue;
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return wallArea;
        }

        // <summary>
        /// returns the wall length related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetWallLength(Compartment compartEntity)
        {
            double wallLangth = 100.0;
            
            return wallLangth;
        }

        /// <summary>
        /// returns the deck height related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetDeckHeight(Compartment compartEntity)
        {
            double deckHeight = 100.0;
           
            return deckHeight;
        }

        /// <summary>
        /// returns the Side wall area related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetSideWallArea(Compartment compartEntity)
        {
            double sideWallArea = 0.0;
            
            try
            {
                PropertyValueDouble propValueDouble = compartEntity.GetPropertyValue("IJCompartAttributes", "SurfaceArea") as PropertyValueDouble;

                if (propValueDouble != null && propValueDouble.PropValue != null)
                {
                    sideWallArea = (double)propValueDouble.PropValue;
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return sideWallArea;
        }

        /// <summary>
        /// returns the bottom area related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetBottomArea(Compartment compartEntity)
        {
            double bottomArea = 0.0;

            try
            {
                PropertyValueDouble propValueDouble = compartEntity.GetPropertyValue("IJCompartAttributes", "SurfaceArea") as PropertyValueDouble;

                if (propValueDouble != null && propValueDouble.PropValue != null)
                {
                    bottomArea = (double)propValueDouble.PropValue;
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return bottomArea;
        }

        /// <summary>
        /// returns the bottom center of gravity of x related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetBottomCGX(Compartment compartEntity)
        {
            double bottomCGX = 0.0;

            try
            {
                PropertyValueDouble propValueDouble = compartEntity.GetPropertyValue("IJCompartAttributes", "CogX") as PropertyValueDouble;

                if (propValueDouble != null && propValueDouble.PropValue != null)
                {
                    bottomCGX = (double)propValueDouble.PropValue;
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return bottomCGX;
        }

        /// <summary>
        /// returns the bottom center of gravity of y related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetBottomCGY(Compartment compartEntity)
        {
            double bottomCGY = 0.0;
            
            try
            {
                PropertyValueDouble propValueDouble = compartEntity.GetPropertyValue("IJCompartAttributes", "CogY") as PropertyValueDouble;

                if (propValueDouble != null && propValueDouble.PropValue != null)
                {
                    bottomCGY = (double)propValueDouble.PropValue;
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return bottomCGY;
        }

        /// <summary>
        /// returns the bottom center of gravity of z related to the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private double GetBottomCGZ(Compartment compartEntity)
        {
            double bottomCGZ = 0.0;
            
            try
            {
                PropertyValueDouble propValueDouble = compartEntity.GetPropertyValue("IJCompartAttributes", "CogZ") as PropertyValueDouble;

                if (propValueDouble != null && propValueDouble.PropValue != null)
                {
                    bottomCGZ = (double)propValueDouble.PropValue;
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return bottomCGZ;
        }
    }
}
