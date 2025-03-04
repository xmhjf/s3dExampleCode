using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Grids.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    
    /// <summary>
    /// ManufacturingPlateCoordinateSystemRule class computes the Coordinate System in range, given the ManufacturingPlate.
    /// It inherits from the class "CoordinateSystemRuleBase", which contains common implementation
    /// across all name rules and provides properties and methods for use in implementation of CoordinateSystem Rules rules. 
    /// </summary>
    public class ManufacturingPlateCoordinateSystemRule : ManufacturingCoordinateSystemRuleBase
    {
        Collection<CoordinateSystem> coordinateSystems = null;

        /// <summary>
        /// Gets the collection of coordinate systems in the range of the entity.
        /// </summary>
        /// <param name="Entity">Manufacturing Profile for which the Frame System has to be computed.</param>
        /// <returns>The Collection of the Coorinate Systems in range on the Manufacturing Profile.</returns> 
        public override Collection<CoordinateSystem> GetCoordinateSystems(BusinessObject Entity)
        {            
            try
            {
                if (Entity == null)
                {
                    throw new ArgumentException("Entity");
                }
              
                //The method "GetCoordinateSystemsInRange" gets all the coordinate systems within a range given the Manufacturing Profile.
                coordinateSystems = ManufacturingCoordinateSystemRuleBase.GetCoordinateSystemsInRange(Entity);               

            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame stackFrame = new System.Diagnostics.StackFrame(true);
                string methodName = stackFrame.GetMethod().Name;
                int lineNumber = stackFrame.GetFileLineNumber();

                //If the rule falis to get the Frame systems, it logs the error information into ErrorLog.
                base.WriteToErrorLog(e, "SMCustomWarningMessages", 3002, "Rules");

            }
            return coordinateSystems;
        }
      
    }
}
