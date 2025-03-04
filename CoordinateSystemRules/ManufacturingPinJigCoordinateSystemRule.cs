

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Grids.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{    
    /// <summary>
    /// PinJigCoordinateSystemRule class computes the Coordinate System in range for the given PinJig.
    /// It inherits from the class "CoordinateSystemRuleBase", which contains common implementation
    /// across all name rules and provides properties and methods for use in implementation of CoordinateSystem Rules rules. 
    /// </summary>
    public class PinJigCoordinateSystemRule : ManufacturingCoordinateSystemRuleBase
    {       

        /// <summary>
        /// Gets the collection of coordinate systems in the range of the entity.
        /// </summary>
        /// <param name="Entity">PinJig for which the Frame System has to be computed.</param>
        /// <returns>The Collection of the Coorinate Systems in range on the PinJig.</returns>
        public override Collection<CoordinateSystem> GetCoordinateSystems(BusinessObject Entity)
        {
            Collection<CoordinateSystem> coordinateSystems = null;
            try
            {
                if (Entity == null)
                {
                    throw new ArgumentException("Entity");
                }

                //The method "GetCoordinateSystemsInRange" gets all the coordinate systems within a range given the PinJig.
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