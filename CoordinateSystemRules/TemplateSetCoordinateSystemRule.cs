using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Grids.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    
    /// <summary>
    /// TemplateSetCoordinateSystemRule class computes the Coordinate System of the given the TemplateSet.
    /// It inherits from the class "CoordinateSystemRuleBase", which contains common implementation
    /// across all name rules and provides properties and methods for use in implementation of CoordinateSystem Rules rules. 
    /// </summary>
    public class TemplateSetCoordinateSystemRule : ManufacturingCoordinateSystemRuleBase
    { 
        /// <summary>
        /// Gets the collection of coordinate systems from the model database of the requested type.
        /// </summary>
        /// <param name="Entity">TemplateSet for which the Frame System has to be computed.</param>
        /// <returns>The Collection of the Coorinate Systems in range on the TemplateSet.</returns>
        public override Collection<CoordinateSystem> GetCoordinateSystems(BusinessObject Entity)
        {   
            Collection<CoordinateSystem> coordinateSystems = null;           

            try
            {
                //To get the Model Database.
                SP3DConnection modelConnection = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel;

                if (Entity == null)
                {
                    throw new ArgumentException("Entity");
                }

                //The method "GetCoordinateSystemsFromDB" gets the collection of coordinate systems from the model database of the requested type.
                coordinateSystems = ManufacturingCoordinateSystemRuleBase.GetCoordinateSystemsFromDB(modelConnection, CoordinateSystem.CoordinateSystemType.Ship);
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