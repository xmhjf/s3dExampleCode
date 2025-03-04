using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Grids.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    
    /// <summary>
    /// ShrinkageCoordinateSystemRule class computes the Coordinate System of the given the Shrinkage.
    /// It inherits from the class "CoordinateSystemRuleBase", which contains common implementation
    /// across all name rules and provides properties and methods for use in implementation of CoordinateSystem Rules rules. 
    /// </summary>
    public class ShrinkageCoordinateSystemRule : ManufacturingCoordinateSystemRuleBase
    {
        /// <summary>
        /// Computes the Frames in range.
        /// </summary>
        /// <param name="Entity">Shrinkage object for which the Frame System has to be computed.</param>
        /// <returns>The Collection of the Coorinate Systems in range on the Shrinkage.</returns>
        public override Collection<CoordinateSystem> GetCoordinateSystems(BusinessObject Entity)
        {
            Collection<CoordinateSystem> coordinateSystems = null;
            try
            {
                //To get the Model Database.                
                SiteManager siteMgr = MiddleServiceProvider.SiteMgr;
                Plant activePlant = siteMgr.ActiveSite.ActivePlant;
                Model modelDB = activePlant.PlantModel;
                BusinessObject rootobj = (BusinessObject)modelDB.RootSystem;
                SP3DConnection modelConnection = rootobj.DBConnection;

                if (Entity == null)
                {
                    throw new ArgumentException("Entity");
                }

                //The method "GetCoordinateSystemsInRange" gets the collection of coordinate systems from the model database of the requested type.
                coordinateSystems = ManufacturingCoordinateSystemRuleBase.GetCoordinateSystemsFromDB(modelConnection, CoordinateSystem.CoordinateSystemType.Ship);
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame stackFrame = new System.Diagnostics.StackFrame(true);
                string methodName = stackFrame.GetMethod().Name;
                int lineNumber = stackFrame.GetFileLineNumber();

                //If the rule falis to get the Frame systems, it logs the error information into ErrorLog.
                base.WriteToErrorLog(e, "SMCustomWarningMessages",3002,"Rules");

            }
            return coordinateSystems;
        }
    }
}