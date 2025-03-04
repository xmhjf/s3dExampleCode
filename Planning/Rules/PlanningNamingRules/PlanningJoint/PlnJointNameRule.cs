using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Planning
{
    public class PlnJointNameRule : NameRuleBase
    {
        private const string strCountFormat = "{0:0000000}";

        public override void ComputeName(BusinessObject oEntity, System.Collections.ObjectModel.ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                long count;
                string locationId, entityName;
                string[] delimiter = {" "}; 
                string plnJointTempName = string.Empty;
                string phyConnTempName = string.Empty;

                PlanningJoint plnJoint = oEntity as PlanningJoint;
                if (plnJoint != null)
                {
                    plnJointTempName = GetTypeString(oEntity);

                    if (plnJointTempName == null)
                    {
                        plnJointTempName = string.Empty;
                    }
                }

                PhysicalConnection phyConn = oParents[0] as PhysicalConnection;
                if (phyConn != null)
                {
                    phyConnTempName = phyConn.Name;

                    if (phyConnTempName == null)
                    {
                        phyConnTempName = string.Empty;
                    }
                }
                
                string[] plnJointTempNames = plnJointTempName.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);
                string strPartname = string.Join("", plnJointTempNames);
                GetCountAndLocationID(strPartname, out count, out locationId);

                if (locationId != null)
                {
                    entityName = phyConnTempName + "-W-" + locationId + "-" + string.Format(strCountFormat, count);
                }
                else
                {
                    entityName = phyConnTempName + "-W" + string.Format(strCountFormat, count);
                }
                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlnJointNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }	
        }

        public override System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oNamingParents = new Collection<BusinessObject>();
            try
            {
                PlanningJoint plnJoint = oEntity as PlanningJoint;              

                if (plnJoint != null)
                {
                    
                    PhysicalConnection phyConn = plnJoint.PhysicalConnection;
                    if (phyConn != null)
                    {
                        oNamingParents.Add(phyConn);
                    }
                }                
            }
            catch(Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlnJointNameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oNamingParents;
        }
    }
}
