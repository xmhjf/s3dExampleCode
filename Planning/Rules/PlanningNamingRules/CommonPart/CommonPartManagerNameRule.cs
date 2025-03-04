using System;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Planning
{
    public class CommonPartManagerNameRule : NameRuleBase
    {
        /// <summary>
        /// Computes a name for the given entity. 
        /// </summary>
        /// <param name="businessObject">The business object whose name is being computed. </param>
        /// <param name="parents">Naming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject businessObject, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            PropertyValue propValue, propValues = null;
            String parentString = string.Empty;

            try
            {
                if (parents != null && activeEntity!=null)
                {
                    propValue = parents[0].GetPropertyValue("IJNamedItem", "Name");
                    parentString = propValue.ToString();
                    if (GetNamingParentsString(activeEntity) != parentString)
                    {
                        SetNamingParentsString(activeEntity, parentString);
                    }
                }

                //This namerule sets the name of the assembly as the Name of the Manager.
                if (businessObject != null)
                {
                    propValues = businessObject.GetPropertyValue("IJNamedItem", "Name");
                    if (propValues.ToString() != parentString)
                    {
                        businessObject.SetPropertyValue(parentString, "IJNamedItem", "Name");
                    }
                }

            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("CommonPartManagerNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }

        }

        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            Collection<BusinessObject> oRetColl = new Collection<BusinessObject>();

            BusinessObject oRelatedAssembly = null;
            try
            {
                if (entity != null)
                {
                    RelationCollection oRelatedAssemblies = entity.GetRelationship("CPMgrHierarchy", "CPAssembly");

                    if (oRelatedAssemblies.TargetObjects.Count > 0)
                    {
                        oRelatedAssembly = oRelatedAssemblies.TargetObjects.ElementAt(0);

                        oRetColl.Add(oRelatedAssembly);
                    }
                }

            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("CommonPartManagerNameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }

            return oRetColl;
        }


    }
}
