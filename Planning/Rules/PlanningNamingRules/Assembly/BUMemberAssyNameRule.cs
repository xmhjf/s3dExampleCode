using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Planning
{
    public class BUMemberAssyNameRule : NameRuleBase
    {
        //********************************************************************
        // Description:
        // Assembly's name will match with name of the Built-up Member if any.
        //*******************************************************************
        public override void ComputeName(BusinessObject oEntity, ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                if (oEntity == null)
                {
                    throw new ArgumentNullException();
                }

                string parentName = null;

                if (oParents.Count == 1)
                {
                    parentName = Convert.ToString(oParents[0].GetPropertyValue("IJNamedItem", "Name"));
                }
                else
                {
                    return;
                }

                string entityName = Convert.ToString(oEntity.GetPropertyValue("IJNamedItem", "Name"));

                if (entityName == "New Assembly" || string.IsNullOrEmpty(entityName) || (string.Compare(entityName, parentName) != 0))
                {
                    oEntity.SetPropertyValue(parentName, "IJNamedItem", "Name");                    
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("BUMemberAssyNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }   

        }

        //''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //Description
        // There will be only one naming parent: Built-up member, if any, related to the assembly
        //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> namingParents = new Collection<BusinessObject>();
            try
            {
                BusinessObject parent=null;
                RelationCollection targetRelations = oEntity.GetRelationship("ModuleToAssembly", "Module");

                if (targetRelations != null && targetRelations.TargetObjects != null)
                {
                    if (targetRelations.TargetObjects.Count ==1)
                    {
                        parent = targetRelations.TargetObjects[0];
                    }
                }

                if (parent != null)
                {
                    if (parent.SupportsInterface("ISPSDesignedMember") == true)
                    {
                        namingParents.Add(parent);
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("BUMemberAssyNameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return namingParents;
        }
    }
}
