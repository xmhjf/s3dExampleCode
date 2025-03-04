using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Planning
{
    public class PlnJointFolderNameRule : NameRuleBase
    {     
        public override void ComputeName(BusinessObject oEntity, System.Collections.ObjectModel.ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                string entityName = string.Empty;
                int plnJointFolderType;
                PlanningJointFolder plnJoingFolder = oEntity as PlanningJointFolder;                

                if (plnJoingFolder != null)
                {
                    plnJointFolderType = plnJoingFolder.FolderType;

                    switch (plnJointFolderType)
                    {
                        case ((int)PlanningJointFolderType.Root):
                            {
                                entityName = "Welds";
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                break;
                            }
                        case ((int)PlanningJointFolderType.Postponed):
                            {
                                entityName = "Postponed";
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                break;
                            }
                        case ((int)PlanningJointFolderType.Unsequenced):
                            {
                                entityName = "Unsequenced";
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                break;
                            }
                        case ((int)PlanningJointFolderType.Unassigned):
                            {
                                entityName = "Unassigned";
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                break;
                            }
                        case ((int)PlanningJointFolderType.Subsequent):
                            {
                                entityName = "Subsequent";
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                break;
                            }
                        case ((int)PlanningJointFolderType.Joining):
                            {
                                foreach (BusinessObject parent in oParents)
                                {
                                    PropertyValue propValue = parent.GetPropertyValue("IJNamedItem", "Name");
                                    entityName = "Joining " + Convert.ToString(propValue);
                                    oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                }
                                break;
                            }
                        default :
                            {
                                // this is for custom planning joint folder.it will get the display name from DB and set it  name to the folder
                                bool found = false;
                                GetAttribute(oEntity, "FolderType", out found, out entityName);
                                oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                break;
                            }
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlnJointFolderNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }             
        }

        public override System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oNamingParents = new Collection<BusinessObject>();
            try
            {
                PlanningJointFolder plnJointFolder = oEntity as PlanningJointFolder;

                if (plnJointFolder != null)
                {
                    BusinessObject assyChild = plnJointFolder.ConnectedPart as BusinessObject;                    
                    if (assyChild != null)
                    {
                        oNamingParents.Add(assyChild);
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlnJointFolderNameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oNamingParents;
        }

        private void GetAttribute(BusinessObject entity, string AttributeName, out bool attrfound, out string longStringValue)
        {
            longStringValue = null;
            attrfound = false;
            try
            {
                ReadOnlyDictionary<InterfaceInformation> interfacesInfo = entity.ClassInfo.Interfaces;
                if (interfacesInfo != null)
                {
                    foreach (var keyValuePair in interfacesInfo)
                    {
                        InterfaceInformation interfaceInfo = keyValuePair.Value;

                        if (interfaceInfo != null)
                        {
                            PropertyInformation propertyInfo = interfaceInfo.GetPropertyInfo(AttributeName);

                            if (propertyInfo != null)
                            {
                                PropertyValue propertyValue = entity.GetPropertyValue(propertyInfo);
                                List<CodelistItem> codeListMembers = propertyInfo.CodeListInfo.CodelistMembers;

                                if (codeListMembers != null)
                                {
                                    foreach (var item in codeListMembers)
                                    {
                                        string PropertyName = Convert.ToString(propertyValue);
                                        if (string.Compare(item.ShortDisplayName, PropertyName) == 0)
                                        {
                                            CodelistItem requiredCodelistItem = item;
                                            longStringValue = requiredCodelistItem.DisplayName;
                                            attrfound = true;
                                            return;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("DefaultGroupNameRule.GetAttribute: Error encountered (" + e.Message + ")");
            }

        }
    }
}
