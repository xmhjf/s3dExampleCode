//*********************************************************************************************
//Copyright (C) 2004, Intergraph Corporation. All rights reserved.                            
//Abstract:                                                                                   
//    Common Naming Rule for Cable Marker                                                     
//Description:                                                                                
//    This naming rule will name the marker after the parent to which the feature holding     
//    the marker is related.  This will be either a conduit run or a cableway.                
//           Marker Name = RunName & TypeName & Counter                                       
//                   where,                                                                   
//                       RunName is the name of the cableway or conduit run                   
//                       TypeName is the user name for the cable marker object class          
//                       Counter is a numberic counter with format 0000                     
//Notes:                                                                                      
//History                                                                                     
//    Vijay          05/26/2015      CR-CP-269176  Route Cable VB6 Rules need to be replaced  
//*********************************************************************************************


using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Route.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ingr.SP3D.Content.Cable.Rules
{
    public class CommonRule : NameRuleBase
    {
        /// <summary>
        /// This function will return the parent object of the cable marker such that the NamedEntity relationship can be created between the marker and it's naming parent.
        /// The relationship will, subsequently, be used to main-the proper name on the marker in the event the name of the run is modified.
        /// </summary>
        /// <param name="businessObject">The business object whose name is being computed.</param>
        /// <param name="parents">The parentsNaming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject businessObject, System.Collections.ObjectModel.ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            if (businessObject == null)
            {
                throw new ArgumentNullException("businessObject");
            }
            if (activeEntity == null)
            {
                throw new ArgumentNullException("activeEntity");
            }
            try
            {
                string baseName = "CableMarker";
                string childName = string.Empty;
                string parentName = string.Empty;
                string namedParentsString = string.Empty;
                long count = 0;
                string location = string.Empty;
                namedParentsString = GetNamingParentsString(activeEntity);
                if (parents.Count > 0)
                {
                    foreach (BusinessObject item in parents)
                    {
                        INamedItem parentNamedItem = item as INamedItem;
                        if (parentNamedItem != null)
                        {
                            parentName = parentNamedItem.Name;
                        }
                        if (childName.Length == 0)
                        {
                            childName = parentName;
                        }
                        else
                        {
                            childName = childName + "-" + parentName;
                        }
                        parentNamedItem = null;
                    }
                    if (childName + "-" + baseName != namedParentsString)
                    {
                        SetNamingParentsString(activeEntity, childName + "-" + baseName);
                        GetCountAndLocationID(childName + "-" + baseName, out count, out location);
                        childName = childName + "-" + baseName + "-" + string.Format("{0:000}", count);
                        SetName(businessObject, childName);
                    }
                }
                else
                {
                    if (baseName != namedParentsString)
                    {
                        SetNamingParentsString(activeEntity, baseName);
                        GetCountAndLocationID(baseName, out count, out location);
                        childName = baseName + "-" + string.Format("{0:000}", count);
                        SetName(businessObject, childName);
                    }
                }
            }
            catch (Exception)
            {
                throw new Exception("Unexpected error: CableMarkerNameRules.CommonRule.ComputeName");
            }
        }
        /// <summary>
        /// This function will return the parent object of the cable marker such that the NamedEntity      
        /// relationship can be created between the marker and its naming parent.  
        /// The relationship will, subsequently, be used to main the proper name on the marker in the event the name of the run is modified.                
        /// </summary>
        /// <param name="entity">>BusinessObject for which naming parents are required.</param>
        /// <returns>ReadOnlyCollection of BusinessObjects</returns>
        public override System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            Collection<BusinessObject> namingParentsColl = new Collection<BusinessObject>();
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }
            try
            {
                if (entity.SupportsInterface("IJRteMarker"))
                {
                    RelationCollection relCol = entity.GetRelationship("MarkedFeatures", "MarkedFeature");
                    if (relCol != null && relCol.TargetObjects.Count > 0)
                    {
                        RouteFeature routeFeat = (RouteFeature)relCol.TargetObjects[0];
                        RouteRun pathRun = (RouteRun)routeFeat.SystemParent;
                        namingParentsColl.Add(pathRun);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return namingParentsColl;
        }
    }
}
