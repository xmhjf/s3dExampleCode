using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Civil.Middle;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Civil
{
    public class TrenchPartDefaultNamingRule : NameRuleBase
    {
        /// <summary>
        /// Computes name for the trenchpart based on trenchpart groupname and parent feature names.
        /// </summary>
        /// <param name="businessObject">The business object whose name is being computed.</param>
        /// <param name="parents">Naming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject businessObject, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {

            if (businessObject == null)
            {
                throw new CmnArgumentNullException("The businessObject to be named is null");
            }
            if (parents == null)
            {
                throw new CmnArgumentNullException("The naming parents of the business object to be named are null");
            }

            if (activeEntity == null)
            {
                throw new CmnArgumentNullException("The name rule active entity associated to the business object to be named is null");
            }

            try
            {
                if (parents.Count() > 0)
                {

                    string groupName=string.Empty;
                   
                    TrenchPart trenchPart = businessObject as TrenchPart;

                    if (trenchPart != null)
                    {
                        groupName = trenchPart.GroupName;
                    }

                    string parentNameFromBO = parents.ElementAt(0).ToString();
                    string parentNameFromAE = base.GetNamingParentsString(activeEntity);

                    if (groupName != null && parentNameFromAE != null)
                    {
                        if (!(string.Equals(parentNameFromAE,parentNameFromBO,StringComparison.Ordinal)))
                        {   
                            parentNameFromAE = parentNameFromBO;

                            string newNameForBO = groupName + "-" + parentNameFromBO;

                            base.SetNamingParentsString(activeEntity, parentNameFromAE);
                            base.SetName(businessObject, newNameForBO);

                        }
                    }
                }
            }
            catch (CmnException ex)
            {
                MiddleServiceProvider.ErrorLogger.Log("Unexpected error:  CivilNamingRules.TrenchPartDefaultNamingRule.ComputeName.",1);
                throw ex;
            }
        }

        /// <summary>
        /// Gets the naming parents.
        /// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="entity">BusinessObject for which naming parents are required.</param>
        /// <returns>Returns the collection of naming parents.</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {

            if (entity == null)
            {
                throw new CmnArgumentNullException("The entity to get the parents is null");
            }
            Collection<BusinessObject> oParentsColl = new Collection<BusinessObject>();
            
            try
            {
                TrenchPart trenchPart = (TrenchPart)entity;
                ISystem systemParent =  trenchPart.SystemParent;
                
                if (systemParent != null)
                {
                    oParentsColl.Add((BusinessObject)systemParent);
                }
            }
            catch (CmnException ex)
            {
                MiddleServiceProvider.ErrorLogger.Log("Unexpected error:  CivilNamingRules.TrenchPartDefaultNamingRule.GetNamingParents.",1);
                throw ex;
            }
            return oParentsColl;
        }
    }
}
