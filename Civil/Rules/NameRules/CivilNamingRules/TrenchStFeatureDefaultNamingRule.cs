using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Civil.Middle;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Civil
{
    public class TrenchStFeatureDefaultNamingRule : NameRuleBase
    {
        /// <summary>
        /// Gets the naming parents.
        /// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="bussinessObject">BusinessObject for which naming parents are required.</param>
        /// <returns>Returns the collection of naming parents.</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject bussinessObject)
        {

            if (bussinessObject == null)
            {
                throw new CmnArgumentNullException("The bussinessObject to get parents is null");
            }
            Collection<BusinessObject> parentsColl = new Collection<BusinessObject>();
           

            try
            {
                TrenchFeature trenchFeature = bussinessObject as TrenchFeature;

                if (trenchFeature != null)
                {
                    ISystem systemParent = trenchFeature.SystemParent;

                    if (systemParent != null)
                    {
                        parentsColl.Add((BusinessObject)systemParent);
                    }
                }
            }
            catch (CmnException ex)
            {
                MiddleServiceProvider.ErrorLogger.Log("Unexpected error:  CivilNamingRules.TrenchStFeatureDefaultNamingRule.GetNamingParents",1);
                throw ex;
            }
            return parentsColl;
        }
        
        /// <summary>
        /// Computes a name for the given entity. 
        /// </summary>
        /// <param name="businessObject">The business object whose name is being computed. </param>
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

                if (parents.Count > 0)
                {
                    string parentNameFromBO;
                    string parentNameFromAE;

                    parentNameFromBO = parents.ElementAt(0).ToString();

                    parentNameFromAE = base.GetNamingParentsString(activeEntity);

                    if (!string.Equals(parentNameFromAE,parentNameFromBO,StringComparison.Ordinal))
                    {   
                        long counter;
                        string locationID;

                        parentNameFromAE = parentNameFromBO;

                        base. GetCountAndLocationID(parentNameFromBO, 1, out counter, out locationID);
                        string sequenceNo = counter.ToString("D4");
                        string newNameForBO = parentNameFromBO + "-" + sequenceNo;

                        base.SetNamingParentsString(activeEntity, parentNameFromAE);
                        base.SetName(businessObject, newNameForBO); 
                    }
                }
            }
            catch (CmnException ex)
            {
                MiddleServiceProvider.ErrorLogger.Log("Unexpected error:  CivilNamingRules.TrenchStFeatureDefaultNamingRule.ComputeName",1);
                throw ex;
            }
        }

    }
}
