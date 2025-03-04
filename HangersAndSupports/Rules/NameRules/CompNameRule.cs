using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    /// <summary>
    /// CatalogDefaultNameRule computes name for a designed member or a smart occurance object. 
    /// NameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class CompNameRule : NameRuleBase
    {
        /// <summary>
        /// Computes the name for the given entity. 
        /// </summary>
        /// <param name="businessObject">The business object whose name is being computed.</param>
        /// <param name="parents">The parentsNaming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject businessObject, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            if (businessObject == null)
            {
                throw new ArgumentNullException("businessObject");
            }
            if (activeEntity == null)
            {
                throw new ArgumentNullException("activeEntity");
            }

            string codeListName = string.Empty;

            if (businessObject is SupportComponent)
            {
                SupportComponent oSuppComp = (SupportComponent)businessObject;
                Part oPart = (Part)oSuppComp.GetRelationship("madeFrom", "part").TargetObjects[0];
                codeListName = oPart.PartNumber;
                string namingParentsString = GetNamingParentsString(activeEntity);

                if (codeListName != null && namingParentsString != null)
                {
                    //Name only needs to be recomputed and if the naming parent string is different from the existing one. 
                    if (!codeListName.Equals(namingParentsString))
                    {
                        //Set the naming parents string to obtained code list name. 
                        SetNamingParentsString(activeEntity, codeListName);

                        long counter = 0;
                        string location = string.Empty;

                        //Get the running count for the business object and location ID from the NameGeneratorService. 
                        GetCountAndLocationID(codeListName, out counter, out location);
                        if (!string.IsNullOrEmpty(location))
                        {
                            codeListName = codeListName + "-" + location;
                        }

                        //Set the child name accordingly. 
                        //Counter has to be padded to the left with zeros for correct formatting. 
                        string childName = codeListName + "-C" + counter.ToString().Trim();


                        //Set the name of the business object. 
                        SetName(businessObject, childName);
                    }
                    else
                        SetName(businessObject, businessObject.ToString());
                }
            }
            
        }

        /// <summary>
        /// Gets the naming parents from naming rule.
        /// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="oEntity">BusinessObject for which naming parents are required.</param>
        /// <returns>ReadOnlyCollection of BusinessObjects.</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            Collection<BusinessObject> parentsColl = new Collection<BusinessObject>();
            if (entity is SupportComponent)
            {
                SupportComponent oHgrSupportComp = (SupportComponent)entity;
                BusinessObject oSupport = (BusinessObject)oHgrSupportComp.SystemParent;
                if (oSupport == null)
                {
                    BusinessObject oConfigProjRoot = (BusinessObject)MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.RootSystem;
                    parentsColl.Add(oConfigProjRoot);
                }
                else
                     parentsColl.Add(oSupport);
            }
            return parentsColl;
        }

    }
}
