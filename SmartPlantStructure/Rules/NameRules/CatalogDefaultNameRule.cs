using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// CatalogDefaultNameRule computes name for a designed member or a smart occurance object. 
    /// NameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class CatalogDefaultNameRule : NameRuleBase
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

            //If the object is rule based Designed Member, use the rule's Part Description to compose the Name. 
            if (businessObject is MemberPart)
            {
                MemberPart memberPart = (MemberPart)businessObject;
                //If member part is a designed member get the part cross section name and assign it to code list name. 
                if (memberPart.DesignedMember)
                {
                    CrossSection partCrossSection = memberPart.CrossSection;
                    codeListName = partCrossSection.Name;
                }
            }
            else
            {
                string relationshipName = "SOtoSI_R";
                string roleName = "toSI_ORIG";

                //Use the relationship name and role name to get the Relationship. 
                RelationCollection relationCollection = businessObject.GetRelationship(relationshipName, roleName);
                if (relationCollection != null)
                {
                    ReadOnlyCollection<BusinessObject> targetObjectCollection = relationCollection.TargetObjects;
                    if (targetObjectCollection.Count > 0)
                    {
                        //If the target object is IPart, get the part number on the part and assign it to code list name. 
                        if (targetObjectCollection[0] is IPart)
                        {
                            IPart part = (IPart)targetObjectCollection[0];
                            codeListName = part.PartNumber;
                        }
                    }
                }
            }

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
                        location = location + "-";
                    }

                    //Set the child name accordingly. 
                    //Counter has to be padded to the left with zeros for correct formatting. 
                    string childName = codeListName + "-" + location + counter.ToString().PadLeft(4, '0');

                    //Set the name of the business object. 
                    SetName(businessObject, childName);
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
            return new Collection<BusinessObject>();            
        }
    }
}
