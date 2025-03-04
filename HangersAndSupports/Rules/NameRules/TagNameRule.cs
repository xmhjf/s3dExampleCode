using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    /// <summary>
    /// MemberTypeNameRule computes name for a member (MemberPart or MemebrSystem) according to its MemebrType property. 
    /// NameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class TagNameRule : NameRuleBase
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
            if (businessObject == null)
            {
                throw new ArgumentNullException("businessObject");
            }
            if (activeEntity == null)
            {
                throw new ArgumentNullException("activeEntity");
            }

            if (businessObject is Ingr.SP3D.Support.Middle.Support)
            {
                Ingr.SP3D.Support.Middle.Support oHgrSupport = (Ingr.SP3D.Support.Middle.Support)businessObject;
                IPart oSupportDefinition = oHgrSupport.SupportDefinition;
                if (oSupportDefinition != null)
                {
                    string codeListName = oSupportDefinition.PartNumber;
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
                            string childName = codeListName + "-" + counter.ToString().Trim();

                            //Set the name of the business object. 
                            SetName(businessObject, childName);
                        }
                        else
                            SetName(businessObject, businessObject.ToString());
                    }
                }
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
            return new Collection<BusinessObject>();
        }
    }
}
