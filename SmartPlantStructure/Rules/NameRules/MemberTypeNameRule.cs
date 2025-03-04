using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// MemberTypeNameRule computes name for a member (MemberPart or MemebrSystem) according to its MemebrType property. 
    /// NameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class MemberTypeNameRule : NameRuleBase
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

            //Get the code list info from metadata. MemberType code list is named ‘StructuralMemberType’ and is defined in CommonSchema namespace. 
            MetadataManager metadataMgr = businessObject.DBConnection.MetadataMgr;
            string codeListTableName = "StructuralMemberType";
            string codeListNameSpace = "CMNSCH";
            CodelistInformation memTypeCodeListInfo = metadataMgr.GetCodelistInfo(codeListTableName, codeListNameSpace);

            //Get the code list name of member type to display a codelist item by name, display name, and value. 
            IMemberType memberTypeBO = businessObject as IMemberType;
            if (memberTypeBO != null)
            {
                CodelistItem codelistItem = memTypeCodeListInfo.GetCodelistItem(memberTypeBO.Type);
                string codeListName = codelistItem.ShortDisplayName;

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

                        //Counter has to be padded to the left with zeros for correct formatting. 
                        string childName = codeListName + "-" + location + counter.ToString().PadLeft(4, '0');

                        //Set the name of the business object. 
                        SetName(businessObject, childName);
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
