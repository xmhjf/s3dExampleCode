using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Civil
{
    public class TrenchRunDefaultNamingRule:NameRuleBase
    {
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
                throw new CmnArgumentNullException("The entity to get parents is null");
            }

            Collection<BusinessObject> oParentsColl = new Collection<BusinessObject>();

            try
            {
                BusinessObject oParent = GetParent(HierarchyTypes.System, entity); // Get the System Parent
                if (oParent != null)
                {
                    oParentsColl.Add(oParent); //Add the Parent to the ParentColl
                }
            }
            catch (CmnException ex)
            {
                MiddleServiceProvider.ErrorLogger.Log("Unexpected error:  CivilNamingRules.TrenchRunDefaultNamingRule.GetNamingParents.",1);
                throw ex;
            }
            return oParentsColl;
        }

        /// <summary>
        /// Computes a name for the given entity.
        ///         /// Earlier name computation was happend with appending location id.Now as part of TR-CP-291496	below steps are performed 
        /// to append the locationid for existing runs on UpdateName & also computing the name with location id with new runs created.
        /// STEP1: Check if the location id is already appended to trench run name.
        /// Step2:If the previous parent string is different from current parent string and it is differ only by missing locatationId then
        /// append the location id preserving the trenchrun counter, if not differ just by location id then modify the location id.
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
                // only if there are any parents we need to access the elements in parent collection
                if (parents.Count > 0)
                {
                    string locationID = string.Empty;
                    long counter=0;
                   
                    string parentNameFromBO = parents.ElementAt(0).ToString();
                    string parentNameFromAE = base.GetNamingParentsString(activeEntity);

                    base.GetCountAndLocationID(parentNameFromBO, 1, out counter, out locationID);

                    INamedItem trechRunNamedItem = businessObject as INamedItem;
                    string oldBOName=string.Empty;
                    if (trechRunNamedItem != null)
                    {
                        oldBOName = trechRunNamedItem.Name;
                    }
                    bool isLocationIdAppended = false;
                    string oldCounter = string.Empty;

                    // Checks if old name is appended with location id 
                    isLocationIdAppended = IsLocationIdAppended(oldBOName, locationID, out oldCounter);

                    if (!isLocationIdAppended || (!string.Equals(parentNameFromAE, parentNameFromBO,StringComparison.Ordinal)))
                    {
                        if (!string.IsNullOrEmpty(locationID))
                        {
                            locationID = locationID + "-";
                        }
                        string newNameForBO = string.Empty;
                        // 1st condition to check should be (parentStringFromAE != parentStringFromBO) becuase if the parent is 
                        //changed for the trenchrun which has no location id appended then same old counter is maintained 
                        //instead of generation of new counter ex if parent system of GenericSystem-0008 
                        //is changed to some other system let say Gen2 then new name will be Gen2-1-0008
                        if (!string.Equals(parentNameFromAE, parentNameFromBO, StringComparison.Ordinal))
                        {
                            string sequenceNo = counter.ToString("D4");
                            parentNameFromAE = parentNameFromBO;
                            newNameForBO = parentNameFromBO + "-" + locationID + sequenceNo;
                        }
                        else
                        {
                            newNameForBO = parentNameFromBO + "-" + locationID + oldCounter;
                        }

                        //Set the name of the business object. 
                       base. SetName(businessObject, newNameForBO);
                       base.SetNamingParentsString(activeEntity, parentNameFromAE);
                    }
                }
            }
            catch (CmnException ex)
            {
                MiddleServiceProvider.ErrorLogger.Log("Unexpected error:  CivilNamingRules.TrenchRunDefaultNamingRule.ComputeName.",1);
                throw ex;
            }
        }

        /// <summary>
        ///     Checks if the locationId is appended to the existing name or not.
        ///     
        ///     (1)Get the old name of the trench run
        ///     (2)Split the trench run with "-" delimeter,This will return the array of string.
        ///     (3)Now compare the string present at index (ArrayCount-2) with current location id :Assumption : Location id 
        ///             will always at lastbutone location in array
        ///     (4) if the location id is same it means there is location id already appended.
        ///
        /// </summary>
        /// <param name="nameToCheckLocID">The business name to check the location id exists or not.</param>
        /// <param name="currentLocID">current location id.</param>
        /// <param name="existingCounter">out argument to preserve the counter if only location id needs to be appended.</param>
        /// <returns>Return True if location id is appended else return false.</returns>
        private bool IsLocationIdAppended(string nameToCheckLocID, string currentLocID, out string existingCounter)
        {
            bool isLocationIdAppend= false;
            existingCounter = string.Empty;
            
            try
            {  
                existingCounter = string.Empty;
                // if old name is empty of location id is null then no need appending of location id
                // return true 
                if (string.IsNullOrEmpty(nameToCheckLocID) || string.IsNullOrEmpty(currentLocID)) 
                {
                    return true;
                }

                string[] nameArray;
                int countOfNameArray;

                nameArray = nameToCheckLocID.Split('-');

                countOfNameArray = nameArray.Count();

                if (countOfNameArray > 1)//Check for the strings whose count of array will be more than 1
                {
                    // since array is 0 based hence decrease by 2 to get last but one location
                    // Last but one element after splitting with '-' should be location id, 
                    //if is not equal, old name doesn't have location id need to append with location id 
                    if (nameArray[countOfNameArray - 2] == currentLocID)
                    {
                        isLocationIdAppend = true;
                    }
                    else
                    {
                        // preserve the  location id last but which is counter
                        existingCounter = nameArray[countOfNameArray - 1];
                    }

                }
            }
            catch (Exception)
            {
                // if there is any exception assume location id is appended as nothing needs to be updated
                isLocationIdAppend = true;
            }
            return isLocationIdAppend;
        }

    }
}
