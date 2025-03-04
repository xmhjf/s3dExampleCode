using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Planning
{
    public class CommonNameRule : NameRuleBase
    {
        private const string format = "{0:000000}";

        //********************************************************************
        // Description:
        // Creates a name for the object passed in. The name is based on the  string "A" and an Index.
        // The Index is unique for the Asembly.
        // It is assumed that all Naming Parents and the Object implement IJNamedItem.
        // The Naming Parents are added in AddNamingParents() of the same interface.
        // Both these methods are called from naming rule semantic.
        //********************************************************************
        public override void ComputeName(BusinessObject oEntity, ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                if (oEntity == null)
                {
                    throw new ArgumentNullException();
                }

                string partName = GetTypeString(oEntity);
                string oEntityName = partName;
                string parentname = null;
                string[] delimiter = { " " };
                parentname = GetNamingParentsString(oActiveEntity);                
                partName = string.Join("", partName.Split(delimiter, StringSplitOptions.RemoveEmptyEntries));

                if (string.Compare(partName, parentname) != 0 &&
                   (string.Compare(oEntityName, "New Assembly") == 0 ||
                   string.IsNullOrEmpty(oEntityName)))
                {
                    SetNamingParentsString(oActiveEntity, partName);
                    long count;
                    string locationId;
                    GetCountAndLocationID(partName, out count, out locationId);
                    string childname = "A" + locationId + string.Format(format, count);
                    oEntity.SetPropertyValue(childname, "IJNamedItem", "Name");               
                }
                else
                {
                    SetNamingParentsString(oActiveEntity, partName);
                }  
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("CommonNameRule.ComputeName: Error encountered (" + e.Message + ")");
            } 
          
        }

        //''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //Description
        // All the Naming Parents that need to participate in an objects naming are added here to the
        // IJElements collection. The parent added here is root of assembly hierarchy and is used in computing
        // the name of the object in ComputeName() of the same interface. Both these methods are called from
        // naming rule semantic.
        //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> namingParents = new Collection<BusinessObject>();
            try
            {
                BusinessObject grandParent = Block.GetRootBlock(oEntity.DBConnection);

                if (grandParent != null)
                {
                     namingParents.Add(grandParent);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("CommonNameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            } 

            return namingParents;
        }        
    }
}
