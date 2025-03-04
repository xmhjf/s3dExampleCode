using Ingr.SP3D.Common.Middle;
using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Planning
{   
    public class AssemblyNameRule : NameRuleBase
    {
        private const string CountFormat = "{0:000000}";

        //********************************************************************
        // Description:
        // Creates a name for the object passed in. The name is based on the  string "A" and an Index.
        // The Index is unique for the Asembly.
        // It is assumed that all Naming Parents and the Object implement IJNamedItem.
        // The Naming Parents are added in AddNamingParents() of the same interface.
        // Both these methods are called from naming rule semantic.
        //********************************************************************        
        public override void ComputeName(BusinessObject oEntity, System.Collections.ObjectModel.ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                string[] delimiter = {" "};
                string assemblyType = string.Empty, entityNewName = string.Empty, activeEntityName = null;
                string partName = string.Empty, locationID;
                long count;                

                if (oEntity == null)
                {                    
                    throw new ArgumentNullException();
                }

                Assembly assembly = oEntity as Assembly;
                activeEntityName = GetNamingParentsString(oActiveEntity);                

                if (assembly != null)
                {
                    string assemblyTempName = GetTypeString(oEntity);

                    //Use part Name for basename, remove blanks:
                    if (assemblyTempName != null)
                    {                        
                        partName = string.Join("", assemblyTempName.Split(delimiter, StringSplitOptions.RemoveEmptyEntries));
                    }

                    if (oEntity is Assembly)
                    {
                        assemblyType = "A";
                    }
                    else if (oEntity is AssemblyBlock)
                    {
                        assemblyType = "AB";
                    }

                    if (string.Compare(partName, activeEntityName) != 0)
                    {
                        if (string.Compare(assembly.Name, "New Assembly") == 0 || string.IsNullOrEmpty(assembly.Name))
                        {
                            // If a Name Rule ID was given at the creation of the database it will be returned
                            // into strLocationID.  If not it will be a null string and not print out
                            GetCountAndLocationID(partName, out count, out locationID);                            
                            entityNewName = assemblyType + locationID + string.Format(CountFormat, count);
                            oEntity.SetPropertyValue(entityNewName, "IJNamedItem", "Name");
                        }
                        
                        SetNamingParentsString(oActiveEntity, partName);                            
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("AssemblyNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }             
        }

        //''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //Description
        // All the Naming Parents that need to participate in an objects naming are added here to the
        // IJElements collection. The parent added here is root of assembly hierarchy and is used in computing
        // the name of the object in ComputeName() of the same interface. Both these methods are called from
        // naming rule semantic.
        //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        public override System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oNamingParents = new Collection<BusinessObject>();
            try
            {
                Block rootBlock = Block.GetRootBlock(oEntity.DBConnection);

                if (rootBlock != null)
                {
                    oNamingParents.Add(rootBlock);
                }                
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("AssemblyNameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oNamingParents;
        }       
    }
}
