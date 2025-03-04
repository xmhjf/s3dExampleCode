using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Planning
{
    public class BlockNameRule : NameRuleBase
    {
        private const string seperator = "-";

        //********************************************************************
        // Overall description:
        // Creates a name for the block object passed in. Creation only takes place if the
        // currnet name is empty (""). BlockName is based on parents name + index.
        // E.g. Block with parentname B0.2.1 gets name B0.2.1.X where X is unique.
        // No checks are performes as to wheather the blockname allready exists.
        //
        // The ComputeName method is called by semantic.
        //
        //
        // It is assumed that all Naming Parents and the Object implement IJNamedItem.
        // The Naming Parents are added in AddNamingParents() of the same interface.
        // Both these methods are called from naming rule semantic.
        //********************************************************************
        public override void ComputeName(BusinessObject oEntity, ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                Block block = oEntity as Block;                
                string [] delimiter ={"-"};
                string parentName = string.Empty;


                //get name from parent & IJNamedItem from this and set new name
                    //check if block allready has a name
                if (block != null && string.IsNullOrEmpty(block.Name))
                {
                    // get parent of this block
                    IAssemblyChild assemblyChild = oEntity as IAssemblyChild;

                    if (assemblyChild != null)
                    {
                        Block parentBlock = assemblyChild.AssemblyParent as Block;

                        if (parentBlock != null)
                        {
                            // get max sub block index, N, which exist below the parent e.g. Bx.y.z.N
                            ReadOnlyCollection<Block> siblingBlocks = block.GetSiblingBlocks();
                            long subBlockIndex = 0;

                            if (siblingBlocks != null)
                            {
                                foreach (IAssemblyChild item in siblingBlocks)
                                {
                                    Block tempBlock = item as Block;
                                    if (tempBlock != null && !string.IsNullOrEmpty(tempBlock.Name) && tempBlock.Name.StartsWith("B0"))
                                    {
                                        string[] subNames = tempBlock.Name.Split(delimiter, StringSplitOptions.RemoveEmptyEntries);

                                        if (subNames.Length >= 2)
                                        {
                                            long index = Convert.ToInt64(subNames[(subNames.Length - 1)]);

                                            if (index > subBlockIndex)
                                            {
                                                subBlockIndex = index;
                                            }
                                        }
                                    }
                                }
                            }

                            parentName = parentBlock.Name;
                            string partName = GetTypeString(oEntity);

                            if (string.IsNullOrEmpty(partName))
                            {
                                partName = string.Empty;
                            }

                            long count;
                            string LocationID, entityName = string.Empty;
                            GetCountAndLocationID(partName, out count, out LocationID);

                            if (parentName != string.Empty)
                            {
                                if (LocationID != null)
                                {
                                    entityName = parentName + "." + LocationID + seperator + (subBlockIndex + 1);
                                    oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                }
                                else
                                {
                                    entityName = parentName + seperator + (subBlockIndex + 1);
                                    oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                                }
                            }
                            else
                            {
                                oEntity.SetPropertyValue("Unspecified", "IJNamedItem", "Name");
                            }
                        }
                        else
                        {
                            oEntity.SetPropertyValue("Unspecified", "IJNamedItem", "Name");
                        }
                    }
                }
            }            
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("BlockNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }   
        }

        //''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //Description
        // All the Naming Parents that need to participate in an objects naming are added here
        // to the IJElements collection. The parent added here is root of assembly hierarchy
        // and is used in computing the name of the object in ComputeName() of the same interface. Both these methods are called from
        // naming rule semantic.
        //
        // For Blocks do NOT add a naming parent as no semantic should be involved.
        //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            
            return new Collection<BusinessObject>(); 
        }
    }
}
