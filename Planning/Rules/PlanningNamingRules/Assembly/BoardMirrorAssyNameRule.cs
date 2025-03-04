using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Planning
{
    public class BoardMirrorAssyNameRule : NameRuleBase
    {
        private const string format = "{0:000000}";

        //********************************************************************
        // Description:
        // Creates a name for the object passed in.
        //   It sets the same name as the symmetrically related object.
        //   If no Symmetrically related object exists, it sets the name in same process as Assembly Name Rule as follows.
        //The name is based on the  string "A" and an Index.
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

                string parentName = null;
                Ingr.SP3D.Planning.Middle.Assembly assembly = oEntity as Ingr.SP3D.Planning.Middle.Assembly;
                Ingr.SP3D.Planning.Middle.Assembly symmetricalAssembly = GetSymmetricalAssembly(assembly);

                if (symmetricalAssembly != null)
                {
                    string symmetricalAssyName = CalculateSymmetricalAssemblyName(symmetricalAssembly.Name);
                    oEntity.SetPropertyValue(symmetricalAssyName, "IJNamedItem", "Name");
                }

                parentName = GetNamingParentsString(oActiveEntity);              
                string partName = GetTypeString(oEntity);
                string[] delimiter = { " " };
                string assemblyType = string.Empty;               
                partName = string.Join("", partName.Split(delimiter, StringSplitOptions.RemoveEmptyEntries));

                if (oEntity is Assembly)
                {
                    assemblyType = "A";
                }
                else
                {
                    assemblyType = "AB";
                }

                if (string.Compare(partName, parentName) != 0)
                {
                    if (assembly != null && (assembly.Name == "New Assembly" || string.IsNullOrEmpty(assembly.Name)))
                    {
                        string locationId;
                        long count;
                        string childname = string.Empty;
                        // If a Name Rule ID was given at the creation of the database it will be returned
                        // into strLocationID.  If not it will be a null string and not print out
                        GetCountAndLocationID(partName, out count, out locationId);
                        childname = assemblyType + locationId + string.Format(format, count);
                        oEntity.SetPropertyValue(childname, "IJNamedItem", "Name");
                    }                   
                    SetNamingParentsString(oActiveEntity, partName);                       
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("BoardMirrorAssyNameRule.ComputeName: Error encountered (" + e.Message + ")");
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
           Collection<BusinessObject> oNamingParents = new Collection<BusinessObject>();
           try
           {
               Assembly tempAssembly = oEntity as Assembly;
               BusinessObject namingParent = GetSymmetricalAssembly(tempAssembly);

               if (namingParent == null)
               {
                   namingParent = Block.GetRootBlock(oEntity.DBConnection);
               }

               if (namingParent != null)
               {
                   oNamingParents.Add(namingParent);
               }
           }
           catch (Exception e)
           {
               MiddleServiceProvider.ErrorLogger.Log("BoardMirrorAssyNameRule.GetNamingParents : Error encountered (" + e.Message + ")");
           }  
           return oNamingParents;
        }

        //**************************************************************************
        // Method : GetSymmetricalAssembly
        //
        // Description:
        //   Gets the assembly reelated to the passed in by the relation PlnMirrorAsm. Returns NOTHING
        //   if there is no mirror assembly.
        //
        // Arguments:
        //   [in] ByVal oAssy As GSCADAsmHlpers.IJAssembly, the assembly to get the relationhelper from
        //
        // Return Values:
        //   the symmetrical assembly,
        //   if no symmetrical assembly returns "nothing"
        //**************************************************************************
        private Ingr.SP3D.Planning.Middle.Assembly  GetSymmetricalAssembly(Ingr.SP3D.Planning.Middle.Assembly oAssembly)
        {
            Ingr.SP3D.Planning.Middle.Assembly symmetricalAssembly = null;

            if (oAssembly != null)
            {
                //Get the collection of relations
                RelationCollection destinationRelations = oAssembly.GetRelationship("PlnMirrorAsm", "PlnMirrorAsm_DEST");

                //check that there is only one relationship
                if (destinationRelations.TargetObjects.Count > 0)
                {
                    symmetricalAssembly = destinationRelations.TargetObjects[0] as Ingr.SP3D.Planning.Middle.Assembly;
                }
            }
            //return the symmetrical assembly, might be nothing
            return symmetricalAssembly;
        }

        //**************************************************************************
        // Function : CalculateSymmetricalAssemblyName
        //
        // Abstract
        //   Returns the name of an assembly by replacing the positional denotation last letter of the 1st token of
        //   sSourceAssemblyName with "S" (or "P") if sSourceAssemblyName contains only one '-' delimiter.
        //
        //   Example: sSourceAssemblyName    CalculateSymmetricalAssemblyName
        //            -------------------------------------------------------
        //            B11P-F101A             B11S-F101A  (replaces P with S)
        //            B11S-G10B              B11P-G10B   (replaces S with P)
        //            F32C-F101A             F32C-F101A  (returns the same name since the last letter is not "P" or "S")
        //            F101A                  F101A       (returns the same name since no '-' delimiter)
        //**************************************************************************
        private string CalculateSymmetricalAssemblyName(string sourceAssemblyName)
        {
            string assemblyName = null;
            string token1, token2 = string.Empty, token1LastChar = string.Empty;
            string[] delimiter = { "-" };
            string[] tempNames = null;

            if (!string.IsNullOrEmpty(sourceAssemblyName))
            {
                tempNames = sourceAssemblyName.Split(delimiter,StringSplitOptions.None);
            }

            if (tempNames != null && tempNames.Length == 2)
            {
                token1 = tempNames[0];
                token2 = tempNames[1];
                token1LastChar = token1.ToUpper().Substring(token1.Length - 1);

                if (!string.IsNullOrEmpty(token1LastChar))
                {
                    string token1ExcludeLastChar = string.Empty;

                    if (token1LastChar == "P" || token1LastChar == "S")
                    {
                        token1ExcludeLastChar = token1LastChar.Substring(0, (token1LastChar.Length - 1));

                        if (token1LastChar == "P")
                        {
                            assemblyName = token1ExcludeLastChar + "S-" + token2;
                        }
                        else
                        {
                            assemblyName = token1ExcludeLastChar + "P-" + token2;
                        }
                    }
                }
            }
            else
            {
                assemblyName = sourceAssemblyName;
            }
            
            return assemblyName;
        }      
    }
}
