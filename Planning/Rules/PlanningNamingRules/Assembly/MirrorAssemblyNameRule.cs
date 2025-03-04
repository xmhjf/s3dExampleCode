using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Planning
{
   public class MirrorAssemblyNameRule : NameRuleBase
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
                string parentName = null;

                if (oEntity == null)
                {
                    throw new ArgumentNullException();
                }

                Ingr.SP3D.Planning.Middle.Assembly assembly = oEntity as Ingr.SP3D.Planning.Middle.Assembly;
                Ingr.SP3D.Planning.Middle.Assembly symmetricalAssembly = GetSymmetricalAssembly(assembly);

                if (symmetricalAssembly != null)
                {                   
                    oEntity.SetPropertyValue(symmetricalAssembly.Name, "IJNamedItem", "Name");
                } 
              
                parentName = GetNamingParentsString(oActiveEntity);               
                string partName = GetTypeString(oEntity);
                string[] delimiter = { " " };
                string typeOfAssembly = string.Empty;
                partName = string.Join(" ", partName.Split(delimiter, StringSplitOptions.None));

                if (oEntity is Assembly)
                {
                    typeOfAssembly = "A";
                }
                else
                {
                    typeOfAssembly = "AB";
                }

                if (string.Compare(partName, parentName) != 0)
                {
                    if (assembly != null && (assembly.Name == "New Assembly" || string.IsNullOrEmpty(assembly.Name)))
                    {
                        string locationId;
                        long count;
                        string entityName = string.Empty;
                        GetCountAndLocationID(partName, out count, out locationId);
                        entityName = typeOfAssembly + locationId + string.Format(format, count);
                        oEntity.SetPropertyValue(entityName, "IJNamedItem", "Name");
                    }

                    SetNamingParentsString(oActiveEntity, partName);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("MirrorAssemblyNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }           
        }

       //'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                MiddleServiceProvider.ErrorLogger.Log("MirrorAssemblyNameRule.GetNamingParents : Error encountered (" + e.Message + ")");
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
        private Ingr.SP3D.Planning.Middle.Assembly GetSymmetricalAssembly(Ingr.SP3D.Planning.Middle.Assembly oAssembly)
        {
            Ingr.SP3D.Planning.Middle.Assembly symmetricalAssembly = null;

            if (oAssembly != null)
            {
                RelationCollection destinationRelations = oAssembly.GetRelationship("PlnMirrorAsm", "PlnMirrorAsm_DEST");

                if (destinationRelations.TargetObjects.Count > 0)
                {
                    symmetricalAssembly = destinationRelations.TargetObjects[0] as Ingr.SP3D.Planning.Middle.Assembly;
                }
            }

            return symmetricalAssembly;
        }
    }
}
