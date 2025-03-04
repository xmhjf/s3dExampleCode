using Ingr.SP3D.Common.Middle;
using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Content.Structure;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// This is an example class from which a new plate system naming rule can be derived.  It
    /// implements a single required interface, INamingSolverBase, which will be used by the
    /// relevant naming semantics to name a Molded Forms business object
    /// </summary>
    /// 

    // The WrapperProgID is a required attribute and should always be set to "MarineRuleWrappers.NamingSolverWrapper"
    [WrapperProgID("MarineRuleWrappers.NamingSolverWrapper")]
    public class PlnSeamNameRule : INamingSolverBase
    {
        public string ComputeName(BusinessObject oEntity, BusinessObject oActiveEntity, string strIdx)
        {
            string computedName = string.Empty;
            try
            {
                string baseName = string.Empty;
                string childName = string.Empty;                
                SeamBase seamBase = oEntity as SeamBase;
                BlockCuttingSurface blockCuttingSurface =null;                 

                if (seamBase != null)
                {
                    BusinessObject intersectingObj = seamBase.IntersectingObject;
                    RelationCollection relationColl = intersectingObj.GetRelationship("OverlapSurface", "OverlapSurface_DEST");

                    if (relationColl != null && relationColl.TargetObjects.Count > 0)
                    {
                        blockCuttingSurface = relationColl.TargetObjects[0] as BlockCuttingSurface;
                    }                  
                }

                if (blockCuttingSurface != null)
                {
                    ReadOnlyCollection<Block> connectedBlocks = blockCuttingSurface.ConnectedBlocks;
                    string block1Name = connectedBlocks[0].Name;
                    string block2Name = connectedBlocks[1].Name;

                    if (oEntity.SupportsInterface("IJSeam") == true)
                    {
                        baseName = block1Name + "-" + block2Name + "-" + "Pseam";
                    }
                    else if (oEntity.SupportsInterface("IJSeamPoint") == true)
                    {
                        baseName = block1Name + "-" + block2Name + "-" + "Pseampoint";
                    }
                }

                if (oActiveEntity != null)
                {
                    childName = Convert.ToString(oActiveEntity.GetPropertyValue("IJNamedItem", "Name"));
                    childName = "-" + strIdx;
                }
                computedName = baseName + childName;
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlnSeamNameRule.ComputeName: Error encountered (" + e.Message + ")");
            }   
            return computedName;
        }

        public Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> namingParents = new Collection<BusinessObject>();
            try
            {
                SeamBase seamBase = oEntity as SeamBase;

                if (seamBase != null)
                {
                    ISystem systemParent = seamBase.SystemParent;
                   
                    if (systemParent != null)
                    {
                        namingParents.Add((BusinessObject)systemParent);
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlnSeamNameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }  
           
            return namingParents;
        }
    }

}
