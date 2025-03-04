using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Structure
{
    public class PlateSysNCNamingRule : PlateSysNamingRule
    {
        public override string ComputeName(Ingr.SP3D.Common.Middle.BusinessObject oEntity, Ingr.SP3D.Common.Middle.BusinessObject oActiveEntity, string strIdx)
        {
            string strComputedName = "<Error>";
            try
            {
                PlateSystemBase oPlateSysBase = (PlateSystemBase)oEntity;

                // A leaf plate system is named in the same manner as the standard plate system naming rule.
                // Defer naming to the base implementation
                if (!oPlateSysBase.IsRoot)
                {
                    strComputedName = base.ComputeName(oEntity, oActiveEntity, strIdx);
                }
                else
                {
                    // Root plate system has the following format:
                    // <ship> <block> <root> <plane> <index>
                    // - ship: Name of the ship
                    // - block: Block extracted from system folder name
                    // - root: Root extracted from naming category
                    // - plane: intersecting plane
                    // - index: Naming index created by name generator
                    // Get the ship name (from the active site)
                    string strShipName = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.Name;
                    string strBlock = ((BusinessObject)oPlateSysBase.SystemParent).ToString();
                    string strRoot = GetNamingCategoryString(oPlateSysBase);
                    string strPlane = GetNamingPlaneString(oPlateSysBase);
                    strComputedName = String.Format("{0}-{1}-{2}-{3}",
                        strShipName, strBlock, strRoot, strIdx);
                }
            }
            catch (Exception e)
            {              
                MiddleServiceProvider.ErrorLogger.Log("PlateSysNCNamingRule.ComputeName: Error encountered (" + e.Message + ")");
            }
            return strComputedName;
        }

        public override System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oNamingParents = new Collection<BusinessObject>();
            try
            {
                PlateSystemBase oPlateSysBase = (PlateSystemBase)oEntity;                

                // if a root system, add its containing system as a naming parent.
                if (oPlateSysBase.IsRoot == true)
                {                    
                    oNamingParents.Add((BusinessObject)((ISystemChild)oEntity).SystemParent);
                }
                else
                {
                    // leaf systems: let the base class handle it
                    oNamingParents = base.GetNamingParents(oEntity);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlateSysNCNamingRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }

            return oNamingParents;
        }
        /// <summary>
        /// Returns the short naming category string based on the plate system's type and naming category value
        /// </summary>
        /// <param name="oPlateSysBase">The plate system for which to get the naming category string</param>
        /// <returns>The short naming category string</returns>
        protected string GetNamingCategoryString(PlateSystemBase oPlateSysBase)
        {
            string strNamingCategoryString = "Unknown";
            MetadataManager oMetaDataMgr =
                MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.MetadataMgr;
            try
            {
                string strCodeListTable = GetNamingCategoryCodelistString(oPlateSysBase);

                if (strCodeListTable != String.Empty)
                {
                    CodelistItem codelistInfo = oMetaDataMgr.GetCodelistInfo(strCodeListTable, "UDP").GetCodelistItem
                            (oPlateSysBase.NamingCategory);

                    if (codelistInfo != null)strNamingCategoryString = codelistInfo.ShortDisplayName;
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlateSysNCNamingRule.GetNamingCategoryString: Error encountered (" + e.Message + ")");
            }
            return strNamingCategoryString;
        }

        /// <summary>
        /// This method returns the codelist table for which a plate's naming category string
        /// value can be derived.  The codelist table returned depends on the current plate type.
        /// 
        /// </summary>
        /// <param name="oPlateSysBase">A plate system</param>
        /// <returns>The codelist name for the argument plate system's naming category</returns>
        protected string GetNamingCategoryCodelistString(PlateSystemBase oPlateSysBase)
        {
            string strNamingCategoryCodelistString = String.Empty;

            try
            {
                switch (oPlateSysBase.Type)
                {
                    case PlateType.BracketPlate:                                        // Brackets
                        strNamingCategoryCodelistString = "BracketCategory";
                        break;

                    case PlateType.CollarPlate:                                         // Collars
                        strNamingCategoryCodelistString = "CollarCategory";
                        break;

                    case PlateType.DeckPlate:                                           // Decks
                        strNamingCategoryCodelistString = "DeckCategory";
                        break;

                    case PlateType.EdgeReinforcement:                                   // Edge reinforcements
                        strNamingCategoryCodelistString = "EdgeReinforcementCategory";
                        break;

                    case PlateType.FlangePlate:                                         // Flange plates
                        strNamingCategoryCodelistString = "FlangeCategory";
                        break;

                    case PlateType.GeneralPlate:                                        // General plate
                        strNamingCategoryCodelistString = "GeneralPlateCategory";
                        break;

                    case PlateType.Hull:                                                // Hulls (shell category)
                        strNamingCategoryCodelistString = "ShellCategory";
                        break;

                    case PlateType.LBulkheadPlate:                                      // Longitudinal bulkheads
                        strNamingCategoryCodelistString = "LongitudinalBulkheadCategory";
                        break;

                    case PlateType.LongitudinalTube:                                    // Longitudinal tubes
                        strNamingCategoryCodelistString = "LongitudinalTubeCategory";
                        break;

                    case PlateType.TransverseTube:                                      // Transverse tubes
                        strNamingCategoryCodelistString = "TransverseTubeCategory";
                        break;

                    case PlateType.TBulkheadPlate:                                      // Transverse bulkheads
                        strNamingCategoryCodelistString = "TransverseBulkheadCategory";
                        break;

                    case PlateType.TubePlate:                                           // Tubes
                        strNamingCategoryCodelistString = "TubeCategory";
                        break;
                    
                    case PlateType.VerticalTube:                                        // Vertical tube
                        strNamingCategoryCodelistString = "VerticalTubeCategory";
                        break;
                    
                    
                    case PlateType.Standalone:                                          // No codelist for these
                    case PlateType.Sphere:                                              // Standalone value same
                    case PlateType.Torus:                                               // as unspecified
                        break;
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlateSysNCNamingRule.GetNamingCategoryCodelistString: Error encountered (" + e.Message + ")");
            }
            return strNamingCategoryCodelistString;
        }
    }
}
