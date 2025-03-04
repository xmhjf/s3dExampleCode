using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    public class ProfileSysNCNamingRule : ProfileSysNamingRule
    {
        public override string ComputeName(BusinessObject oEntity, BusinessObject oActiveEntity, string strIdx)
        {
            string strComputedName = "<Error>";
            try
            {
                StiffenerSystemBase oProfileSysBase = (StiffenerSystemBase)oEntity;

                // A leaf Profile system is named in the same manner as the standard Profile system naming rule.
                // Defer naming to the base implementation
                if (!oProfileSysBase.IsRootSystem)
                {
                    strComputedName = base.ComputeName(oEntity, oActiveEntity, strIdx);
                }
                else
                {
                    // Name takes the form of:
                    // <Parent system name>-<Profile naming category string>-<Naming Index>
                    string strParentSys = ((BusinessObject)oProfileSysBase.SystemParent).ToString();
                    string strNamingCategory = this.GetNamingCateogryString(oProfileSysBase);
                    strComputedName = String.Format("{0}-{1}-{2}",
                        strParentSys, strNamingCategory, strIdx);
                }
            }
            catch (Exception e)
            {              
                MiddleServiceProvider.ErrorLogger.Log("ProfileSysNCNamingRule.ComputeName: Error encountered (" + e.Message + ")");
            }
            return strComputedName;
        }

        public override System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oNamingParents = new Collection<BusinessObject>();
            try
            {
                StiffenerSystemBase oProfileSysBase = (StiffenerSystemBase)oEntity;                

                // if a root system, add its containing system as a naming parent.
                if (oProfileSysBase.IsRootSystem == true)
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
                MiddleServiceProvider.ErrorLogger.Log("ProfileSysNCNamingRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }

            return oNamingParents;
        }

        /// <summary>
        /// Returns the short naming category string based on the plate system's type and naming category value
        /// </summary>
        /// <param name="oPlateSysBase">The plate system for which to get the naming category string</param>
        /// <returns>The short naming category string</returns>
        protected string GetNamingCateogryString(StiffenerSystemBase oStiffenerSys)
        {
            string strNamingCategoryString = "Unknown";
            MetadataManager oMetaDataMgr =
                MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.MetadataMgr;
            try
            {
                string strCodeListTable = GetNamingCategoryCodelistString(oStiffenerSys);

                if (strCodeListTable != String.Empty)
                {
                    CodelistItem codelistInfo = oMetaDataMgr.GetCodelistInfo(strCodeListTable, "UDP").GetCodelistItem
                            (oStiffenerSys.NamingCategory);
                    if (codelistInfo != null)strNamingCategoryString = codelistInfo.ShortDisplayName;
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("ProfileSysNCNamingRule.GetNamingCategoryString: Error encountered (" + e.Message + ")");
            }
            return strNamingCategoryString;
        }

        protected string GetNamingCategoryCodelistString(StiffenerSystemBase oStiffenerSys)
        {
            string strNamingCategoryCodelistString = String.Empty;
            try
            {
                switch (oStiffenerSys.StiffenerType)
                {
                    case StiffenerType.Axial:
                        {
                            strNamingCategoryCodelistString = "AxialProfileCategory";
                            break;
                        }
                    case StiffenerType.AxialOnBuiltupFlange:
                        {
                            strNamingCategoryCodelistString = "AxialOnBUMemberFlangeCategory";
                            break;
                        }
                    case StiffenerType.AxialOnBuiltupWeb:
                        {
                            strNamingCategoryCodelistString = "AxialOnBUMemberWebCategory";
                            break;
                        }
                    case StiffenerType.EdgeReinforcement:
                        {
                            strNamingCategoryCodelistString = "EdgeReinforcementCategory";
                            break;
                        }
                    case StiffenerType.General:
                        {
                            strNamingCategoryCodelistString = "ProfileGeneralCategory";
                            break;
                        }
                    case StiffenerType.Longitudinal:
                        {
                            strNamingCategoryCodelistString = "LongitudinalProfileCategory";
                            break;
                        }
                    case StiffenerType.NonAxialOnBuiltup:
                        {
                            strNamingCategoryCodelistString = "NonAxialOnBUMemberCategory";
                            break;
                        }
                    case StiffenerType.Ring:
                        {
                            strNamingCategoryCodelistString = "RingProfileCategory";
                            break;
                        }
                    case StiffenerType.Transversal:
                        {
                            strNamingCategoryCodelistString = "TransverseProfileCategory";
                            break;
                        }
                    case StiffenerType.Vertical:
                        {
                            strNamingCategoryCodelistString = "VerticalProfileCategory";
                            break;
                        }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("ProfileSysNCNamingRule.GetNamingCategoryCodelistString: Error encountered (" + e.Message + ")");
            }

            return strNamingCategoryCodelistString;
        }
    }   
}
