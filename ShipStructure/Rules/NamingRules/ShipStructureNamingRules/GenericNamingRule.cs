using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// This is an example class from which a new generic naming rule can be derived.  It
    /// implements a single required interface, INamingSolverBase, which will be used by the
    /// relevant naming semantics to name a Molded Forms business object
    /// 
    /// This naming rule is relatively simple.  The name is formed by the combination of the
    /// business object's design parent, its type string, and a naming index.
    /// </summary>
    /// 

    // The WrapperProgID is a required attribute and should always be set to "MarineRuleWrappers.NamingSolverWrapper"
    [WrapperProgID("MarineRuleWrappers.NamingSolverWrapper")]
    public class GenericNamingRule : INamingSolverBase
    {

        #region INamingSolverBase Members

        public string ComputeName(BusinessObject oEntity, BusinessObject oActiveEntity, string strIdx)
        {
            string strComputedName = "<Error>";
            try
            {
                // Get the type string of the business object
                string strBOTypeStr = oEntity.ClassInfo.DisplayName;

                // get the type string of the design parent
                string strDesignParentTypeStr = GetDesignParent(oEntity).ClassInfo.DisplayName;

                // get the name of the design parent
                string strDesignParentStr = GetDesignParent(oEntity).ToString();

                // Special case: this object's parent is the same type as
                // its design parent (such as the case with leaf logical connection)
                if (strBOTypeStr == strDesignParentTypeStr)
                {
                    // Look at the last character in the design parent's name.
                    // If it ends in a digit, a dot ('.') separates that digit
                    // from the naming index we will to append to the end of the
                    // computed string
                    if (Char.IsDigit(strDesignParentStr[strDesignParentStr.Length - 1]))
                    {
                        strIdx = "." + strIdx;
                    }

                    // Don't repeat the type string since the design parent is the
                    // same type.  Just make the type string blank
                    strBOTypeStr = "";
                }
                else
                {
                    switch (strBOTypeStr)
                    {

                        case "StructConnection":            // Naming a logical connection
                            strBOTypeStr = "LC";
                            break;

                        case "AssemblyConnection":          // Naming an assembly connection
                            strBOTypeStr = "AC";
                            break;

                        case "StructPhysicalConnection":    // Naming a physical connection
                            strBOTypeStr = "PC";
                            break;

                        // special case: the type string of a seam should reflect its type.
                        // defer this to a private method
                        case "Seam":
                        case "SeamPoint":
                            strBOTypeStr = GetSeamTypeString(oEntity) + strBOTypeStr;
                            break;
                    }
                }
                // output string takes the following format:
                // <design parent>-<type string><index>
                strComputedName = String.Format("{0}{1}{2}",
                    strDesignParentStr,
                    //If the BO type string is blank, don't insert the hyphen ('-')
                    (strBOTypeStr == String.Empty ? String.Empty : "-" + strBOTypeStr),
                    strIdx);

                return (strComputedName);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("GenericNamingRule.ComputeName: Error encountered (" + e.Message + ")");
            }

            return strComputedName;
        }

        public Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {


            // Return the design parent in a new collection.
            Collection<BusinessObject> oRetColl = new Collection<BusinessObject>();

            try
            {
                oRetColl.Add(GetDesignParent(oEntity));
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("GenericNamingRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oRetColl;
        }
        #endregion

        private BusinessObject GetDesignParent(BusinessObject oEntity)
        {
            // use the ShpStrDesignHierarchy interface to find this business object's design parent.
            // We have to use this in lieu of ISystemChild because some FCBOs do not yet have BO wrappers
            RelationCollection oHierarchyColl = oEntity.GetRelationship
                ("ShpStrDesignHierarchy", "ShpStrDesignParent");
            if (oHierarchyColl.TargetObjects.Count > 0)
            {
                return oHierarchyColl.TargetObjects[0];
            }
            else throw new Exception("Could not find a suitable design parent");
        }

        private string GetSeamTypeString(BusinessObject oEntity)
        {
            string strSeamType = "X";

            SeamBase oSeam = (SeamBase)oEntity;

            switch (oSeam.SeamType)
            {
                case SeamType.Design:       // Design seam
                    strSeamType = "D";
                    break;
                
                case SeamType.Planning:     // Planning seam
                    strSeamType = "P";
                    break;

                case SeamType.Intersection: // Intersection seam
                    strSeamType = "I";
                    break;

                case SeamType.Straking:     // Straking seam
                    strSeamType = "S";
                    break;
            }
            return strSeamType;
        }
    }
}