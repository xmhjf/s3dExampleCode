using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;
using System.Collections.ObjectModel;
namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// This is an example class from which a new profile system naming rule can be derived.  It
    /// implements a single required interface, INamingSolverBase, which will be used by the
    /// relevant naming semantics to name a Molded Forms business object
    /// </summary>
    /// 

    // The WrapperProgID is a required attribute and should always be set to "MarineRuleWrappers.NamingSolverWrapper"
    [WrapperProgID("MarineRuleWrappers.NamingSolverWrapper")]
    public class ProfileSysNamingRule : INamingSolverBase
    {
        #region INamingSolverBase Members

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oEntity">Business object to be named</param>
        /// <param name="oActiveEntity">Naming active entity associated with the business object 
        /// to be named.  If no active entity exists, this will contain the design parent of
        /// the the business object to be named</param>
        /// <param name="strIdx">String index (1-based) uniquely identifying the business object based
        /// within its associated naming active entity or its design parent.  This could be optionally
        /// incorporated into the returned string</param>
        /// <returns>A string value used to name the business object</returns>
        virtual public string ComputeName(BusinessObject oEntity, BusinessObject oActiveEntity, string strIdx)
        {
            string strComputedName = "<Error>";            
            
            try
            {
                StiffenerSystemBase oProfileSys = (StiffenerSystemBase)oEntity;

                // Name takes the form of:
                // <Parent system name>-<Naming Index><Profile type string>-<Location ID>
                strComputedName = String.Format("{0}-{1}-{2}",
                    ((ISystemChild)oEntity).SystemParent,
                    strIdx + (oProfileSys.IsRootSystem ? 
                        GetProfileTypeString(oProfileSys) :
                        String.Empty
                    ),
                    MiddleServiceProvider.SiteMgr.ActiveSite.ActiveLocation.NamingString
                );

                string strProfileType = GetProfileTypeString(oProfileSys);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("ProfileSysNamingRule.ComputeName: Error encountered (" + e.Message + ")");
            }

            return strComputedName;
        }

        /// <summary>
        /// When a naming rule references one or more objects in order to calculate
        /// the name of another object, these should be specified via the GetNamingParents method.
        /// 
        /// In this rule, a root plate system has its defining plane as its only naming parent.  A
        /// leaf plate system will return its design parent
        /// </summary>
        /// <param name="oEntity">The business object to be named</param>
        /// <returns>A collection of naming parents for which an active entity will be created
        /// and associated</returns>
        virtual public System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oRetColl = new Collection<BusinessObject> ();

            try
            {

                // Profile system uses the stiffened plate (its design parent)
                // as a naming reference if it is a root system and the root profile
                // system if it is a leaf profile system.  In either case, the 
                // naming parent is also the design parent.
                oRetColl.Add((BusinessObject)(((ISystemChild)oEntity).SystemParent));
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("ProfileSysNamingRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oRetColl;
        }

        #endregion

        /// <summary>
        /// Private method returns a short string reflecting the profile type
        /// </summary>
        /// <param name="oProfileSys">Profile system for which to check the profile type</param>
        /// <returns>Short string used to name the profile system</returns>
        protected string GetProfileTypeString(StiffenerSystemBase oProfileSys)
        {
            string strProfileType = "Unknown Profile Type";

            switch (oProfileSys.StiffenerType)
            {
                case StiffenerType.Vertical:                // Vertical
                    strProfileType = "V";
                    break;

                case StiffenerType.Transversal:             // Transverse
                    strProfileType = "T";
                    break;

                case StiffenerType.Longitudinal:            // Longitudinal
                    strProfileType = "L";
                    break;

                case StiffenerType.Ring:                    // Ring
                    strProfileType = "R";
                    break;

                case StiffenerType.Axial:                   // Axial
                    strProfileType = "A";
                    break;

                case StiffenerType.AxialOnBuiltupWeb:       // Axial on builtup web
                    strProfileType = "W";
                    break;

                case StiffenerType.AxialOnBuiltupFlange:    // Axial on builtup flange
                    strProfileType = "F";
                    break;

                case StiffenerType.NonAxialOnBuiltup:       // Nonaxial on builtup
                    strProfileType = "R";
                    break;

                case StiffenerType.General:                 // General
                    strProfileType = "G";
                    break;

                default:                                    // Everything else (edge reinforcement)
                    strProfileType = "X";
                    break;
            }
            return strProfileType;
        }
    }
}
