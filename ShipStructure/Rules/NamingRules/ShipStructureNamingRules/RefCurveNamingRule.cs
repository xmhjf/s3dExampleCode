using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Grids.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// This is an example class from which a new reference curve naming rule can be derived.  It
    /// implements a single required interface, INamingSolverBase, which will be used by the
    /// relevant naming semantics to name a Molded Forms business object
    /// </summary>
    /// 

    // The WrapperProgID is a required attribute and should always be set to "MarineRuleWrappers.NamingSolverWrapper"
    [WrapperProgID("MarineRuleWrappers.NamingSolverWrapper")]
    public class RefCurveNamingRule : INamingSolverBase
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
        public string ComputeName(BusinessObject oEntity, BusinessObject oActiveEntity, string strIdx)
        {
            string strComputedName = "<Error>";            
            
            try
            {
                string strRefCurveType = GetRefCurveTypeString(oEntity);

                // We name differently based on whether or not the ref curve type is set to Grid.
                // However, because we don't (as of this writing) have a BO wrapper for ref curves on
                // surface, we can't find the value directly.  We have the type string (returned via the
                // codelist value from GetRefCurveTypeString), but because the value can be changed,
                // we should not assume that its value is always going to be "Grid".

                // Because of that, query the metadata and get the string value for "Grid" so we
                // can reliably compare the values.
                Catalog oCatalog = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog;
                CodelistItem gridCodeListItem = oCatalog.MetadataMgr.GetCodelistInfo
                    ("RefCrveOnSurfTypes", "STRUCT").GetCodelistItem(4);

                string strGridTypeString = gridCodeListItem != null ? gridCodeListItem.ShortDisplayName : "Grid";


                // Reference plane may be null if this is an imported reference plane
                GridPlaneBase refPlane = (GridPlaneBase)GetRefPlane(oEntity);

                if (refPlane != null && (strGridTypeString == strRefCurveType))
                {                    
                    // Grids have the naming convention CS:RefPlane:Type-NameIndex
                    strComputedName = refPlane.Axis.CoordinateSystem + 
                        ":" + refPlane.Name + ":" + strRefCurveType + "-" + strIdx;
                }
                else
                {
                    // all other reference curves are a concatenation of the ref curve type
                    // and the naming index.
                    strComputedName = strRefCurveType + "-" + strIdx;
                }

            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("RefCurveNamingRule.ComputeName: Error encountered (" + e.Message + ")");
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
        public System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oRetColl = new Collection<BusinessObject> ();

            try
            {
                // Reference curves are indexed according to their plate system (the design parent)
                // as well (in the case of grid type), the coordinate system and reference plane.
                //
                // Because the reference curve type may change at any time, and the naming parents can
                // only be configured at the time the naming rule is assigned, add all of these naming 
                // references now
                //
                // Note that only reference curves placed by intersection will have a reference plane
                // and by extension a coordinate system.  Check the definition of the reference curve
                // before trying to get the plane information.



                BusinessObject oRefPlane = GetRefPlane(oEntity);                                
                if (oRefPlane != null)
                {
                    oRetColl.Add (oRefPlane);   // add the reference plane
                    try
                    {
                        oRetColl.Add (((GridPlaneBase)oRefPlane).Axis.CoordinateSystem);
                    }
                    catch (Exception e)
                    {
                        MiddleServiceProvider.ErrorLogger.Log ("RefCurveNamingRule.GetNamingParents: Error encountered while " +
                            "Getting ref curve coordinate system: " + e.Message);
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("RefCurveNamingRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oRetColl;
        }

        #endregion

        /// <summary>
        /// Given a reference curve
        /// </summary>
        /// <param name="oBO"></param>
        /// <returns></returns>
        private BusinessObject GetRefPlane(BusinessObject oBO)
        {
            BusinessObject oAE = GetLandingCurveAE(oBO);

            if (oAE == null) return null;

            // Grid plane, if it exists, is the origin of the LandCrvIntParent1 relationship
            ReadOnlyCollection<BusinessObject> oIntParentTargets = 
                oAE.GetRelationship("LandCrvIntParent1", "LandCrvIntParent1_ORIG").TargetObjects;

            // No targets?  return null
            if (oIntParentTargets.Count == 0) return null;

            return oIntParentTargets[0];
        }

        /// <summary>
        /// Given a business object, get the active entity responsible for creating the
        /// landing curve
        /// </summary>
        /// <param name="oBO">BusinessObject (seam, reference curve, etc.) containing a landing curve</param>
        /// <returns>The active entity if it has one or NULL otherwise</returns>
        private BusinessObject GetLandingCurveAE(BusinessObject oBO)
        {
            // From the BusinessObject traverset the StructToGeometry relationship.  This should yield
            // the landing curve
            ReadOnlyCollection<BusinessObject> oStructToGeomTargets =
                oBO.GetRelationship("StructToGeometry", "StructToGeometry_DEST").TargetObjects;

            // No target object?  Return NULL
            if (oStructToGeomTargets.Count == 0) return null;

            // From the landing curve, get the geometry generation result
            ReadOnlyCollection<BusinessObject> oGeomGenTargets =
            oStructToGeomTargets[0].GetRelationship("GeometryGeneration_RSLT1",
                "GeometryGeneration_RSLT1_DEST").TargetObjects;

            // No target object?  Return NULL
            if (oGeomGenTargets.Count == 0) return null;

            return oGeomGenTargets[0];
        }

        private string GetRefCurveTypeString(BusinessObject oRefCurve)
        {
            string strResult = "Unknown";

            try
            {
                // At the time of this writing, there was not a reference curve BO wrapper.
                // Work around this by accessing the attributes indirectly via BusinessObject.

                // Since the reference curve type is codelisted, the GetPropertyValue method also
                // has the added benefit of returning the short type string for us, which is exactly
                // what we want.
                
                strResult = oRefCurve.GetPropertyValue("IJRefCurveOnSurface", "CurveType").ToString();
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("RefCurveNamingRule.GetRefCurveTypeString: Error encountered (" + e.Message + ")");
            }

            return strResult;
        }
    }
}
