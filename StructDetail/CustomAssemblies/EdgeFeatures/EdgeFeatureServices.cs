//--------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  EdgeFeatureServices.cs
//
//Abstract
//	EdgeFeatureServices is a helper class to have common method implementation for .NET selector rule, parameter rule and definition of the edge features.
//--------------------------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Helper class to have common method implementation for .NET selector rule, parameter rule and definition of the edge features.
    /// </summary>
    internal static class EdgeFeatureServices
    {
        #region Internal Methods

        /// <summary>
        /// Gets the reference side or anti reference bevel depths.
        /// </summary>
        /// <param name="edgeFeature">The edge feature for which the bevel depth will be calculated.</param>
        /// <param name="physicalConnection">Physical connection with respect to which bevel depth is calculated.</param>
        /// <param name="isReferenceSide">Boolean value to determine if the value needs to be calculated for reference side or anti reference side.</param>
        /// <returns></returns>
        internal static double GetBevelDepth(Feature edgeFeature, PhysicalConnection physicalConnection, bool isReferenceSide)
        {
            double bevelDepth = 0.0;

            int physicalConnectionSubType = StructHelper.GetIntProperty(physicalConnection.Part.PartClass, MarineSymbolConstants.IJSmartClass, MarineSymbolConstants.PartClassSubType);

            double teeMountingAngle = physicalConnection.GetConnectionAngle();

            if (physicalConnectionSubType == MarineSymbolConstants.PartClassSubTypeTeeWeld)
            {
                if (isReferenceSide)
                {
                    bevelDepth = GetReferenceSideBevelDepthForTeeWeld(physicalConnection, teeMountingAngle);
                }
                else
                {
                    bevelDepth = GetAntiReferenceSideBevelDepthForTeeWeld(physicalConnection, teeMountingAngle);
                }
            }
            else if (physicalConnectionSubType == MarineSymbolConstants.PartClassSubTypeButtWeld)
            {
                //Get the cut out part
                INamedItem partWithEdgeFeature = (INamedItem)edgeFeature.EdgePort.Connectable;

                //Checking in case of ButtWelds, whether edge feature is placed on Ref Part or Non Ref Part                
                string referencePartName = ((PropertyValueString)(physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.RefPartName))).PropValue;

                double cornerButtMountingAngle = (Math.PI / 2.0 - Math.Abs(teeMountingAngle) < Math3d.FitTolerance) ? 0.0 : teeMountingAngle;

                bool isEdgeFeatureOnReferencePart = false;
                if (string.Compare(partWithEdgeFeature.Name, referencePartName, true) == 0)
                {
                    isEdgeFeatureOnReferencePart = true;
                }

                if (isReferenceSide)
                {
                    bevelDepth = GetReferenceSideBevelDepthForButtWeld(physicalConnection, cornerButtMountingAngle, isEdgeFeatureOnReferencePart);
                }
                else
                {
                    bevelDepth = GetAntiReferenceSideBevelDepthForButtWeld(physicalConnection, cornerButtMountingAngle, isEdgeFeatureOnReferencePart);
                }
            }
            return bevelDepth;
        }

        #endregion Internal Methods

        #region Private Methods

        /// <summary>
        /// Get the reference side bevel depth for Tee weld.
        /// </summary>
        /// <param name="physicalConnection">Physical connection for which reference side bevel depth is calculated.</param>
        /// <param name="teeMountingAngle">The tee mounting angle.</param>
        private static double GetReferenceSideBevelDepthForTeeWeld(PhysicalConnection physicalConnection, double teeMountingAngle)
        {
            //Defining variables to retrieve bevel parameters
            double refSideFirstBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideFirstBevelDepth);
            double refSideFirstBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideFirstBevelAngle);
            int refSideFirstBevelMethod = StructHelper.GetIntProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideFirstBevelMethod);

            double refSideSecondBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideSecondBevelDepth);
            double refSideSecondBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideSecondBevelAngle);
            int refSideSecondBevelMethod = StructHelper.GetIntProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideSecondBevelMethod);

            if (teeMountingAngle > (Math.PI - StructHelper.DISTTOL))
            {
                teeMountingAngle = 2.0 * Math.PI - teeMountingAngle;
            }

            //Changing Bevel angle parameters from varying method to constant
            if (refSideFirstBevelMethod == (int)BevelMethod.Varying && refSideFirstBevelDepth > MarineSymbolConstants.CompareTolerance)
            {
                refSideFirstBevelAngle = Math.PI / 2.0 + refSideFirstBevelAngle - teeMountingAngle;
            }

            if (refSideSecondBevelMethod == (int)BevelMethod.Varying && refSideSecondBevelDepth > MarineSymbolConstants.CompareTolerance)
            {
                refSideSecondBevelAngle = Math.PI / 2.0 + refSideSecondBevelAngle - teeMountingAngle;
            }

            //if noseOrientationAngle is zero then rootGapCorrection will be automatically incorporated in RefSideBevelDepth or AntiRefSideBevelDepth                
            double noseOrientationAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.NoseOrientationAngle);
            double rootGapCorrection = Math.Sin(noseOrientationAngle) > MarineSymbolConstants.CompareTolerance ? (physicalConnection.RootGap) / Math.Sin(noseOrientationAngle) : 0.0;

            //Since Bevel depths are calculated as if bevel portions are projected along bounded part
            //Another correction is required to take bevel projections perpendicular to bounded part and mounting angle into account                
            double refSideBevelDepthCorrection = 0.0;
            if (!StructHelper.AreEqual(teeMountingAngle, Math.PI / 2.0, StructHelper.DISTTOL))
            {
                refSideBevelDepthCorrection = -(refSideFirstBevelDepth + refSideSecondBevelDepth) / Math.Tan(teeMountingAngle);
            }

            //RefSide Bevel Depth
            return GetAdjustedBevelDepth(refSideFirstBevelDepth, refSideFirstBevelAngle, refSideSecondBevelDepth, refSideSecondBevelAngle, rootGapCorrection, refSideBevelDepthCorrection);
        }

        /// <summary>
        /// Get the anti-reference bevel depth for Tee weld.
        /// </summary>
        /// <param name="physicalConnection">Physical connection for which anti reference side bevel depth is calculated.</param>
        /// <param name="teeMountingAngle">The tee mounting angle.</param>
        /// <returns></returns>
        private static double GetAntiReferenceSideBevelDepthForTeeWeld(PhysicalConnection physicalConnection, double teeMountingAngle)
        {
            double refSideSecondBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideSecondBevelDepth);

            double antiRefSideFirstBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth);
            double antiRefSideFirstBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle);
            int antiRefSideFirstBevelMethod = StructHelper.GetIntProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelMethod);

            double antiRefSideSecondBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideSecondBevelDepth);
            double antiRefSideSecondBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideSecondBevelAngle);
            int antiRefSideSecondBevelMethod = StructHelper.GetIntProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideSecondBevelMethod);

            //Changing Bevel angle parameters from varying method to constant
            if (antiRefSideFirstBevelMethod == (int)BevelMethod.Varying && antiRefSideFirstBevelDepth > MarineSymbolConstants.CompareTolerance)
            {
                antiRefSideFirstBevelAngle = antiRefSideFirstBevelAngle + teeMountingAngle - Math.PI / 2.0;
            }

            if (antiRefSideSecondBevelMethod == (int)BevelMethod.Varying && antiRefSideSecondBevelDepth > MarineSymbolConstants.CompareTolerance)
            {
                antiRefSideSecondBevelAngle = antiRefSideSecondBevelAngle + teeMountingAngle - Math.PI / 2.0;
            }

            if (teeMountingAngle > (Math.PI - StructHelper.DISTTOL))
            {
                teeMountingAngle = 2.0 * Math.PI - teeMountingAngle;
            }

            //if noseOrientationAngle is zero then rootGapCorrection will be automatically incorporated in RefSideBevelDepth or AntiRefSideBevelDepth                
            double noseOrientationAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.NoseOrientationAngle);
            double rootGapCorrection = Math.Sin(noseOrientationAngle) > MarineSymbolConstants.CompareTolerance ? (physicalConnection.RootGap) / Math.Sin(noseOrientationAngle) : 0.0;

            //Since Bevel depths are calculated as if bevel portions are projected along bounded part
            //Another correction is required to take bevel projections perpendicular to bounded part and mounting angle into account                
            double antiRefSideBevelDepthCorrection = 0.0;
            if (!StructHelper.AreEqual(teeMountingAngle, Math.PI / 2.0, StructHelper.DISTTOL))
            {
                antiRefSideBevelDepthCorrection = (antiRefSideFirstBevelDepth + antiRefSideSecondBevelDepth) / Math.Tan(teeMountingAngle);
            }

            //AntiRefSide Bevel Depth
            return GetAdjustedBevelDepth(antiRefSideFirstBevelDepth, antiRefSideFirstBevelAngle, refSideSecondBevelDepth, antiRefSideSecondBevelAngle, rootGapCorrection, antiRefSideBevelDepthCorrection);
        }

        /// <summary>
        /// Get the reference side bevel depth for Butt weld at edge feature.
        /// </summary>
        /// <param name="physicalConnection">Physical connection for which reference side bevel depth is calculated.</param>
        /// <param name="cornerButtMountingAngle">Corner butt mounting angle which is used to calculate the angle between plate and nose.</param>
        /// <param name="isEdgeFeatureOnReferencePart">Boolean to determine whether the edge feature is placed on reference part or non reference part.</param>
        /// <returns></returns>
        private static double GetReferenceSideBevelDepthForButtWeld(PhysicalConnection physicalConnection, double cornerButtMountingAngle, bool isEdgeFeatureOnReferencePart)
        {
            return GetBevelDepthForButtWeld(physicalConnection, cornerButtMountingAngle, isEdgeFeatureOnReferencePart, true);
        }

        /// <summary>
        /// Get the anti-reference side bevel depth for Butt weld at edge feature.
        /// </summary>
        /// <param name="physicalConnection">Physical connection for which anti reference side bevel depth is calculated.</param>
        /// <param name="cornerButtMountingAngle">Corner butt mounting angle which is used to calculate the angle between plate and nose.</param>
        /// <param name="isEdgeFeatureOnReferencePart">Boolean to determine whether the edge feature is placed on reference part or non-reference part.</param>
        /// <returns></returns>
        private static double GetAntiReferenceSideBevelDepthForButtWeld(PhysicalConnection physicalConnection, double cornerButtMountingAngle, bool isEdgeFeatureOnReferencePart)
        {
            return GetBevelDepthForButtWeld(physicalConnection, cornerButtMountingAngle, isEdgeFeatureOnReferencePart, false);
        }

        /// <summary>
        /// Get the bevel depth for Butt weld at edge feature.
        /// </summary>
        /// <param name="physicalConnection">Physical connection for which reference side and anti-reference side bevel depths are calculated.</param>
        /// <param name="cornerButtMountingAngle">Corner Butt mounting angle.</param>
        /// <param name="isEdgeFeatureOnReferencePart">Boolean to determine whether the edge feature is placed on reference part or non-reference part.</param>
        /// <param name="isReferenceSide">Boolean value to determine if the value needs to be calculated for reference side or anti reference side.</param>
        private static double GetBevelDepthForButtWeld(PhysicalConnection physicalConnection, double cornerButtMountingAngle, bool isEdgeFeatureOnReferencePart, bool isReferenceSide)
        {
            double bevelDepth = 0.0;

            double refSideFirstBevelDepth = 0.0, refSideFirstBevelAngle = 0.0, antiRefSideFirstBevelDepth = 0.0, antiRefSideFirstBevelAngle = 0.0, refSideSecondBevelDepth = 0.0, refSideSecondBevelAngle = 0.0, antiRefSideSecondBevelDepth = 0.0, antiRefSideSecondBevelAngle = 0.0;
            double noseOrientationAngle = 0.0;
            double rootGap = physicalConnection.RootGap;

            if (isEdgeFeatureOnReferencePart)
            {
                //Defining variables to retrieve bevel parameters
                refSideFirstBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideFirstBevelDepth);
                refSideFirstBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideFirstBevelAngle);
                antiRefSideFirstBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth);
                antiRefSideFirstBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle);
                refSideSecondBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideSecondBevelDepth);
                refSideSecondBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.RefSideSecondBevelAngle);
                antiRefSideSecondBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideSecondBevelDepth);
                antiRefSideSecondBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.AntiRefSideSecondBevelAngle);
                noseOrientationAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUARefPartBevelParams, DetailingCustomAssembliesConstants.NoseOrientationAngle);
            }
            else
            {
                refSideFirstBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRRefSideFirstBevelDepth);
                refSideFirstBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRRefSideFirstBevelAngle);
                antiRefSideFirstBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRAntiRefSideFirstBevelDepth);
                antiRefSideFirstBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRAntiRefSideFirstBevelAngle);
                refSideSecondBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRRefSideSecondBevelDepth);
                refSideSecondBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRRefSideSecondBevelAngle);
                antiRefSideSecondBevelDepth = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRAntiRefSideSecondBevelDepth);
                antiRefSideSecondBevelAngle = StructHelper.GetDoubleProperty(physicalConnection, DetailingCustomAssembliesConstants.IJUANonRefPartBevelParams, DetailingCustomAssembliesConstants.NRAntiRefSideSecondBevelAngle);
            }

            if (cornerButtMountingAngle > Math.PI / 2.0 - StructHelper.DISTTOL)
            {
                cornerButtMountingAngle = Math.PI - cornerButtMountingAngle;
            }

            //Angle between Plates and Nose
            double angle = (Math.PI / 2.0 - (cornerButtMountingAngle / 2.0));

            //if noseOrientationAngle is zero then rootGapCorrection will be automatically incorporated in RefSideBevelDepth or AntiRefSideBevelDepth                    
            double rootGapCorrection = Math.Sin(noseOrientationAngle) > MarineSymbolConstants.CompareTolerance ? rootGap / Math.Sin(noseOrientationAngle) : 0.0;

            //Since Bevel depths are calculated as if bevel portions are projected along bounded part
            //Another correction is required to take bevel projections perpendicular to bounded part and mounting angle into account
            double refSideBevelDepthCorrection = 0.0, antiRefSideBevelDepthCorrection = 0.0;
            if (!(StructHelper.AreEqual(angle, 0.0, StructHelper.DISTTOL) || StructHelper.AreEqual(angle, Math.PI / 2.0, StructHelper.DISTTOL)))
            {
                refSideBevelDepthCorrection = Math.Abs((refSideFirstBevelDepth + refSideSecondBevelDepth) / Math.Tan(angle));
                antiRefSideBevelDepthCorrection = Math.Abs((antiRefSideFirstBevelDepth + antiRefSideSecondBevelDepth) / Math.Tan(angle));
            }

            //checking if Base in inside face of knuckle or not
            if (physicalConnection.BoundedObject is PlatePartBase && physicalConnection.BoundingObject is PlatePartBase)
            {
                if (PlateKnuckle.GetInsideFaceContext(physicalConnection).HasFlag(ContextTypes.Base))
                {
                    refSideFirstBevelDepth = -refSideFirstBevelDepth;
                }
                else
                {
                    antiRefSideFirstBevelDepth = -antiRefSideFirstBevelDepth;
                }
            }
            else
            {
                //We will use approximation
                if (isEdgeFeatureOnReferencePart)
                {
                    if (refSideBevelDepthCorrection > antiRefSideBevelDepthCorrection - MarineSymbolConstants.CompareTolerance)
                    {
                        antiRefSideBevelDepthCorrection = refSideBevelDepthCorrection;
                    }
                    else
                    {
                        refSideBevelDepthCorrection = antiRefSideBevelDepthCorrection;
                    }
                }

                if (isReferenceSide)
                {
                    //RefSide Bevel Depth
                    bevelDepth = GetAdjustedBevelDepth(refSideFirstBevelDepth, refSideFirstBevelAngle, refSideSecondBevelDepth, refSideSecondBevelAngle, rootGapCorrection, refSideBevelDepthCorrection);
                }
                else
                {
                    //Anti-RefSide Bevel Depth
                    bevelDepth = GetAdjustedBevelDepth(antiRefSideFirstBevelDepth, antiRefSideFirstBevelAngle, antiRefSideSecondBevelDepth, antiRefSideSecondBevelAngle, rootGapCorrection, antiRefSideBevelDepthCorrection);
                }
            }

            return bevelDepth;
        }

        /// <summary>
        /// Get the adjusted bevel depth.
        /// </summary>
        /// <param name="firstBevelDepth">The ref side or anti-ref side first bevel depth.</param>
        /// <param name="firstBevelAngle">The ref side or anti-ref side first bevel angle.</param>
        /// <param name="secondBevelDepth">The ref side or anti-ref side second bevel depth.</param>
        /// <param name="secondBevelAngle">The ref side or anti-ref side second bevel angle.</param>
        /// <param name="rootGapCorrection">The root gap correction.</param>
        /// <param name="bevelDepthCorrection">The ref side or anti-ref side bevel correction.</param>
        /// <returns></returns>
        private static double GetAdjustedBevelDepth(double firstBevelDepth, double firstBevelAngle, double secondBevelDepth, double secondBevelAngle, double rootGapCorrection, double bevelDepthCorrection)
        {
            double bevelDepth = 0.0;

            if (StructHelper.AreEqual(firstBevelAngle, Math.PI / 2.0, StructHelper.DISTTOL))
            {
                bevelDepth = bevelDepth + firstBevelDepth;
            }
            else
            {
                bevelDepth = bevelDepth + firstBevelDepth * Math.Tan(firstBevelAngle);
            }

            if (StructHelper.AreEqual(secondBevelAngle, Math.PI / 2.0, StructHelper.DISTTOL))
            {
                bevelDepth = bevelDepth + secondBevelDepth;
            }
            else
            {
                bevelDepth = bevelDepth + secondBevelDepth * Math.Tan(secondBevelAngle);
            }

            bevelDepth = bevelDepth + rootGapCorrection + bevelDepthCorrection;

            return bevelDepth;
        }

        #endregion Private Methods
    }
}