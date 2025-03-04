//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Assembly Range class 
//                
//                 
//                    
//
//      History:
//      April 24th, 2015   Created by VIBGYOR
//
//-----------------------------------------------------------------------------------

using System;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// AssemblyRange rule class Implements PlnAssemblyRangeRuleBase
    /// </summary>
    public class AssemblyRange : PlnAssemblyRangeRuleBase
    {
        #region Methods

        /// <summary>
        /// Computes Assembly range based on input data
        /// </summary>
        /// <param name="rangeData"></param>
        /// <returns>void</returns>
        public override void Evaluate(RangeData rangeData)
        {
            #region Input Error Handling

            if (rangeData == null || rangeData.Assembly == null)
            {
                throw new ArgumentNullException("rangeData Information is null");
            }

            #endregion //Input Error Handling

            try
            {
                Vector rangeVecX, rangeVecY, rangeVecZ;
                ComputeAssemblyRangeData(rangeData, out rangeVecX, out rangeVecY, out rangeVecZ);

                // Add Assembly Shrinkage, Assembly Margin to Range
                AddMarginAndShrinkageToAssemblyRange(rangeData, rangeVecX, rangeVecY, rangeVecZ);
                
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        /// <summary>
        /// Add Assembly Margin And Shrinkage To Assembly Range
        /// </summary>
        /// <param name="rangeData"></param>
        /// <param name="rangeVecX"></param>
        /// <param name="rangeVecY"></param>
        /// <param name="rangeVecZ"></param>
        /// <returns></returns>
        private void AddMarginAndShrinkageToAssemblyRange(RangeData rangeData, Vector rangeVecX, Vector rangeVecY, Vector rangeVecZ)
        {
            // Get the Range Type IDs from the Assembly Range table in catalog
            int rangeInput1 = 0, rangeInput2 = 0, rangeInput3 = 0, rangeInput4 = 0, rangeInput5 = 0;
            RunQueryOnAssemblyRangeLookUp(rangeData.RangeTypeID, ref rangeInput1, ref rangeInput2, ref rangeInput3, ref rangeInput4, ref rangeInput5);

            // Apply Assembly Shrinkage, Assembly Margin
            if (rangeInput3 == 1 | rangeInput4 == 1)
            {
                rangeVecX.Length = 1;
                rangeVecY.Length = 1;
                rangeVecZ.Length = 1;

                int count = 0, index = 0;
                ReadOnlyCollection<BusinessObject> mfgObjColl = null;

                // Apply Assembly Shrinkage
                if (rangeInput3 == 1)
                {
                    RelationCollection relCollection = rangeData.Assembly.GetRelationship("StrMfgHierarchy", "IsStrMfgChildOf");
                    mfgObjColl = relCollection.TargetObjects;
                    count = mfgObjColl.Count;
                        
                    for (index = 0; index < count; index++)
                    {
                        if (mfgObjColl[index].GetType() == typeof(Shrinkage))
                        {
                            Shrinkage shrinkageBO = mfgObjColl[index] as Shrinkage;

                            if (shrinkageBO != null)
                            {
                                // Get the Assembly Shrinkage factors
                                double primaryFactor = shrinkageBO.PrimaryFactor;
                                double secondaryFactor = shrinkageBO.SecondaryFactor;
                                double tertiaryFactor = shrinkageBO.TertiaryFactor;

                                // Calculate the shrinkage factor to be applied in box orintation
                                double xDirShrFactor = CalculateFactorInRangeBoxVecDirection(rangeVecX, primaryFactor, secondaryFactor, tertiaryFactor);
                                double yDirShrFactor = CalculateFactorInRangeBoxVecDirection(rangeVecY, primaryFactor, secondaryFactor, tertiaryFactor);
                                double zDirShrFactor = CalculateFactorInRangeBoxVecDirection(rangeVecZ, primaryFactor, secondaryFactor, tertiaryFactor);

                                // Calculate new dimensions based on shrinkage factors in box orientation
                                rangeData.Length = rangeData.Length * xDirShrFactor;
                                rangeData.Width = rangeData.Width * yDirShrFactor;
                                rangeData.Height = rangeData.Height * zDirShrFactor;
                            }
                        }
                    }
                }

                // Apply Assembly Margin
                if (rangeInput4 == 1)
                {
                    RelationCollection relCollection = rangeData.Assembly.GetRelationship("StrMfgHierarchy", "IsStrMfgChildOf");
                    mfgObjColl = null;
                    mfgObjColl = relCollection.TargetObjects;
                    count = mfgObjColl.Count;

                    for (index = 0; index < count; index++)
                    {
                        if (mfgObjColl[index].GetType() == typeof(MarginAssembly))
                        {
                            MarginAssembly assemblyMargin = mfgObjColl[index] as MarginAssembly;

                            if (assemblyMargin != null)
                            {
                                Collection<Margin> assyMarginChildren = assemblyMargin.Children;

                                if (assyMarginChildren != null)
                                {
                                    double marginValue = 0, maxMarginVal = 0;
                                    int marginCnt = assyMarginChildren.Count;
                                    TopologyPort relatedPort = null;

                                    for (int idx = 0; idx < marginCnt; idx++)
                                    {
                                        Margin margin = assyMarginChildren[idx];
                                        marginValue = margin.StartValue;

                                        if (maxMarginVal < marginValue)
                                        {
                                            relatedPort = margin.Port;
                                            maxMarginVal = marginValue;
                                        }
                                    }

                                    // Get the port normal to apply margin on the range box in the port normal direction
                                    ISurface surfaceBody = relatedPort;
                                    
                                    if (surfaceBody != null)
                                    {
                                        Position cog = surfaceBody.Centroid();
                                        Vector portNormalVec = surfaceBody.OutwardNormalAtPoint(cog);

                                        portNormalVec.Length = 1;

                                        // Compute Margin value in each direction of range box
                                        double rangeXDirMargin = CalculateMarginValInRangeBoxVecDirection(rangeVecX, portNormalVec, maxMarginVal);
                                        double rangeYDirMargin = CalculateMarginValInRangeBoxVecDirection(rangeVecY, portNormalVec, maxMarginVal);
                                        double rangeZDirMargin = CalculateMarginValInRangeBoxVecDirection(rangeVecZ, portNormalVec, maxMarginVal);

                                        // Calculate new dimensions based on Margin values in box orientation
                                        rangeData.Length = rangeData.Length + rangeXDirMargin;
                                        rangeData.Width = rangeData.Width + rangeYDirMargin;
                                        rangeData.Height = rangeData.Height + rangeZDirMargin;
                                    }                                    
                                }
                            }
                        }
                    }
                }
            }
        }

        private double CalculateFactorInRangeBoxVecDirection(Vector rangeVec, double xFactor, double yFactor, double zFactor)
        {
            Vector xVector = new Vector(1, 0, 0);
            Vector yVector = new Vector(0, 1, 0);
            Vector zVector = new Vector(0, 0, 1);

            // Calculate dot product
            double dotProductX = System.Math.Abs(rangeVec.Dot(xVector));
            double dotProductY = System.Math.Abs(rangeVec.Dot(yVector));
            double dotProductZ = System.Math.Abs(rangeVec.Dot(zVector));

            // Get the component of the length in three directions...
            // Use the above cos of angle and multiple it with total lengths to get the component in each of the directions (dCompTotalLength)
            // dxCompTotalLength = x component of length plus added shrinkage in x direction
            double dXCompTotalLength = dotProductX + (dotProductX * xFactor);
            double dYCompTotalLength = dotProductY + (dotProductY * yFactor);
            double dZCompTotalLength = dotProductZ + (dotProductZ * zFactor);

            // Total primary port length after adding shrinkage = sqrt(sum of squares of components)
            return System.Math.Sqrt((dXCompTotalLength * dXCompTotalLength) + (dYCompTotalLength * dYCompTotalLength) + (dZCompTotalLength * dZCompTotalLength));
        }

        private double CalculateMarginValInRangeBoxVecDirection(Vector rangeBoxVec, Vector portNormalVec, double maxMarginVal)
        {
            // dproduct is value of cos(angle between primary dir and Port normal vector)
            double dotProduct = System.Math.Abs(rangeBoxVec.Dot(portNormalVec));

            // Return component of length plus added Margin in Range box vec direction
            return dotProduct * maxMarginVal;
        }

        # endregion
    }
}
