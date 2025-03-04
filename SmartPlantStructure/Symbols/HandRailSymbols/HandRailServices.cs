//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//Abstract : Service methods related to handrail symbols
//
//History:
//  Feb 16, 2015    Ninad   DI-CP-267808  Implement content changes for support of drop of Handrail  
//
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
using System;
using Ingr.SP3D.Content.Structure;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Provides service methods that are used by handrail symbols.
    /// </summary>
    public class HandrailServices
    {

        /// <summary>
        /// Get the maximum possible distance covered by the cross-section based on the selected cardinal point and the angle of rotation.
        /// </summary>
        /// <param name="crossSection">handrail crosssection</param>
        /// <param name="cardinalPoint">cardinal point</param>
        /// <param name="crossSectionAngle">rotation angle</param>
        /// <param name="sectionWidth">width</param>
        /// <param name="sectionHeight">Section height</param>
        /// <param name="deltaHeight">delta height</param>
        /// <param name="centerX">X-coordinate for center</param>
        /// <param name="centerY">Y-coordinate for center</param>
        /// <remarks></remarks>
        static internal double GetMaximumProjectedDimensionForSection(CrossSection crossSection, int cardinalPoint, double crossSectionAngle, ref double centerX, ref double centerY)
        {
            if (crossSection == null)
            {
                throw new ArgumentNullException("crossSection");
            }
            double deltaHeight = 0.0;
            double cogX = 0;
            double cogY = 0;
            double xp = 0;
            double yp = 0;
            double topY = 0.0;
            double bottomY = 0.0;
            double bottomRightX = 0.0;
            double bottomLeftX = 0.0;
            double centroidX = 0;
            double centroidY = 0;
            double sectionWidth = 0;
            double sectionHeight = 0;

            // Get toprail section depth and width   
            if ((crossSection != null))
            {
                sectionHeight = crossSection.Depth;
                sectionWidth = crossSection.Width;
            }

            //Getting the centroid of the given cross section. This will be used only when cardinal points 10, 11, 12, 13 and 14 
            if (crossSection.SupportsInterface(SPSSymbolConstants.IStructCrossSectionDesignProperties))
            {
                cogX = StructHelper.GetDoubleProperty(crossSection, SPSSymbolConstants.IStructCrossSectionDesignProperties, SPSSymbolConstants.CentroidX);
                cogY = StructHelper.GetDoubleProperty(crossSection, SPSSymbolConstants.IStructCrossSectionDesignProperties, SPSSymbolConstants.CentroidY);
            }
            //Getting the shear center of the given cross section. This will be used only when cardinal points 15 
            if (crossSection.SupportsInterface(SPSSymbolConstants.IJUAL))
            {
                xp = StructHelper.GetDoubleProperty(crossSection, SPSSymbolConstants.IJUAL, SPSSymbolConstants.xp);
                yp = StructHelper.GetDoubleProperty(crossSection, SPSSymbolConstants.IJUAL, SPSSymbolConstants.yp);
            }
            // Checking weather the data is got from the custom properties,  'if not then assume the points are lie at distance half of either cross section width or depth or both. 
            if (StructHelper.AreEqual(cogX, 0.0) == true)
            {
                cogX = sectionWidth / 2;
            }
            if (StructHelper.AreEqual(cogY, 0.0) == true)
            {
                cogY = sectionHeight / 2;
            }
            if (StructHelper.AreEqual(xp, 0.0) == true)
            {
                xp = sectionWidth / 2;
            }
            if (StructHelper.AreEqual(yp, 0.0) == true)
            {
                yp = sectionHeight / 2;
            }
            switch (cardinalPoint)
            {

                case 1:
                    // Bottom Left   
                    bottomRightX = sectionWidth;
                    topY = sectionHeight;
                    centroidX = sectionWidth / 2;
                    centroidY = sectionHeight / 2;
                    break; 
                case 2:
                    // Bottom Center    
                    bottomRightX = sectionWidth / 2;
                    bottomLeftX = -sectionWidth / 2;
                    topY = sectionHeight;
                    centroidX = 0;
                    centroidY = sectionHeight / 2;
                    break; 
                case 3:
                    //Bottom Right   
                    bottomLeftX = -sectionWidth;
                    topY = sectionHeight;
                    centroidX = -sectionWidth / 2;
                    centroidY = sectionHeight / 2;
                    break; 
                case 4:
                    //Center Left    
                    bottomRightX = sectionWidth;
                    topY = sectionHeight / 2;
                    bottomY = -sectionHeight / 2;
                    centroidX = sectionWidth / 2;
                    centroidY = 0;
                    break; 
                case 5:
                    //Center   
                    bottomRightX = sectionWidth / 2;
                    bottomLeftX = -sectionWidth / 2;
                    topY = sectionHeight / 2;
                    bottomY = -sectionHeight / 2;
                    centroidX = 0;
                    centroidY = 0;
                    break; 
                case 6:
                    //Center Right    
                    bottomLeftX = -sectionWidth;
                    topY = sectionHeight / 2;
                    bottomY = -sectionHeight / 2;
                    centroidX = -sectionWidth / 2;
                    centroidY = 0;
                    break; 
                case 7:
                    //Top Left   
                    bottomRightX = sectionWidth;
                    bottomY = -sectionHeight;
                    centroidX = sectionWidth / 2;
                    centroidY = -sectionHeight / 2;
                    break; 
                case 8:
                    //Top Center    
                    bottomRightX = sectionWidth / 2;
                    bottomLeftX = -sectionWidth / 2;
                    bottomY = -sectionHeight;
                    centroidX = 0;
                    centroidY = -sectionHeight / 2;
                    break; 
                case 9:
                    //Top Right    
                    bottomLeftX = -sectionWidth;
                    bottomY = -sectionHeight;
                    centroidX = -sectionWidth / 2;
                    centroidY = -sectionHeight / 2;
                    break; 
                case 10:
                    //Centroid    
                    bottomRightX = sectionWidth - cogX;
                    bottomLeftX = -cogX;
                    topY = sectionHeight - cogY;
                    bottomY = -cogY;
                    centroidX = sectionWidth / 2 - cogX;
                    centroidY = sectionHeight / 2 - cogY;
                    break; 
                case 11:
                    //Centroid Bottom    
                    bottomRightX = sectionWidth - cogX;
                    bottomLeftX = -cogX;
                    topY = sectionHeight;
                    centroidX = sectionWidth / 2 - cogX;
                    centroidY = sectionHeight / 2;
                    break; 
                case 12:
                    //Centroid Left   
                    bottomRightX = sectionWidth;
                    topY = sectionHeight - cogY;
                    bottomY = -cogY;
                    centroidX = sectionWidth / 2;
                    centroidY = sectionHeight / 2 - cogY;
                    break; 
                case 13:
                    //Centroid Right  
                    bottomLeftX = -sectionWidth;
                    topY = sectionHeight - cogY;
                    bottomY = -cogY;
                    centroidX = -sectionWidth / 2;
                    centroidY = sectionHeight / 2 - cogY;
                    break; 
                case 14:
                    //Centroid Top   
                    bottomRightX = sectionWidth - cogX;
                    bottomLeftX = -cogX;
                    bottomY = -sectionHeight;
                    centroidX = sectionWidth / 2 - cogX;
                    centroidY = -sectionHeight / 2;
                    break; 
                case 15:
                    //Shear Center    
                    bottomRightX = sectionWidth - xp;
                    bottomLeftX = -xp;
                    topY = sectionHeight - yp;
                    bottomY = -yp;
                    centroidX = sectionWidth / 2 - xp;
                    centroidY = sectionHeight / 2 - yp;
                    break; 
            }
            //Apply rotation to get the actual width and height
            double height = topY * Math.Cos(crossSectionAngle) - bottomRightX * Math.Sin(crossSectionAngle);
            double width = bottomY * Math.Cos(crossSectionAngle) - bottomLeftX * Math.Sin(crossSectionAngle);
            //get the center of the cross section after applying rotation angle.
            centerX = centroidX * Math.Cos(crossSectionAngle) + centroidY * Math.Sin(crossSectionAngle);
            centerY = centroidY * Math.Cos(crossSectionAngle) - centroidX * Math.Sin(crossSectionAngle);
            //return the maximum value amongest
            if (height < width)
            {
                deltaHeight = width;
            }
            else
            {
                deltaHeight = height;
            }

            return deltaHeight;
        }

        #region "Create Rail Curve"
        /// <summary>
        /// This method creates the geometry for the given rail curve.
        /// </summary>
        /// <param name="traceCurve">Handrail path curve.</param>
        /// <param name="railBeginOffset">The rail begin offset.</param>
        /// <param name="railEndOffset">The rail end offset.</param>
        /// <param name="beginTreatmentType">Handrail begin treatment type.</param>
        /// <param name="endTreatmentType">Handrail end treatment type.</param>
        /// <returns>Geometry corresponding to the given rail curve.</returns>
        /// <exception cref="Ingr.SP3D.Common.Middle.CmnException">
        /// </exception>
        static internal ComplexString3d CreateRailCurve(ComplexString3d traceCurve, double railBeginOffset, double railEndOffset, int beginTreatmentType, int endTreatmentType)
        {
            // Check if the arguments are null
            if (traceCurve == null)
            {
                throw new ArgumentNullException("traceCurve");
            }

            Collection<ICurve> curves = null;
            ComplexString3d tempComplexString = null;

            // Validate tracecurve
            if (!IsTraceCurveValid(traceCurve, out curves))
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves, "Handrail path for creation of rail is invalid as it does not have any geometry."));
            }

            //offset the first and last segments of the input curve
            if (beginTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT |
                endTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
            {
                Line3d firstLineSegmentOfEndTreatemnt = null;
                Line3d lastLineSegmentOfEndTreatemnt = null;
                Arc3d firstArcSegmentOfEndTreatemnt = null;
                Arc3d lastArcSegmentOfEndTreatment = null;

                ICurve firstSegment = curves[0];
                int curveCount = curves.Count - 1;
                ICurve lastSegment = curves[curveCount];

                //if the last segment of the handrail is a line then we need to see if 
                //the first segment is either a line or an arc and cast it to the appropriate object.
                //This code applies the same to one or multiple segment handrails.  If it is only a one
                //segment handrail then lastSegment is pointing to the same item in the collection as firstSegment thus
                //allowing us to utilize the same code regardless of number of segments.
                if (lastSegment is ILine)
                {
                    if (beginTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
                    {
                        if (firstSegment is ILine)
                        {
                            firstLineSegmentOfEndTreatemnt = (Line3d)firstSegment;
                            firstLineSegmentOfEndTreatemnt.TrimAtStart(railBeginOffset);
                            //True
                            curves[0] = firstLineSegmentOfEndTreatemnt;
                        }
                        else
                        {
                            firstArcSegmentOfEndTreatemnt = (Arc3d)firstSegment;
                            firstArcSegmentOfEndTreatemnt.TrimAtStart(railBeginOffset);
                            curves[0] = firstArcSegmentOfEndTreatemnt;
                        }
                    }

                    //since the last segment is a line we cast to Line3d and if a circular end treatment is applied
                    //then we change the segment in the collection to the new returned ICurve from the GetEndCurve method
                    //this basically changes the end point of the curve.  If the handrail is only a one segment handrail then
                    //this object in the collection is the same object as the object in index 0 since curveCount will be 0.
                    //This change in code allow us to reuse the code regardless if we have one or multiple segments in the handrail.
                    lastLineSegmentOfEndTreatemnt = (Line3d)lastSegment;
                    if (endTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
                    {
                        lastLineSegmentOfEndTreatemnt.TrimAtEnd(railEndOffset);
                        curves[curveCount] = lastLineSegmentOfEndTreatemnt;
                    }
                    //if the lastSegment is an arc then the same logic applies.
                }
                else if (lastSegment is Arc3d)
                {
                    if (beginTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
                    {
                        if (firstSegment is ILine)
                        {
                            firstLineSegmentOfEndTreatemnt = (Line3d)firstSegment;
                            firstLineSegmentOfEndTreatemnt.TrimAtStart(railBeginOffset);
                            curves[0] = firstLineSegmentOfEndTreatemnt;
                        }
                        else
                        {
                            firstArcSegmentOfEndTreatemnt = (Arc3d)firstSegment;
                            firstArcSegmentOfEndTreatemnt.TrimAtStart(railBeginOffset);
                            curves[0] = firstArcSegmentOfEndTreatemnt;
                        }
                    }

                    lastArcSegmentOfEndTreatment = (Arc3d)lastSegment;
                    if (endTreatmentType == SPSSymbolConstants.CIRCULAR_END_TREATMENT)
                    {
                        lastArcSegmentOfEndTreatment.TrimAtEnd(railEndOffset);
                        curves[curveCount] = lastArcSegmentOfEndTreatment;
                    }
                }

                if (lastLineSegmentOfEndTreatemnt == null & lastArcSegmentOfEndTreatment == null)
                {
                    throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrInvalidPath, "Handrail path segment is not a line nor an arc. Edit handrail path and redefine all path segments to be either a line or an arc."));
                }

                tempComplexString = new ComplexString3d(curves);
            }
            else
            {
                tempComplexString = new ComplexString3d(curves);
            }
            

            return tempComplexString;
        }
        #endregion

        #region "Create Toprail Curve"

        /// <summary>
        /// This method offsets the given toprail curve and creates the corresponding geometry.
        /// </summary>
        /// <param name="traceCurve">Handrail path curve.</param>
        /// <param name="height">Handrail height.</param>
        /// <param name="orientation">Handrail orientation.</param>
        /// <param name="toprailBeginOffset">The toprail begin offset.</param>
        /// <param name="toprailEndOffset">The toprail end offset.</param>
        /// <param name="beginTreatmentType">Handrail begin treatment type.</param>
        /// <param name="endTreatmentType">Handrail end treatment type.</param>
        /// <returns>Geometry corresponding to the given toprail.</returns>
        /// <exception cref="System.ArgumentNullException">Raised if any of the given arguments is null.</exception>
        /// <exception cref="Ingr.SP3D.Common.Middle.CmnException">Raised if the given handrail orientation is not supported. </exception>
        static internal ComplexString3d CreateToprailCurve(ComplexString3d traceCurve, double height, HandrailPostOrientation orientation, 
                                                            double toprailBeginOffset, double toprailEndOffset, int beginTreatmentType, int endTreatmentType)
        {
            // Check if the arguments are null
            if (traceCurve == null)
            {
                throw new ArgumentNullException("traceCurve");
            }

            Collection<ICurve> curves = null;
            // Validate tracecurve
            if (!IsTraceCurveValid(traceCurve, out curves))
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves, "Handrail path for creation of top rail is invalid as it does not have any geometry."));
            }
            // Initialize local variables
            ComplexString3d offsetCurve = null;

            //if orientation is perpendicular, then create another curve with the given offset normal to curve plane
            if ((int)orientation == (int)SPSSymbolConstants.PERPENDICULAR_ORIENTATION)
            {
                offsetCurve = (ComplexString3d)traceCurve.GetOffsetCurve(height, OffsetDirection.Normal);
            }
            else if (orientation == SPSSymbolConstants.VERTICAL_ORIENTATION)
            {
                //Translate vertically with an offset equal to height
                Vector translation = new Vector(0.0, 0.0, height);
                offsetCurve = (ComplexString3d)traceCurve.GetOffsetCurve(translation);
            }
            else
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrInvalidOrientationType, "The given handrail orientation is not supported: ") + orientation.ToString());
                //stop evaluating
            }

            // Create the rail curve from offset curve
            return CreateRailCurve(offsetCurve, toprailBeginOffset, toprailEndOffset, beginTreatmentType, endTreatmentType);
        }
        #endregion

        #region "Create Midrail Curves"

        /// <summary>
        /// This method creates collection of geometries for midrails given the trace curve.
        /// </summary>
        /// <param name="traceCurve">Handrail path curve.</param>
        /// <param name="orientation">Handrail orientation.</param>
        /// <param name="numberOfMidrails">Number of midrails.</param>
        /// <param name="midrailSpacing">Midrails spacing.</param>
        /// <param name="topOfMidrailDimension">Top of midrail dimension value.</param>
        /// <param name="topOfToePlateDimension">Top of toe plate dimension value.</param>
        /// <param name="lowestMidrailHeight">Lowest midrail height.</param>
        /// <param name="lastMidrailHeight">Last midrail height.</param>
        /// <param name="beginTreatmentType">Handrail begin treatment type.</param>
        /// <param name="endTreatmentType">Handrail end treatment type.</param>
        /// <param name="bottomRailBeginOffset">The bottom rail begin offset.</param>
        /// <param name="bottomRailEndOffset">The bottom rail end offset.</param>
        /// <returns>Collection of geometries corresponding to the given rail curve.</returns>
        /// <exception cref="System.ArgumentNullException">Raised if any of the given arguments is null.</exception>
        /// <exception cref="Ingr.SP3D.Common.Middle.CmnException">Raised if the given handrail orientation is not supported.</exception>
        static internal Collection<ComplexString3d> CreateMidrailCurves(ComplexString3d traceCurve, HandrailPostOrientation orientation,
                                                                        int numberOfMidrails, double midrailSpacing, double topOfMidrailDimension,
                                                                        double topOfToePlateDimension, ref double lowestMidrailHeight, double lastMidrailHeight, int beginTreatmentType,
                                                                        int endTreatmentType, double bottomRailBeginOffset, double bottomRailEndOffset)
        {
            // Check if the arguments are null
            if (traceCurve == null)
            {
                throw new ArgumentNullException("traceCurve");
            }

            // Validate tracecurve
            Collection<ICurve> curves = null;
            if (!IsTraceCurveValid(traceCurve, out curves))
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves, "Handrail path for creation of mid rail is invalid as it does not have any geometry."));
            }

            // Initialize local variables
            Collection<ComplexString3d> midRails = new Collection<ComplexString3d>();
            double height = topOfMidrailDimension;
            int midrailsPlaced = 0;

            for (int i = 1; i <= numberOfMidrails; i++) 
            {
                // Continue if current height  of the rail is non-zero and greater than toe-plate dimension.
                ComplexString3d tempComplexString = null;
                if (height > topOfToePlateDimension & height > 0) 
                {
                    if ((int)orientation == (int)SPSSymbolConstants.PERPENDICULAR_ORIENTATION) 
                    {
                        tempComplexString = (ComplexString3d)traceCurve.GetOffsetCurve(height, OffsetDirection.Normal);
                    }
                    else if (orientation == SPSSymbolConstants.VERTICAL_ORIENTATION) 
                    {
                        //Translate the vector vertically
                        Matrix4X4 transformMatrix = new Matrix4X4();
                        Vector translation = new Vector(0.0, 0.0, height);
                        tempComplexString = (ComplexString3d)traceCurve.GetOffsetCurve(translation);
                    } 
                    else 
                    {
                        throw new CmnException("The given post orientation is not supported. Post orientation type :" + orientation.ToString());
                        //stop evaluating
                    }

                    if ((tempComplexString != null))
                        tempComplexString.GetCurves(out curves);

                    if ((curves != null)) 
                    {
                        if (Math.Abs(height - lastMidrailHeight) > StructHelper.MEDIUMDISTTOL)
                        {
                            tempComplexString = new ComplexString3d(curves);

                        } 
                        else 
                        {
                            //Get the mid rail by offsetting the path curve and apply endtreatments for the last mid rail (trim ends)
                            tempComplexString = CreateRailCurve(tempComplexString, bottomRailBeginOffset, bottomRailEndOffset, beginTreatmentType, endTreatmentType);
                        }

                        midRails.Add(tempComplexString);
                        midrailsPlaced = midrailsPlaced + 1;
                        if (Math.Abs(midrailSpacing) >= StructHelper.SMALLDISTTOL)
                        {
                            height = height - midrailSpacing;
                        } 
                        else 
                        {
                            return midRails;
                        }
                    }
                }
            }

            // In case of muultiple midrails, calculate the lowest midrail height which will be used later to connect end treatment
            // In case of multiple midrails, adjust the lowest midrails height for the placement of next one.
            if (numberOfMidrails > 1 & midrailsPlaced > 1)
            {
                lowestMidrailHeight = height + midrailSpacing;
            }
            return midRails;
        }

        #endregion

        #region "Create ToePlate Curve"

        /// <summary>
        /// This method offsets the given toe plate curve and creates the corresponding geometry.
        /// </summary>
        /// <param name="traceCurve">Handrail path curve.</param>
        /// <param name="topOfToePlateDimension">Top of toe plate dimension value.</param>
        /// <param name="orientation">Handrail orientation.</param>
        /// <returns>Geometry corresponding to the given toe plate curve.</returns>
        /// <exception cref="System.ArgumentNullException">Raised if any of the given arguments is null.</exception>
        /// <exception cref="Ingr.SP3D.Common.Middle.CmnException">Raised if the given handrail orientation is not supported. </exception>
        static internal ComplexString3d CreateToePlate(ComplexString3d traceCurve, double topOfToePlateDimension, int orientation)
        {
            // Check if the arguments are null
            if(traceCurve == null)
        {
                throw new ArgumentNullException("traceCurve");
            }
            ComplexString3d offsetTraceCurve = new ComplexString3d(traceCurve);

            // Translate the curve normal to the curve plane
            if (topOfToePlateDimension > 0.0)
            {
                if ((int)orientation == (int)SPSSymbolConstants.PERPENDICULAR_ORIENTATION)
                {
                    // For this case, we can only support planar curves and path that does not contain arcs
                    if (traceCurve.Scope == CurveScopeType.Planar |
                        !traceCurve.HasCurveGeometryType(GeometryType.Arc3d))
                    {
                        offsetTraceCurve = (ComplexString3d)traceCurve.GetOffsetCurve(topOfToePlateDimension, OffsetDirection.Normal);
                    }
                    else
                    {
                        throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrInvalidOrientationType, "The given handrail orientation is not supported: ") + orientation.ToString());
                    }
                }
                else
                {
                    Vector translation = new Vector(0.0, 0.0, topOfToePlateDimension);
                    offsetTraceCurve = (ComplexString3d)traceCurve.GetOffsetCurve(translation);
                }
            }
            return offsetTraceCurve;
        }
        #endregion

        #region "Add Handrail Output"

        /// <summary>
        /// This method builds the handrail output from given curve an add it to SimplePhysical or Centerline aspect.
        /// </summary>
        /// <param name="currentAspect">The current aspect.</param>
        /// <param name="traceCurve">Handrail trace curvet.</param>
        /// <param name="section">Cross section.</param>
        /// <param name="sectionCP">Cross section cardinal point.</param>
        /// <param name="sectionAngle">Cross section angle.</param>
        /// <param name="isMirror">Is output need to be mirrored.</param>
        /// <param name="outputPrefix">Output name.</param>
        /// <param name="handrailMemberType">Handrail member type.</param>
        /// <param name="outputCreatedCount">Number of output created.</param>
        /// <param name="SweepOptions">The sweep options.</param>
        /// <returns>
        /// Collection of surfaces corresponding to the given cross section.
        /// </returns>
        static internal Collection<ISurface> AddHandrailOutput(AspectDefinition currentAspect, BusinessObject traceCurve, CrossSection section,
                                                int sectionCP, double sectionAngle, bool isMirror, string outputPrefix,
                                                int handrailMemberType, ref int outputCreatedCount, SweepOptions SweepOptions)
        {
            string outputName = null;
            //Dim handrailHelper As New HandrailHelper
            SP3D.Structure.Middle.Services.CrossSectionServices crossSectionServices = new SP3D.Structure.Middle.Services.CrossSectionServices();
            Collection<ISurface> surfaces = default(Collection<ISurface>);
            string aspectName = "";
            string aspectDesc = "";
            int aspectId = 0;
            currentAspect.GetAspectInfo(out aspectName, out aspectDesc, out aspectId);

            if (aspectId == (int)Ingr.SP3D.Common.Middle.AspectID.SimplePhysical)
            {
                if (handrailMemberType == (int)HandrailHelper.HandrailMemberType.ToePlate)
                {
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(section, (ICurve)traceCurve, sectionCP, isMirror, sectionAngle, SweepOptions, SweepOrientation.NormalAlongZ);
                }
                else
                {
                    surfaces = crossSectionServices.GetProjectionSurfacesFromCrossSection(section, (ICurve)traceCurve, sectionCP, isMirror, sectionAngle, SweepOptions);
                }

                if ((surfaces != null))
                {

                    for (int index = 0; index <= surfaces.Count - 1; index++)
                    {
                        if (outputCreatedCount == 0)
                        {
                            outputName = outputPrefix;
                        }
                        else
                        {
                            outputName = outputPrefix + outputCreatedCount;
                        }

                        currentAspect.Outputs.Add(outputName, surfaces[index]);
                        outputCreatedCount = outputCreatedCount + 1;
                    }
                }
            }
            else if (aspectId == (int)Ingr.SP3D.Common.Middle.AspectID.Centerline)
            {
                //traceCurve is not persistent. so read it to make a persistent one.
                ICurve outTraceCurve = (ICurve)traceCurve;

                if ((outTraceCurve != null))
                {
                    if (outputCreatedCount == 0)
                    {
                        outputName = outputPrefix;
                    }
                    else
                    {
                        outputName = outputPrefix + outputCreatedCount;
                    }
                    currentAspect.Outputs.Add(outputName, outTraceCurve);
                    outputCreatedCount = outputCreatedCount + 1;
                }

            }
            return surfaces;
        }

        #endregion

        #region "Get Orientation Angle"

        /// <summary>
        /// Gets the orientation angle for aligning the circular treatment
        /// </summary>
        /// <param name="section">The section.</param>
        /// <param name="isStart">Boolean indicating whether start is true.</param>
        /// <returns></returns>
        static internal double GetOrientationAngle(CrossSection section, bool isStart)
        {
            double sectionAngle = 0;
            // for symmetric about both axes
            bool symmetryAboutX = false;
            bool symmetryAboutY = false;
            symmetryAboutX = StructHelper.GetBoolProperty(section, "IStructCrossSectionDesignProperties", "IsSymmetricAboutX");
            symmetryAboutY = StructHelper.GetBoolProperty(section, "IStructCrossSectionDesignProperties", "IsSymmetricAboutY");
            //if we use toprail section angle for creating circular treatment, it is not properly aligned
            //For assymetric sectionls like "L" the treatment angle goes 90 and 270 to the w.r.t toprail at start and end respectively.
            if (!symmetryAboutX & !symmetryAboutY)
            {
                if (isStart)
                {
                    sectionAngle = sectionAngle + 3 * Math.PI / 2;
                }
                else
                {
                    sectionAngle = sectionAngle + Math.PI / 2;
                }
            }
            else
            {
                //sections like 2L places exactly opposite to the toprail section angle.
                if (symmetryAboutY)
                {
                    sectionAngle = sectionAngle + Math.PI;
                }
            }
            return sectionAngle;
        }

        #endregion

        /// <summary>
        /// Determines whether given trace curve is valid by checking if it is null and curve count is 0.
        /// </summary>
        /// <param name="traceCurve">The trace curve.</param>
        /// <returns></returns>
        /// <exception cref="Ingr.SP3D.Common.Middle.CmnException">
        /// </exception>
        private static bool IsTraceCurveValid(ComplexString3d traceCurve, out Collection<ICurve> curves)
        {
            if (traceCurve == null)
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves, "Handrail path is invalid as it does not have any geometry."));
            }

            // Initialize
            curves = new Collection<ICurve>();

            // Get trace curves
            if ((traceCurve != null))
            {
                traceCurve.GetCurves(out curves);
            }

            // If curves object is null or if there are 0 curves, throw exception
            if (curves == null)
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves, "Invalid handrail path. The path should contain atleast one curve."));
            }

            if (curves.Count == 0)
            {
                throw new CmnException(HandRailSymbolsLocalizer.GetString(HandRailSymbolsResourceIDs.ErrNoInputCurves, "Invalid handrail path. The path should contain atleast one curve."));
            }

            return true;
        }
    }
}
