//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Ship direction mark rule creates ship direction marks on Template.    
//
//      Author:  
//
//      History:
//      June 27th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Ship Direction mark Rule.
    /// </summary>
    public class ShipDirectionMark : MarkingRule
    {
        /// <summary>
        /// Creates ship direction marks.
        /// </summary>
        /// <param name="markingInfo">The marking info.</param>
        public override void Evaluate(MarkingInformation markingInfo)
        {
            try
            {
                if (markingInfo == null)
                    throw new ArgumentNullException("Input markingInfo is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                TemplateSet mfgTemplateSet = null;
                if (markingInfo.ManufacturingPart != null)
                {
                    mfgTemplateSet = (TemplateSet)markingInfo.ManufacturingPart;
                }

                Template mfgTemplate = null;
                foreach (BusinessObject businessObject in markingInfo.ReferenceObjects)
                {
                    if (businessObject is Template)
                    {
                        mfgTemplate = (Template)businessObject;
                        break;
                    }
                }

                if (mfgTemplate == null)
                    return;

                ComplexString3d bottomLine = null, UVLine = null, unfoldedBottomLine = null, unfoldedUVLine = null, unfoldedTopLine = null;
                Curve3d referenceCurve = null;

                if (mfgTemplateSet != null)
                {
                    if ((mfgTemplateSet.Type == TemplateSetType.ProfileEdge) || (mfgTemplateSet.Type == TemplateSetType.Tube))
                    {
                        //Get the template bottom line from the reference objects collection     
                        foreach (BusinessObject businessObject in markingInfo.ReferenceObjects)
                        {
                            if (businessObject is ManufacturingGeometry)
                            {
                                ManufacturingGeometry mfgGeometry = (ManufacturingGeometry)businessObject;
                                if (mfgGeometry.Context == ManufacturingContext.Model)
                                {
                                    if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateLocationMark)
                                    {
                                        bottomLine = mfgGeometry.Geometry;
                                    }
                                    else if (mfgGeometry.GeometryType == ManufacturingGeometryType.UVMark)
                                    {
                                        UVLine = mfgGeometry.Geometry;
                                    }
                                }
                                else //if (mfgGeometry.Context == ManufacturingContext.Unfolded)
                                {
                                    if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateLocationMark)
                                    {
                                        unfoldedBottomLine = mfgGeometry.Geometry;
                                    }
                                    else if (mfgGeometry.GeometryType == ManufacturingGeometryType.UVMark)
                                    {
                                        unfoldedUVLine = mfgGeometry.Geometry;
                                    }
                                    else if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateTopLine)
                                    {
                                        unfoldedTopLine = mfgGeometry.Geometry;
                                    }
                                }
                            }
                            else if (businessObject is Curve3d)
                            {
                                referenceCurve = (Curve3d)businessObject;
                            }
                        }
                    }
                }

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                List<ManufacturingGeometry> shipDirectionMarks = null;

                double shipDirectionMarkPrimaryLength = Convert.ToDouble(markingInfo.GetArguments("ShipDirectionMarkPrimaryLength").FirstOrDefault().Value);
                double shipDirectionMarkSecondaryLength = Convert.ToDouble(markingInfo.GetArguments("ShipDirectionMarkSecondaryLength").FirstOrDefault().Value);

                if (mfgTemplateSet != null)
                {
                    if ((mfgTemplateSet.Type == TemplateSetType.Plate) || (mfgTemplateSet.Type == TemplateSetType.ProfileFace))
                    {
                        shipDirectionMarks = CreateShipDirectionMarks(mfgTemplate, shipDirectionMarkPrimaryLength, shipDirectionMarkSecondaryLength);
                    }
                    else if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                    {
                        shipDirectionMarks = CreateShipDirectionMarksForEdge(bottomLine, unfoldedBottomLine, UVLine, unfoldedUVLine, shipDirectionMarkPrimaryLength, shipDirectionMarkSecondaryLength);
                    }
                    else // if (mfgTemplateSet.Type == TemplateSetType.Tube)
                    {
                        shipDirectionMarks = CreateShipDirectionMarksForTube(mfgTemplateSet, unfoldedBottomLine, unfoldedTopLine, referenceCurve, shipDirectionMarkPrimaryLength, shipDirectionMarkSecondaryLength);
                    }
                }


                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(shipDirectionMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3022, "Call to TemplateSet Ship Direction Mark Rule failed with the error" + e.Message);
            }

        }

        #region Private Methods

        private List<ManufacturingGeometry> CreateShipDirectionMarks(Template mfgTemplate, double shipDirectionMarkPrimaryLength, double shipDirectionMarkSecondaryLength)
        {
            if (mfgTemplate == null)
            {
                throw new NullReferenceException("The Input mfgTemplate is null");

            }

            List<ManufacturingGeometry> shipDirectionMarks = new List<ManufacturingGeometry>();

            Position baseControlPosition = mfgTemplate.Report.GetPosition(TemplatePositionType.BaseControl);
            Position baseControlTopLinePosition = mfgTemplate.Report.GetPosition(TemplatePositionType.BaseControlOnTopLine);

            //Get the template top line
            ManufacturingGeometry topLine = mfgTemplate.TopLine;
            if (topLine == null)
                return shipDirectionMarks;

            ComplexString3d topLineGeometry = mfgTemplate.TopLine.Geometry;

            Position startPosition = null, endPosition = null;
            topLineGeometry.EndPoints(out startPosition, out endPosition);
            Vector topLineVector = endPosition.Subtract(startPosition);

            Vector UVDirection = mfgTemplate.XAxis;
            double dotProduct = topLineVector.Dot(UVDirection);

            if (dotProduct < 0.0)
            {
                Position tempPosition = startPosition;
                startPosition = endPosition;
                endPosition = tempPosition;              
            }

            Vector primaryMarkVector = startPosition.Subtract(endPosition);
            primaryMarkVector.Length = (topLineGeometry.Length) * 0.25;

            Vector sideLineVector = baseControlPosition.Subtract(baseControlTopLinePosition);
            double distance = baseControlPosition.DistanceToPoint(baseControlTopLinePosition);
            sideLineVector.Length = distance * 0.25;

            Vector resultVector = primaryMarkVector.Add(sideLineVector);
           
            Position topLineMidPosition = topLineGeometry.PointAtDistanceAlong(topLineGeometry.Length * 0.5);

            Position shipDirectionPosition = topLineMidPosition.Offset(resultVector);

            primaryMarkVector.Length = 1.0;
            Vector normalVector = new Vector(0.0, 1.0, 0.0);

            Plane3d plane = new Plane3d(new Position(), normalVector);           

            double minDistance = 0.0; Position positionOnPlane = null, positionOnCurve = null;
            plane.DistanceBetween(topLineGeometry, out minDistance, out positionOnPlane, out positionOnCurve);         

            bool centerLine = false;
            if (Math.Abs(minDistance -0.0) < 0.0001)
            {
                centerLine = true;
            }

            ComplexString3d mark = null;           

            mark = base.GetLineAtPosition(shipDirectionPosition, primaryMarkVector, shipDirectionMarkPrimaryLength, false, true);

            //Create the geometry for the mark
            ManufacturingGeometry shipDirectionMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.DirectionMark, mark);
            
            // Set the name
            string directionName, oppositeDirectionName;
            TemplateSetHelper.GetDirectionNames(primaryMarkVector, shipDirectionPosition, centerLine, out directionName, out oppositeDirectionName);
            shipDirectionMarkGeometry.MarkingInfoName = directionName;

            shipDirectionMarks.Add(shipDirectionMarkGeometry);

            Vector secondaryMarkVector = primaryMarkVector.Cross(mfgTemplate.Report.Plane.Normal);

            if (secondaryMarkVector.Dot(sideLineVector) < 0.0)
            {
                secondaryMarkVector.Length = -1.0;
            }
            else
            {
                secondaryMarkVector.Length = 1.0;
            }

            mark = base.GetLineAtPosition(shipDirectionPosition, secondaryMarkVector, shipDirectionMarkSecondaryLength, false, true);

            //Create the geometry for the mark
            shipDirectionMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.DirectionMark, mark);

            // Set the name
            TemplateSetHelper.GetDirectionNames(secondaryMarkVector, shipDirectionPosition, centerLine, out directionName, out oppositeDirectionName);
            shipDirectionMarkGeometry.MarkingInfoName = directionName;

            shipDirectionMarks.Add(shipDirectionMarkGeometry);

            return shipDirectionMarks;
        }

        private List<ManufacturingGeometry> CreateShipDirectionMarksForEdge(ComplexString3d bottomLine, ComplexString3d unfoldedBottomLine,
                                                                            ComplexString3d UVLine, ComplexString3d unfoldedUVLine,
                                                                            double shipDirectionMarkPrimaryLength, double shipDirectionMarkSecondaryLength)
        {
            List<ManufacturingGeometry> shipDirectionMarks = new List<ManufacturingGeometry>();

            Position bottomLineMidPosition = bottomLine.PointAtDistanceAlong(bottomLine.Length * 0.5);

            Position unfoldedBottomLineMidPosition = unfoldedBottomLine.ProjectPoint(unfoldedBottomLine.Centroid);
            unfoldedBottomLineMidPosition = unfoldedBottomLine.PointAtDistanceAlong(unfoldedBottomLine.Length * 0.5);

            Vector UVector = bottomLine.TangentAtPoint(bottomLineMidPosition);
            UVector.Length = 1.0;

            Vector UVectorUnfolded = unfoldedBottomLine.TangentAtPoint(unfoldedBottomLineMidPosition);
            UVectorUnfolded.Length = 1.0;

            Position startPosition = null, endPosition = null;
            UVLine.EndPoints(out startPosition, out endPosition);
            Vector VVector = UVLine.TangentAtPoint(startPosition);
            VVector.Length = 1.0;

            unfoldedUVLine.EndPoints(out startPosition, out endPosition);
            Vector VVectorUnfolded = unfoldedUVLine.TangentAtPoint(startPosition);
            VVectorUnfolded.Length = 1.0;

            Vector tempVector = VVectorUnfolded.Add(UVectorUnfolded);
            tempVector.Length = 0.3;

            Position markPosition = unfoldedBottomLineMidPosition.Offset(tempVector);

            ComplexString3d mark = base.GetLineAtPosition(markPosition, UVectorUnfolded, shipDirectionMarkPrimaryLength, false, false);

            //Create the geometry for the mark
            ManufacturingGeometry shipMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.DirectionMark, mark);
            
            //Set the name
            string directionName, oppositeDirectionName;
            TemplateSetHelper.GetDirectionNames(UVector, bottomLineMidPosition, false, out directionName, out oppositeDirectionName);
            shipMarkGeometry.MarkingInfoName = directionName;

            shipDirectionMarks.Add(shipMarkGeometry);

            mark = base.GetLineAtPosition(markPosition, VVectorUnfolded, shipDirectionMarkSecondaryLength, false, false);

            //Create the geometry for the mark
            shipMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.DirectionMark, mark);
            TemplateSetHelper.GetDirectionNames(VVector, bottomLineMidPosition, false, out directionName, out oppositeDirectionName);
            shipMarkGeometry.MarkingInfoName = directionName;

            shipDirectionMarks.Add(shipMarkGeometry);

            return shipDirectionMarks;
        }

        private List<ManufacturingGeometry> CreateShipDirectionMarksForTube( TemplateSet mfgTemplateSet, ComplexString3d unfoldedBottomLine,
                                                                                ComplexString3d unfoldedTopLine, Curve3d referenceCurve,
                                                                                double shipDirectionMarkPrimaryLength, double shipDirectionMarkSecondaryLength)
        {
            List<ManufacturingGeometry> shipDirectionMarks = new List<ManufacturingGeometry>();

            Vector markVector = new Vector(0.0, -1.0, 0.0);

            Position inputPosition = new Position((0.4 * referenceCurve.Length), 0.0, 0.0);

            ComplexString3d verticalMark = base.GetLineAtPosition(inputPosition, markVector, 5.0, true, false);

            double distance = 0.0;
            Position srcPosition = null, inPosition = null;
            unfoldedTopLine.DistanceBetween((ICurve)verticalMark, out distance, out srcPosition, out inPosition);

            markVector.Length = 0.1;
            Position markPosition = srcPosition.Offset(markVector);           

            ManufacturingGeometry baseControlLine = mfgTemplateSet.GetBaseControlLine();

            IPlane centerPlane = new Plane3d(new Position(), mfgTemplateSet.BasePlane.Normal);

            Collection<BusinessObject> colCurves = null;
            GeometryIntersectionType eIntersectCode;
            baseControlLine.Geometry.Intersect((ISurface)centerPlane, out colCurves, out eIntersectCode);

            ComplexString3d mark = base.GetLineAtPosition(markPosition, markVector, shipDirectionMarkSecondaryLength, false, false);

            //Create the geometry for the mark
            ManufacturingGeometry shipMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.DirectionMark, mark);
            
            //Set the name
            string directionName, oppositeDirectionName;
            TemplateSetHelper.GetDirectionNames(mfgTemplateSet.BasePlane.Normal, mfgTemplateSet.BasePlane.RootPoint, false, out directionName, out oppositeDirectionName);
            shipMarkGeometry.MarkingInfoName = directionName;

            shipDirectionMarks.Add(shipMarkGeometry);

            return shipDirectionMarks;
        }

        #endregion Private Methods
    }
}
