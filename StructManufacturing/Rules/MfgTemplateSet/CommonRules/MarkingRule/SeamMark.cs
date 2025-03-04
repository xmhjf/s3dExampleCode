//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Seam mark rule creates seam marks on Template.    
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
    /// Seam mark Rule.
    /// </summary>
    public class SeamMark : MarkingRule
    {
        /// <summary>
        /// Creates seam marks.
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

                if (mfgTemplate == null || mfgTemplateSet == null)
                    return;

                ComplexString3d bottomLine = null, topLine = null;
                Curve3d referenceCurve = null;

                if ((mfgTemplateSet.Type == TemplateSetType.ProfileEdge) || (mfgTemplateSet.Type == TemplateSetType.Tube))
                {
                    //Get the template bottom line from the reference objects collection     
                    foreach (BusinessObject businessObject in markingInfo.ReferenceObjects)
                    {
                        if (businessObject is ManufacturingGeometry)
                        {
                            ManufacturingGeometry mfgGeometry = (ManufacturingGeometry)businessObject;
                            if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateLocationMark)
                            {
                                bottomLine = mfgGeometry.Geometry;
                            }
                            else if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateTopLine)
                            {
                                topLine = mfgGeometry.Geometry;
                            }
                        }
                        else if (businessObject is Curve3d)
                        {
                            referenceCurve = (Curve3d)businessObject;
                        }
                    }

                    if (bottomLine == null)
                        return;
                }



                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                List<ManufacturingGeometry> seamMarks = new List<ManufacturingGeometry>();

                Position markPosition = null, startPosition = null, endPosition = null;
                ComplexString3d lowerSeamMark = null, upperSeamMark = null;                
                BusinessObject representedEntity = null;
                ManufacturingGeometry lowerSeamMarkGeometry = null, upperSeamMarkGeometry = null;

                double seamMarkLength = Convert.ToDouble(markingInfo.GetArguments("SeamMarkLength").FirstOrDefault().Value);

                ////////////////////// Lower Seam Mark //////////////////////  
                #region Lower Seam Mark

                if ((mfgTemplateSet.Type == TemplateSetType.Plate) || (mfgTemplateSet.Type == TemplateSetType.ProfileFace))
                {
                    //Get the template bottom line
                    bottomLine = mfgTemplate.BottomLine.Geometry;

                    //Get the mark position
                    markPosition = mfgTemplate.Report.GetPosition(TemplatePositionType.LowerSeam);

                    //Get the mark
                    lowerSeamMark = GetSeamMarkAtPosition(markPosition, seamMarkLength, bottomLine, mfgTemplate);

                    //Get the entity representing the mark 
                    if (mfgTemplateSet.Type == TemplateSetType.Plate)
                    {
                        representedEntity = mfgTemplate.Report.GetBoundary(TemplateSide.Lower);
                    }

                    //Create the geometry for the mark
                    if (representedEntity == null)
                    {
                        lowerSeamMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.SeamMark, lowerSeamMark);
                    }
                    else
                    {
                        lowerSeamMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.SeamMark, representedEntity, lowerSeamMark);
                    }

                    representedEntity = null;

                }
                else //if ((mfgTemplateSet.Type == TemplateSetType.ProfileEdge) || (mfgTemplateSet.Type == TemplateSetType.Tube))
                {

                    if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                    {
                        //Get the mark position   
                        if (bottomLine != null)
                        {
                            bottomLine.EndPoints(out startPosition, out endPosition);
                            markPosition = startPosition;
                        }
                    }
                    else // if (mfgTemplateSet.Type == TemplateSetType.Tube)
                    {
                        // If there is no extension then exit
                        if (topLine != null && referenceCurve != null)
                        {
                            if ((topLine.Length - referenceCurve.Length) < 0.001)
                                return;

                            //create at 180 degree from origin
                            markPosition = new Position((0.5 * referenceCurve.Length), 0.0, 0.0);
                        }
                    }

                    //Get the mark
                    if (markPosition != null)
                    {
                        lowerSeamMark = GetSeamMarkAtPosition(markPosition, seamMarkLength, bottomLine);

                        //Create the geometry for the mark
                        lowerSeamMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.SeamMark, lowerSeamMark);
                    }
                }


                if (lowerSeamMarkGeometry != null)
                {
                    lowerSeamMarkGeometry.MarkingInfoName = "Seam Mark";

                    //Add to the list
                    seamMarks.Add(lowerSeamMarkGeometry);
                }

                #endregion Lower Seam Mark

                ////////////////////// Upper Seam Mark //////////////////////
                #region Upper Seam Mark

                if ((mfgTemplateSet.Type == TemplateSetType.Plate) || (mfgTemplateSet.Type == TemplateSetType.ProfileFace))
                {
                    //Get the template bottom line
                    bottomLine = mfgTemplate.BottomLine.Geometry;

                    //Get the mark position
                    markPosition = mfgTemplate.Report.GetPosition(TemplatePositionType.UpperSeam);

                    //Get the mark
                    upperSeamMark = GetSeamMarkAtPosition(markPosition, seamMarkLength, bottomLine, mfgTemplate);

                    //Get the entity representing the mark 
                    if (mfgTemplateSet.Type == TemplateSetType.Plate)
                    {
                        representedEntity = mfgTemplate.Report.GetBoundary(TemplateSide.Upper);
                    }

                    //Create the geometry for the mark
                    if (representedEntity == null)
                    {
                        upperSeamMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.SeamMark, upperSeamMark);
                    }
                    else
                    {
                        upperSeamMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.SeamMark, representedEntity, upperSeamMark);
                    }
                }
                else //if ((mfgTemplateSet.Type == TemplateSetType.ProfileEdge) || (mfgTemplateSet.Type == TemplateSetType.Tube))
                {

                    if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                    {
                        markPosition = endPosition;
                    }
                    else // if (mfgTemplateSet.Type == TemplateSetType.Tube)
                    {
                        // create at -180 degree from origin
                        if(referenceCurve != null)
                            markPosition = new Position((-0.5 * referenceCurve.Length), 0.0, 0.0);
                    }

                    if (markPosition != null)
                    {
                        //Get the mark
                        upperSeamMark = GetSeamMarkAtPosition(markPosition, seamMarkLength, bottomLine);

                        //Create the geometry for the mark
                        upperSeamMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.SeamMark, upperSeamMark);
                    }
                                    
                }

                if (upperSeamMarkGeometry != null)
                {
                    upperSeamMarkGeometry.MarkingInfoName = "Seam Mark";
                    //Add to the list
                    seamMarks.Add(upperSeamMarkGeometry);
                }

                #endregion Upper Seam Mark

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(seamMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e,  3021, "Call to TemplateSet Seam Mark Rule failed with the error" + e.Message);
            }

        }

        #region Private Methods

        private ComplexString3d GetSeamMarkAtPosition(Position markPosition, double seamMarkLength, Curve3d bottomCurve)
        {
            ComplexString3d mark = null;
            Vector markVector = new Vector(0.0, -1.0, 0.0);

            mark = base.GetLineAtPosition(markPosition, markVector, 5.0, false, false);

            double minimumDistance = 0.0;
            Position srcPosition = null, inPosition = null;
            mark.DistanceBetween(bottomCurve, out minimumDistance, out srcPosition, out inPosition);

            markPosition.Set(inPosition.X, inPosition.Y, inPosition.Z);

            Vector tangentVector = bottomCurve.TangentAtPoint(markPosition);
            Vector zVector = new Vector(0.0, 0.0, 1.0);
            markVector = tangentVector.Cross(zVector);

            mark = base.GetLineAtPosition(markPosition, markVector, seamMarkLength, true, false);

            return mark;
        }

        private ComplexString3d GetSeamMarkAtPosition(Position markPosition, double seamMarkLength, Curve3d bottomCurve, Template mfgTemplate)
        {

            Vector tangentVector = bottomCurve.TangentAtPoint(markPosition);            
            Vector planeNormal = mfgTemplate.Report.Plane.Normal;
            Vector markVector = tangentVector.Cross(planeNormal);

            ComplexString3d mark = base.GetLineAtPosition(markPosition, markVector, seamMarkLength, true, true);

            return mark;
        }

        #endregion Private Methods
    }
}
