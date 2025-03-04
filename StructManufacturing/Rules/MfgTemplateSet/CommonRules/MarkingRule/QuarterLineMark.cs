//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Quarterline mark rule creates quarter line mark on Template.    
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
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Quarter line mark Rule.
    /// </summary>
    public class QuarterLineMark : MarkingRule
    {
        /// <summary>
        /// Creates quarter line mark.
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

                ComplexString3d unfoldedbottomLine = null, unfoldedTopLine = null;
                Curve3d referenceCurve = null;


                if (mfgTemplateSet.Type == TemplateSetType.Tube)
                {
                    //Get the unfolded template bottom line and UV line from the reference objects collection     
                    foreach (BusinessObject businessObject in markingInfo.ReferenceObjects)
                    {
                        if (businessObject is ManufacturingGeometry)
                        {
                            ManufacturingGeometry mfgGeometry = (ManufacturingGeometry)businessObject;
                            if (mfgGeometry.Context == ManufacturingContext.Unfolded)
                            {
                                if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateLocationMark)
                                {
                                    unfoldedbottomLine = mfgGeometry.Geometry;
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

                    if (unfoldedbottomLine == null || unfoldedTopLine == null)
                        return;
                }
                else
                    return;
 

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                List<ManufacturingGeometry> quarterLineMarks = new List<ManufacturingGeometry>();                
                ManufacturingGeometry baseControlLine = mfgTemplateSet.GetBaseControlLine();

                double minimumDistance = 0.0;
                Position positionOnBaseControl = null, positionOnRefCurve = null;
                baseControlLine.Geometry.DistanceBetween(referenceCurve, out minimumDistance, out positionOnBaseControl, out positionOnRefCurve);

                Vector tangentVector = referenceCurve.TangentAtPoint(positionOnRefCurve);
                tangentVector.Length = 1.0;

                Vector centerVector = referenceCurve.Centroid.Subtract(positionOnRefCurve);

                if(centerVector.Dot(tangentVector) < 0.0)
                {
                    tangentVector.Length = -1.0;
                }

                string directionReference, directionReferenceOpposite;
                TemplateSetHelper.GetDirectionNames(tangentVector, positionOnRefCurve, false, out directionReference, out directionReferenceOpposite);

                //TODO - Need .NET equivalent of oRuleHelper.GetPointAlongCurveAtDistance(oRefCurveWB, oRefBCLPos, 0.25 * dRefCurveLen, oRefBCLPos)
                Position rightPosition = referenceCurve.PointAtDistanceAlong(positionOnRefCurve, 0.25 * (referenceCurve.Length));

                tangentVector = referenceCurve.TangentAtPoint(rightPosition);
                centerVector = referenceCurve.Centroid.Subtract(rightPosition);
                if (centerVector.Dot(tangentVector) < 0.0)
                {
                    tangentVector.Length = -1.0;
                }

                string directionRight, directionLeft;
                TemplateSetHelper.GetDirectionNames(tangentVector, rightPosition, false, out directionRight, out directionLeft);                              
        
                Position markPosition = new Position();

                //Create Quarter curve at Origin
                ManufacturingGeometry quarterLineMark = GetTrimmedQuarterLineAtPosition(markPosition, unfoldedbottomLine, unfoldedTopLine, directionReference);
                quarterLineMarks.Add(quarterLineMark);

                //create Quarter curve at 90 degree from origin
                markPosition.Set((0.25 * referenceCurve.Length), 0.0, 0.0);
                quarterLineMark = GetTrimmedQuarterLineAtPosition(markPosition, unfoldedbottomLine, unfoldedTopLine, directionLeft);
                quarterLineMarks.Add(quarterLineMark);

                //create Quarter curve at -90 degree from origin
                markPosition.Set((-0.25 * referenceCurve.Length), 0.0, 0.0);
                quarterLineMark = GetTrimmedQuarterLineAtPosition(markPosition, unfoldedbottomLine, unfoldedTopLine, directionRight);
                quarterLineMarks.Add(quarterLineMark);

                //create Quarter curve at 180 degree from origin
                markPosition.Set((0.5 * referenceCurve.Length), 0.0, 0.0);
                quarterLineMark = GetTrimmedQuarterLineAtPosition(markPosition, unfoldedbottomLine, unfoldedTopLine, directionReferenceOpposite);
                quarterLineMarks.Add(quarterLineMark);

                // create Quarter curve at -180 degree from origin
                markPosition.Set((-0.5 * referenceCurve.Length), 0.0, 0.0);
                quarterLineMark = GetTrimmedQuarterLineAtPosition(markPosition, unfoldedbottomLine, unfoldedTopLine, directionReferenceOpposite);
                quarterLineMarks.Add(quarterLineMark);

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(quarterLineMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e,  3020, "Call to TemplateSet Quarter Line Mark Rule failed with the error" + e.Message);
            }
        }

        private ManufacturingGeometry GetTrimmedQuarterLineAtPosition(Position markPosition, ComplexString3d bottomCurve, ComplexString3d topCurve, string direction)
        {
            Vector markVector = new Vector(0.0, -1.0, 0.0);
            ICurve quarterLineMark = base.GetLineAtPosition(markPosition, markVector, 5.0, false, false);

            double minimumDistance = 0.0;
            Position trimPosition1 = null, trimPosition2 = null, positionOnCurve = null;

            quarterLineMark.DistanceBetween(topCurve, out minimumDistance, out trimPosition1, out positionOnCurve);
            quarterLineMark.DistanceBetween(bottomCurve, out minimumDistance, out trimPosition2, out positionOnCurve);

            ComplexString3d trimmedQuarterLineMark = (ComplexString3d) quarterLineMark;

            base.TrimCurve(ref trimmedQuarterLineMark, trimPosition1, trimPosition2);

            ManufacturingGeometry quarterLine = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.ReferenceMark, trimmedQuarterLineMark);
            quarterLine.MarkingInfoName = direction;

            return quarterLine;

        }

    }
}
