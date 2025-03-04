//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Sight Line mark rule creates sight line mark on Template.    
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
    /// Sight line mark rule.
    /// </summary>
    public class SightLineMark : MarkingRule
    {
        /// <summary>
        /// Creates sight line mark.
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

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                List<ManufacturingGeometry> sightLineMarks = new List<ManufacturingGeometry>();
                Position markPosition = null, startPosition = null, endPosition = null;

                double sightLineMarkLength = Convert.ToDouble(markingInfo.GetArguments("SightMarkLength").FirstOrDefault().Value);
                double sightMarkOffset = Convert.ToDouble(markingInfo.GetArguments("SightMarkOffset").FirstOrDefault().Value);

                //Get the template top line
                ManufacturingGeometry topLine = mfgTemplate.TopLine;
                if (topLine != null)
                {

                    ComplexString3d topLineGeometry = mfgTemplate.TopLine.Geometry;

                    topLineGeometry.EndPoints(out startPosition, out endPosition);

                    Vector topLineVector = endPosition.Subtract(startPosition);
                    Vector lowerToUpperDirectionVector = mfgTemplateSet.GetLowerToUpperDirection(mfgTemplate.GroupNumber);

                    //Get the base control on top line position
                    Position baseControlOnTopLinePosition = mfgTemplate.Report.GetPosition(TemplatePositionType.BaseControlOnTopLine);

                    if (lowerToUpperDirectionVector.Dot(topLineVector) > 0.0)
                    {
                        //TODO - Need .NET equivalent of oMfgRuleHelper.GetPointAlongCurveAtDistance(oTopLineWB, oBCTLPosition, sightMarkOffset, oEndPos)

                        //markPosition = topLineCurve.PointAtDistanceAlong(baseControlOnTopLinePosition, sightMarkOffset);
                        markPosition = baseControlOnTopLinePosition; // Temporary code :: Remove this line when above is available
                    }
                    else
                    {
                        //markPosition = topLineCurve.PointAtDistanceAlong(startPosition, sightMarkOffset);
                        markPosition = baseControlOnTopLinePosition; // Temporary code :: Remove this line when above is available
                    }

                    // Get the direction to create the mark
                    Vector tangentVector = topLineGeometry.TangentAtPoint(markPosition);
                    Vector planeNormal = mfgTemplate.Report.Plane.Normal;
                    Vector markVector = tangentVector.Cross(planeNormal);

                    //Get the mark
                    ComplexString3d sightLineMark = base.GetLineAtPosition(markPosition, markVector, sightLineMarkLength, true, true);

                    //Get the entity representing the mark 
                    BusinessObject representedEntity = mfgTemplateSet.GetBaseControlLine(mfgTemplate.GroupNumber);

                    //Create the geometry for the mark
                    if (representedEntity != null)
                    {
                        ManufacturingGeometry sightLineMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.SightLineMark, representedEntity, sightLineMark);
                        sightLineMarkGeometry.MarkingInfoName = "Sight Line Mark";

                        //Add to the list
                        sightLineMarks.Add(sightLineMarkGeometry);
                    }
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(sightLineMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3023, "Call to TemplateSet Sight Line Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
