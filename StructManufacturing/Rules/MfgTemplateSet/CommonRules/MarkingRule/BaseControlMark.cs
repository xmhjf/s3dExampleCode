//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Base Control mark rule creates base control mark on Template.    
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
    /// Base Control mark Rule.
    /// </summary>
    public class BaseControlMark : MarkingRule
    {
        /// <summary>
        /// Creates base control mark.
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

                if ((mfgTemplateSet.Type == TemplateSetType.ProfileEdge) || (mfgTemplateSet.Type == TemplateSetType.Tube))
                {
                    //Get the template bottom line from the reference objects collection     
                    foreach (BusinessObject businessObject in markingInfo.ReferenceObjects)
                    {
                        if (businessObject is ManufacturingGeometry)
                        {
                            ManufacturingGeometry mfgGeometry = (ManufacturingGeometry)businessObject;
                            if (mfgGeometry.Context == ManufacturingContext.Unfolded)
                            {
                                if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateLocationMark)
                                {
                                    bottomLine = mfgGeometry.Geometry;
                                }
                                else if (mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateTopLine)
                                {
                                    topLine = mfgGeometry.Geometry;
                                }
                            }
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

                List<ManufacturingGeometry> baseControlMarks = new List<ManufacturingGeometry>();

                Position markPosition = null;
                ComplexString3d baseControlMark = null;                           
                ManufacturingGeometry baseControlMarkGeometry = null;
                Vector markVector = null;

                double baseControlMarkLength = Convert.ToDouble(markingInfo.GetArguments("BaseControlMarkLength").FirstOrDefault().Value);

                if ((mfgTemplateSet.Type == TemplateSetType.Plate) || (mfgTemplateSet.Type == TemplateSetType.ProfileFace))
                {
                    //Get the template bottom line
                    bottomLine = mfgTemplate.BottomLine.Geometry;

                    //Get the mark position
                    markPosition = mfgTemplate.Report.GetPosition(TemplatePositionType.BaseControl);

                    //Get the mark
                    if (bottomLine != null)
                    {
                        Vector tangentVector = bottomLine.TangentAtPoint(markPosition);
                        Vector planeNormal = mfgTemplate.Report.Plane.Normal;
                        markVector = tangentVector.Cross(planeNormal);
                    }

                    baseControlMark = base.GetLineAtPosition(markPosition, markVector, baseControlMarkLength, true, true);

                    //Get the entity representing the mark 
                    BusinessObject representedEntity = mfgTemplateSet.GetBaseControlLine(mfgTemplate.GroupNumber);

                    //Create the geometry for the mark
                    baseControlMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.BaseLineMark, representedEntity, baseControlMark);

                }
                else //if ((mfgTemplateSet.Type == TemplateSetType.ProfileEdge) || (mfgTemplateSet.Type == TemplateSetType.Tube))
                {
                    if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                    {
                        if (bottomLine != null)
                        {
                            //Get the mark position   
                            markPosition = bottomLine.PointAtDistanceAlong(bottomLine.Length * 0.5);
                            markVector = bottomLine.TangentAtPoint(markPosition);

                        }

                    }
                    else // if (mfgTemplateSet.Type == TemplateSetType.Tube)
                    {
                        markVector = new Vector(0.0, -1.0, 0.0);
                        markPosition = new Position();

                        ComplexString3d tempMark = base.GetLineAtPosition(markPosition, markVector, 5.0, false, false);

                        double minimumDistance = 0.0;
                        Position srcPosition = null, inPosition = null;
                        tempMark.DistanceBetween(bottomLine, out minimumDistance, out srcPosition, out inPosition);

                        //Get the mark position   
                        markPosition = inPosition;
                    }

                    //Get the mark  
                    baseControlMark = base.GetLineAtPosition(markPosition, markVector, baseControlMarkLength, true, false);

                    //Create the geometry for the mark
                    baseControlMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.BaseLineMark, baseControlMark);
                }
 

                if (baseControlMarkGeometry != null)
                {
                    baseControlMarkGeometry.MarkingInfoName = "Base Control Mark";
                    //Add to the list
                    baseControlMarks.Add(baseControlMarkGeometry);
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(baseControlMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3016, "Call to TemplateSet Base Control Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
