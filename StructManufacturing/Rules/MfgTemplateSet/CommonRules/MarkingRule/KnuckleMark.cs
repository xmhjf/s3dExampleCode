//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Knuckle mark rule creates knuckle mark on Template.    
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
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Knuckle mark Rule.
    /// </summary>
    public class KnuckleMark : MarkingRule
    {
        /// <summary>
        /// Creates knuckle mark.
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

                if (mfgTemplateSet != null)
                {
                    if (!((mfgTemplateSet.Type == TemplateSetType.Plate) || (mfgTemplateSet.Type == TemplateSetType.ProfileFace)))
                    {
                        return;
                    }
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

                List<ManufacturingGeometry> knuckleMarks = new List<ManufacturingGeometry>();

                double knuckleMarkLength = Convert.ToDouble(markingInfo.GetArguments("KnuckleMarkLength").FirstOrDefault().Value);
                ComplexString3d bottomLine = mfgTemplate.BottomLine.Geometry;

                IManufacturable parentPart = mfgTemplateSet.DetailedPart;
                Plate platePart = null;
                //ReadOnlyCollection<PlateKnuckle> plateKnuckles = null;                
                ReadOnlyCollection<ProfileKnuckle> profileKnuckles = null;
                ReadOnlyCollection<Curve3d> knucklesOnMarkingSide = null;
                ReadOnlyCollection<double> knuckleAngles = null;

                List<Position> knuckles = null; //TODO - oWireBodyUtils.GetKnucklePoints(oBottomLineWB)

                if (parentPart is Plate)
                {

                    platePart = (Plate)parentPart;
                    PlateSystem plateSystem = (PlateSystem)platePart.SystemParent;
                    if (plateSystem.Type == PlateType.Hull)
                        return;

                    ////plateKnuckles = plateSystem.Knuckles;

                    if (knuckles == null)
                        return;
                    
                }
                else if (parentPart is Profile)
                {

                    base.GetBendProfileKnuckleData((ProfilePart)parentPart, SectionFaceType.Web_Left, out profileKnuckles, out knucklesOnMarkingSide, out knuckleAngles);

                    if (knuckles == null)
                    {
                        knuckles = new List<Position>();
                    }

                    if (profileKnuckles != null)
                    {
                        foreach (Curve3d knuckleCurve in knucklesOnMarkingSide)
                        {
                            double minimumDistance = 0.0;
                            Position srcPosition = null, inPosition = null;
                            knuckleCurve.DistanceBetween(bottomLine, out minimumDistance, out srcPosition, out inPosition);

                            knuckles.Add(inPosition);
                        }
                    }

                    if (knuckles.Count == 0)
                        return;
                }
                else
                    return;
                

                foreach (Position knucklePosition in knuckles)
                {
                    // Get the direction to create the mark
                    Vector tangentVector = bottomLine.TangentAtPoint(knucklePosition);
                    Vector planeNormal = mfgTemplate.Report.Plane.Normal;
                    Vector markVector = tangentVector.Cross(planeNormal);

                    //Get the mark
                    ComplexString3d knuckleMark = base.GetLineAtPosition(knucklePosition, markVector, knuckleMarkLength, true, true);

                    //Create the geometry for the mark
                    ManufacturingGeometry referenceMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, ManufacturingGeometryType.KnuckleMark, knuckleMark);
                    referenceMarkGeometry.MarkingInfoName = "Knuckle Mark";

                    //Add to the list
                    knuckleMarks.Add(referenceMarkGeometry);
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(knuckleMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3019, "Call to TemplateSet Knuckle Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
