//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Reference mark rule creates reference mark on Template.    
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
    /// Reference mark Rule.
    /// </summary>
    public class ReferenceMark : MarkingRule
    {
        /// <summary>
        /// Creates reference mark.
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
                    if (mfgTemplateSet.Type != TemplateSetType.Plate)
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

                List<ManufacturingGeometry> referenceMarks = new List<ManufacturingGeometry>();
                double referenceMarkLength = Convert.ToDouble(markingInfo.GetArguments("ReferenceMarkLength").FirstOrDefault().Value);

                if (mfgTemplateSet != null)
                {
                    IManufacturable parentPart = mfgTemplateSet.DetailedPart;
                    if (parentPart is PlatePartBase)
                    {
                        PlatePartBase platePart = (PlatePartBase)parentPart;
                        if (platePart.Type == PlateType.Hull)
                            return;
                    }
                    else
                        return;
                }

                ComplexString3d bottomLine = mfgTemplate.BottomLine.Geometry;                

                ReadOnlyCollection<ReferenceCurveBase> referenceCurves  = null;

                if (mfgTemplateSet.Orientation.Side == 5130) //BaseSide
                {
                    //TODO - Get Reference curves on base side -  //oSDPlateWrapper.ReferenceCurves PlateBaseSide, oReferenceCurvesCol
                   
                }
                else
                {
                    // Get Reference curves on offset side
                }

                if (referenceCurves == null)
                    return;

                // ReferenceCurves is null for the time being. So Following Code is dead code and is commented out
                // -------------------------------------------------------------------------------------------------

                //ManufacturingGeometryType geometryType = ManufacturingGeometryType.GeneralMark;

                //foreach (ReferenceCurveBase referenceCurve in referenceCurves)
                //{
                //    if (referenceCurve is PlateKnuckle)
                //    {
                //        PlateKnuckle plateKnuckle = (PlateKnuckle)referenceCurve;

                //        if (plateKnuckle.Type == ReferenceCurveType.Knuckle)
                //        {
                //            geometryType = ManufacturingGeometryType.KnuckleMark;
                //        }
                //        else if (plateKnuckle.Type == ReferenceCurveType.Reference)
                //        {
                //            geometryType = ManufacturingGeometryType.GeneralMark;
                //        }
                        
                //    }
                   
                //    Position markPosition = null, refCurvePosition = null;
                //    double minDistance = 0.0;

                //    bottomLine.DistanceBetween((ICurve)referenceCurve, out minDistance, out markPosition, out refCurvePosition);

                //    if (minDistance < 0.001)
                //    {
                //        // Get the direction to create the mark
                //        Vector tangentVector = bottomLine.TangentAtPoint(markPosition);
                //        Vector planeNormal = mfgTemplate.Report.Plane.Normal;
                //        Vector markVector = tangentVector.Cross(planeNormal);

                //        //Get the mark
                //        ComplexString3d referenceMark = base.GetLineAtPosition(markPosition, markVector, referenceMarkLength, true, false);

                //        //Create the geometry for the mark
                //        ManufacturingGeometry referenceMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Model, geometryType, referenceMark);
                //        referenceMarkGeometry.MarkingInfoName = referenceCurve.Name;

                //        //Add to the list
                //        referenceMarks.Add(referenceMarkGeometry);
                //    }
                //}       
                
                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs
                // ReferenceCurves is null for the time being. So Following Code is dead code and is commented out
                //markingInfo.manufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(referenceMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3020, "Call to TemplateSet Reference Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
