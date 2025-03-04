//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Fitting mark rule creates fitting mark on Template.    
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
    /// Fitting mark Rule.
    /// </summary>
    public class FittingMark : MarkingRule
    {
        /// <summary>
        /// Creates fitting mark.
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

                ComplexString3d unfoldedTopLine = null;

                if (mfgTemplateSet.Type == TemplateSetType.Tube)
                {
                    //Get the unfolded template top line from the reference objects collection     
                    foreach (BusinessObject businessObject in markingInfo.ReferenceObjects)
                    {
                        if (businessObject is ManufacturingGeometry)
                        {
                            ManufacturingGeometry mfgGeometry = (ManufacturingGeometry)businessObject;
                            if ((mfgGeometry.GeometryType == ManufacturingGeometryType.TemplateTopLine) &&
                                (mfgGeometry.Context == ManufacturingContext.Unfolded))
                            {
                                unfoldedTopLine = mfgGeometry.Geometry;
                                break;
                            }
                        }
                    }
                    if (unfoldedTopLine == null)
                        return;  

                }
                else
                    return;



                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                List<ManufacturingGeometry> fittingMarks = new List<ManufacturingGeometry>();

                double fittingMarkOffset = Convert.ToDouble(markingInfo.GetArguments("FittingMarkOffset").FirstOrDefault().Value);

                Position startPosition = null, endPosition = null;

                unfoldedTopLine.EndPoints(out startPosition, out endPosition);

                //Offset the top curve by the offset to create the fitting mark
                Vector offsetVector = new Vector(0.0, -1.0, 0.0);
                offsetVector.Length = fittingMarkOffset;

                Position startOffsetPosition = null, endOffsetPosition = null;
                startOffsetPosition = startPosition.Offset(offsetVector);
                endOffsetPosition = endPosition.Offset(offsetVector);

                Line3d markLine = new Line3d(startOffsetPosition, endOffsetPosition);
                Collection<ICurve> lineCol = new Collection<ICurve>();
                lineCol.Add(markLine);
                ComplexString3d fittingMark = new ComplexString3d(lineCol);                

                //Create the geometry for the mark
                ManufacturingGeometry fittingMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.FittingMark, fittingMark);
                fittingMarkGeometry.MarkingInfoName = "Fitting Mark";

                //Add to the list
                fittingMarks.Add(fittingMarkGeometry);

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(fittingMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 2004, "Call to TemplateSet Fitting Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
