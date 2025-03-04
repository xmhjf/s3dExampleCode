//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Custom mark rule creates custom mark on Template.    
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
    /// Custom mark Rule.
    /// </summary>
    public class CustomMark : MarkingRule
    {
        /// <summary>
        /// Creates custom mark.
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

                ComplexString3d unfoldedbottomLine = null, unfoldedUVLine = null;

                if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
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
                                else if (mfgGeometry.GeometryType == ManufacturingGeometryType.UVMark)
                                {
                                    unfoldedUVLine = mfgGeometry.Geometry;
                                }
                            }
                        }
                        if (unfoldedbottomLine == null || unfoldedUVLine == null)
                            return;
                    }
                }
                else
                    return;



                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                List<ManufacturingGeometry> customMarks = new List<ManufacturingGeometry>(); 
                
                Position startPosition = null, endPosition = null;
                unfoldedUVLine.EndPoints(out startPosition, out endPosition);

                Vector directionVector = startPosition.Subtract(endPosition);

                Position midPosition = unfoldedbottomLine.PointAtDistanceAlong(unfoldedbottomLine.Length * 0.5);

                //TODO - Need .NET equivalent of oMfgRuleHelper.GetPointAlongCurveAtDistance(oBottomWB, oMidPos, 0.2, Nothing)
                Position markPosition = midPosition; // unfoldedBottomLineCurve.PointAtDistanceAlong(midPosition, 0.2);

                Vector markVector = new Vector(directionVector.X, directionVector.Y, directionVector.Z);
                markVector.Length = -1.0; //Thickness direction is opposite to UV mark direction

                ComplexString3d customMark = base.GetLineAtPosition(markPosition, markVector, 0.05, false, false);

                //Create the geometry for the mark
                ManufacturingGeometry customMarkGeometry = new ManufacturingGeometry(ManufacturingContext.Unfolded, ManufacturingGeometryType.ReferenceMark, customMark);
                customMarkGeometry.MarkingInfoName = "Base Control Mark";
                customMarkGeometry.Hidden = true;

                //Add to the list
                customMarks.Add(customMarkGeometry);

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(customMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3017, "Call to TemplateSet Custom Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
