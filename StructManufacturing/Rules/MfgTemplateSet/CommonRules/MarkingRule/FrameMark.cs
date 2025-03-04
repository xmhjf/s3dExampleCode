//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Frame mark rule creates frame mark on Template.    
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
    /// Frame mark Rule.
    /// </summary>
    public class FrameMark : MarkingRule
    {
        /// <summary>
        /// Creates frame marks.
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

                List<ManufacturingGeometry> frameMarks = new List<ManufacturingGeometry>();
                ReadOnlyCollection<ManufacturingGeometry> marks = null;

                if (mfgTemplateSet != null)
                {
                    if ((mfgTemplateSet.Type == TemplateSetType.Plate) || (mfgTemplateSet.Type == TemplateSetType.ProfileFace))
                    {

                        marks = mfgTemplate.Report.GetMarks(ManufacturingGeometryType.XFrameMark);
                        foreach (ManufacturingGeometry mark in marks)
                        {
                            frameMarks.Add(mark);
                        }

                        marks = mfgTemplate.Report.GetMarks(ManufacturingGeometryType.YFrameMark);
                        foreach (ManufacturingGeometry mark in marks)
                        {
                            frameMarks.Add(mark);
                        }

                        marks = mfgTemplate.Report.GetMarks(ManufacturingGeometryType.ZFrameMark);
                        foreach (ManufacturingGeometry mark in marks)
                        {
                            frameMarks.Add(mark);
                        }
                    }
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(frameMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3018, "Call to TemplateSet Frame Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
