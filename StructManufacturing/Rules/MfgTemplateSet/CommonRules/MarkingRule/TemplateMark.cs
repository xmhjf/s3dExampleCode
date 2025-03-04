//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Template mark rule creates template mark on Template.    
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
    /// Template mark Rule.
    /// </summary>
    public class TemplateMark : MarkingRule
    {
        /// <summary>
        /// Creates template marks.
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

                List<ManufacturingGeometry> templateMarks = new List<ManufacturingGeometry>();
                ReadOnlyCollection<ManufacturingGeometry> marks = null;

                if (mfgTemplateSet != null)
                {
                    if ((mfgTemplateSet.Type == TemplateSetType.Plate) || (mfgTemplateSet.Type == TemplateSetType.ProfileFace))
                    {
                        marks = mfgTemplate.Report.GetMarks(ManufacturingGeometryType.TemplateMark);
                        foreach (ManufacturingGeometry mark in marks)
                        {
                            templateMarks.Add(mark);
                        }
                    }
                }

                // ***Note: Below attributes for slots are defined in XML. If needed, user can change their values in the XML.
                    
                //int slotGeometryType = 1;   // 1 - SlotAsCut; 2 - SlotAsMark; 0 - Ignore.;
                //int slotAtTop = 1;          // 1 - TopSlotForPrimary; 2 - TopSlotSecondSecondary;
                //int slotAtEdgeType = 0;     // 0 - EdgeSlotAsSlot; 1- EdgeSlotAsTrim; 2-EdgeSlotAsIgnore;
                //int slotLocationType = 0;   // 0 - SlotAtCenter; 1- SlotTowardsUpper; 2-SlotTowardsLower;
                //double slotMarginAtLower = 0.0;
                //double slotMarginAtUpper = 0.0;
                //double slotMarginAtBottom = 0.0;

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                markingInfo.ManufacturingGeometries = new ReadOnlyCollection<ManufacturingGeometry>(templateMarks);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3024, "Call to TemplateSet Template Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
