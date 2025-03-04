//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Provides implementation for plate template service.
//               - implements CreateBaseplane
//               - implements creationof control line.
//               - Implements creation of template out contours.
//               - implements creation of template plane
//               
//
//      Author:  Satish Kumar Adimula
//
//      History:
//      November 23 rd, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;
using Ingr.SP3D.Grids.Middle;
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Exceptions;

namespace Ingr.SP3D.Content.Manufacturing
{
    public class ProfileEdgeTemplateGeometryRule : EdgeTemplateGeometryRuleBase
    {
        /// <summary>
        /// Creates the out contours for profile face template set
        /// </summary>
        /// <param name="inputInfo"></param>

        public override ReadOnlyCollection<TemplateContourInformation> CreateTemplates(TemplateSetInformation inputInfo, ref Plane3d basePlane, IComplexString baseControlLine)
        {

            if (inputInfo == null)
            {
                throw new CmnNullArgumentException("inputInfo");
            }

            TemplateSetEdgeInformation edgeInfo = (TemplateSetEdgeInformation)inputInfo;
            if (edgeInfo == null)
            {
                throw new CmnNullArgumentException("edgeInfo");
            }

            BusinessObject profilePart = (BusinessObject)edgeInfo.ManufacturingParent;

            if (profilePart == null)
            {
                throw new CmnNullArgumentException("profilePart");
            }

            ReadOnlyCollection<TemplateContourInformation> templateOutContours = null;

            try
            {
                templateOutContours = base.GetTemplateContours(edgeInfo);
                if (templateOutContours == null)
                {
                    base.WriteToErrorLog("Template Service routine Template Contours: Failed to Create Template Contours", "RULES");
                }
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Template Contours: Failed to Create Template Contours", e.Message);
            }

            return templateOutContours;
        }

        /// <summary>
        /// Creates the template plane
        /// </summary>
        /// <param name="inputInfo">Input templateSet information</param>
        /// <param name="templatePosition">Template position along it's base control line.</param>
        /// <param name="basePlane">Input base plane of the templateSet.</param>

        public override Plane3d CreateTemplatePlane(TemplateSetInformation inputInfo, Position templatePosition, Plane3d basePlane)
        {
            return null;
            //throw new NotImplementedException("Creation of template plane is not supported for profile edge templates");
        }
    }
}
