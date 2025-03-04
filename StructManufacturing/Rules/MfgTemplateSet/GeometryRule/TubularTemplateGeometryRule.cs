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

    /// <summary>
    /// 
    /// </summary>
    public class TubularTemplateGeometryRule : TubularTemplateGeometryRuleBase
    {
        private enum BasePlaneType
        {
            BySystem = 7100,
            NormalToAxis = 7101,
            Global = 7102,
            UserDefined = 7103
        };


        private enum TemplateSetType
        {
            ShortDistanceBCL = 7100,
            LongDistanceBCL = 7101,
            ModelOrigin = 7102,
            GlobalXMax = 7103,
            GlobalXMin = 7104,
            GlobalYMax = 7105,
            GlobalYMin = 7106,
            GlobalZMax = 7107,
            GlobalZMin = 7108
        };


        /// <summary>
        /// Creates the base plane for tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set information</param>
        /// <param name="userDefinePlane">Boolean indicates whether base plane is user defined or not.</param>
        public override Plane3d CreateBasePlane(TemplateSetInformation inputInfo, out bool userDefinedPlane)
        {
            userDefinedPlane = false;
            if (inputInfo == null)
            {
                throw new CmnNullArgumentException("inputInfo");
            }

            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }
            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;
            if (parentMemeber == null)
            {
                throw new ArgumentNullException("parentMemeber");
            }

            ManufacturingOutputBase mfgOutBase = tubeInfo.ManufacturingPart as ManufacturingOutputBase;
            CodelistItem codeItem = null;
            string basePlaneType = "";
            if (mfgOutBase != null)
            {
                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessTube", "BasePlane");
                basePlaneType = codeItem.ShortDisplayName;
            }
            

            Plane3d basePlane = null;

            try
            {
                switch (basePlaneType)
                {
                    //codereview: private enums for code lists.Done
                    case "BySystem":
                        {
                            //codereview: replace transient class with function overload.done                       
                            basePlane = base.GetBasePlane(BasePlaneConstructionType.BySystem, tubeInfo);
                            break;
                        }
                    case "NormalToAxis":
                        {
                            basePlane = base.GetBasePlane(BasePlaneConstructionType.NormalToAxis, tubeInfo);
                            break;
                        }
                    case "Global":
                        {
                            basePlane = base.GetBasePlane(BasePlaneConstructionType.Global, tubeInfo);
                            break;
                        }
                    case "UserDefined":
                        {
                            userDefinedPlane = true;
                            basePlane = null;
                            break;
                        }                        
                }

                if (basePlane == null)
                {
                    base.WriteToErrorLog("Template Service routine Create Base Plane: Failed to Create Base Plane", "RULES");
                }
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Create Base Plane: Failed to Create Base Plane", e.Message);
            }

            return basePlane;
        }


        /// <summary>
        /// Creates the control line for tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set tube information</param>

        public override IComplexString CreateControlLine(TemplateSetInformation inputInfo, Plane3d basePlane)
        {
            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }
            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;
            if (parentMemeber == null)
            {
                throw new CmnNullArgumentException("parentMemeber");
            }

            string templetSetType = "";
            ManufacturingOutputBase mfgOutBase = tubeInfo.ManufacturingPart as ManufacturingOutputBase;
            if (mfgOutBase != null)
            {
                CodelistItem codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessTube", "Type");
                templetSetType = codeItem.ShortDisplayName;
            }           

            IComplexString baseControlLine = null;
            TubeType type = TubeType.GlobalXMax;

            if (templetSetType == "ShortDistanceBCL")
            {
                type = (TubeType.ShortDistanceBCL);
            }
            else if (templetSetType == "LongDistanceBCL")
            {
                type = (TubeType.LongDistanceBCL);
            }
            else if (templetSetType == "ShortStringValue")
            {
                type = (TubeType.ModelOrigin);
            }
            else if (templetSetType == "Global_X_Max")
            {
                type = (TubeType.GlobalXMax);
            }
            else if (templetSetType == "Global_X_Min")
            {
                type = (TubeType.GlobalXMin);
            }
            else if (templetSetType == "Global_Y_Max")
            {
                type = (TubeType.GlobalYMax);
            }
            else if (templetSetType == "Global_Y_Min")
            {
                type = (TubeType.GlobalYMin);
            }
            else if (templetSetType == "Global_Z_Max")
            {
                type = (TubeType.GlobalZMax);
            }
            else if (templetSetType == "Global_Z_Min")
            {
                type = (TubeType.GlobalZMin);
            }
            else if (templetSetType == "ModelOrigin")
            {
                type = TubeType.ModelOrigin;

            }

            //codereview: replace transient class with function overload.done
            try
            {
                baseControlLine = base.GetBaseControlLine(type, tubeInfo, basePlane);
                if (baseControlLine == null)
                {
                    base.WriteToErrorLog("Template Service routine Create BaseControl Line: Failed to create Base Control Line", "RULES");
                }
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Create BaseControl Line: Failed to create Base Control Line", e.Message);
            }

            return baseControlLine;
        }


        /// <summary>
        /// Creates the out contours for tube memeber template ser
        /// </summary>
        /// <param name="inputInfo">Input template set tube information</param>

        public override ReadOnlyCollection<TemplateContourInformation> CreateTemplates(TemplateSetInformation inputInfo, ref Plane3d basePlane, IComplexString baseControlLine)
        {
            TemplateSetTubeInformation tubeInfo = inputInfo as TemplateSetTubeInformation;

            ReadOnlyCollection<TemplateContourInformation> templateOutContours = null;
            try
            {
                templateOutContours = base.GetTemplateContours(tubeInfo, basePlane, baseControlLine);
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
        /// Creates unfold out contour for a tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set information</param>

        public override ReadOnlyCollection<ManufacturingGeometry> CreateTemplateUnfoldContour(TemplateSetInformation inputInfo, Plane3d basePlane)
        {

            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }

            ReadOnlyCollection<ManufacturingGeometry> unfoldOutContour = null;
            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;
            if (parentMemeber == null)
            {
                throw new CmnNullArgumentException("parentMemeber");
            }

            try
            {
                unfoldOutContour = base.GetUnfoldedTemplateContour(inputInfo, basePlane);
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Template Contours: Failed to Create Template Contours", e.Message);
            }

            return unfoldOutContour;
        }


        /// <summary>
        /// Creates unfold reference curve for a tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set tube information</param>

        public override IComplexString CreateTemplateUnfoldReferenceCurve(TemplateSetInformation inputInfo, Plane3d basePlane)
        {
            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }
            IComplexString unfoldRefCurve = null;
            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;
            if (parentMemeber == null)
            {
                throw new CmnNullArgumentException("parentMemeber");
            }

            try
            {
                unfoldRefCurve = base.GetUnfoldedTemplateReferenceCurve(inputInfo, basePlane);
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Template unfold reference curve: Failed to Create Template Contours", e.Message);
            }

            return unfoldRefCurve;
        }


        /// <summary>
        /// Creates template plane for a tube member template  
        /// </summary>
        /// <param name="inputInfo">Input templateSet information</param>
        /// <param name="templatePosition">Template position along it's base control line.</param>
        /// <param name="basePlane">Input base plane of the templateSet.</param>

        public override Plane3d CreateTemplatePlane(TemplateSetInformation inputInfo, Position templatePosition, Plane3d basePlane)
        {

            return null;
            //throw new NotImplementedException("Creation of template plane not supported");
        }
    }

    /*/// <summary>
    /// </summary>
    public class CopyOfTubeService : TemplateServiceTubeRuleBase
    {
        private enum BasePlaneType
        {
            BySystem = 7100,
            NormalToAxis = 7101,
            Global = 7102,
            UserDefined = 7103
        };


        private enum TemplateSetType
        {
            ShortDistanceBCL = 7100,
            LongDistanceBCL = 7101,
            ModelOrigin = 7102,
            GlobalXMax = 7103,
            GlobalXMin = 7104,
            GlobalYMax = 7105,
            GlobalYMin = 7106,
            GlobalZMax = 7107,
            GlobalZMin = 7108
        };


        /// <summary>
        /// Creates the base plane for tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set information</param>
        /// <param name="userDefinedPlane">Boolean indicates whether base plane is user defined or not.</param>
        public override Plane3d CreateBasePlane(TemplateSetInformation inputInfo, out bool userDefinedPlane)
        {
            userDefinedPlane = false;
            if (inputInfo == null)
            {
                throw new CmnNullArgumentException("inputInfo");
            }

            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }
            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;
            if (parentMemeber == null)
            {
                throw new ArgumentNullException("parentMemeber");
            }

            ManufacturingOutputBase mfgOutBase = inputInfo.ManufacturingPart as ManufacturingOutputBase;
            CodelistItem codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessTube", "BasePlane");
            string basePlaneType = codeItem.ShortDisplayName;

            Plane3d basePlane = null;

            try
            {
                switch (basePlaneType)
                {
                    //codereview: private enums for code lists.Done
                    case "BySystem":
                        {
                            //codereview: replace transient class with function overload.done                       
                            basePlane = base.GetBasePlane(BasePlaneConstructionType.BySystem, tubeInfo);
                            break;
                        }
                    case "NormalToAxis":
                        {
                            basePlane = base.GetBasePlane(BasePlaneConstructionType.NormalToAxis, tubeInfo);
                            break;
                        }
                    case "Global":
                        {
                            basePlane = base.GetBasePlane(BasePlaneConstructionType.Global, tubeInfo);
                            break;
                        }
                }

                if (basePlane == null)
                {
                    base.WriteToErrorLog("Template Service routine Create Base Plane: Failed to Create Base Plane", "RULES");
                }
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Create Base Plane: Failed to Create Base Plane", e.Message);
            }

            return basePlane;
        }


        /// <summary>
        /// Creates the control line for tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set tube information</param>
        /// <param name="basePlane">input base plane of the templaet set.</param>
        /// <returns> Returns the  base control line for a tempalte set.</returns>

        public override IComplexString CreateControlLine(TemplateSetInformation inputInfo, Plane3d basePlane)
        {
            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }
            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;
            if (parentMemeber == null)
            {
                throw new CmnNullArgumentException("parentMemeber");
            }

            IComplexString baseControlLine = null;
            ManufacturingOutputBase mfgOutBase = tubeInfo.ManufacturingPart as ManufacturingOutputBase;
            CodelistItem codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessTube", "Type");
            string templetSetType = codeItem.ShortDisplayName;

            // TemplateSetType templetSetType = (TemplateSetType)tubeInfo.TemplateProcessType;


            TubeType type = TubeType.GlobalXMax;

            if (templetSetType == "ShortStringValue")
            {
                type = (TubeType.ShortDistanceBCL);
            }
            else if (templetSetType == "LongDistanceBCL")
            {
                type = (TubeType.LongDistanceBCL);
            }
            else if (templetSetType == "ShortStringValue")
            {
                type = (TubeType.ModelOrigin);
            }
            else if (templetSetType == "Global_X_Max")
            {
                type = (TubeType.GlobalXMax);
            }
            else if (templetSetType == "Global_X_Min")
            {
                type = (TubeType.GlobalXMin);
            }
            else if (templetSetType == "Global_Y_Max")
            {
                type = (TubeType.GlobalYMax);
            }
            else if (templetSetType == "Global_Y_Min")
            {
                type = (TubeType.GlobalYMin);
            }
            else if (templetSetType == "Global_Z_Max")
            {
                type = (TubeType.GlobalZMax);
            }
            else if (templetSetType == "Global_Z_Min")
            {
                type = (TubeType.GlobalZMin);
            }


            //codereview: replace transient class with function overload.done
            try
            {
                baseControlLine = base.GetBaseControlLine(type, tubeInfo, basePlane);
                if (baseControlLine == null)
                {
                    base.WriteToErrorLog("Template Service routine Create BaseControl Line: Failed to create Base Control Line", "RULES");
                }
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Create BaseControl Line: Failed to create Base Control Line", e.Message);
            }

            return baseControlLine;
        }


        /// <summary>
        /// Creates the out contours for tube memeber template ser
        /// </summary>
        /// <param name="inputInfo">Input template set tube information</param>
        /// <param name="basePlane">input base plane of the templaet set.</param>
        ///  <returns> Returns the template out contours for a template set.</returns>

        public override ReadOnlyCollection<TemplateContourInformation> CreateTemplates(TemplateSetInformation inputInfo, ref Plane3d basePlane, IComplexString baseControlLine)
        {
            TemplateSetTubeInformation tubeInfo = inputInfo as TemplateSetTubeInformation;
            ReadOnlyCollection<TemplateContourInformation> templateOutContours = null;
            try
            {
                templateOutContours = base.GetTemplateContours(tubeInfo, basePlane, baseControlLine);
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
        /// Creates unfold out contour for a tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set information</param>
        /// <returns> Returns the collection of unfolded contours for a templaet set.</returns>

        public override ReadOnlyCollection<ManufacturingGeometry> CreateTemplateUnfoldContour(TemplateSetInformation inputInfo, Plane3d basePlane)
        {

            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }
            ReadOnlyCollection<ManufacturingGeometry> unfoldOutContours = null;

            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;
            if (parentMemeber == null)
            {
                throw new CmnNullArgumentException("parentMemeber");
            }

            try
            {
                unfoldOutContours = base.GetUnfoldedTemplateContour(inputInfo, basePlane);
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Template Contours: Failed to Create Template Contours", e.Message);
            }

            return unfoldOutContours;
        }


        /// <summary>
        /// Creates unfold reference curve for a tube member template set
        /// </summary>
        /// <param name="inputInfo">Input template set tube information.</param>
        /// <returns> Returns the unfolded reference curve for a template set.</returns>

        public override IComplexString CreateTemplateUnfoldReferenceCurve(TemplateSetInformation inputInfo, Plane3d basePlane)
        {
            TemplateSetTubeInformation tubeInfo = (TemplateSetTubeInformation)inputInfo;
            if (tubeInfo == null)
            {
                throw new CmnNullArgumentException("tubeInfo");
            }

            IComplexString unfoldRefCurve = null;
            BusinessObject parentMemeber = (BusinessObject)tubeInfo.ManufacturingParent;

            if (parentMemeber == null)
            {
                throw new CmnNullArgumentException("parentMemeber");
            }

            try
            {
                unfoldRefCurve = base.GetUnfoldedTemplateReferenceCurve(inputInfo, basePlane);
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Template unfold reference curve: Failed to Create Template Contours", e.Message);
            }

            return unfoldRefCurve;
        }


        /// <summary>
        /// Creates template plane for a tube member template  
        /// </summary>
        /// <param name="inputInfo">Input templateSet information</param>
        /// <param name="templatePosition">Template position along it's base control line.</param>
        /// <returns> Returns the template plane at a specified location along the base control line.</returns>

        public override Plane3d CreateTemplatePlane(TemplateSetInformation inputInfo, Position templatePosition, Plane3d basePlane)
        {
            // throw new NotImplementedException("Creation of template plane not supported");

            return null;
        }
    }*/
}

