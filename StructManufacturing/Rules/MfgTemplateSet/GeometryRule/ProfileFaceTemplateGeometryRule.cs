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
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;
using Ingr.SP3D.Grids.Middle;
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Structure.Middle;



namespace Ingr.SP3D.Content.Manufacturing
{
    public class ProfileFaceTemplateGeometryRule : FaceTemplateGeometryRule
    {

        private enum BasePlaneTypeCatalog
        {
            AverageCornersPlane = 5370,
            MostPlanarNatural = 5371,
            BySystem = 5372,
            ParallelAxis = 5373,
            TrueNatural = 5374,
            UserDefined = 5375
        };

        private enum TemplateSetTypeCatalog
        {
            Frame = 5340,
            Perpendicular = 5341,
            UserDefined = 5342,
            AftForward = 5343,
            Even = 5344

        };

        private enum DirectionCatalog
        {
            PerpendicularAxis = 5380,
            AlongAxis = 5381
        };

        private enum ProfileOrientationCatalog
        {
            AlongFrame = 5350,
            NormalToBaseplane = 5351,
            Perpendicular = 5352
        };

        private enum PlateGeometryType
        {
            /// <summary>
            /// Indicates the plate is symmetrical to center line
            /// </summary>

            SymToCenterLine = 0,

            /// <summary>
            /// Indicates the plate is perpendicular to xy plane
            /// </summary>

            PerpendicularToXY = 1,

            /// <summary>
            /// indicates the plate is of Box type
            /// </summary>

            Box = 2,

            /// <summary>
            /// Indicates the  plate is a flat type
            /// </summary>

            FlatPlate = 3,

            /// <summary>
            /// Indicates the plate is a normal plate
            /// </summary>

            NormalPlate = 4
        };


        /// <summary>
        /// Provides the base plane for a profile template set.
        /// </summary>
        /// <param name="inputPlaneInfo">Input template set information</param>
        /// <param name="userDefinePlane">Indicates the whether base plane user defined or not.</param>
        /// <returns></returns>

        public override Plane3d CreateBasePlane(TemplateSetInformation inputInfo, out bool userDefinedPlane)
        {
            userDefinedPlane = false;

            TemplateSetProfileInformation profileInfo = (TemplateSetProfileInformation)inputInfo;

            if (profileInfo == null)
                throw new ArgumentNullException("profileInfo");

            BusinessObject profilePart = profileInfo.ManufacturingParent;

            if (profilePart == null)
            {
                throw new ArgumentNullException("profilePart");
            }

            Plane3d basePlane = null;
            TemplateSet templateSetobj = (TemplateSet)profileInfo.ManufacturingPart;

            ManufacturingOutputBase mfgOutBase = templateSetobj as ManufacturingOutputBase;
            CodelistItem codeItem = null;
            int directionCodeListVal = 0;
            int basePlaneCodeListValue = 0;
            if (mfgOutBase != null)
            {
                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessProfile", "Direction");
                directionCodeListVal = codeItem.Value;

                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessProfile", "BasePlane");
                basePlaneCodeListValue = codeItem.Value;
            }
            
            double planarity = 0.0;
            Vector primaryDirection = null;
            try
            {
                switch ((BasePlaneTypeCatalog)basePlaneCodeListValue)
                {
                    case BasePlaneTypeCatalog.AverageCornersPlane:
                        {
                            Vector profileAxis = base.GetProfileDirection(profileInfo.PartSurface);
                            basePlane = base.GetBalancingPlane(BalanceCreationType.AverageOfCorners, (TopologySurface)profileInfo.PartSurface, profileAxis);
                        }
                        break;

                    case BasePlaneTypeCatalog.MostPlanarNatural:
                        {
                            basePlane = base.GetBalancingPlane((TopologySurface)profileInfo.PartSurface, out planarity);
                        }
                        break;

                    case BasePlaneTypeCatalog.BySystem:
                        {
                            PlateGeometryType plateGeomType = GetPlateType(profileInfo.PartSurface);
                            if (plateGeomType == PlateGeometryType.FlatPlate)
                            {
                                basePlane = base.GetBalancingPlane((TopologySurface)profileInfo.PartSurface, out planarity);
                            }
                            else
                            {
                                if ((DirectionCatalog)directionCodeListVal == DirectionCatalog.PerpendicularAxis)
                                    primaryDirection = base.GetProfileDirection(profileInfo.PartSurface);
                                else
                                {
                                    primaryDirection = base.GetProfileSecondaryDirection(profileInfo.PartSurface);
                                }

                                basePlane = base.GetBalancingPlane(BalanceCreationType.AverageOfCorners, (TopologySurface)profileInfo.PartSurface, primaryDirection);
                            }
                        }
                        break;

                    case BasePlaneTypeCatalog.ParallelAxis:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.ParallelToGlobalAxis, (TopologySurface)profileInfo.PartSurface);
                        }
                        break;

                    case BasePlaneTypeCatalog.TrueNatural:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.TrueNatural, (TopologySurface)profileInfo.PartSurface);
                        }
                        break;

                    case BasePlaneTypeCatalog.UserDefined:
                        {
                            userDefinedPlane = true;
                            basePlane = null;
                        }
                        break;
                }

                if (basePlaneCodeListValue != 5375 && basePlane == null)
                {
                    base.WriteToErrorLog("Template Service routine Create Base Plane: Failed to Create Base Plane", "RULES");
                }

            }
            catch (Exception e)
            {
                base.WriteToErrorLog("Template Service routine Create Base Plane: Failed to Create Base Plane", e.Message);
            }

            return basePlane;
        }


        /// <summary>
        /// Creates the base control line for profile template set
        /// </summary>
        /// <param name="inputInfo">input profile process setting information</param>
        /// <param name="basePlane">input base plan eof the template set.</param>
        public override IComplexString CreateControlLine(TemplateSetInformation inputInfo, Plane3d basePlane)
        {
            if (inputInfo == null)
            {
                throw new CmnNullArgumentException("inputInfo");
            }

            if (basePlane == null)
            {
                throw new CmnNullArgumentException("basePlane");
            }

            TemplateSetProfileInformation profileInfo = (TemplateSetProfileInformation)inputInfo;

            Vector primaryDirection = null;
            IComplexString baseControlLine = null;

            TemplateSet templateSetobj = (TemplateSet)profileInfo.ManufacturingPart;

            ManufacturingOutputBase mfgOutBase = templateSetobj as ManufacturingOutputBase;
            CodelistItem codeItem = null;

            codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessProfile", "Direction");
            int directionCodeListVal = codeItem.Value;


            if ((DirectionCatalog)directionCodeListVal == DirectionCatalog.PerpendicularAxis)
                primaryDirection = base.GetProfileDirection(profileInfo.PartSurface);
            else
            {
                primaryDirection = base.GetProfileDirection(profileInfo.PartSurface);
                primaryDirection = primaryDirection.Cross(basePlane.Normal);
            }

            profileInfo.PrimaryDirection = primaryDirection;

            try
            {
                baseControlLine = base.GetControlLine(ControlLineConstructionType.Default, (IManufacturable)profileInfo.ManufacturingParent, profileInfo.PartSurface, basePlane.Normal, primaryDirection);
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
        /// Creates the template out contourfor profile face templaets
        /// </summary>
        /// <param name="inputInfo">Input templateset information</param>
        /// <return> returns the collection of template out contours.</return>
        public override ReadOnlyCollection<TemplateContourInformation> CreateTemplates(TemplateSetInformation inputInfo, ref Plane3d basePlane, IComplexString baseControlLine)
        {
            if (inputInfo == null)
            {
                throw new CmnNullArgumentException("inputInfo");
            }

            if (basePlane == null)
            {
                throw new CmnNullArgumentException("basePlane");
            }

            if (baseControlLine == null)
            {
                throw new CmnNullArgumentException("baseControlLine");
            }

            ReadOnlyCollection<TemplateContourInformation> templateContoursColl = null;
            TemplateSetSurfaceInformation surfaceSetInfo = inputInfo as TemplateSetSurfaceInformation;
            TemplateSetProfileInformation profileInfo = (TemplateSetProfileInformation)inputInfo;
            BusinessObject profilePart = (BusinessObject)profileInfo.ManufacturingParent;
            TemplateSet templateSet = (TemplateSet)profileInfo.ManufacturingPart;

            ManufacturingOutputBase mfgOutBase = templateSet as ManufacturingOutputBase;
            CodelistItem codeItem = null;

            codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessProfile", "Orientation");
            int orientationCodeListvalue = codeItem.Value;
            TemplateProcessInfo.Orientation templateOrientation = GetOrientation((ProfileOrientationCatalog)orientationCodeListvalue);
            profileInfo.TemplateOrientation = templateOrientation;


            Collection<TemplateData> templateDataCol = new Collection<TemplateData>();

            codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessProfile", "Direction");
            int directionCodeListVal = codeItem.Value;

            Vector primaryDirection = null;
            if ((DirectionCatalog)directionCodeListVal == DirectionCatalog.PerpendicularAxis)
                primaryDirection = base.GetProfileDirection(profileInfo.PartSurface);
            else
            {
                primaryDirection = base.GetProfileDirection(profileInfo.PartSurface);
                primaryDirection = primaryDirection.Cross(basePlane.Normal);
            }

            profileInfo.PrimaryDirection = primaryDirection;

            try
            {
                switch (profileInfo.Type)
                {
                    case "FRAME":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                            break;
                        }
                    case "PERPENDICULAR":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                            break;
                        }
                    case "USERDEFINED":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                            break;
                        }
                    case "AFT/FORWARD":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                            break;
                        }
                    case "EVEN":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                            break;
                        }
                }

            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Template Contours: Failed to Create Template Contours", e.Message);
            }
            return templateContoursColl;
        }


        /// <summary>
        /// Creates the template plane
        /// </summary>
        /// <param name="inputInfo">Input templateset information</param>
        /// <param name="templatePosition">Template position along it's base control line.</param>
        /// <param name="basePlane">Input base plane of the emplateset.</param>
        public override Plane3d CreateTemplatePlane(TemplateSetInformation inputInfo, Position templatePosition, Plane3d basePlane)
        {
            if (inputInfo == null)
                throw new CmnNullArgumentException("inputInfo");

            if (templatePosition == null)
                throw new CmnNullArgumentException("templatePosition");

            if (basePlane == null)
                throw new CmnNullArgumentException("basePlane");

            Plane3d templatePlane = null;
            TemplateSetPlateInformation profileInfo = (TemplateSetPlateInformation)inputInfo;
            BusinessObject platePart = inputInfo.ManufacturingParent;
            if (platePart == null)
            {
                throw new CmnNullArgumentException("platePart");
            }

            ManufacturingOutputBase mfgOutBase = inputInfo.ManufacturingPart as ManufacturingOutputBase;
            CodelistItem codeItem = null;
            Vector primaryDirection = null;

            if (mfgOutBase != null)
            {
                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessProfile", "Direction");
                int directionCodeListVal = codeItem.Value;
               
                if ((DirectionCatalog)directionCodeListVal == DirectionCatalog.PerpendicularAxis)
                    primaryDirection = base.GetProfileDirection(profileInfo.PartSurface);
                else
                {
                    primaryDirection = base.GetProfileDirection(profileInfo.PartSurface);
                    primaryDirection = primaryDirection.Cross(basePlane.Normal);
                }
                profileInfo.PrimaryDirection = primaryDirection;
            }

            templatePlane = (Plane3d)(base.GetTemplatePlane((TemplateSetSurfaceInformation)inputInfo, basePlane, templatePosition, 1));
            return templatePlane;
        }


        private PlateGeometryType GetPlateType(ISurface surfaceBody)
        {
            PlateGeometryType plategeomType = PlateGeometryType.FlatPlate;
            double ratio = 0.0;
            EntityService.GetPlanarRatio(surfaceBody, out ratio);

            if (ratio > 0.33)
                plategeomType = PlateGeometryType.FlatPlate;
            else
                plategeomType = PlateGeometryType.NormalPlate;

            return plategeomType;
        }

        /* private Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo.Orientation GetProfileOrientation(ProfileOrientationCatalog orientation)
         {
             Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo.Orientation templateOrientation = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
             if (orientation == ProfileOrientationCatalog.AlongFrame)
             {
                 templateOrientation = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
             }
             else if (orientation == ProfileOrientationCatalog.NormalToBaseplane)
             {
                 templateOrientation = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo.Orientation.NormalToBasePlane;
             }
             else if (orientation == ProfileOrientationCatalog.Perpendicular)
             {
                 templateOrientation = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo.Orientation.Perpendicular;
             }

             return templateOrientation;
         }*/



        private Content.Manufacturing.TemplateProcessInfo.Orientation GetOrientation(ProfileOrientationCatalog orientation)
        {
            Content.Manufacturing.TemplateProcessInfo.Orientation templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
            if (orientation == ProfileOrientationCatalog.AlongFrame)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
            }
            else if (orientation == ProfileOrientationCatalog.NormalToBaseplane)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.PerpendicularToBasePlane;
            }
            else if (orientation == ProfileOrientationCatalog.Perpendicular)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.Perpendicular;
            }
            return templateOrientation;
        }       
    }
}

