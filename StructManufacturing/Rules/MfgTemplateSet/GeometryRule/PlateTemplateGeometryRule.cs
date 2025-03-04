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
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Structure.Middle;


namespace Ingr.SP3D.Content.Manufacturing
{
   public class PlateTemplateGeometryRule : FaceTemplateGeometryRule
    {

        //private members

        // private TemplateSet mTemplateSet = null;        

        private enum BasePlaneTypeCatalog
        {
            AverageCornersPlane = 5170,
            MostPlanarNatural = 5171,
            BySystem = 5172,
            ParallelAxis = 5173,
            TrueNatural = 5174,
            LowerAftCorners = 5175,
            UpperAftCorners = 5176,
            LowerForeCorners = 5177,
            UpperForeCorners = 5178,
            UserDefined = 5179,
            MidTemplate = 5180
        };

        private enum DirectionCatalog
        {
            Transversal = 5180,
            Longitudinal = 5181,
            Waterline = 5182
        };

        private enum TemplateSetTypeCatalog
        {
            Frame = 5140,
            CenterLine = 5141,
            Perpendicular = 5142,
            StermStern = 5143,
            PerpendicularXY = 5144,
            UserDefined = 5145,
            AftForward = 5146,
            Even = 5147,
            Box = 5148,
            UserDefinedBox = 5149,
            UserDefinedBoxWithEdges = 5150
        };

        private enum PlateOrientationCatalog
        {
            AlongFrame = 5150,
            NormalToBaseplane = 5151,
            Perpendicular = 5152
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
        /// Creates the base plane for a plate part template set.
        /// </summary>
        /// <param name="inputPlaneInfo">Input template set information</param>
        /// <param name="userDefinePlane">Indicates the whether base plane user defined or not.</param>
        /// <returns>returns the base of the template set.</returns>

        public override Plane3d CreateBasePlane(TemplateSetInformation inputInfo, out bool userDefinedPlane)
        {
            userDefinedPlane = false;
            if (inputInfo == null)
            {
                throw new CmnNullArgumentException("inputInfo");
            }

            TemplateSetPlateInformation plateInfo = (TemplateSetPlateInformation)inputInfo;

            BusinessObject platePart = inputInfo.ManufacturingParent;

            CodelistItem codeItem = null;
            int directionCodeListValue = 0;
            int basePlaneCodeListvalue = 0;
            Vector primaryDirection = null;
            TemplateSet templaetSetObj = (TemplateSet)inputInfo.ManufacturingPart;
            ManufacturingOutputBase mfgOutBase = templaetSetObj as ManufacturingOutputBase;
            if (mfgOutBase != null)
            {
                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "Direction");
                directionCodeListValue = codeItem.Value;

                primaryDirection = GetDirection((DirectionCatalog)directionCodeListValue);
                plateInfo.PrimaryDirection = primaryDirection;

                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "BasePlane");
                basePlaneCodeListvalue = codeItem.Value;
            } 

            Collection<TopologySurface> collSurfaces = new Collection<TopologySurface>();
            collSurfaces.Add((TopologySurface)inputInfo.PartSurface);


            Plane3d basePlane = null;

            //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
            //1. Get Inputs
            //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
            double planarity = 0.0;

            try
            {
                switch ((BasePlaneTypeCatalog)basePlaneCodeListvalue)
                {

                    case BasePlaneTypeCatalog.AverageCornersPlane:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.AverageOfCorners, (TopologySurface)plateInfo.PartSurface, primaryDirection);
                        }
                        break;

                    case BasePlaneTypeCatalog.MostPlanarNatural:
                        {
                            basePlane = base.GetBalancingPlane((TopologySurface)plateInfo.PartSurface, out planarity);
                        }
                        break;

                    case BasePlaneTypeCatalog.BySystem:
                        {
                            PlateGeometryType plateGeomType = GetPlateType(plateInfo.PartSurface, plateInfo.Type);
                            if (plateGeomType == PlateGeometryType.SymToCenterLine)
                            {
                                basePlane = base.GetBalancingPlane(BalanceCreationType.AverageOfCorners, (TopologySurface)plateInfo.PartSurface, primaryDirection);
                            }
                            else if (plateGeomType == PlateGeometryType.PerpendicularToXY)
                            {
                                basePlane = base.GetBalancingPlane(BalanceCreationType.PerpendicularToXY, (TopologySurface)plateInfo.PartSurface, primaryDirection);
                            }
                            else if (plateGeomType == PlateGeometryType.FlatPlate)
                            {
                                basePlane = base.GetBalancingPlane((TopologySurface)plateInfo.PartSurface, out planarity);
                            }
                            else if (plateGeomType == PlateGeometryType.NormalPlate)
                            {
                                basePlane = base.GetBalancingPlane(BalanceCreationType.AverageOfCorners, (TopologySurface)plateInfo.PartSurface, primaryDirection);
                            }
                            else if (plateGeomType == PlateGeometryType.Box)
                            {
                                basePlane = base.GetBalancingPlane(BalanceCreationType.ParallelToGlobalAxis, (TopologySurface)plateInfo.PartSurface);
                            }

                        }
                        break;

                    case BasePlaneTypeCatalog.ParallelAxis:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.ParallelToGlobalAxis, (TopologySurface)plateInfo.PartSurface);
                        }
                        break;

                    case BasePlaneTypeCatalog.TrueNatural:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.TrueNatural, (TopologySurface)plateInfo.PartSurface);
                        }
                        break;

                    case BasePlaneTypeCatalog.LowerAftCorners:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.BottomLeftCorners, collSurfaces);
                        }
                        break;
                    case BasePlaneTypeCatalog.UpperAftCorners:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.TopLeftCorners, collSurfaces);
                        }
                        break;

                    case BasePlaneTypeCatalog.LowerForeCorners:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.BottomRightCorners, collSurfaces);
                        }
                        break;
                    case BasePlaneTypeCatalog.UpperForeCorners:
                        {
                            basePlane = base.GetBalancingPlane(BalanceCreationType.TopRightCorners, collSurfaces);
                        }
                        break;

                    case BasePlaneTypeCatalog.UserDefined:
                        {
                            userDefinedPlane = true;
                            basePlane = null;
                        }
                        break;

                    case BasePlaneTypeCatalog.MidTemplate:
                        {

                            int count = plateInfo.PositionInformation.GridPlanes.Count;
                            GridPlaneBase midFrame = null;

                            if (count > 0)
                            {
                                int midFrameNo = (count + 1) / 2;
                                midFrame = plateInfo.PositionInformation.GridPlanes[midFrameNo - 1];
                                basePlane = base.GetBalancingPlane((TopologySurface)plateInfo.PartSurface, midFrame);
                            }

                        }
                        break;
                }

                if (basePlaneCodeListvalue != 5179 && basePlane == null)
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
        /// Creates the base control line for plate template set
        /// </summary>
        /// <param name="inputInfo">input plate process setting information</param>  
        /// <param name="basePlane">input base plane of the templaet set.</param>
        public override IComplexString CreateControlLine(TemplateSetInformation inputInfo, Plane3d basePlane)
        {

            if (inputInfo == null)
            {
                throw new CmnNullArgumentException("inputInfo");
            }

            TemplateSetPlateInformation plateInfo = (TemplateSetPlateInformation)inputInfo;

            if (plateInfo == null)
            {
                throw new CmnNullArgumentException("plateInfo");
            }

            IComplexString baseControlLine = null;
            int directionCodeListValue = 0;
            Vector primaryDirection = null;

            TemplateSet templaetSetObj = (TemplateSet)inputInfo.ManufacturingPart;
            ManufacturingOutputBase mfgOutBase = templaetSetObj as ManufacturingOutputBase;
            if( mfgOutBase != null )
            {
                CodelistItem codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "Direction");
                directionCodeListValue = codeItem.Value;

                primaryDirection = GetDirection((DirectionCatalog)directionCodeListValue);
                plateInfo.PrimaryDirection = primaryDirection;
            }
   
            PlateGeometryType plateGeomType = GetPlateType(plateInfo.PartSurface, plateInfo.Type);
            try
            {
                if (plateGeomType == PlateGeometryType.SymToCenterLine)
                {
                    Vector symDirection = new Vector(0, 1, 0);
                    Position rootPt = new Position(0, 0, 0);
                    Plane3d tempPlane = new Plane3d(rootPt, symDirection);
                    baseControlLine = base.GetControlLine(ControlLineConstructionType.ByPlane, plateInfo.PartSurface, tempPlane);

                }
                else if (plateGeomType == PlateGeometryType.PerpendicularToXY)
                {
                    baseControlLine = base.GetControlLine(ControlLineConstructionType.PerpendicularToXY, (IManufacturable)plateInfo.ManufacturingParent, plateInfo.PartSurface, basePlane.Normal, primaryDirection);
                }
                else if (plateGeomType == PlateGeometryType.Box)
                {
                    baseControlLine = null;
                    return null;
                }
                else
                {
                    plateGeomType = PlateGeometryType.NormalPlate;
                    baseControlLine = base.GetControlLine(ControlLineConstructionType.Default, (IManufacturable)plateInfo.ManufacturingParent, plateInfo.PartSurface, basePlane.Normal, primaryDirection);
                }

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
        /// Creates the out contours such as bottom line,
        /// top line and side lines for a template set
        /// </summary>
        /// <param name="inputInfo">Input template set information</param>
        public override ReadOnlyCollection<TemplateContourInformation> CreateTemplates(TemplateSetInformation inputInfo, ref Plane3d basePlane, IComplexString baseControlLine)
        {

            if (inputInfo == null)
                throw new CmnNullArgumentException("inputInfo");

            TemplateSetPlateInformation plateInfo = (TemplateSetPlateInformation)inputInfo;
            TemplateSetSurfaceInformation surfaceSetInfo = inputInfo as TemplateSetSurfaceInformation;

            if (plateInfo == null)
                throw new CmnNullArgumentException("plateInfo");

            BusinessObject platePart = plateInfo.ManufacturingParent;
            if (platePart == null)
            {
                throw new CmnNullArgumentException("platePart");
            }

            BusinessObject templateSetBo = plateInfo.ManufacturingPart;
            if (templateSetBo == null)
            {
                throw new CmnNullArgumentException("templateSetBo");
            }

            TemplateSet templaetSetObj = (TemplateSet)templateSetBo;
            if (templaetSetObj == null)
            {
                throw new CmnNullArgumentException("templaetSetObj");
            }


            Vector primaryDirection = null;
            Vector secondaryDirection = null;


            ManufacturingOutputBase mfgOutBase = templaetSetObj as ManufacturingOutputBase;
            if (mfgOutBase != null)
            {
                CodelistItem codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "Direction");
                int directionCodeListValue = codeItem.Value;

                primaryDirection = GetDirection((DirectionCatalog)directionCodeListValue);
                secondaryDirection = GetSecondaryDirection(inputInfo, primaryDirection);
                plateInfo.PrimaryDirection = primaryDirection;
                plateInfo.SecondaryDirection = secondaryDirection;

                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "Orientation");
                int orientationCoeListvalue = codeItem.Value;

                TemplateProcessInfo.Orientation templateOrientation = GetOrientation((PlateOrientationCatalog)orientationCoeListvalue);
                plateInfo.TemplateOrientation = templateOrientation;
            }


            ReadOnlyCollection<TemplateContourInformation> templateContoursColl = null;
            plateInfo.AreEqualHeightTemplates = false;

            string templateSetType = templaetSetObj.ProcessSettings;
            if (templateSetType == "FrameEqualHeight_TemplateProcessPlate")
                plateInfo.AreEqualHeightTemplates = true;

            try
            {
                switch (plateInfo.Type)
                {
                    case "FRAME":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                        }
                        break;

                    case "CENTERLINE":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                        }
                        break;

                    case "PERPENDICULAR":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                        }
                        break;

                    case "STEM/STERN":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.StemStern, ref basePlane, baseControlLine);
                        }
                        break;

                    case "PERPENDICULARXY":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.PerpendicularXY, ref basePlane, baseControlLine);
                        }
                        break;

                    case "USERDEFINED":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                        }
                        break;

                    case "AFT/FORWARD":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                        }
                        break;

                    case "EVEN":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Standard, ref basePlane, baseControlLine);
                        }
                        break;

                    case "BOX":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Box, ref basePlane, baseControlLine);
                        }
                        break;

                    case "USERDEFINED BOX":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Box, ref basePlane, baseControlLine);
                        }
                        break;

                    case "USERDEFINED BOX WITH EDGES":
                        {
                            templateContoursColl = base.GetTemplateContours(surfaceSetInfo, Content.Manufacturing.TemplateProcessInfo.ContourType.Box, ref basePlane, baseControlLine);
                        }
                        break;
                }

                if (templateContoursColl == null)
                {
                    base.WriteToErrorLog("Template Service routine Template Contours: Failed to Create Template Contours", "RULES");
                }
            }
            catch (Exception e)
            {
                base.WriteToErrorLog(e, "Template Service routine Template Contours: Failed to Create Template Contours", e.Message);
            }

            return templateContoursColl;
        }


        /// <summary>
        /// Creates the template plane for a template
        /// </summary>
        /// <param name="inputInfo">Input template set information</param>
        /// <param name="templatePosition">Template position along it's base control line.</param>
        /// <param name="inputInfo">Input base plane of the template set.</param> 
        ///  <returns> Returns the template plane at a specified location along the base control line.</returns>
        public override Plane3d CreateTemplatePlane(TemplateSetInformation inputInfo, Position templatePosition, Plane3d basePlane)
        {
            if (inputInfo == null)
                throw new CmnNullArgumentException("inputInfo");

            Plane3d templatePlane = null;
            TemplateSetPlateInformation plateInfo = (TemplateSetPlateInformation)inputInfo;

            BusinessObject platePart = inputInfo.ManufacturingParent;
            if (platePart == null)
            {
                throw new CmnNullArgumentException("platePart");
            }

            CodelistItem codeItem = null;
            int directionCodeListValue = 0;
            ManufacturingOutputBase mfgOutBase = inputInfo.ManufacturingPart as ManufacturingOutputBase;            
            if (mfgOutBase != null)
            {
                codeItem = mfgOutBase.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "Direction");
                directionCodeListValue = codeItem.Value;
                Vector primaryDirection = GetDirection((DirectionCatalog)directionCodeListValue);
                Vector secondaryDirection = GetSecondaryDirection(inputInfo, primaryDirection);
                plateInfo.PrimaryDirection = primaryDirection;
                plateInfo.SecondaryDirection = secondaryDirection;
            }
           
            templatePlane = (Plane3d)(base.GetTemplatePlane((TemplateSetSurfaceInformation)inputInfo, basePlane, templatePosition, 1));
            if (templatePlane == null)
            {
                base.WriteToErrorLog("Template Service routine Template Contours: Failed to Create Bottom Lines", "RULES");
            }

            return templatePlane;
        }

        private Vector GetDirection(DirectionCatalog directionCodeListValue)
        {
            Vector directionVector = new Vector();
            if (directionCodeListValue == DirectionCatalog.Transversal)
            {
                directionVector.X = 1;
                directionVector.Y = 0;
                directionVector.Z = 0;
            }
            else if (directionCodeListValue == DirectionCatalog.Longitudinal)
            {
                directionVector.X = 0;
                directionVector.Y = 1;
                directionVector.Z = 0;
            }
            else if (directionCodeListValue == DirectionCatalog.Waterline)
            {
                directionVector.X = 0;
                directionVector.Y = 0;
                directionVector.Z = 1;
            }
            return directionVector;
        }

        private PlateGeometryType GetPlateType(ISurface surfaceBody, string templateSetType)
        {
            PlateGeometryType plategeomType = PlateGeometryType.FlatPlate;
            if (templateSetType == "STEM / STERN" || templateSetType == "CENTERLINE")  // stem/stern or CenterLine
            {
                plategeomType = PlateGeometryType.SymToCenterLine;
            }
            else if (templateSetType == "PERPENDICULARXY")
            {
                plategeomType = PlateGeometryType.PerpendicularToXY;
            }
            else if (templateSetType == "BOX" || templateSetType == "USERDEFINED BOX" || templateSetType == "USERDEFINED BOX EDGES")
            {
                plategeomType = PlateGeometryType.Box;

            }
            else
            {
                double ratio = 0.0;
                EntityService.GetPlanarRatio(surfaceBody, out ratio);

                if (ratio > 0.33)
                    plategeomType = PlateGeometryType.FlatPlate;
                else
                    plategeomType = PlateGeometryType.NormalPlate;

            }
            return plategeomType;
        }

        private Content.Manufacturing.TemplateProcessInfo.Orientation GetOrientation(PlateOrientationCatalog orientation)
        {
            Content.Manufacturing.TemplateProcessInfo.Orientation templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
            if (orientation == PlateOrientationCatalog.AlongFrame)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
            }
            else if (orientation == PlateOrientationCatalog.NormalToBaseplane)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.PerpendicularToBasePlane;
            }
            else if (orientation == PlateOrientationCatalog.Perpendicular)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.Perpendicular;
            }

            return templateOrientation;
        }

        /*public override Content.Manufacturing.TemplateProcessInfo.Orientation GetOrientation(int orientationCodeListvalue)
        {
            Content.Manufacturing.TemplateProcessInfo.Orientation templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
            if (orientationCodeListvalue == (int)PlateOrientationCatalog.AlongFrame)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.AlongFrame;
            }
            else if (orientationCodeListvalue == (int)PlateOrientationCatalog.NormalToBaseplane)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.NormalToBasePlane;
            }
            else if (orientationCodeListvalue == (int)PlateOrientationCatalog.Perpendicular)
            {
                templateOrientation = Content.Manufacturing.TemplateProcessInfo.Orientation.Perpendicular;
            }
            return templateOrientation;
        }*/


        private Vector GetSecondaryDirection(TemplateSetInformation inputInfo, Vector primaryDirection)
        {
            TemplateSetPlateInformation plateInfo = (TemplateSetPlateInformation)inputInfo;
            Vector secondaryDirection = new Vector();
            if (plateInfo.Type == "BOX" || plateInfo.Type == "USERDEFINED BOX" ||
                 plateInfo.Type == "USERDEFINED BOX WITH EDGES")
            {

                Vector ParallelVec = GetParallelVector(plateInfo.PartSurface);
                Vector secondaryDirVec = ParallelVec.Cross(primaryDirection);
                if (StructHelper.AreEqual(Math.Abs(secondaryDirVec.X), 1) &&
                         StructHelper.AreEqual(Math.Abs(secondaryDirVec.Y), 0) &&
                         StructHelper.AreEqual(Math.Abs(secondaryDirVec.Z), 0))
                {
                    secondaryDirection.Set(1, 0, 0);

                }
                else if (StructHelper.AreEqual(Math.Abs(secondaryDirVec.X), 0) &&
                         StructHelper.AreEqual(Math.Abs(secondaryDirVec.Y), 1) &&
                         StructHelper.AreEqual(Math.Abs(secondaryDirVec.Z), 0))
                {
                    secondaryDirection.Set(0, 1, 0);
                }
                else
                {
                    secondaryDirection.Set(0, 0, 1);
                }
            }

            return secondaryDirection;
        }


        private Vector GetParallelVector(ISurface surfaceBody)
        {
            Position center = null;
            Vector normal = null;
            EntityService.FindApproxCenterAndNormal(surfaceBody, out center, out normal);
            Vector oXVector = new Vector();
            Vector oYVector = new Vector();
            Vector oZVector = new Vector();
            oXVector.Set(1, 0, 0);
            oXVector.Length = 1;
            oYVector.Set(0, 1, 0);
            oYVector.Length = 1;
            oZVector.Set(0, 0, 1);
            oZVector.Length = 1;

            double dX = 0.0, dY = 0.0, dZ = 0.0;
            dX = Math.Abs(oXVector.Dot(normal));
            dY = Math.Abs(oYVector.Dot(normal));
            dZ = Math.Abs(oZVector.Dot(normal));

            Vector outvector = new Vector();
            outvector.Length = 1;

            if ((dX > dY) && (dX > dZ))
                outvector.Set(1, 0, 0);
            else if (dY > dX && dY > dZ)
                outvector.Set(0, 1, 0);
            else
                outvector.Set(0, 0, 1);
            return outvector;
        }
        

    }
}
