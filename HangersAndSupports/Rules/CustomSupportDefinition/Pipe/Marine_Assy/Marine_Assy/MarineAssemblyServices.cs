//-----------------------------------------------------------------------------
// Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
// File
//      MarineAssemblyServices.cs
// Author:     
//      Vijaya     
//  04.Dec.2013   Rajeswari  DI-CP-241804 Modified the code as part of hardening
//
// Abstract:
//     Marine_Assy Commom Methods

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  12.aug.2014     PVK      TR-CP-257296 	Resolve coverity issues found in June 2014 report 
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//  11.Dec.2014     PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//  16-jul-2015     PVK      Resolve coverity issues found in July 2015 report
//  26-Oct-2015     PVK      Resolve coverity issues found in Octpber 2015 report
//  17/12/2015     Ramya     TR 284319	Multiple Record exception dumps are created on copy pasting supports
//  11/04/2016      PVK      Corrected the position of disposing hsMarineServiceDimPart
//-----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Linq;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    public static class MarineAssemblyServices
    {
        /// <summary>
        /// Defines Properties of BeamClamp
        /// </summary>
        public struct BeamCLProperties
        {
            /// <summary>
            /// The Begin Snip Anchor Point1 along Flange
            /// </summary>
            public CodelistItem VertSecCBAncPt1AlFlg;
            /// <summary>
            /// The Begin Snip Anchor Point1 along Web
            /// </summary>
            public CodelistItem VertSecCBAncPt1AlWeb;
            /// <summary>
            /// The Begin Snip Anchor Point2 along Flange
            /// </summary>
            public CodelistItem VertSecCBAncPt2AlFlg;
            /// <summary>
            /// The Begin Snip Anchor Point2 along Web
            /// </summary>
            public CodelistItem VertSecCBAncPt2AlWeb;
            /// <summary>
            /// The Hanger Beam Type
            /// </summary>
            public CodelistItem HgrBeamType1;
            /// <summary>
            /// The Hanger Beam Type 
            /// </summary>
            public CodelistItem HgrBeamType2;
            /// <summary>
            /// The Begin Cap Port Cardinal Point
            /// </summary>
            public CodelistItem CP1;
            /// <summary>
            /// The Neutral Port Cardinal Point
            /// </summary>
            public CodelistItem CP5;
        }

        /// <summary>
        /// Defines 
        /// </summary>
        public struct PADProperties
        {
            /// <summary>
            /// The Pad Part
            /// </summary>
            public string padPart;
            /// <summary>
            /// The Length of the Pad
            /// </summary>
            public double padLength;
            /// <summary>
            /// The thickness of Pad
            /// </summary>
            public double padThickness;
            /// <summary>
            /// The width of Pad
            /// </summary>
            public double padWidth;
        }

        /// <summary>
        /// defines the config index
        /// </summary>
        public struct ConfigIndex
        {
            /// <summary>
            /// Geometry type A
            /// </summary>
            public Plane A;
            /// <summary>
            /// Geometry type B
            /// </summary>
            public Plane B;
            /// <summary>
            /// Geometry type C
            /// </summary>
            public Axis C;
            /// <summary>
            /// Geometry type D
            /// </summary>
            public Axis D;

            public ConfigIndex(Plane a, Plane b, Axis c, Axis d)
            {
                this.A = a;
                this.B = b;
                this.C = c;
                this.D = d;
            }
        }

        /// <summary>
        /// This method Gets the largest pipe diameter.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>double</returns>
        /// <code>  
        /// 
        /// </code>
        public static double GetLargePipeDiameter(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                double[] pipeDiameter = new double[customSupportDefinition.SupportHelper.SupportedObjects.Count];
                string UnitType = string.Empty;
                for (int i = 0; i < customSupportDefinition.SupportHelper.SupportedObjects.Count; i++)
                {
                    PipeObjectInfo pipe = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(i + 1);
                    pipeDiameter[i] = pipe.NominalDiameter.Size;
                    UnitType = pipe.NominalDiameter.Units;
                }
                double largePipeDiameter = pipeDiameter.Max();


                return largePipeDiameter;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetLargePipeDiameter." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method Gets the route angle array and route is vertical or not
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="routeCount"> Number of Routes</param>
        /// <param name="isVerticalRoute">Boolean Value of Vertical route</param>       
        /// <returns>double array</returns>
        /// <code>
        ///  double[] routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(this, routeCount, out isVerticalRoute);
        /// </code>
        public static double[] GetRoutAngleAndIsRouteVertical(CustomSupportDefinition customSupportDefinition, int routeCount, out bool isVerticalRoute)
        {
            try
            {
                string routePortName = string.Empty;
                double[] angle = new double[routeCount];
                isVerticalRoute = false;
                for (int index = 1; index <= routeCount; index++)
                {
                    if (index == 1)
                        routePortName = "Route";
                    else
                        routePortName = "Route_" + index;
                    angle[index - 1] = customSupportDefinition.RefPortHelper.AngleBetweenPorts(routePortName, PortAxisType.X, OrientationAlong.Global_Z);

                    //when pipe is vertical
                    if (angle[index - 1] < (Math.Round(Math.Atan(1) * 4, 3)) / 4)
                    {
                        isVerticalRoute = true;
                        angle[index - 1] = angle[index - 1];
                    }
                    else if (angle[index - 1] > 3 * (Math.Round(Math.Atan(1) * 4, 3)) / 4)
                    {
                        isVerticalRoute = true;
                        angle[index - 1] = (Math.Round(Math.Atan(1) * 4, 3)) - angle[index - 1];
                    }
                    else
                    {
                        isVerticalRoute = false;
                        angle[index - 1] = (Math.Round(Math.Atan(1) * 4, 3)) / 2 - angle[index - 1];
                    }
                }
                return angle;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetRoutAngleAndIsRouteVertical." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method Gets the right side structure angle and Structure is vertical or not.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="isVerticalStruct">Structue Vetical   or not - Boolean</param>
        /// <param name="leftStructAngle">Left structure angle - double</param>       
        /// <param name="rightStructAngle">Right structure angle - double</param>  
        /// <code>      
        ///  MarineAssemblyServices.GetLeftRightStructAngleAndIsVerticalStructure(this, out isVerticalSruct, out leftStructAngle, out rightStructAngle);
        /// </code>
        public static void GetLeftRightStructAngleAndIsVerticalStructure(CustomSupportDefinition customSupportDefinition, out bool isVerticalStruct, out double leftStructAngle, out double rightStructAngle)
        {
            try
            {
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = GetIsLugEndOffsetApplied(customSupportDefinition);

                string[] structPort = new string[2];
                structPort = GetIndexedStructPortName(customSupportDefinition, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];

                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    leftStructAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z);
                    rightStructAngle = leftStructAngle;
                }
                else
                {
                    leftStructAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts(leftStructPort, PortAxisType.X, OrientationAlong.Global_Z);
                    rightStructAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts(rightStructPort, PortAxisType.X, OrientationAlong.Global_Z);
                }

                isVerticalStruct = false;

                //when structure is vertical
                if (leftStructAngle < (Math.Round(Math.Atan(1) * 4, 3)) / 4)
                {
                    isVerticalStruct = true;
                }
                else if (leftStructAngle > 3 * (Math.Round(Math.Atan(1) * 4, 3)) / 4)
                {
                    isVerticalStruct = true;
                    leftStructAngle = (Math.Round(Math.Atan(1) * 4, 3)) - leftStructAngle;
                    rightStructAngle = (Math.Round(Math.Atan(1) * 4, 3)) - rightStructAngle;
                }
                else
                {
                    isVerticalStruct = false;
                    leftStructAngle = (Math.Round(Math.Atan(1) * 4, 3)) / 2 - leftStructAngle;
                    rightStructAngle = (Math.Round(Math.Atan(1) * 4, 3)) / 2 - rightStructAngle;
                }

            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetLeftRightStructAngleAndIsVerticalStructure." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method Gets the route angle array and route is vertical or not
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="routeCount"> Number of Routes</param>
        /// <param name="isVerticalRoute">Boolean Value of Vertical route</param>       
        /// <returns>double array</returns>
        /// <code>
        ///  double[] routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(this, routeCount, out isVerticalRoute);
        /// </code>
        public static double[] GetRoutAngleAndIsRouteVertical(CustomSupportDefinition customSupportDefinition, double PI, int routeCount, out bool isVerticalRoute)
        {
            try
            {
                string routePortName = string.Empty;
                double[] angle = new double[routeCount];
                isVerticalRoute = false;
                for (int index = 1; index <= routeCount; index++)
                {
                    if (index == 1)
                        routePortName = "Route";
                    else
                        routePortName = "Route_" + index;
                    angle[index - 1] = customSupportDefinition.RefPortHelper.AngleBetweenPorts(routePortName, PortAxisType.X, OrientationAlong.Global_Z);
                    //when pipe is vertical
                    if (angle[index - 1] < PI / 4)
                    {
                        isVerticalRoute = true;
                        angle[index - 1] = angle[index - 1];
                    }
                    else if (angle[index - 1] > 3 * PI / 4)
                    {
                        isVerticalRoute = true;
                        angle[index - 1] = PI - angle[index - 1];
                    }
                    else
                    {
                        isVerticalRoute = false;
                        angle[index - 1] = PI / 2 - angle[index - 1];
                    }
                }
                return angle;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetRoutAngleAndIsRouteVertical." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method Gets the right side structure angle and Structure is vertical or not.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="isVerticalStruct">Structue Vetical   or not - Boolean</param>
        /// <param name="leftStructAngle">Left structure angle - double</param>       
        /// <param name="rightStructAngle">Right structure angle - double</param>  
        /// <code>      
        ///  MarineAssemblyServices.GetLeftRightStructAngleAndIsVerticalStructure(this, out isVerticalSruct, out leftStructAngle, out rightStructAngle);
        /// </code>
        public static void GetLeftRightStructAngleAndIsVerticalStructure(CustomSupportDefinition customSupportDefinition, double PI, out bool isVerticalStruct, out double leftStructAngle, out double rightStructAngle)
        {
            try
            {
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = GetIsLugEndOffsetApplied(customSupportDefinition);

                string[] structPort = new string[2];
                structPort = GetIndexedStructPortName(customSupportDefinition, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];

                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    leftStructAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z);
                    rightStructAngle = leftStructAngle;
                }
                else
                {
                    leftStructAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts(leftStructPort, PortAxisType.X, OrientationAlong.Global_Z);
                    rightStructAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts(rightStructPort, PortAxisType.X, OrientationAlong.Global_Z);
                }

                isVerticalStruct = false;
                //when structure is vertical
                if (leftStructAngle < PI / 4)
                {
                    isVerticalStruct = true;
                }
                else if (leftStructAngle > 3 * PI / 4)
                {
                    isVerticalStruct = true;
                    leftStructAngle = PI - leftStructAngle;
                    rightStructAngle = PI - rightStructAngle;
                }
                else
                {
                    isVerticalStruct = false;
                    leftStructAngle = PI / 2 - leftStructAngle;
                    rightStructAngle = PI / 2 - rightStructAngle;
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetLeftRightStructAngleAndIsVerticalStructure." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method Gets the Supporting Type for the Support
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>String SupportingType</returns>
        /// <code>
        ///  supportingType  = MarineAssemblyServices.GetSupportingTypes(this);
        /// </code>
        public static string GetSupportingTypes(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                string supportingType;
                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    supportingType = "Slab";
                else
                {
                    if (customSupportDefinition.SupportHelper.SupportingObjects.Count != 0)
                    {
                        if (customSupportDefinition.SupportHelper.SupportingObjects.Count > 1)
                        {
                            supportingType = "Steel";
                            if ((customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (customSupportDefinition.SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                                supportingType = "Slab";          //Two Slabs

                            else if ((customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (customSupportDefinition.SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member))
                                supportingType = "Slab-Steel";    //Slab then Steel

                            else if ((customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) && (customSupportDefinition.SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                                supportingType = "Steel-Slab";    //Steel then Slab

                            else
                                supportingType = "Steel";
                        }
                        else
                        {
                            if (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab)
                                supportingType = "Slab";
                            else if (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Wall)
                                supportingType = "Wall";
                            else if ((customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                                supportingType = "Steel";    //Steel                      
                            else
                                supportingType = "Slab";
                        }
                    }
                    else
                        supportingType = "Slab";
                }
                return supportingType;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetSupportingTypes." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method Gets BeamClamp Properties
        /// </summary>
        /// <param name="componentDictionary">Collection of keys and values - Dictionary</param>       
        /// <param name="PartName">Name of the Part - String</param>       
        /// <returns>BeamCLProperties</returns>
        /// <code>
        ///   MarineAssemblyServices.BeamCLProperties beamCLProperties = MarineAssemblyServices.GetBeamCLProperties(componentDictionary, VERTSECTION1);    
        /// </code>
        public static BeamCLProperties GetBeamCLProperties(Dictionary<string, SupportComponent> componentDictionary)
        {
            try
            {
                BeamCLProperties beamCLProperties;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                beamCLProperties.VertSecCBAncPt1AlFlg = metadataManager.GetCodelistInfo("hsCutbackAnchorPoint", "UDP").GetCodelistItem(4);
                beamCLProperties.VertSecCBAncPt1AlWeb = metadataManager.GetCodelistInfo("hsCutbackAnchorPoint", "UDP").GetCodelistItem(8);
                beamCLProperties.VertSecCBAncPt2AlFlg = metadataManager.GetCodelistInfo("hsCutbackAnchorPoint", "UDP").GetCodelistItem(6);
                beamCLProperties.VertSecCBAncPt2AlWeb = metadataManager.GetCodelistInfo("hsCutbackAnchorPoint", "UDP").GetCodelistItem(2);
                beamCLProperties.HgrBeamType1 = metadataManager.GetCodelistInfo("HsHgrBeamType", "UDP").GetCodelistItem(1);
                beamCLProperties.HgrBeamType2 = metadataManager.GetCodelistInfo("HsHgrBeamType", "UDP").GetCodelistItem(2);
                beamCLProperties.CP1 = metadataManager.GetCodelistInfo("CrossSectionCardinalPoints", "CMNSCH").GetCodelistItem(1);
                beamCLProperties.CP5 = metadataManager.GetCodelistInfo("CrossSectionCardinalPoints", "CMNSCH").GetCodelistItem(5);
                return beamCLProperties;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetBeamCLProperties of MarineAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method Gets Pad part and its dimensions.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="sectionSize"> Size of the section - string </param>  
        /// <param name="sectionCode"> Section code - string</param>  
        /// <param name="steelStadard"> steel standard - string</param>  
        /// <returns>PADProperties</returns>
        /// <code>
        /// PADProperties padProperties = MarineAssemblyServices.GetPadPartAndDimensions(this, verticalSectionSize, out sectionCode, out steelStd);
        /// </code>
        public static PADProperties GetPadPartAndDimensions(CustomSupportDefinition customSupportDefinition, string sectionSize, out string sectionCode, out string steelStadard)
        {
            try
            {
                PADProperties padProperties;
                sectionCode = string.Empty;
                Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
                bool value = customSupportDefinition.GenericHelper.GetDataByRule("hsMrnSteelStandardName", (BusinessObject)support, out steelStadard);
                PropertyValueCodelist padShapeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPadShape", "PadShape");

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                ReadOnlyCollection<BusinessObject> classItems;
                PartClass auxilaryTable = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_FrmCorresp");
                classItems = auxilaryTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in classItems)
                {
                    if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnFrmCorresp", "SectionSize")).PropValue == sectionSize) && ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnFrmCorresp", "StdName")).PropValue == steelStadard.ToUpper()))
                    {
                        sectionCode = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnFrmCorresp", "Size")).PropValue;
                        break;
                    }
                }
                padProperties.padPart = string.Empty;
                padProperties.padLength = padProperties.padThickness = padProperties.padWidth = 0.0;
                if (padShapeCodeList.PropValue == 2)
                {
                    padProperties.padLength = GetDataByCondition("hsMrnSrv_TriangPadDim", "IJUAhsMrnSrvAnglePad", "L", "IJUAhsMrnSrvAnglePad", "SectionSize", sectionCode);
                    padProperties.padThickness = GetDataByCondition("hsMrnSrv_TriangPadDim", "IJUAhsMrnSrvAnglePad", "t", "IJUAhsMrnSrvAnglePad", "SectionSize", sectionCode);
                    padProperties.padPart = "hsMrn_TrianglPlate_01";
                }
                if (padShapeCodeList.PropValue == 1)
                {
                    padProperties.padLength = GetDataByCondition("hsMrnSrv_RectPadDim", "IJUAhsMrnSrvRectPad", "L", "IJUAhsMrnSrvRectPad", "SectionSize", sectionCode);
                    padProperties.padThickness = GetDataByCondition("hsMrnSrv_RectPadDim", "IJUAhsMrnSrvRectPad", "t", "IJUAhsMrnSrvRectPad", "SectionSize", sectionCode);
                    padProperties.padWidth = GetDataByCondition("hsMrnSrv_RectPadDim", "IJUAhsMrnSrvRectPad", "W", "IJUAhsMrnSrvRectPad", "SectionSize", sectionCode);
                    padProperties.padPart = "hsMrn_RectPlate_01";
                }
                if (padShapeCodeList.PropValue == 3)
                {
                    padProperties.padLength = GetDataByCondition("hsMrnSrv_RndPadDim", "IJUAhsMrnSrvRndPad", "D", "IJUAhsMrnSrvRndPad", "SectionSize", sectionCode);
                    padProperties.padThickness = GetDataByCondition("hsMrnSrv_RndPadDim", "IJUAhsMrnSrvRndPad", "t", "IJUAhsMrnSrvRndPad", "SectionSize", sectionCode);
                    padProperties.padPart = "hsMrn_RoundPlate_01";
                }

                return padProperties;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetPadPartAndDimensions." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method sets the Pad properties.
        /// </summary>
        /// <param name="componentDictionary">Collection of keys and values - Dictionary</param>  
        /// <param name="support">Support</param>  
        ///  <param name="padProperties">Pad properties</param>  
        ///   <param name="PartNames">Part Names- Array</param>        
        /// <code>       
        ///  MarineAssemblyServices.SetPadProperties(componentDictionary, support, padProperties, padPartNames);
        /// </code>
        public static void SetPadProperties(Dictionary<string, SupportComponent> componentDictionary, Ingr.SP3D.Support.Middle.Support support, PADProperties padProperties, string[] PartNames)
        {
            try
            {
                PropertyValueCodelist padShapeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPadShape", "PadShape");
                for (int index = 0; index < PartNames.Length; index++)
                {
                    if (PartNames[index] != null && componentDictionary.ContainsKey(PartNames[index]))//componentDictionary[PartNames[index]] != null)
                    {

                        componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength, "IJUAhsLength1", "Length1");
                        componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padThickness, "IJUAhsThickness1", "Thickness1");


                        if (padShapeCodeList.PropValue == 1) //Rectangle
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padWidth, "IJUAhsWidth1", "Width1");
                        else if (padShapeCodeList.PropValue == 2)   //Triangle
                        {
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength, "IJUAhsWidth1", "Width1");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength, "IJUAhsTLCorner", "TLCornerX");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength, "IJUAhsTLCorner", "TLCornerY");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength / 6, "IJUAhsStandardPort2", "P2xOffset");
                            componentDictionary[PartNames[index]].SetPropertyValue(-padProperties.padLength / 6, "IJUAhsStandardPort2", "P2yOffset");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength / 6, "IJUAhsStandardPort1", "P1xOffset");
                            componentDictionary[PartNames[index]].SetPropertyValue(-padProperties.padLength / 6, "IJUAhsStandardPort1", "P1yOffset");
                        }
                        else if (padShapeCodeList.PropValue == 3)  //Round
                        {
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength, "IJUAhsWidth1", "Width1");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength / 2, "IJUAhsTLCorner", "TLCornerRadius");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength / 2, "IJUAhsTRCorner", "TRCornerRadius");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength / 2, "IJUAhsBLCorner", "BLCornerRadius");
                            componentDictionary[PartNames[index]].SetPropertyValue(padProperties.padLength / 2, "IJUAhsBRCorner", "BRCornerRadius");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in SetPadProperties  of MarineAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method set  BeamClamp Properties on part occurences.
        /// </summary>
        /// <param name="componentDictionary">Collection of keys and values - Dictionary</param>  
        /// <param name="support">Support</param>  
        ///  <param name="beamCLProperties">Beam Clamp properties</param>  
        ///   <param name="PartNames">Part Names- Array</param>  
        /// <code>     
        ///   MarineAssemblyServices.IntializeBeamProperties(componentDictionary, support, beamCLProperties, partNames);
        /// </code>
        public static void IntializeBeamProperties(Dictionary<string, SupportComponent> componentDictionary, Ingr.SP3D.Support.Middle.Support support, BeamCLProperties beamCLProperties, string[] PartNames)
        {
            try
            {
                for (int index = 0; index < PartNames.Length; index++)
                {
                    if (PartNames[index] != null && componentDictionary[PartNames[index]] != null)
                    {
                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.HgrBeamType1.Value, "IJOAHsHgrBeamType", "HgrBeamType");
                        componentDictionary[PartNames[index]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                        componentDictionary[PartNames[index]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.CP5.Value, "IJOAhsSteelCP", "CP3");
                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.CP1.Value, "IJOAhsSteelCP", "CP1");
                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.CP5.Value, "IJOAhsSteelCP", "CP4");
                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.CP1.Value, "IJOAhsSteelCP", "CP2");
                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.CP1.Value, "IJOAhsSteelCP", "CP6");
                        componentDictionary[PartNames[index]].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[PartNames[index]].SetPropertyValue(0.0, "IJOAhsEndCap", "EndCapRotZ");

                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlFlg.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                        componentDictionary[PartNames[index]].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlFlg.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        componentDictionary[PartNames[index]].SetPropertyValue(0.0, "IJOAhsCutback", "CutbackBeginAngle");
                        componentDictionary[PartNames[index]].SetPropertyValue(0.0, "IJOAhsCutback", "CutbackEndAngle");
                    }
                }
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in IntializeBeamProperties of MarineAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method Gets the route port
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="support">Support</param>  
        /// <param name="slopedRoute">Route is sloped - boolean</param>
        /// <param name="slopedSteel">Steel is sloped - boolean</param>
        /// <param name="isVerticalRoute">Route is Vertical or not - boolean</param>
        /// <param name="isVerticalStruct"> structure is Vertical or not - boolean</param>
        /// <returns>String routePort</returns>
        /// <code>
        ///  routePort = MarineAssemblyServices.GetRoutePort(this, support, slopedRoute, slopedSteel, isVerticalRoute, isVerticalSruct);      
        /// </code>
        public static string GetRoutePort(CustomSupportDefinition customSupportDefinition, Ingr.SP3D.Support.Middle.Support support, bool slopedRoute, bool slopedSteel, bool isVerticalRoute, bool isVerticalStruct)
        {
            try
            {
                PropertyValueCodelist orientCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmOrient", "FrmOrient");
                string fromOrient = orientCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(orientCodeList.PropValue).DisplayName;
                string routePort = string.Empty;
                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                        if (slopedRoute == true)
                            if (isVerticalRoute == true)
                                routePort = "BBSV_High";
                            else
                                if (isVerticalStruct == true)
                                    routePort = "BBSV_High";
                                else
                                    routePort = "BBSR_High";
                        else
                            routePort = "BBSR_High";
                    else if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                        if (slopedRoute == true)
                            routePort = "BBR_High";
                        else if (slopedSteel == true)
                            if (isVerticalRoute == true)
                                routePort = "BBSR_High";
                            else
                                if (isVerticalStruct == true)
                                    routePort = "BBSR_High";
                                else
                                    routePort = "BBSV_High";
                        else
                            routePort = "BBSR_High";
                }
                else                                //for Place By Point case
                {
                    if (slopedRoute == true)
                        if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                            if (isVerticalRoute == true)
                                routePort = "BBR_High";
                            else
                                routePort = "BBRV_High";
                        else
                            routePort = "BBR_High";
                    else if (slopedSteel == true)
                        if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                        {
                            CreateBoundingBox(customSupportDefinition, "MarineBBX", false);
                            routePort = "MarineBBX_High";
                        }//for a route which is not vertical
                        else
                            routePort = "BBRV_High";
                    else
                        if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                            if (isVerticalRoute == true)
                            {
                                CreateBoundingBox(customSupportDefinition, "MarineBBX", isVerticalRoute);
                                routePort = "MarineBBX_High";                   //for a route which is vertical
                            }
                            else
                                routePort = "BBR_High";
                        else
                            routePort = "BBR_High";
                }

                return routePort;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetRoutePort." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method gets the bounding box Dimensions.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="routeCount">Number of routes.</param>       
        /// <code>
        /// boundingBox = MarineAssemblyServices.GetBoundingBoxDimensions(this, routeCount);    
        /// </code>
        public static BoundingBox GetBoundingBoxDimensions(CustomSupportDefinition customSupportDefinition, int routeCount)
        {
            try//CreateBBXGivenTwoVectors
            {
                BoundingBox boundingBox;
                double[] routeAngle = new double[routeCount], dimensions = new double[2];
                bool isVerticalRoute, isVerticalSruct, slopedSteel = false;
                double leftStructAngle = 0.0, rightStructAngle = 0.0;
                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = customSupportDefinition.BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = customSupportDefinition.BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                double PI = Math.Atan(1) * 4;
                routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(customSupportDefinition, PI, routeCount, out isVerticalRoute);
                MarineAssemblyServices.GetLeftRightStructAngleAndIsVerticalStructure(customSupportDefinition, PI, out isVerticalSruct, out leftStructAngle, out rightStructAngle);

                try
                {
                    if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(leftStructAngle, 3), 0) == false || Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(rightStructAngle, 3), 0) == false)    //for sloped steel
                        slopedSteel = true;
                }
                catch
                { }
                if (slopedSteel == true)
                    boundingBox = customSupportDefinition.BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedVertical);

                return boundingBox;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetBoundingBoxDimensions." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface1.</param>
        /// <param name="conditionProperty1">The condition property1.</param>        
        /// <param name="referenceValue1">The reference value1.</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///   padProperties.padLength = GetDataByCondition("hsMrnSrv_TriangPadDim", "IJUAhsMrnSrvAnglePad", "L", "IJUAhsMrnSrvAnglePad", "SectionSize", sectionCode);          
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, string referencevalue)
        {
            IEnumerable<BusinessObject> parts = null;
            try
            {
                double propertyValue = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (partClass.PartClassType.Equals("HgrServiceClass"))
                    parts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    parts = partClass.Parts;

                parts = parts.Where(part => (string)((PropertyValueString)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue == referencevalue);
                if (parts.Count() > 0)
                    propertyValue = ((double)((PropertyValueDouble)parts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                return propertyValue;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByCondition of MarineAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (parts is IDisposable)
                {
                    ((IDisposable)parts).Dispose(); // This line will be executed
                }
            }
        }

        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="isOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public static String[] GetIndexedStructPortName(CustomSupportDefinition customSupportDefinition, Boolean[] isOffsetApplied)
        {
            try
            {
                String[] structurePort = new String[2];
                int structureCount = customSupportDefinition.SupportHelper.SupportingObjects.Count;
                int i;

                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Structure";
                }
                else
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Struct_2";

                    if (structureCount > 1)
                    {
                        if (customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        {
                            for (i = 0; i <= 1; i++)
                            {
                                double angle = 0;
                                if (customSupportDefinition.SupportHelper.SupportingObjects.Count != 0)
                                {
                                    if ((customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && isOffsetApplied[i] == false)
                                    {
                                        angle = customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                                    }
                                }
                                //the port is the right structure port
                                if (Math.Abs(angle) < (Math.Round(Math.Atan(1) * 4, 3)) / 2.0)
                                {
                                    if (i == 0)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }
                                //the port is the left structure port
                                else
                                {
                                    if (i == 1)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }
                            }
                        }
                    }
                    else
                        structurePort[1] = "Structure";
                }
                //switch the OffsetApplied flag
                if (structurePort[0] == "Struct_2")
                {
                    Boolean flag = isOffsetApplied[0];
                    isOffsetApplied[0] = isOffsetApplied[1];
                    isOffsetApplied[1] = flag;
                }
                return structurePort;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetIndexedStructPortName." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public static Boolean[] GetIsLugEndOffsetApplied(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                Collection<BusinessObject> structureObjects = customSupportDefinition.SupportHelper.SupportingObjects;
                Boolean[] isOffsetApplied = new Boolean[2];

                //first route object is set as primary route object

                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;

                if (structureObjects != null)
                {
                    if (structureObjects.Count >= 2)
                    {
                        for (int index = 0; index <= 1; index++)
                        {
                            //check if offset is to be applied
                            const double routeStructAngle = Math.PI / 180.0;
                            Boolean varRuleApplied = true;

                            if (customSupportDefinition.SupportHelper.SupportingObjects.Count != 0)
                            {
                                if (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member)
                                {
                                    //if angle is within 1 degree, regard as parallel case
                                    //Also check for Sloped structure                                
                                    MemberPart memberPart = (MemberPart)customSupportDefinition.SupportHelper.SupportingObjects[index];
                                    ICurve memberCurve = memberPart.Axis;

                                    Vector supportedVector = new Vector();
                                    Vector supportingVector = new Vector();

                                    if (customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                                    {
                                        Position startLocation = new Position(customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).StartLocation);
                                        Position endLocation = new Position(customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).EndLocation);
                                        supportedVector = new Vector(endLocation - startLocation);
                                    }
                                    if (memberCurve is ILine)
                                    {
                                        ILine line = (ILine)memberCurve;
                                        supportingVector = line.Direction;
                                    }

                                    double angle = GetAngleBetweenVectors(supportingVector, supportedVector);
                                    double refAngle1 = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - (Math.Round(Math.Atan(1) * 4, 3)) / 2;
                                    double refAngle2 = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                    if (angle < (refAngle1 + 0.001) & angle > (refAngle1 - 0.001))
                                        angle = angle - Math.Abs(refAngle1);
                                    else if (angle < (refAngle2 + 0.001) & angle > (refAngle2 - 0.001))
                                        angle = angle - Math.Abs(refAngle2);
                                    else
                                        angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    if (Math.Abs(angle) < routeStructAngle || Math.Abs(angle - (Math.Round(Math.Atan(1) * 4, 3))) < routeStructAngle)
                                        varRuleApplied = false;
                                }
                            }
                            isOffsetApplied[index] = varRuleApplied;
                        }
                    }
                }

                return isOffsetApplied;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///       ContentHelper contentHelper = new ContentHelper();
        ///       double value;
        ///       value = contentHelper. GetAngleBetweenVectors(vector1, vector2 );
        ///</code>

        public static double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double arcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dotProd), 1) == false)
                {
                    arcCos = (Math.Round(Math.Atan(1) * 4, 3)) / 2 - Math.Atan(dotProd / Math.Sqrt(1 - dotProd * dotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, -1) == true)
                {
                    arcCos = (Math.Round(Math.Atan(1) * 4, 3));
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, 1) == true)
                {
                    arcCos = 0;
                }
                return arcCos;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetAngleBetweenVectors of MarineAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method Gets the UBolt Parts
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>String UBolt Parts</returns>
        /// <code>
        ///   uBoltPart = MarineAssemblyServices.GetUboltPart(this);  
        /// </code>
        public static String[] GetUboltPart(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
                SupportedHelper supportedhlpr = new SupportedHelper(support);
                int routeCount = support.SupportedObjects.Count;
                String[] uBoltPart = new String[routeCount];
                String[] uBoltPart1 = new String[routeCount];
                double[] pipeDiameter = new double[routeCount];
                string unitType = string.Empty;
                string tempPipeAttachment = string.Empty;
                int[] UBolt = new int[routeCount];
                Collection<object> uBoltPart1Collection = null;
                IEnumerable<BusinessObject> hsMarineServiceDimPart = null;

                for (int i = 0; i < routeCount; i++)
                {
                    PipeObjectInfo pipe = (PipeObjectInfo)supportedhlpr.SupportedObjectInfo(i + 1);
                    pipeDiameter[i] = pipe.NominalDiameter.Size;
                    unitType = pipe.NominalDiameter.Units;
                }

                bool uBoltOption = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOpt", "PipeAttachOption")).PropValue;
                if (uBoltOption == true)
                {
                    bool varUBoltPart = customSupportDefinition.GenericHelper.GetDataByRule("hsMrnRLPAttachSel", (BusinessObject)support, out uBoltPart1Collection);
                    if (uBoltPart1Collection != null)
                    {
                        if (uBoltPart1Collection[0] == null)
                        {
                            for (int routeIndex = 0; routeIndex < routeCount; routeIndex++)
                            {
                                uBoltPart1[routeIndex] = (string)uBoltPart1Collection[routeIndex + 1];
                                uBoltPart[routeIndex] = uBoltPart1[routeIndex] + pipeDiameter[routeIndex];
                            }
                        }
                        else
                        {
                            for (int routeIndex = 0; routeIndex < routeCount; routeIndex++)
                            {
                                uBoltPart1[routeIndex] = (string)uBoltPart1Collection[routeIndex];
                                uBoltPart[routeIndex] = uBoltPart1[routeIndex] + pipeDiameter[routeIndex];
                            }
                        }
                    }
                }
                else if (uBoltOption == false)
                {
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        tempPipeAttachment = "U" + routeIndex;
                        UBolt[routeIndex - 1] = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPAttach", tempPipeAttachment)).PropValue;
                        if (UBolt[routeIndex - 1] == -1)
                            UBolt[routeIndex - 1] = 1;

                        if (!(UBolt[routeIndex - 1] == 3))
                        {
                            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                            PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachSel");
                            hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                            hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachType")).PropValue == UBolt[routeIndex - 1]));
                            if (hsMarineServiceDimPart.Count() > 0)
                                uBoltPart[routeIndex - 1] = (string)((PropertyValueString)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachPart")).PropValue;

                            uBoltPart1[routeIndex - 1] = uBoltPart[routeIndex - 1];
                            uBoltPart[routeIndex - 1] = uBoltPart[routeIndex - 1] + pipeDiameter[routeIndex - 1];
                        }
                    }
                    GC.SuppressFinalize(hsMarineServiceDimPart); // to resolve Coverity Issue (disposing hsMarineServiceDimPart)
                }               
                return uBoltPart;
                
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetUboltPart." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method Gets the BraceOffset
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="sectionSize"> SectionSize</param>       
        /// <param name="horizontalOffset"> Horizontel Offset</param>       
        /// <param name="verticalOffset"> Vertical Offset</param>       
        /// <returns>Horizontel OffSet and Vertical OffSet</returns>
        /// <code>
        ///   MarineAssemblyServices.GetBraceOffset(this, sSectionSize, ref dBraceHorOff, ref dBraceVerOff);
        /// </code>
        public static void GetBraceOffset(CustomSupportDefinition customSupportDefinition, string sectionSize, ref double horizontalOffset, ref double verticalOffset)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
                string steelStanded = string.Empty;
                string SectionSize = string.Empty;
                bool value = customSupportDefinition.GenericHelper.GetDataByRule("hsMrnSteelStandardName", (BusinessObject)support, out steelStanded);

                //get the section size code (Like L1, L2 etc) from Correspondance sheet
                steelStanded = steelStanded.ToUpper();

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_FrmCorresp");
                IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "SectionSize")).PropValue == sectionSize) && (string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnFrmCorresp", "StdName")).PropValue == steelStanded);
                if (hsMarineServiceDimPart.Count() > 0)
                    SectionSize = (string)((PropertyValueString)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnFrmCorresp", "Size")).PropValue;

                //get Width Offset and Height Offset
                PartClass hsMarineBrace = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_BraceOffset");
                IEnumerable<BusinessObject> hsMarineBracePart = null;
                hsMarineBracePart=hsMarineBrace.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                hsMarineBracePart = hsMarineBracePart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvBraceOff", "SectionSize")).PropValue == SectionSize));
                if (hsMarineBracePart.Count() > 0)
                {
                    horizontalOffset = (double)((PropertyValueDouble)hsMarineBracePart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvBraceOff", "HorizontalOff")).PropValue;
                    verticalOffset = (double)((PropertyValueDouble)hsMarineBracePart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvBraceOff", "VerticalOff")).PropValue;
                }
                GC.SuppressFinalize(hsMarineBracePart);
                GC.SuppressFinalize(hsMarineServiceDimPart);
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetBraceOffset." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method Gets the SteelStandard
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>String StellStandard</returns>
        /// <code>
        ///  string steelStanded = MarineAssemblyServices.GetSteelStandard(this);  
        /// </code>
        public static string GetSteelStandard(CustomSupportDefinition customSupportDefinition, string steelStanded)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
                string standardSteel = string.Empty;

                if (steelStanded == "JIS-2005")
                    standardSteel = "Japan-2005";
                else if (steelStanded == "Russian")
                    standardSteel = "Russia";
                else if (steelStanded == "GB")
                    standardSteel = "CHINA-2006";
                else if (steelStanded == "ICHA-2000")
                    standardSteel = "Chile-2000";
                else if (steelStanded == "BS5950-1:2000")
                    standardSteel = "BS";
                else if (steelStanded == "AUST-OneSteel-05")
                    standardSteel = "AUST-05";
                else if (steelStanded == "AISC-METRIC")
                    standardSteel = "AISC-Metric";
                else
                    standardSteel = steelStanded;

                return standardSteel;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetSteelStandard." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method Gets the Width and Height of BoundingBox and Left and Right PipeDiameter
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="numberofRoute"> Number Of Routes</param>   
        /// <param name="boundingBoxWidth"> Returns BoundingBox Width</param>  
        /// <param name="boundingBoxHeight"> Returns BoundingBox Height</param>   
        /// <param name="leftPipeDiameter"> Returns Left PipeDiameter</param>   
        /// <param name="rightPipeDiameter"> Returns Right PipeDiameter</param>   
        /// <code>
        ///  MarineAssemblyServices.GetBBXDimAndPipeDia(this, routeCount, ref boundingBoxWidth, ref boundingBoxHeight, ref leftPipeDiameter, ref rightPipeDiameter);  
        /// </code>
        public static void GetBoundingBoxDimensionsAndPipeDiameter(CustomSupportDefinition customSupportDefinition, int numberofRoute, ref double boundingBoxWidth, ref double boundingBoxHeight, ref double leftPipeDiameter, ref double rightPipeDiameter)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
                //load standard bounding box.
                customSupportDefinition.BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;
                string refPlane = string.Empty;
                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    refPlane = "BBSR";
                    boundingBox = customSupportDefinition.BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    refPlane = "BBR";
                    boundingBox = customSupportDefinition.BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }
                boundingBoxWidth = boundingBox.Width;
                boundingBoxHeight = boundingBox.Height;
                leftPipeDiameter = 0.0;
                rightPipeDiameter = 0.0;

                //Get the Route collection
                //get route object collection on the reference plane
                if (customSupportDefinition.SupportHelper.SupportedObjects.Count == 1)
                {
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(1);
                    leftPipeDiameter = pipeInfo.OutsideDiameter - 2 * pipeInfo.InsulationThickness;
                    rightPipeDiameter = leftPipeDiameter;
                }
                else
                {
                    int routeIndex = customSupportDefinition.BoundingBoxHelper.GetBoundaryRouteIndex(refPlane, BoundingBoxEdge.Right);
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(routeIndex);
                    leftPipeDiameter = pipeInfo.OutsideDiameter - 2 * pipeInfo.InsulationThickness;

                    routeIndex = customSupportDefinition.BoundingBoxHelper.GetBoundaryRouteIndex(refPlane, BoundingBoxEdge.Left);
                    pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(routeIndex);
                    rightPipeDiameter = pipeInfo.OutsideDiameter - 2 * pipeInfo.InsulationThickness;
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetBBXDimAndPipeDiameter." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method Create Note for SupportComponent
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="noteName"> Note Name</param>   
        /// <param name="supportComponent"> SupportComponent</param>  
        /// <param name="portName"> Port Name</param>   
        /// <code>
        ///     MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[uBolt[uBoltCount[routeIndex - 1]]], "Route");
        /// </code>
        public static void CreateDimensionNote(CustomSupportDefinition customSupportDefinition, string noteName, SupportComponent supportComponent, string portName)
        {
            try
            {
                Note note = customSupportDefinition.CreateNote(noteName, supportComponent, portName);
                note.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList = note.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);       //3 means fabrication
                note.SetPropertyValue(codeList, "IJGeneralNote", "Purpose");
                note.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in CreateDimensionNote." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;                
            }
        }
        /// <summary>
        /// Creates the bounding box based on the given inputs.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="boundinBoxName"> The name of the bounding boxe</param>   
        /// <param name="isRouteVertical"> Route is vertical or not.</param>  
        /// <code>
        ///     MarineAssemblyServices.CreateBoundingBox(this, "MarineBBX", false);
        /// </code>
        public static void CreateBoundingBox(CustomSupportDefinition customSupportDefinition, string boundingBoxName, bool isRouteVertical)
        {
            try
            {
                //Create Vectors to define the plane of the BBX
                Vector globalZ, bbX, bbZ, bbY;
                //Create Vectors to define the plane of the BBX
                if (isRouteVertical == false)
                {
                    //Vertical Plane Normal Along Route - Z Axis Towards Structure
                    globalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);//Get Global Z
                    bbX = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, globalZ);//Project Route X-Axis into Horizontal Plane
                    bbZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteToStructure);// Project Vector From Route to Structure into the BBX Plane
                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(bbZ, bbX, boundingBoxName, false, false, true, false);
                }
                else
                {
                    // Vertical Plane Normal Along Route - Z Axis Towards Structure
                    bbY = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY);//Get Global Z
                    bbX = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, bbY);//Project Route X-Axis into Horizontal Plane
                    bbZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteToStructure);//Project Vector From Route to Structure into the BBX Plane
                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(bbZ, bbX, boundingBoxName, false, false, true, false);
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in CreateBoundingBox." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }


        /// <summary>
        /// This method Gets the Section Size
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>String Section Size</returns>
        /// <code>
        ///   sStStandard = MarineAssemblyServices.GetSectionSize(this);  
        /// </code>
        public static string GetSectionSize(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
                bool value;
                bool sectionFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSection", "SectionFromRule")).PropValue;
                PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionSize");
                string sectionSize = string.Empty;
                if (sectionFromRule == true)
                    value = customSupportDefinition.GenericHelper.GetDataByRule("hsMrnRLSectionSize", (BusinessObject)support, out sectionSize);
                else
                {
                    sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodelist.PropValue).ShortDisplayName;
                    if (sectionSize.ToUpper() == "NONE" || string.IsNullOrEmpty(sectionSize))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "MarineAssemblyServices.cs", 1);
                        return sectionSize;
                    }
                }
                return sectionSize;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetSectionSize." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method returns the direct angle between Route and structur ports in Radians
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="routePortName">string-route Port Name.</param>
        /// <param name="structPortName">string-structure Port Name.</param>
        /// <param name="axisType">PortAxisType-axis Type.</param>
        /// <returns>double</returns>        
        /// <code>
        ///   byPointAngle1=GetRouteStructConfigAngle("Route", "Structure", PortAxisType.Y);
        ///</code>
        public static double GetRouteStructConfigAngle(CustomSupportDefinition customSupportDefinition, String routePortName, String structPortName, PortAxisType axisType)
        {
            try
            {

                //get the appropriate axis
                Vector[] vecAxis = new Vector[2];

                Matrix4X4 routeMatrix = customSupportDefinition.RefPortHelper.PortLCS(routePortName);
                Position routepoint = routeMatrix.Origin;

                switch (axisType)
                {
                    case PortAxisType.X:
                        {
                            vecAxis[0] = routeMatrix.XAxis;
                            break;
                        }
                    case PortAxisType.Y:
                        {
                            vecAxis[0] = routeMatrix.ZAxis.Cross(routeMatrix.XAxis);
                            break;
                        }
                    case PortAxisType.Z:
                        {
                            vecAxis[0] = routeMatrix.ZAxis;
                            break;
                        }
                }
                Matrix4X4 structMatrix = customSupportDefinition.RefPortHelper.PortLCS(structPortName);
                Position structPoint = structMatrix.Origin;
                vecAxis[1] = structPoint.Subtract(routepoint);

                return GetAngleBetweenVectors(vecAxis[0], vecAxis[1]);

            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetRouteStructConfigAngle." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}

