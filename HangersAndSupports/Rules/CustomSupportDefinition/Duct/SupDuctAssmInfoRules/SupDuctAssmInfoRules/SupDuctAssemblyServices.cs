//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      SupDuctAssemblyServices.cs
// Author:     
//      Vijay 
//
// Abstract:
//     SupDuctAssemblyServices Commom Methods
//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  22-Jan-2015     PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//-----------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Support.Middle.Hidden;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    public static class SupDuctAssemblyServices
    {

        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="customSupportDefinition">The custom support definition.</param>
        /// <param name="isOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = SupDuctAssemblyServices.GetIndexedStructPortName(bIsOffsetApplied);
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
                                if ((customSupportDefinition.SupportHelper.SupportingObjects.Count != 0))
                                {
                                    if (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member & isOffsetApplied[i] == false)
                                        angle = customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                                }
                                //the port is the right structure port
                                if (Math.Abs(angle) < Math.PI / 2.0)
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
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///      double value = SupDuctAssemblyServices.GetAngleBetweenVectors(vector1, vector2);
        ///</code>

        public static double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double arcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dotProd), 1) == false)
                {
                    arcCos = Math.PI / 2 - Math.Atan(dotProd / Math.Sqrt(1 - dotProd * dotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, -1) == true)
                {
                    arcCos = Math.PI;
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, 1) == true)
                {
                    arcCos = 0;
                }
                return arcCos;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetAngleBetweenVectors."+ ". Error:" + e.Message, e);
                throw e1;
            }

        }
        /// <summary>
        /// This method get/set all the occurrence attributes defined on the duct structural support assemblies.
        /// </summary>
        /// <param name="customSupportDefinition">The custom support definition.</param>
        /// <returns>Object of Collection</returns>
        /// <code>
        ///     object[] attributeCollection = SupDuctAssemblyServices.GetDuctStructuralASMAttributes(this);
        /// </code>
        public static object[] GetDuctStructuralASMAttributes(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                double D, G;
                bool leftPad, rightPad;
                if ((customSupportDefinition.SupportHelper.SupportStatus & 4) > 0)
                {
                    //(RECOMPUTE PROCESS)retrieve the occurrence value
                    //retrieve D and G
                    D = (double)((PropertyValueDouble)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrOccStructuralAssembly", "D")).PropValue;
                    G = (double)((PropertyValueDouble)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrOccStructuralAssembly", "Gap")).PropValue;
                    leftPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "LeftPad")).PropValue;
                    rightPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "RightPad")).PropValue;
                }
                else
                {
                    Collection<Object> supAttrib;
                    DuctObjectInfo ductInfo = (DuctObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(1);

                    if (ductInfo.CrossSectionShape == CrossSectionShape.Rectangular)
                    {
                        customSupportDefinition.GenericHelper.GetDataByRule("HgrSupDuctAngleOffset", customSupportDefinition.SupportHelper.Support, out supAttrib);
                        D = (double)supAttrib[0];
                        G = (double)supAttrib[1];
                    }
                    else
                    {
                        customSupportDefinition.GenericHelper.GetDataByRule("HgrSupDuctClamp_D", customSupportDefinition.SupportHelper.Support, out D);
                        customSupportDefinition.GenericHelper.GetDataByRule("HgrSup_G", customSupportDefinition.SupportHelper.Support, out G);
                    }
                    //(INITIAL PLACEMENT)recalculate all the support attributes
                    //based on predefined rule
                    //Set D and G
                    customSupportDefinition.SupportHelper.Support.SetPropertyValue(D, "IJUAHgrOccStructuralAssembly", "D");
                    customSupportDefinition.SupportHelper.Support.SetPropertyValue(G, "IJUAHgrOccStructuralAssembly", "Gap");
                    leftPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "LeftPad")).PropValue;
                    rightPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "RightPad")).PropValue;
                }
                object[] AttributeColl = new object[4];
                AttributeColl[0] = D;
                AttributeColl[1] = G;
                AttributeColl[2] = leftPad;
                AttributeColl[3] = rightPad;
                return AttributeColl;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetDuctStructuralASMAttributes." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = SupDuctAssemblyServices.GetIsLugEndOffsetApplied(this);  
        /// </code>
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

                            if ((customSupportDefinition.SupportHelper.SupportingObjects.Count != 0))
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
        /// This method get the Route Sruct Configuration
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="isOffsetApplied">boolen values wheather offset is applied or not</param>
        /// <returns>Route Struct Configuration</returns>
        /// <code>
        ///     int routeStructConfiguration = SupDuctAssemblyServices.GetDuctLSupportConfiguration(this, isOffsetApplied);
        /// </code>
        public static int GetDuctLSupportConfiguration(CustomSupportDefinition customSupportDefinition, Boolean[] isOffsetApplied)
        {
            try
            {
                string topStructPort = string.Empty;
                string bottomStructPort = string.Empty;
                string routeRefPort = string.Empty;
                int routeStructConfiguration;
                double angle;

                //default value for one structure input
                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;

                topStructPort = "Structure";
                bottomStructPort = "Struct_2";

                routeStructConfiguration = -1;

                //if first input is a slab, offset is needed
                bool[] getOffsetApplied = GetIsLugEndOffsetApplied(customSupportDefinition);
                isOffsetApplied[0] = getOffsetApplied[0];
                isOffsetApplied[1] = getOffsetApplied[1];

                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routeRefPort = "BBSR_Low";
                else
                    routeRefPort = "BBR_Low";

                if (!(isOffsetApplied[0]))
                {
                    angle = customSupportDefinition.RefPortHelper.PortConfigurationAngle(routeRefPort, topStructPort, PortAxisType.Y);

                    if (Math.Abs(angle) < (4 * Math.Atan(1.0)) / 2.0)
                        //the member is on the right side of the route
                        routeStructConfiguration = 1;
                    else
                        routeStructConfiguration = 2;
                }

                //check the second input
                if (customSupportDefinition.SupportHelper.SupportingObjects.Count == 2)
                {
                    isOffsetApplied[1] = false;

                    if (routeStructConfiguration == -1)
                    {
                        //check whether it is on the left  or right
                        angle = customSupportDefinition.RefPortHelper.PortConfigurationAngle(routeRefPort, bottomStructPort, PortAxisType.Y);

                        if (Math.Abs(angle) < (4 * Math.Atan(1)) / 2.0)
                            //the top member is on the left side of the route
                            routeStructConfiguration = 2;
                        else
                            routeStructConfiguration = 1;
                    }
                }
                return routeStructConfiguration;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetDuctLSupportConfiguration." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition.</param>      
        /// <returns>String</returns>        
        /// <code>
        ///       parts.Add(new PartInfo(RIGHTPAD, SupportStructAssemblyServices.GetPartByPadSelection(this, parts[0]))); 
        ///</code>

        public static string GetPadPartNameByCrossSectionType(CustomSupportDefinition customSupportDefinition, PartInfo hgrBeam)
        {
            try
            {
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                string partSelectionRule = string.Empty;
                SupportComponentUtils supportComponentUtils = new SupportComponentUtils();

                PartClass partClass;
                if (hgrBeam.PartSelectionRule == "")
                {
                    partClass = (PartClass)catalogBaseHelper.GetPartClass(hgrBeam.PartClassOrPartNo);
                    hgrBeam.PartSelectionRule = partClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                }
                Part hangerBeam = supportComponentUtils.GetPartFromPartClass(hgrBeam.PartClassOrPartNo, hgrBeam.PartSelectionRule, customSupportDefinition.SupportHelper.Support);
                CrossSection crossSection = (CrossSection)hangerBeam.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                string sectionName = crossSection.Name, padPartNumber = string.Empty, sectionType = crossSection.CrossSectionClass.Name, padPart = string.Empty, padPartClass = string.Empty;

                partClass = (PartClass)catalogBaseHelper.GetPartClass("HgrServ_PadsSelection");
                ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject padItem in classItems)
                {
                    if (((string)((PropertyValueString)padItem.GetPropertyValue("IJUAHngServ_PadSelection", "SectionName")).PropValue).ToUpper() == sectionName.ToUpper())
                    {
                        padPartNumber = (string)((PropertyValueString)padItem.GetPropertyValue("IJUAHngServ_PadSelection", "Part")).PropValue;
                        break;
                    }
                }
                if (sectionType == "L")
                    padPartClass = "TriangularPad";
                else if (sectionType == "CS")
                    padPartClass = "CircularPad";
                else if (sectionType == "W" || sectionName == "CS")
                {
                    double width;
                    if (padPartNumber == string.Empty)
                    {
                        width = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                        PartClass circularPadClass = (PartClass)catalogBaseHelper.GetPartClass("HgrServ_PadsSelection");
                        ReadOnlyCollection<BusinessObject> circularPads = circularPadClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        foreach (BusinessObject padItem in circularPads)
                        {
                            if (((double)((PropertyValueDouble)padItem.GetPropertyValue("IJUAHngServ_PadSelection", "SectionWidthFrom")).PropValue) >= width)
                            {
                                padPartNumber = (string)((PropertyValueString)padItem.GetPropertyValue("IJUAHngServ_PadSelection", "Part")).PropValue;
                                break;
                            }
                        }
                    }
                    padPartClass = "RectangularPad";
                }
                else
                    padPartClass = "TriangularPad";

                partClass = (PartClass)catalogBaseHelper.GetPartClass(padPartClass);
                foreach (Part part in partClass.Parts)
                {
                    if (part.PartNumber == padPartNumber)
                    {
                        padPart = part.PartNumber;
                        break;
                    }
                }
                return padPart;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetPartByPadSelection." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
