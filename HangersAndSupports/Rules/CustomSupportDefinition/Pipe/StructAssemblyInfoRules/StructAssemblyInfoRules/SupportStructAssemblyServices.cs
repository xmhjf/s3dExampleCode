//-----------------------------------------------------------------------------
// Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
// File
//     SupportStructAssemblyServices.cs
// Author:     
//      Vijaya     
//
// Abstract:
//    Convert SupportStructAssemblyInfoRules   Commom Methods
//   Change History:
//   dd.mmm.yyyy     who       change description
//   -----------     ---       ------------------
//  31-Oct-2014     PVK       TR-CP-260301	Resolve coverity issues found in August 2014 report
//  22/01/2015      PVK       TR-CP-264951  Resolve coverity issues found in November 2014 report 
//  09.Feb.2015     Siva      TR-CP-261379    When a Conduit support type is is changed, it is not re-placed correctly
//-----------------------------------------------------------------------------
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Route.Middle;
using Ingr.SP3D.Support.Middle.Hidden;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{

    public static class SupportStructAssemblyServices
    {
        /// <summary>
        /// This method Gets the This method get/set all the occurrence attributes defined on the pipe structural support assemblies
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>            
        /// <returns>object array</returns>
        /// <code>
        ///   object[] attributeCollection = SupportStructAssemblyServices.GetPipeStructuralASMAttributes(this);
        /// </code>
        public static object[] GetPipeStructuralASMAttributes(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                const int SUPPORTRECOMPUTE = 4;
                double d=0, gap=0;
                bool leftPad, rightPad;
                object[] attributes = new object[4];
                if ((customSupportDefinition.SupportHelper.SupportStatus & SUPPORTRECOMPUTE) > 0)
                {
                    //=====================================
                    //(INITIAL PLACEMENT)recalculate all the support attributes
                    //based on predefined rule
                    //=====================================
                    //Set D and G

                    d = (double)((PropertyValueDouble)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrOccStructuralAssembly", "D")).PropValue;
                    gap = (double)((PropertyValueDouble)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrOccStructuralAssembly", "Gap")).PropValue;
                    //retrieve pad configuration
                    leftPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "LeftPad")).PropValue;
                    rightPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "RightPad")).PropValue;
                }
                else
                {
                    //=============================
                    //(RECOMPUTE PROCESS)retrieve the occurrence value
                    //=============================
                    //retrieve D and G                                                                                                  

                    Collection<object> valueCollection = null, gapValues = null;
                    Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
                    bool value = customSupportDefinition.GenericHelper.GetDataByRule("HgrSupAngleByLF", (BusinessObject)support, out valueCollection);
                    if (valueCollection != null)
                        d = (double)valueCollection[3];
                    support.SetPropertyValue(d, "IJUAHgrOccStructuralAssembly", "D");
                    value = customSupportDefinition.GenericHelper.GetDataByRule("HgrSup_G", (BusinessObject)support, out gapValues);
                    if (gapValues != null)
                        gap = (double)gapValues[0];
                    support.SetPropertyValue(gap, "IJUAHgrOccStructuralAssembly", "Gap");
                    leftPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "LeftPad")).PropValue;
                    rightPad = (bool)((PropertyValueBoolean)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrPadOcc", "RightPad")).PropValue;
                }
                attributes[0] = d;
                attributes[1] = gap;
                attributes[2] = leftPad;
                attributes[3] = rightPad;
                return attributes;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetPipeStructuralASMAttributes." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method returns the pad part number by given cross section.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition.</param> 
        /// <param name="hgrBeam">Hanger beam PartInfo.</param> 
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
                if (string.IsNullOrEmpty(hgrBeam.PartSelectionRule))
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
                if (sectionType.Equals("L"))
                    padPartClass = "TriangularPad";
                else if (sectionType.Equals("CS"))
                    padPartClass = "CircularPad";
                else if (sectionType.Equals("W") || sectionName.Equals("CS"))
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
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public static int GetLSupportConfiguration(CustomSupportDefinition customSupportDefinition, ref bool[] isOffSetApplied, ref double[] LugOffset)
        {
            try
            {
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
                double[] boxOffset = GetBoundaryObjectDimension(customSupportDefinition, boundingBox);
                string topStructPort = "Structure", bottomStructPort = "Struct_2", routeRefPort;
                double angle;
                int routeStructConfig = -1;

                //if first input is a slab, offset is needed               
                bool[] isOffsetApplied = new bool[2];
                isOffSetApplied[0] = true;
                isOffSetApplied[1] = true;
                isOffsetApplied = GetIsLugEndOffsetApplied(customSupportDefinition);
                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routeRefPort = "BBSR_Low";
                else
                    routeRefPort = "BBR_Low";
                if (!isOffsetApplied[0])
                {
                    //check whether it is on the left  or right
                    angle = GetRouteStructConfigAngle(customSupportDefinition, routeRefPort, topStructPort, PortAxisType.Y);

                    if (Math.Abs(angle) < (4 * Math.Atan(1)) / 2)
                        //the member is on the right side of the route
                        routeStructConfig = 1;
                    else
                        routeStructConfig = 2;
                }
                //check the second input
                if (customSupportDefinition.SupportHelper.SupportingObjects.Count == 2)
                {
                    isOffsetApplied[1] = false;

                    if (routeStructConfig == -1)
                    {
                        //check whether it is on the left  or right
                        angle = GetRouteStructConfigAngle(customSupportDefinition, routeRefPort, bottomStructPort, PortAxisType.Y);

                        if (Math.Abs(angle) < (4 * Math.Atan(1)) / 2)
                            //the top member is on the left side of the route
                            routeStructConfig = 2;
                        else
                            routeStructConfig = 1;

                    }
                }
                //define appropriate offset value
                double d = (double)((PropertyValueDouble)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAHgrOccStructuralAssembly", "D")).PropValue;
                Collection<object> angleConfig = null;
                BusinessObject support = customSupportDefinition.SupportHelper.Support;
                bool varUBoltPart = customSupportDefinition.GenericHelper.GetDataByRule("HgrSupAngleByLF", support, out angleConfig);
                if (angleConfig != null)
                {
                    LugOffset[0] = d / 2 - boxOffset[0] / 2;
                    LugOffset[1] = (double)angleConfig[3] - boxOffset[3] / 2;
                }

                if (routeStructConfig == 1)
                {
                    LugOffset[0] = boxOffset[3];
                    LugOffset[1] = boxOffset[2];
                }

                return routeStructConfig;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetLSupportConfiguration of MarineAssemblyServices class" + ". Error:" + e.Message, e);
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
                Vector[] vectorAxis = new Vector[2];

                Matrix4X4 routeMatrix = customSupportDefinition.RefPortHelper.PortLCS(routePortName);
                Position routepoint = routeMatrix.Origin;
                switch (axisType)
                {
                    case PortAxisType.X:
                        {
                            vectorAxis[0] = routeMatrix.XAxis;
                            break;
                        }
                    case PortAxisType.Y:
                        {
                            vectorAxis[0] = routeMatrix.ZAxis.Cross(routeMatrix.XAxis);
                            break;
                        }
                    case PortAxisType.Z:
                        {
                            vectorAxis[0] = routeMatrix.ZAxis;
                            break;
                        }
                }
                Matrix4X4 structMatrix = customSupportDefinition.RefPortHelper.PortLCS(structPortName);
                Position structPoint = structMatrix.Origin;
                vectorAxis[1] = structPoint.Subtract(routepoint);

                return GetAngleBetweenVectors(vectorAxis[0], vectorAxis[1]);

            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetRouteStructConfigAngle." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
        public static double[] GetBoundaryObjectDimension(CustomSupportDefinition customSupportDefinition, BoundingBox boundingBox)
        {
            try
            {
                ReadOnlyCollection<BusinessObject> routeObjectCollection;
                double[] boundaryObjectSize = new double[5];
                BoundingBoxEdge[] type = new BoundingBoxEdge[4];
                type[0] = BoundingBoxEdge.Bottom;
                type[1] = BoundingBoxEdge.Top;
                type[2] = BoundingBoxEdge.Left;
                type[3] = BoundingBoxEdge.Right;
                double pipeRadius = 0.0;
                for (int i = 0; i < 4; i++)
                {
                    routeObjectCollection = boundingBox.SupportedObjectsAtEdge(type[i]);
                    IPipePathFeature pipeInfo = routeObjectCollection[0] as IPipePathFeature;
                    NominalDiameter pipeDiameter = new NominalDiameter();
                    if (pipeInfo != null)
                    {
                        pipeDiameter.Size = pipeInfo.NPD.Size;
                        pipeRadius = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    }
                    IConduitPathFeature conduitInfo = routeObjectCollection[0] as IConduitPathFeature;
                    if (conduitInfo != null)
                    {
                        pipeDiameter.Size = conduitInfo.NCD.Size;
                        pipeRadius = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    }
                    IRouteFeatureWithCrossSection ductInfo = routeObjectCollection[0] as IRouteFeatureWithCrossSection;

                    if (ductInfo != null)
                        pipeRadius = ductInfo.Width;
                    boundaryObjectSize[i] = pipeRadius;
                }
                return boundaryObjectSize;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetBoundaryObjectDimension." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
        ///     SupportStructAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[uBolt[uBoltCount[routeIndex - 1]]], "Route");
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
                                if (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member & isOffsetApplied[i] == false)
                                    angle = customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
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
    }
}
