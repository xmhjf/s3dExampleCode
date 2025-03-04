//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   GenericAssemblyServices.cs
//   Author       : Rajeswari
//   Creation Date: 03-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 03-Sep-2013  Rajeswari,Hema CR-CP-224494  Convert HS_Generic_Assy to C# .Net
// 22-Jan-2015   PVK           TR-CP-264951  Resolve coverity issues found in November 2014 report  
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class GenericAssemblyServices
    {
        /// <summary>
        /// defines the config index
        /// </summary>
        public struct ConfigIndex
        {
            /// <summary>
            /// Geometry Plane A
            /// </summary>
            public Plane A;
            /// <summary>
            /// Geometry Plane B
            /// </summary>
            public Plane B;
            /// <summary>
            /// Geometry Axis C
            /// </summary>
            public Axis C;
            /// <summary>
            /// Geometry Axis D
            /// </summary>
            public Axis D;
            /// <summary>
            /// Structure to hold Plane and axis values of a config Index. 
            /// </summary>
            /// <param name="a"> Plane a- Middle Plane</param>
            /// <param name="b">Plane b- Middle Plane</param>
            /// <param name="c">Axis c- Middle Axis</param>
            /// <param name="d">Axis d- Middle Axis</param>
            ///<example>
            /// <code>
            /// GenericAssemblyServices.ConfigIndex braceConfigIdx1 = new GenericAssemblyServices.ConfigIndex();
            /// braceConfigIdx1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
            /// </code>
            /// </example>    
            public ConfigIndex(Plane a, Plane b, Axis c, Axis d)
            {
                this.A = a;
                this.B = b;
                this.C = c;
                this.D = d;
            }
        }      
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns>String value</returns>
        /// <example>
        /// <code>
        ///     string sAngPadPart = GenericAssemblyServices.GetDataByConditionString("GenServ_AnglePadDim", "IJUAHgrGenServAnglePad", "PadPart", "IJUAHgrGenServAnglePad", "SectionSize", sectionSize);
        /// </code>
        /// </example>        
        public static string GetDataByConditionString(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, string referencevalue)
        {
            IEnumerable<BusinessObject> partsCollection = null;
            try
            {
                string result = string.Empty;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (partClass.PartClassType.Equals("HgrServiceClass"))
                    partsCollection = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    partsCollection = partClass.Parts;

                partsCollection = partsCollection.Where(part => (string)((PropertyValueString)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue == referencevalue);
                if (partsCollection.Count() > 0)
                    result = ((string)((PropertyValueString)partsCollection.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                return result;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetDataByConditionString of GenericAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (partsCollection is IDisposable)
                {
                    ((IDisposable)partsCollection).Dispose(); // This line will be executed
                }
            }
        }

        /// <summary>
        /// This method Creates a note to a port on a SupportComponent
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="noteName">Note Name</param>   
        /// <param name="supportComponent">an abstract class for Support Components</param>  
        /// <param name="portName"> Port Name</param>   
        /// <code>
        ///     GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[uboltPartKeys[indexText - 1]], "Route");
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
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="minimumReferencevalue">The minimum reference value.</param>
        /// <param name="maximumReferencevalue">The maximum reference value</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>     
        ///     double W = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "W", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, double minimumReferencevalue, double maximumReferencevalue)
        {
            IEnumerable<BusinessObject> partsCollection = null;
            try
            {
                double result = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (partClass.PartClassType.Equals("HgrServiceClass"))
                    partsCollection = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    partsCollection = partClass.Parts;

                partsCollection = partsCollection.Where(part => (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue > minimumReferencevalue && (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue < maximumReferencevalue);
                if (partsCollection.Count() > 0)
                    result = ((double)((PropertyValueDouble)partsCollection.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                return result;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetDataByCondition GenericAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (partsCollection is IDisposable)
                {
                    ((IDisposable)partsCollection).Dispose(); // This line will be executed
                }
            }
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>        
        ///     double L = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeI_Dim", "IJUAHgrGenSrvIBrkDim", "L", "IJUAHgrGenSrvIBrkDim", "LBarSize", sectionSize);
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, string referenceValue)
        {
            IEnumerable<BusinessObject> partsCollection = null;
            try
            {
                double result;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (partClass.PartClassType.Equals("HgrServiceClass"))
                    partsCollection = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    partsCollection = partClass.Parts;

                partsCollection = partsCollection.Where(part => (string)((PropertyValueString)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue == referenceValue);
                if (partsCollection.Count() > 0)
                    result = ((double)((PropertyValueDouble)partsCollection.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                else
                    result = 0;
                return result;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetDataByCondition of GenericAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (partsCollection is IDisposable)
                {
                    ((IDisposable)partsCollection).Dispose(); // This line will be executed
                }
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
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dotProd) , 1)==false)
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
                CmnException e1 = new CmnException("Error in GetAngleBetweenVectors." + ". Error:" + e.Message, e);
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
        /// <summary>
        /// This method returns the maximum distance between any two ports.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="distanceType">distanceType -PortDistanceType.</param>       
        /// <returns>double</returns>        
        /// <code>
        /// double extremePipesDist = GenericAssemblyServices.GetPipesMaximumDistance(this, PortDistanceType.Vertical);
        ///</code>
        public static double GetPipesMaximumDistance(CustomSupportDefinition customSupportDefinition,PortDistanceType distanceType)
        {
            try
            {
                double maxDistance = 0;

                //Get Route object
                int routeCount = customSupportDefinition.SupportHelper.SupportedObjects.Count;

                //Get the distance between Extreme Left and Extreme Right Pipese

                string[] pipePortName = new string[routeCount];

                for (int i = 1; i <= routeCount; i++)
                {
                    if (i == 1)
                        pipePortName[i - 1] = "Route";
                    else
                        pipePortName[i - 1] = "Route_" + i;
                }

                double[,] pipesPortsDistance = new double[20, 20];

                for (int i = 1; i <= routeCount; i++)
                {
                    for (int j = i + 1; j <= routeCount; j++)
                    {
                        pipesPortsDistance[i - 1, j - 2] = customSupportDefinition.RefPortHelper.DistanceBetweenPorts(pipePortName[i - 1], pipePortName[j - 1], distanceType);
                    }
                }

                for (int i = 1; i <= routeCount; i++)
                {
                    for (int j = i + 1; j <= routeCount; j++)
                    {
                        if (maxDistance < pipesPortsDistance[i - 1, j - 2])
                            maxDistance = pipesPortsDistance[i - 1, j - 2];
                    }
                }
                return maxDistance;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetPipesMaximumDistance." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }      
        /// <summary>
        /// This method Compares the two double variables with some tolerance
        /// </summary>
        /// <param name="leftVariable">the left double cariable</param>
        /// <param name="rightVariale">the right double cariable</param>
        /// <param name="tolerance">the tolerance value</param>
        /// <code>
        ///     GenericAssemblyServices.CmpDblLessThan(overhang, hgrOverHang1);
        /// </code>
        public static Boolean CmpDblLessThan(double leftVariable, double rightVariale, double tolerance = 0.0000001)
        {
            bool isdblLessThan;
            if (leftVariable < (rightVariale + tolerance))
                isdblLessThan = true;
            else
                isdblLessThan = false;

            return isdblLessThan;
        }
        /// <summary>
        /// This method Compares the two double variables with some tolerance
        /// </summary>
        /// <param name="leftVariable">the left double cariable</param>
        /// <param name="rightVariale">the right double cariable</param>
        /// <param name="tolerance">the tolerance value</param>
        /// <code>
        ///     GenericAssemblyServices.CmpDblEqual(overhang, hgrOverHang1);
        /// </code>
        public static Boolean CmpDblEqual(double leftVariable, double rightVariale, double tolerance = 0.0000001)
        {
            bool isdblEqual;
            if ((leftVariable >= (rightVariale - tolerance)) && (leftVariable <= (rightVariale + tolerance)))
                isdblEqual = true;
            else
                isdblEqual = false;

            return isdblEqual;
        }
        /// <summary>
        /// This method gets the maximum vertical distance between route and structure
        /// </summary>
        ///  <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="distanceType">distanceType -PortDistanceType.</param>
        /// <code>
        ///  double largeDistRouteStruct = GenericAssemblyServices.GetMaximumRouteStructDistance(this,PortDistanceType.Vertical);
        /// </code>
        public static double GetMaximumRouteStructDistance(CustomSupportDefinition customSupportDefinition, PortDistanceType distanceType)
        {
            try
            {
                // Get Route object
                int iNumRoutes = customSupportDefinition.SupportHelper.SupportedObjects.Count;
                // Get the distance between Extreme Left and Extreme Right Pipes
                int i, j;
                string pipePortName = string.Empty;
                double[] structRouteVertDist = new double[0];
                for (i = 1; i <= iNumRoutes; i++)
                {
                    Array.Resize(ref structRouteVertDist, i);
                    if (i == 1)
                        pipePortName = "Route";
                    else
                        pipePortName = "Route_" + i;

                    structRouteVertDist[i - 1] = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Structure", pipePortName, distanceType);
                }
                double temp;
                for (i = 1; i <= iNumRoutes; i++)
                {
                    for (j = i + 1; j <= iNumRoutes; j++)
                    {
                        if (structRouteVertDist[j - 1] > structRouteVertDist[i - 1])
                        {
                            temp = structRouteVertDist[i - 1];
                            structRouteVertDist[i - 1] = structRouteVertDist[j - 1];
                            structRouteVertDist[j - 1] = temp;
                        }
                    }
                }

                return structRouteVertDist[0];
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetMaximumRouteStructDistance." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method is to trim the insulation from a pipe run to get the pipe dia only
        /// </summary>
        /// <param name="route"> the pipeobject</param>
        /// <param name="totalOD"> total diameter of pipe</param>
        /// <code>
        ///     GenericAssemblyServices.GetPipeODwithoutInsulatation(pipeDiameter[idxRoute - 1], pipeinfo);
        /// </code>
        public static double GetPipeODwithoutInsulatation(double totalOD, PipeObjectInfo route)
        {
            try
            {
                double insulationThickness = route.InsulationThickness;
                return totalOD - 2 * insulationThickness;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetPipeODwithoutInsulatation GenericAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
        }  
    }
}


