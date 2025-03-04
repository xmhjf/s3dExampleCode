//-----------------------------------------------------------------------------
// Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//      TrayShipAseemblyServices.cs
// Author:     
//      Vijay,Manikanth  
//
// Abstract:
//     TrayShip Commom Methods
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
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    public static class TrayShipAseemblyServices
    {
        /// <summary>
        /// This method Gets the property value from the BusinessObject
        /// </summary>
        /// <param name="businessObject"> Can be support or supportComponent or Part or supportDefinition</param>
        /// <param name="collectionOfInterfaces">collection of interfaces</param>
        /// <param name="defaultValue"> returns default value given</param>
        /// <param name="propertyName">given propertyname </param>
        /// <returns>propValue</returns>
        /// <code>
        ///     GetProertyFromObject(businessObject, collectionOfInterfaces, propertyName, defaultValue)
        /// </code>
        public static PropertyValue GetProertyFromObject(BusinessObject businessObject, string[] collectionOfInterfaces, string propertyName, object defaultValue)
        {
            PropertyValue propValue = null;
            foreach (string interfacename in collectionOfInterfaces)
            {
                if (businessObject.SupportsInterface(interfacename))
                {
                    businessObject.GetPropertyValue(interfacename, propertyName);
                }
            }
            return propValue;
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
        /// This method Sets the property value on the BusinessObject
        /// </summary>
        /// <param name="businessObject"> Can be support or supportComponent or Part or supportDefinition</param>
        /// <param name="collectionOfInterfaces">collection of interfaces</param>
        /// <param name="propertyName">given propertyname </param>
        /// <param name="value"> int-value to be set</param>
        /// <code>
        ///     SetPropertyFromObject(businessObject, collectionOfInterfaces, propertyName, value)
        /// </code>
        public static void SetPropertyFromObject(BusinessObject businessObject, string[] collectionOfInterfaces, string propertyName, int value)
        {
            foreach (string interfacename in collectionOfInterfaces)
            {
                if (businessObject.SupportsInterface(interfacename))
                {
                    businessObject.SetPropertyValue(value, interfacename, propertyName);
                }
            }
        }
        /// <summary>
        /// This method Sets the property value on the BusinessObject
        /// </summary>
        /// <param name="businessObject"> Can be support or supportComponent or Part or supportDefinition</param>
        /// <param name="collectionOfInterfaces">collection of interfaces</param>
        /// <param name="propertyName">given propertyname </param>
        /// <param name="value">-double-value to be set</param>
        /// <code>
        ///     SetPropertyFromObject(businessObject, collectionOfInterfaces, propertyName, value)
        /// </code>
        public static void SetPropertyFromObject(BusinessObject businessObject, string[] collectionOfInterfaces, string propertyName, double value)
        {
            foreach (string interfacename in collectionOfInterfaces)
            {
                if (businessObject.SupportsInterface(interfacename))
                {
                    businessObject.SetPropertyValue(value, interfacename, propertyName);
                }
            }
        }


        /// <summary>
        /// This method Sets the property value on the BusinessObject
        /// </summary>
        /// <param name="businessObject"> Can be support or supportComponent or Part or supportDefinition</param>
        /// <param name="collectionOfInterfaces">collection of interfaces</param>
        /// <param name="propertyName">given propertyname </param>
        /// <param name="value">-string-value to be set</param>
        /// <code>
        ///     SetPropertyFromObject(businessObject, collectionOfInterfaces, propertyName, value)
        /// </code>
        public static void SetPropertyFromObject(BusinessObject businessObject, string[] collectionOfInterfaces, string propertyName, string value)
        {
            foreach (string interfacename in collectionOfInterfaces)
            {
                if (businessObject.SupportsInterface(interfacename))
                {
                    businessObject.SetPropertyValue(value, interfacename, propertyName);
                }
            }
        }
        /// <summary>
        /// This method Sets the property value on the BusinessObject
        /// </summary>
        /// <param name="businessObject"> Can be support or supportComponent or Part or supportDefinition</param>
        /// <param name="collectionOfInterfaces">collection of interfaces</param>
        /// <param name="propertyName">given propertyname </param>
        /// <param name="value">-bool-value to be set</param>
        /// <code>
        ///     SetPropertyFromObject(businessObject, collectionOfInterfaces, propertyName, value)
        /// </code>
        public static void SetPropertyFromObject(BusinessObject businessObject, string[] collectionOfInterfaces, string propertyName, bool value)
        {
            foreach (string interfacename in collectionOfInterfaces)
            {
                if (businessObject.SupportsInterface(interfacename))
                {
                    businessObject.SetPropertyValue(value, interfacename, propertyName);
                }
            }
        }
        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="isOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public static String[] GetIndexedStructPortName(CustomSupportDefinition customSupportDefinition, Boolean[] isOffsetApplied)
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
        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
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
                                    double refAngle1 = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - Math.PI / 2;
                                    double refAngle2 = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                    if (angle < (refAngle1 + 0.001) & angle > (refAngle1 - 0.001))
                                        angle = angle - Math.Abs(refAngle1);
                                    else if (angle < (refAngle2 + 0.001) & angle > (refAngle2 - 0.001))
                                        angle = angle - Math.Abs(refAngle2);
                                    else
                                        angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    if (Math.Abs(angle) < routeStructAngle || Math.Abs(angle - Math.PI) < routeStructAngle)
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
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied Method." +". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method returns the direct angle between Route and structur ports in Radians
        /// </summary>
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
            catch (Exception ex)
            {
                throw ex;
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
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public static double GetMaxRouteStructDistance(CustomSupportDefinition customSupportDefinition, int iNumRoutes, ref int[] iTrayIndex)
        {
            //Get the distance between Extreme Left and Extreme Right Pipese
            string sTrayPortName = string.Empty;
            double[] dStructRouteHorDist = new double[iNumRoutes + 1];

            for (int index = 1; index <= iNumRoutes; index++)
            {
                if (index == 1)
                    sTrayPortName = "Route";
                else
                    sTrayPortName = "Route_" + index;
                dStructRouteHorDist[index] = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Structure", sTrayPortName, PortDistanceType.Vertical);
                iTrayIndex[index] = index;
            }

            double temp;
            int tmpIndex;

            for (int i = 1; i <= iNumRoutes; i++)
            {
                for (int j = i + 1; j <= iNumRoutes; j++)
                {
                    if (dStructRouteHorDist[j] > dStructRouteHorDist[i])
                    {
                        temp = dStructRouteHorDist[i];
                        tmpIndex = iTrayIndex[i];
                        dStructRouteHorDist[i] = dStructRouteHorDist[j];
                        iTrayIndex[i] = iTrayIndex[j];
                        dStructRouteHorDist[j] = temp;
                        iTrayIndex[j] = tmpIndex;
                    }
                }
            }
            return dStructRouteHorDist[1];
        }
    }
}
