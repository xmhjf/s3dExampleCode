
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_GD_GEN.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_GD_GEN
//   Author       :Vijaya
//   Creation Date:09.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  09.Apr.2013     Vijaya   CR-CP-224484-Initial Creation
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Linq;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class Assy_GD_GEN : CustomSupportDefinition
    {
        //Constants
        private const string CSECTION1 = "nCSection1"; 
        private const string CSECTION2 = "nCSection2";
        private const string LSECTION1 = "nLSection1";
        private const string LSECTION2 = "nLSection2";

        string lSection;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    lSection = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecSize2")).PropValue;

                    parts.Add(new PartInfo(CSECTION1, "Utility_CUTBACK_C1_1"));
                    parts.Add(new PartInfo(CSECTION2, "Utility_CUTBACK_C1_1"));
                    parts.Add(new PartInfo(LSECTION1, lSection));
                    parts.Add(new PartInfo(LSECTION2, lSection));

                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 4;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeRadius = pipeInfo.OutsideDiameter / 2.0;

                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                bool isVerticalRoute;

                if (Math.Round(Math.Abs(routeAngle), 3) < Math.Round(Math.PI / 4, 3))
                    isVerticalRoute = true;
                else if (Math.Round(Math.Abs(routeAngle), 3) > Math.Round(3 * Math.PI / 4, 3))
                    isVerticalRoute = true;
                else
                    isVerticalRoute = false;

                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                double gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrVesselGuide", "Gap")).PropValue;

                string cSectionStandard = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecStd")).PropValue;
                string cSectionType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecType")).PropValue;
                string cSectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecSize")).PropValue;
                string lSectionStandard = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecStd2")).PropValue;
                string lSectionType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecType2")).PropValue;
                string lSectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecSize2")).PropValue;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                CatalogStructHelper catalogHelper = new CatalogStructHelper();
                CrossSection cCrosssection = catalogHelper.GetCrossSection(cSectionStandard, cSectionType, cSectionSize.Remove(0, 10));

                //Get the C Section Data
                double cWidth = cCrosssection.Width;
                double cFlangeThickness = (double)((PropertyValueDouble)cCrosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                double webThickness = (double)((PropertyValueDouble)cCrosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tw")).PropValue;
                double cDepth = cCrosssection.Depth;

                //Get the L Section Width
                CrossSection lCrosssection = catalogHelper.GetCrossSection(lSectionStandard, lSectionType, lSectionSize.Remove(0, 10));
                double lWidth = lCrosssection.Width;

                double lSectionLength = pipeRadius * 2 + 2 * cWidth + 2 * gap;
                double cSectionLength = length + pipeRadius + gap + lWidth;
                double cSectionLength1 = cSectionLength;
                double cSectionLength2 = cSectionLength;

                //Get the structure surface
                bool isBBRVLow_yGlobal = false;
                bool useBBRVlowReverese = false;

                BusinessObject SupportingFace = (BusinessObject)support.SupportingFaces.First();

                Matrix4X4 BBRVMatrix;
                Vector BBRVYReverse = new Vector(0,0,0) , normalprojectedVector = new Vector(0,0,0);
                Position projectedVector = new Position(0,0,0), leftPosition = new Position(0, 0, 0), rightPosition = new Position(0, 0, 0);
                double cutbackAngle1 = 0, cutbackAngle2 = 0;
                bool portExists;// = false;

                if (support.SupportingFaces != null)
                {
                    if (support.SupportingFaces.Count > 0)// if struct object is there, find the exact distance and the cutback angles
                    {
                        Matrix4X4 routeMatrix = RefPortHelper.PortLCS("Route");
                        try
                        {
                            BBRVMatrix = RefPortHelper.PortLCS("BBRV_Low");
                            portExists = true;
                        }
                        catch
                        {
                            portExists = false;
                            BBRVMatrix = new Matrix4X4();
                            BBRVMatrix.Set(new Position(0, 0, 0), new Vector(0, 0, 0), new Vector(0, 0, 0));
                        }
                        double offsetDist = Math.Round(gap + (pipeInfo.OutsideDiameter + cWidth) / 2, 7);

                        if (Configuration == 1 || Configuration == 3)
                        {
                            Vector localVetor = routeMatrix.YAxis;
                            localVetor.Length = offsetDist;
                            leftPosition = routeMatrix.Origin.Offset(localVetor);

                            localVetor.Length = -offsetDist;
                            rightPosition = routeMatrix.Origin.Offset(localVetor);
                        }
                        else if (Configuration == 2 || Configuration == 4)
                        {
                            if (portExists)
                            {
                                int local = Math.Abs((int)BBRVMatrix.YAxis.Z);//using this double value in the if condition not working properly, so taking a local Integer.
                                if (local == 1)
                                {
                                    isBBRVLow_yGlobal = true;
                                    BBRVYReverse = new Vector(-BBRVMatrix.YAxis.X, -BBRVMatrix.YAxis.Y, -BBRVMatrix.YAxis.Z);

                                    // try to get a projected point on the Structure in the BBRV-Y direction      
                                    try
                                    {
                                        SupportingHelper.GetProjectedPointOnSurface(routeMatrix.Origin, BBRVMatrix.YAxis, SupportingFace, out projectedVector, out normalprojectedVector);
                                    }
                                    catch
                                    {
                                        if (projectedVector == null)
                                        {
                                            SupportingHelper.GetProjectedPointOnSurface(routeMatrix.Origin, BBRVYReverse, SupportingFace, out projectedVector, out normalprojectedVector);
                                            if (projectedVector == null)
                                                useBBRVlowReverese = false;
                                            else
                                                useBBRVlowReverese = true;
                                        }
                                        else
                                            useBBRVlowReverese = false;
                                    }
                                }
                                if (isBBRVLow_yGlobal)
                                {
                                    BBRVMatrix.ZAxis.Length = offsetDist;
                                    leftPosition = routeMatrix.Origin.Offset(BBRVMatrix.ZAxis);
                                    BBRVMatrix.ZAxis.Length = -offsetDist;
                                    rightPosition = routeMatrix.Origin.Offset(BBRVMatrix.ZAxis);
                                }
                                else
                                {
                                    BBRVMatrix.YAxis.Length = offsetDist;
                                    leftPosition = routeMatrix.Origin.Offset(BBRVMatrix.YAxis);
                                    BBRVMatrix.YAxis.Length = -offsetDist;
                                    rightPosition = routeMatrix.Origin.Offset(BBRVMatrix.YAxis);
                                }
                            }
                        }
                        if (Configuration == 1 || Configuration == 3)
                            SupportingHelper.GetProjectedPointOnSurface(leftPosition, routeMatrix.ZAxis, SupportingFace, out projectedVector, out normalprojectedVector);
                        else if (Configuration == 2 || Configuration == 4)
                            if (portExists)
                            {
                                if (isBBRVLow_yGlobal)
                                    if (useBBRVlowReverese)
                                        SupportingHelper.GetProjectedPointOnSurface(leftPosition, BBRVYReverse, SupportingFace, out projectedVector, out normalprojectedVector);
                                    else
                                        SupportingHelper.GetProjectedPointOnSurface(leftPosition, BBRVMatrix.YAxis, SupportingFace, out projectedVector, out normalprojectedVector);
                                else
                                    SupportingHelper.GetProjectedPointOnSurface(leftPosition, BBRVMatrix.ZAxis, SupportingFace, out projectedVector, out normalprojectedVector);
                            }
                        if (projectedVector != null)
                        {
                            cSectionLength1 = projectedVector.DistanceToPoint(leftPosition) + pipeRadius + gap + lWidth;

                            if (Configuration == 1 || Configuration == 3)
                                cutbackAngle1 = (3 * Math.PI / 2) - normalprojectedVector.Angle(routeMatrix.ZAxis, routeMatrix.XAxis);
                            else if (Configuration == 2 || Configuration == 4)
                                if (portExists)
                                {
                                    if (isBBRVLow_yGlobal)
                                        if (useBBRVlowReverese)
                                            cutbackAngle1 = (3 * Math.PI / 2) - normalprojectedVector.Angle(BBRVYReverse, BBRVMatrix.XAxis);
                                        else
                                            cutbackAngle1 = -Math.PI / 2 + normalprojectedVector.Angle(BBRVMatrix.YAxis, BBRVMatrix.XAxis);
                                    else
                                        cutbackAngle1 = (3 * Math.PI / 2) - normalprojectedVector.Angle(BBRVMatrix.ZAxis, BBRVMatrix.XAxis);
                                }
                        }
                        if (Configuration == 1 || Configuration == 3)
                            SupportingHelper.GetProjectedPointOnSurface(rightPosition, routeMatrix.ZAxis, SupportingFace, out projectedVector, out normalprojectedVector);
                        else if (Configuration == 2 || Configuration == 4)
                            if (portExists)
                            {
                                if (isBBRVLow_yGlobal)
                                    if (useBBRVlowReverese)
                                        SupportingHelper.GetProjectedPointOnSurface(rightPosition, BBRVYReverse, SupportingFace, out projectedVector, out normalprojectedVector);
                                    else
                                        SupportingHelper.GetProjectedPointOnSurface(rightPosition, BBRVMatrix.YAxis, SupportingFace, out projectedVector, out normalprojectedVector);
                                else
                                    SupportingHelper.GetProjectedPointOnSurface(rightPosition, BBRVMatrix.ZAxis, SupportingFace, out projectedVector, out normalprojectedVector);
                            }

                        if (projectedVector != null)
                        {
                            cSectionLength2 = projectedVector.DistanceToPoint(rightPosition) + pipeRadius + gap + lWidth;

                            if (Configuration == 1 || Configuration == 3)
                                cutbackAngle2 = -Math.PI / 2 + normalprojectedVector.Angle(routeMatrix.ZAxis, routeMatrix.XAxis);
                            else if (Configuration == 2 || Configuration == 4)
                                if(portExists)
                                {
                                if (isBBRVLow_yGlobal)
                                    if (useBBRVlowReverese)
                                        cutbackAngle2 = -Math.PI / 2 + normalprojectedVector.Angle(BBRVYReverse, BBRVMatrix.XAxis);
                                    else
                                        cutbackAngle2 = (3 * Math.PI / 2) - normalprojectedVector.Angle(BBRVMatrix.YAxis, BBRVMatrix.XAxis);
                                else
                                    cutbackAngle2 = -Math.PI / 2 + normalprojectedVector.Angle(BBRVMatrix.ZAxis, BBRVMatrix.XAxis);
                                }
                        }
                    }
                }
                string CBOM = "AISC " + (cSectionSize.Remove(0,10)).Trim() + ", Length: " + cSectionLength;

                componentDictionary[CSECTION1].SetPropertyValue(cSectionLength1, "IJOAHgrUtility_GENERIC_W", "L");
                componentDictionary[CSECTION1].SetPropertyValue(cWidth, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                componentDictionary[CSECTION1].SetPropertyValue(cDepth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                componentDictionary[CSECTION1].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                componentDictionary[CSECTION1].SetPropertyValue(webThickness, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                componentDictionary[CSECTION1].SetPropertyValue(Math.Round(Math.PI / 2,7), "IJOAHgrUtility_CUTBACK", "ANGLE");
                componentDictionary[CSECTION1].SetPropertyValue(cutbackAngle1, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                componentDictionary[CSECTION1].SetPropertyValue(CBOM, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");

                componentDictionary[CSECTION2].SetPropertyValue(cSectionLength2, "IJOAHgrUtility_GENERIC_W", "L");
                componentDictionary[CSECTION2].SetPropertyValue(cWidth, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                componentDictionary[CSECTION2].SetPropertyValue(cDepth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                componentDictionary[CSECTION2].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                componentDictionary[CSECTION2].SetPropertyValue(webThickness, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                componentDictionary[CSECTION2].SetPropertyValue(Math.Round(Math.PI / 2, 7), "IJOAHgrUtility_CUTBACK", "ANGLE");
                componentDictionary[CSECTION2].SetPropertyValue(cutbackAngle2, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                componentDictionary[CSECTION2].SetPropertyValue(CBOM, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");

                componentDictionary[LSECTION1].SetPropertyValue(lSectionLength, "IJUAHgrOccLength", "Length");
                componentDictionary[LSECTION2].SetPropertyValue(lSectionLength, "IJUAHgrOccLength", "Length");

                //Create Joints
                if (Configuration == 1 || Configuration == 2)
                {
                    //Add a Joint between the Second L and the second C
                    JointHelper.CreateRigidJoint(LSECTION2, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -cWidth + lSectionLength, -cDepth / 2, lWidth);

                    //Add a Joint between the First L and the second C
                    JointHelper.CreateRigidJoint(LSECTION1, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, cWidth, -cDepth / 2, -pipeRadius * 2 - gap * 2 - lWidth);
                }
                else if (Configuration == 3 || Configuration == 4)
                {
                    //Add a Joint between the Second L and the second C
                    JointHelper.CreateRigidJoint(LSECTION2, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -cWidth + lSectionLength, -cDepth / 2, -pipeRadius * 2 - gap * 2 - lWidth);

                    //Add a Joint between the First L and the second C
                    JointHelper.CreateRigidJoint(LSECTION1, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, cWidth, -cDepth / 2, lWidth);
                }

                if (Configuration == 1 || Configuration == 3)
                {
                    //Add a Joint between the First C and the route
                    JointHelper.CreateRigidJoint(CSECTION1, "StartStructure", "-1", "Route", Plane.XY, Plane.NegativeZX, Axis.X, Axis.Z, pipeRadius + gap, 0, pipeRadius + gap + lWidth);

                    //Add a Joint between the Second C and the route
                    JointHelper.CreateRigidJoint(CSECTION2, "StartStructure", "-1", "Route", Plane.XY, Plane.ZX, Axis.X, Axis.Z, pipeRadius + gap, 0, pipeRadius + gap + lWidth);
                }
                else if (Configuration == 2 || Configuration == 4)
                {
                    if (isVerticalRoute == true)
                    {
                        //Add a Joint between the First C and the route
                        JointHelper.CreateRigidJoint(CSECTION1, "StartStructure", "-1", "Route", Plane.XY, Plane.NegativeZX, Axis.X, Axis.Z, pipeRadius + gap, 0, pipeRadius + gap + lWidth);

                        //Add a Joint between the Second C and the route
                        JointHelper.CreateRigidJoint(CSECTION2, "StartStructure", "-1", "Route", Plane.XY, Plane.ZX, Axis.X, Axis.Z, pipeRadius + gap, 0, pipeRadius + gap + lWidth);
                    }
                    else if (isBBRVLow_yGlobal)
                    {
                        if (useBBRVlowReverese)
                        {
                            //Add a Joint between the First C and the route
                            JointHelper.CreateRigidJoint(CSECTION1, "StartStructure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 2 * pipeRadius + gap, 0, 2 * pipeRadius + gap + lWidth);

                            //Add a Joint between the Second C and the route
                            JointHelper.CreateRigidJoint(CSECTION2, "StartStructure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, gap, 0, 2 * pipeRadius + gap + lWidth);
                        }
                        else
                        {
                            //Add a Joint between the First C and the route
                            JointHelper.CreateRigidJoint(CSECTION1, "StartStructure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 2 * pipeRadius + gap, 0, gap + lWidth);

                            //Add a Joint between the Second C and the route
                            JointHelper.CreateRigidJoint(CSECTION2, "StartStructure", "-1", "BBRV_Low", Plane.XY, Plane.XY, Axis.X, Axis.Y, gap, 0, gap + lWidth);
                        }
                    }
                    else
                    {
                        //Add a Joint between the First C and the route
                        JointHelper.CreateRigidJoint(CSECTION1, "StartStructure", "-1", "BBRV_Low", Plane.XY, Plane.NegativeZX, Axis.X, Axis.Z, 2 * pipeRadius + gap, 0, gap + lWidth);

                        //Add a Joint between the Second C and the route
                        JointHelper.CreateRigidJoint(CSECTION2, "StartStructure", "-1", "BBRV_Low", Plane.XY, Plane.ZX, Axis.X, Axis.Z, gap, 0, gap + lWidth);
                    }
                }

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(CSECTION1, 1)); // partindex, routeindex

                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    //Return the collection of Structure connection information.
                    return structConnections;//support is not connecting to any structure so we have nothing to return
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        /// <summary>
        /// This method returns the Mirrored configuration Value
        /// </summary>
        /// <param name="CurrentMirrorToggleValue">int - Toggle Value.</param>
        /// <param name="eMirrorPlane">MirrorPlane - eMirrorPlane.</param>
        /// <returns>int</returns>        
        /// <code>
        ///     MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane);
        ///</code>
        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {
                if (eMirrorPlane != MirrorPlane.YZPlane)
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 3;
                    else if (CurrentMirrorToggleValue == 2)
                        return 4;
                    else if (CurrentMirrorToggleValue == 3)
                        return 1;
                    else if (CurrentMirrorToggleValue == 4)
                        return 2;
                    else
                        return CurrentMirrorToggleValue;
                }
                else
                    return CurrentMirrorToggleValue;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Mirrored Configuration." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}



