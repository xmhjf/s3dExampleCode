//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ClipHoldClamp.cs
//   CabletrayAssy,Ingr.SP3D.Content.Support.Rules.CW_StructSup.cs
//   Author       :  MK
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013       MK        CR-CP-224477 - Converted CabletrayAssemblies to C# .Net
//  22-Jan-2015     PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Support.Middle;
using System.Collections.Generic;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle.Hidden;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class CW_StructSup : CustomSupportDefinition
    {
        double noOfParts, tierSpacing, trayOffset, beamOffset;
        int  tierNo, idxLegBegin, idxLegEnd, idxConnectionBegin, idxPlateBegin, idxConnectionEnd, idxPlateEnd;
        string[] part = new string[10];
        int legPairNumber = 2;
        public override Collection<PartInfo> Parts
        {

            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    SupportComponentUtils supportComponentUtils = new SupportComponentUtils();
                    tierNo = (int)((PropertyValueInt)support.GetPropertyValue("IJUAHgrCWOcc", "TierNo")).PropValue;
                    tierSpacing = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrCWOcc", "TierSpacing")).PropValue;
                    trayOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrCWOcc", "TrayOffset")).PropValue;
                    beamOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrCWOcc", "BeamOffset")).PropValue;
                    if (tierSpacing == 0)
                        tierSpacing = 0.5;

                    if (tierNo == 0)
                        tierNo = 1;
                    //retrieve attributes from cableway support object                    
                    idxLegBegin = 1;
                    idxLegEnd = 2 * legPairNumber;
                    idxConnectionBegin = idxLegEnd + 1;
                    idxConnectionEnd = idxConnectionBegin + 2 * legPairNumber - 1;
                    idxPlateBegin = idxConnectionEnd + 1;
                    idxPlateEnd = idxPlateBegin + tierNo - 1;
                    noOfParts = 2 * legPairNumber + 2 * legPairNumber + tierNo;
                    string[] partClass = new string[Convert.ToInt32(noOfParts) + 4];
                    partClass[2] = partClass[3] = partClass[1] = partClass[4] = "HgrBeam";
                    for (int i = idxLegBegin; i <= idxLegEnd; i++)
                    {
                        partClass[i] = "HgrBeam";
                    }
                    for (int i = idxConnectionBegin; i <= idxConnectionEnd; i++)
                    {
                        partClass[i] = "Connection";
                    }
                    for (int i = idxPlateBegin; i <= idxPlateEnd; i++)
                    {
                        partClass[i] = "HgrSupFlatPlate";
                    }
                    for (int i = 1; i <= partClass.Length - 1; i++)
                    {
                        Part flatPlate;
                        part[i] = "part" + i;

                        if (i >= idxLegBegin && i <= idxLegEnd)
                        {
                            flatPlate = supportComponentUtils.GetPartFromPartClass("HgrBeam", "HgrCrossSectionByCW", support);
                            parts.Add(new PartInfo(part[i], flatPlate.ToString()));
                        }
                        else if (i >= idxConnectionBegin && i <= idxConnectionEnd)
                        {
                            flatPlate = supportComponentUtils.GetPartFromPartClass("Connection", "", support);
                            parts.Add(new PartInfo(part[i], flatPlate.ToString()));
                        }
                        else
                        {
                            flatPlate = supportComponentUtils.GetPartFromPartClass("HgrSupFlatPlate", "HgrPartByCW", support);
                            parts.Add(new PartInfo(part[i], flatPlate.ToString()));
                        }
                    }
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
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                //ReadOnlyCollection<IPort> supportingPorts = SupportHelper.SupportingObjectPorts;
                BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                BoundingBox boundingBox;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double dWidth = boundingBox.Width;
                double dHeight = boundingBox.Height;
                double pipeRadius = 0;
                string beginCap = null;
                string endCap = null;
                int portconfiguration = 0;
                CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                IPort portStruct = (IPort)support.SupportingFaces[0];
                PortConfiguration portStructCofiguration = RefPortHelper.PortConfiguration(1, 1);
                portconfiguration = (int)portStructCofiguration; 
                double structOrient = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Z, "WORLD", PortAxisType.Z, OrientationAlong.Direct);
                Plane plane1 = new Plane();
                Plane plane2 = new Plane();
                Axis axis1 = new Axis();
                Axis axis2 = new Axis();
                if (structOrient < Math.PI / 4.0 || structOrient > 3 * Math.PI / 4)
                {
                    if (portconfiguration == 1)
                    {
                        plane1 = Plane.XY;
                        plane2 = Plane.NegativeXY;
                        axis1 = Axis.X;
                        axis2 = Axis.Y;
                        beginCap = "BeginCap";
                        endCap = "EndCap";
                    }
                    else
                    {
                        plane1 = Plane.XY;
                        plane2 = Plane.XY;
                        axis1 = Axis.X;
                        axis2 = Axis.Y;
                        beginCap = "EndCap";
                        endCap = "BeginCap";
                    }
                    
                    double lugOffset;
                    if (Configuration == 1)
                        lugOffset = pipeRadius;
                    else
                        lugOffset = 2 * pipeRadius;

                    string strBOverLength = "BeginOverLength";
                    string strEOverLength = "EndOverLength";
                    for (int i = idxLegBegin; i <= idxLegEnd; i++)
                    {
                        if (i <= 2)
                            (componentDictionary[part[i]]).SetPropertyValue(beamOffset, "IJUAHgrOccOverLength", strBOverLength);
                        else
                            (componentDictionary[part[i]]).SetPropertyValue(beamOffset, "IJUAHgrOccOverLength", strEOverLength);
                    }
                    double width = 1.1 * dWidth;
                    double length = 4 * dWidth;
                    double thickness = 0.5 * dHeight;
                    support.SetPropertyValue(width, "", "Width");
                    support.SetPropertyValue(length, "", "Length");
                    support.SetPropertyValue(thickness, "", "Thickness");                  
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint("-1", "BBR_Low", part[idxPlateBegin], "Left", Plane.XY, Plane.XY, Axis.X, Axis.X, -thickness / 2 - trayOffset, (width - dWidth) / 2 + dWidth);
                    else
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", part[idxPlateBegin], "Left", Plane.XY, Plane.XY, Axis.X, Axis.X, -thickness / 2 - trayOffset, (width - dWidth) / 2 + dWidth, 0);

                    int seq;
                    double offset = 0;
                    double distD = 0;
                    double spaceD = 0;
                    double distW = 0;
                    double spaceW = 0;
                    if (legPairNumber > 1)
                    {
                        distD = (length - 2.0 * beamOffset - width) / (legPairNumber - 1);
                        spaceD = length / 2.0 - beamOffset;
                        distW = (length - 2.0 * beamOffset - length) / (legPairNumber - 1);
                        spaceW = length / 2.0 - beamOffset;
                    }
                    if (portconfiguration == 1)
                    {
                        spaceD = spaceD - cableInfo.Width;
                        spaceW = spaceW - length;
                    }
                    for (int i = idxConnectionBegin; i <= idxConnectionBegin + legPairNumber - 1; i++)
                    {
                        seq = i - idxConnectionBegin;
                        offset = spaceD - seq * distD;
                        JointHelper.CreateRigidJoint(part[idxPlateBegin], "Left", part[i], "Connection", Plane.YZ, Plane.ZX, Axis.Y, Axis.X, offset, 0, 0);
                        JointHelper.CreateRevoluteJoint(part[i], "Connection", part[i - 2 * legPairNumber], beginCap, Axis.Z, Axis.Z);
                    }
                    for (int i = idxConnectionBegin + legPairNumber; i <= idxConnectionEnd; i++)
                    {
                        seq = i - (idxConnectionBegin + legPairNumber);
                        offset = spaceW - seq * distW;
                        JointHelper.CreateRigidJoint(part[idxPlateBegin], "Left", part[i], "Connection", Plane.YZ, Plane.ZX, Axis.Y, Axis.X, offset, 0, 0);
                        JointHelper.CreateRevoluteJoint(part[i], "Connection", part[i - 2 * legPairNumber], beginCap, Axis.Z, Axis.Z);
                    }
                    for (int i = idxLegBegin; i <= idxLegEnd; i++)
                    {
                        JointHelper.CreateCylindricalJoint(part[i], beginCap, part[i], endCap, Axis.X, Axis.X, 0);
                    }
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (legPairNumber > 1)
                            //JointHelper.CreatePrismaticJoint("-1", "Structure", Idx_Leg_Begin, EndCap, StructLegConfig, 0, Length / 2 - BeamOffset);
                            JointHelper.CreatePrismaticJoint("-1", "Structure", part[idxLegBegin], endCap, plane1, plane2, axis1, axis2, 0, length / 2 - beamOffset);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "Structure", part[idxLegBegin], endCap, plane1, plane2, axis1, axis2, 0, 0);
                    }
                    double angleZ = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Z, "WORLD", PortAxisType.Z, OrientationAlong.Direct);
                    double tanZ = Math.Abs(Math.Tan(angleZ));
                    for (int i = idxPlateBegin + 1; i <= idxPlateEnd; i++)
                    {
                        offset = tanZ * ((thickness + tierSpacing) * (i - idxPlateBegin));
                        JointHelper.CreateRigidJoint(part[idxPlateBegin], "Left", part[i], "Left", Plane.XY, Plane.XY, Axis.X, Axis.X, (tierSpacing + thickness) * (i - idxPlateBegin), 0, offset);
                    }
                    for (int i = idxLegBegin; i <= idxLegEnd; i++)
                    {
                        JointHelper.CreateGlobalAxesAlignedJoint(part[i], "EndCap", Axis.Y, Axis.Z);
                    }
                }
                else
                {
                    plane1 = Plane.XY;
                    plane2 = Plane.NegativeXY;
                    axis1 = Axis.X;
                    axis2 = Axis.Y;
                    beginCap = "BeginCap";
                    endCap = "EndCap";
                    CableTrayObjectInfo cableInfo2 = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);                                                                           
                    double lugOffset;
                    if (Configuration == 1)
                        lugOffset = pipeRadius;
                    else
                        lugOffset = 2 * pipeRadius;
                    string strBOverLength = "BeginOverLength";
                    string strEOverLength = "EndOverLength";


                    for (int i = idxLegBegin; i <= idxLegEnd; i++)
                    {
                        if (i <= 2)
                            (componentDictionary[part[i]]).SetPropertyValue(beamOffset, "IJUAHgrOccOverLength", strBOverLength);
                        else
                            (componentDictionary[part[i]]).SetPropertyValue(beamOffset, "IJUAHgrOccOverLength", strEOverLength);
                    }
                    double width = 1.1 * dWidth;
                    double length = 4 * dWidth;
                    double thickness = 0.5 * dHeight;

                    support.SetPropertyValue(width, "", "Width");
                    support.SetPropertyValue(length, "", "Length");
                    support.SetPropertyValue(thickness, "", "Thickness");

                
                    int seq;
                    double offset = 0;
                    double distD = 0;
                    double spaceD = 0;
                    double distW = 0;
                    double spaceW = 0;
                    for (int i = idxConnectionBegin; i <= idxConnectionBegin + legPairNumber - 1; i++)
                    {
                        seq = i - idxConnectionBegin;
                        offset = spaceD - seq * distD;
                        JointHelper.CreateRigidJoint(part[idxPlateBegin], "Left", part[i], "Connection", Plane.YZ, Plane.ZX, Axis.Y, Axis.X, offset, 0, 0);
                        JointHelper.CreateRevoluteJoint(part[i], "Connection", part[i - 2 * legPairNumber], beginCap, Axis.Z, Axis.Z);
                    }

                    for (int i = idxConnectionBegin + legPairNumber; i <= idxConnectionEnd; i++)
                    {
                        seq = i - (idxConnectionBegin + legPairNumber);
                        offset = spaceW - seq * distW;
                        JointHelper.CreateRigidJoint(part[idxPlateBegin], "Left", part[i], "Connection", Plane.YZ, Plane.ZX, Axis.Y, Axis.X, offset, 0, 0);
                        JointHelper.CreateRevoluteJoint(part[i], "Connection", part[i - 2 * legPairNumber], beginCap, Axis.Z, Axis.Z);
                    }
                    for (int i = idxLegBegin; i <= idxLegEnd; i++)
                    {
                        JointHelper.CreateCylindricalJoint(part[i], beginCap, part[i], endCap, Axis.X, Axis.X, 0);
                    }
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRigidJoint(part[idxPlateBegin], "Left", part[1], "Connection", Plane.YZ, Plane.ZX, Axis.Y, Axis.X, offset, 0, 0);
                    }

                    double AngleZ = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Z, "WORLD", PortAxisType.Z, OrientationAlong.Direct);
                    double tanZ = Math.Abs(Math.Tan(AngleZ));
                    for (int i = idxPlateBegin + 1; i <= idxPlateEnd; i++)
                    {
                        offset = tanZ * ((thickness + tierSpacing) * (i - idxPlateBegin));
                        JointHelper.CreateRigidJoint(part[idxPlateBegin], "Left", part[i], "Left", Plane.XY, Plane.XY, Axis.X, Axis.X, (tierSpacing + thickness) * (i - idxPlateBegin), 0, offset);
                    }
                    for (int i = idxLegBegin; i <= idxLegEnd; i++)
                    {
                        JointHelper.CreateGlobalAxesAlignedJoint(part[i], "EndCap", Axis.Y, Axis.Z);
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
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 1;
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

                    routeConnections.Add(new ConnectionInfo(part[5], 1)); // partindex, routeindex
                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
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

                    structConnections.Add(new ConnectionInfo(part[1], 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(part[2], 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(part[3], 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(part[4], 1)); // partindex, routeindex

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
                }
            }
        }
    }
}