//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HoldDownClamp.cs
//   CabletrayAssy,Ingr.SP3D.Content.Support.Rules.HoldDownClamp
//   Author       :  MK
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013       MK        CR-CP-224477 - Converted CabletrayAssemblies to C# .Net
//  30/04/2015       Chethan   TR-CP-271643  Rec exc. minidumps created while placing support by ref. on cabletray 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Support.Middle.Hidden;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class HoldDownClamp : CustomSupportDefinition
    {
        private const string CTHOLDDOWNCLAMP1 = "CTHOLDDOWNCLAMP1";
        private const string CTHOLDDOWNCLAMP2 = "CTHOLDDOWNCLAMP2";
        private const string G4G1461011 = "G4G1461011"; //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1461012 = "G4G1461012"; //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1460011 = "G4G1460011"; //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
        private const string G4G1460012 = "G4G1460012"; //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
        private const string HGRBEAM = "HGRBEAM";   //GeneralProfileSymbols.HgrBeam (HgrBeam)

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    SupportComponentUtils supportComponentUtils = new SupportComponentUtils();
                    Collection<object> colllection = new Collection<object>();
                    PartClass ctClipHoldClamp = (PartClass)catalogBaseHelper.GetPartClass("CTHoldDownClamp");
                    string strPartSelection = ctClipHoldClamp.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    PartClass HgrBeam = (PartClass)catalogBaseHelper.GetPartClass("HgrBeam");
                    string strPartSelection1 = HgrBeam.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();

                    parts.Add(new PartInfo(CTHOLDDOWNCLAMP1, "CTHoldDownClamp", strPartSelection));
                    parts.Add(new PartInfo(CTHOLDDOWNCLAMP2, "CTHoldDownClamp", strPartSelection));
                    parts.Add(new PartInfo(G4G1461011, "G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));
                    parts.Add(new PartInfo(G4G1461012, "G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));
                    parts.Add(new PartInfo(G4G1460011, "G4G_1460_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));
                    parts.Add(new PartInfo(G4G1460012, "G4G_1460_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));
                    parts.Add(new PartInfo(HGRBEAM, "HgrBeam", strPartSelection1));
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
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========
                CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                double depth = cableInfo.Depth;
                double width = cableInfo.Width;
                double radius = cableInfo.BendRadius;

                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }
                double beamWidth;

                if (SupportHelper.SupportingObjects.Count>0)
                {
                    beamWidth = SupportingHelper.SupportingObjectInfo(1).Width;
                }
                else
                {
                    beamWidth = 0;
                }


                if (beamWidth <= 0)
                    beamWidth = width;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                (componentDictionary[CTHOLDDOWNCLAMP1]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDDOWNCLAMP1]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTHOLDDOWNCLAMP2]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDDOWNCLAMP2]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[G4G1461011]).SetPropertyValue((0.75 * beamWidth), "IJUAHgrOccGeometry", "Width");
                (componentDictionary[G4G1461012]).SetPropertyValue((0.75 * beamWidth), "IJUAHgrOccGeometry", "Width");
                double eOverLength, bOverLength;
                Collection<object> colllection = new Collection<object>();
                bool value = GenericHelper.GetDataByRule("HgrSupStructOffset", (componentDictionary[HGRBEAM]), out colllection);
                double lugOffset = 2 * (double)(colllection[0]);
                bOverLength = eOverLength = lugOffset;
                (componentDictionary[HGRBEAM]).SetPropertyValue((bOverLength), "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HGRBEAM]).SetPropertyValue((eOverLength), "IJUAHgrOccOverLength", "BeginOverLength");


                string strBBName = string.Empty, strBBLow = string.Empty, strBBHigh = string.Empty;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    strBBName = "BBSR";
                    strBBLow = "BBSR_Low";
                    strBBHigh = "BBSR_High";
                }
                else
                {
                    strBBName = "BBR";
                    strBBLow = "BBR_Low";
                    strBBHigh = "BBR_High";
                }
                BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                BoundingBox boundingBox;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double dWidth = boundingBox.Width;

                //Create the Flexible (Cylindrical) Joint between the ports of the rod_1
                JointHelper.CreateCylindricalJoint(G4G1460011, "RodTop", G4G1460011, "RodBottom", Axis.Z, Axis.Z, 0);
                //Create the Flexible (Cylindrical) Joint between the ports of the rod_2
                JointHelper.CreateCylindricalJoint(G4G1460012, "RodTop", G4G1460012, "RodBottom", Axis.Z, Axis.Z, 0);

                //Create the Joint between the RteLow Reference Port and the Bottom Symbol
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    //Beam Structure... Add a Prismatic Joint
                    JointHelper.CreatePlanarJoint("-1", "BBSR_Low", HGRBEAM, "BeginCap", Plane.ZX, Plane.XY, -(double)(colllection[0]));
                else
                    //Plate Structure... Add a Rigid Joint
                    JointHelper.CreatePlanarJoint("-1", "BBR_Low", HGRBEAM, "BeginCap", Plane.ZX, Plane.XY, -(double)(colllection[0]));


                //Create the Plane Joint between the RteHigh Reference Port
                //and the Right Bottom Symbol
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_High", HGRBEAM, "EndCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -depth, (double)(colllection[0]));
                else
                    JointHelper.CreateRigidJoint("-1", "BBR_High", HGRBEAM, "EndCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -depth, (double)(colllection[0]), 0);

                JointHelper.CreatePrismaticJoint(HGRBEAM, "BeginCap", HGRBEAM, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                JointHelper.CreateRigidJoint(G4G1460011, "RodBottom", HGRBEAM, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, (double)(colllection[0]), (double)(colllection[0]), -(double)(colllection[0]));
                JointHelper.CreateRigidJoint(G4G1460012, "RodBottom", HGRBEAM, "EndCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, (double)(colllection[0]), -(double)(colllection[0]), -(double)(colllection[0]));

                JointHelper.CreateSphericalJoint(G4G1460011, "RodTop", G4G1461011, "Rod");
                JointHelper.CreateSphericalJoint(G4G1460012, "RodTop", G4G1461012, "Rod");

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePrismaticJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0);
                    JointHelper.CreateTranslationalJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                }
                else
                {
                    JointHelper.CreateTranslationalJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                    JointHelper.CreateTranslationalJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                }

                JointHelper.CreatePrismaticJoint(CTHOLDDOWNCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                JointHelper.CreatePrismaticJoint(CTHOLDDOWNCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);

                JointHelper.CreatePlanarJoint(HGRBEAM, "BeginCap", CTHOLDDOWNCLAMP1, "Structure", Plane.YZ, Plane.YZ, (double)(colllection[0]));
                JointHelper.CreatePlanarJoint(HGRBEAM, "EndCap", CTHOLDDOWNCLAMP2, "Structure", Plane.YZ, Plane.NegativeYZ, (double)(colllection[0]));

            }
            // return the collection of ocpmmittedmtd joints,
            catch (Exception excepion)
            {
                Type myType = this.GetType();
                CmnException exception1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + excepion.Message, excepion);
                throw exception1;
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

                    routeConnections.Add(new ConnectionInfo(CTHOLDDOWNCLAMP1, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(CTHOLDDOWNCLAMP2, 1)); // partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(G4G1461011, 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(G4G1461012, 1)); // partindex, routeindex


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