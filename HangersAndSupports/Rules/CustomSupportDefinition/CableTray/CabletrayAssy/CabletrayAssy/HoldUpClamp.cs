//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HoldUpClamp.cs
//   CabletrayAssy,Ingr.SP3D.Content.Support.Rules.HoldUpClamp
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
    public class HoldUpClamp : CustomSupportDefinition
    {
        private const string CTHOLDUPCLAMP1 = "CTHOLDUPCLAMP1";
        private const string CTHOLDUPCLAMP2 = "CTHOLDUPCLAMP2";
        private const string G4G1461011 = "G4G1461011"; //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1461012 = "G4G1461012"; //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1460011 = "G4G1460011"; //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
        private const string G4G1460012 = "G4G1460012"; //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)

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
                    PartClass ctHoldSideClamp = (PartClass)catalogBaseHelper.GetPartClass("CTHoldUpClamp");
                    string partselection = ctHoldSideClamp.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    Part g4gPart1 = supportComponentUtils.GetPartFromPartClass("G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize", support);
                    Part ctHoldSideClampPart1 = supportComponentUtils.GetPartFromPartClass("CTHoldUpClamp", partselection, support);
                    Part ctHoldSideClampPart2 = supportComponentUtils.GetPartFromPartClass("CTHoldUpClamp", partselection, support);
                    PartClass G4G146101 = (PartClass)catalogBaseHelper.GetPartClass("G4G_1460_01");
                    string g4gPart2Part = string.Empty;
                    Part g4gPart2 = null;
                    foreach (BusinessObject businessObject in G4G146101.Parts)
                    {
                        if (((double)((PropertyValueDouble)businessObject.GetPropertyValue("IJUAHgrDia", "RodSizeDia")).PropValue) >= (((double)((PropertyValueDouble)ctHoldSideClampPart1.GetPropertyValue("IJUAHgrCTHdUpClamp", "RodDiameter")).PropValue)))
                        {
                            g4gPart2 = (Part)businessObject;
                            g4gPart2Part = g4gPart2.ToString();
                            break;
                        }
                    }
                    parts.Add(new PartInfo(CTHOLDUPCLAMP1, ctHoldSideClampPart1.ToString()));
                    parts.Add(new PartInfo(CTHOLDUPCLAMP2, ctHoldSideClampPart2.ToString()));
                    parts.Add(new PartInfo(G4G1461011, g4gPart1.ToString()));
                    parts.Add(new PartInfo(G4G1461012, g4gPart1.ToString()));
                    parts.Add(new PartInfo(G4G1460011, g4gPart2Part));
                    parts.Add(new PartInfo(G4G1460012, g4gPart2Part));
                    return parts;


                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
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

                double depth = cableInfo.Depth;
                double width = cableInfo.Width;
                double radius = cableInfo.BendRadius;
                double beamWidth;
                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                if (SupportHelper.SupportingObjects.Count>0)
                {
                    beamWidth = SupportingHelper.SupportingObjectInfo(1).Width;
                }
                else
                {
                    beamWidth = 0;
                }

                // If unable to retrieve BeamWidth, Structure could be a slab.
                // Use the cable tray width to set the clamp dimensions.
                if (beamWidth <= 0)
                    beamWidth = width;

                (componentDictionary[CTHOLDUPCLAMP1]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDUPCLAMP1]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTHOLDUPCLAMP1]).SetPropertyValue(depth / 4.0, "IJUAHgrCTOffset", "TrayBeamWidth");

                (componentDictionary[CTHOLDUPCLAMP2]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDUPCLAMP2]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTHOLDUPCLAMP2]).SetPropertyValue(depth / 4.0, "IJUAHgrCTOffset", "TrayBeamWidth");

                (componentDictionary[G4G1461011]).SetPropertyValue((0.75 * beamWidth), "IJUAHgrOccGeometry", "Width");
                (componentDictionary[G4G1461012]).SetPropertyValue((0.75 * beamWidth), "IJUAHgrOccGeometry", "Width");

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================


                //======================================================
                //Create Joints
                //======================================================


                //Create the Flexible (Cylindrical) Joint between the ports of the rod_1
                JointHelper.CreateCylindricalJoint(G4G1460011, "RodTop", G4G1460011, "RodBottom", Axis.Z, Axis.Z, 0);

                //Create the Flexible (Cylindrical) Joint between the ports of the rod_2
                JointHelper.CreateCylindricalJoint(G4G1460012, "RodTop", G4G1460012, "RodBottom", Axis.Z, Axis.Z, 0);


                //Add a Joint between cable tray and Clamp
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePrismaticJoint(CTHOLDUPCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    JointHelper.CreatePrismaticJoint(CTHOLDUPCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);
                }
                else
                {
                    JointHelper.CreateRigidJoint(CTHOLDUPCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(CTHOLDUPCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                }

                //Add a Spherical Joint between cable tray Clamp and Bottom of Rod
                JointHelper.CreateSphericalJoint(CTHOLDUPCLAMP1, "Structure", G4G1460011, "RodBottom");
                JointHelper.CreateSphericalJoint(CTHOLDUPCLAMP2, "Structure", G4G1460012, "RodBottom");


                //Add a Spherical Joint between Beam Clamp and Top of Rod
                JointHelper.CreateSphericalJoint(G4G1460011, "RodTop", G4G1461011, "Rod");
                JointHelper.CreateSphericalJoint(G4G1460012, "RodTop", G4G1461012, "Rod");


                //Add a Vertical Joint to the Rod Z axis
                JointHelper.CreateGlobalAxesAlignedJoint(G4G1460011, "RodBottom", Axis.Z, Axis.Z);
                JointHelper.CreateGlobalAxesAlignedJoint(G4G1460012, "RodBottom", Axis.Z, Axis.Z);


                //Add a Joint between the Supporting Object (Beam or Plate) and the Beam Clamp
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    //Beam Structure ... Add a Prismatic Joint between Beam and Beam Clamp
                    if (Configuration == 1)
                    {
                        JointHelper.CreatePrismaticJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreatePrismaticJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreatePrismaticJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0);
                        JointHelper.CreatePrismaticJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0);
                    }
                }
                else
                {
                    //Plate Structure ... Add a Translation Joint Between Plate and Beam Clamp
                    if (Configuration == 1)
                    {
                        JointHelper.CreateTranslationalJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0);
                        JointHelper.CreateTranslationalJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0);
                    }
                    else
                    {
                        JointHelper.CreateTranslationalJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                        JointHelper.CreateTranslationalJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                    }
                }
            }
            catch (Exception exception)
            {
                Type myType = this.GetType();
                CmnException exception1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
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

                    routeConnections.Add(new ConnectionInfo(CTHOLDUPCLAMP1, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(CTHOLDUPCLAMP2, 1)); // partindex, routeindex

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