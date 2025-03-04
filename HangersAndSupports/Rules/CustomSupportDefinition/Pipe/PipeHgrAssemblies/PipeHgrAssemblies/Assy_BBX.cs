//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_BBX.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_BBX
//   Author       :  Vijay
//   Creation Date:  15-04-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-04-2013     Vijay   CR-CP-224484  Convert HS_Assembly to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class Assy_BBX : CustomSupportDefinition
    {
        private const string LOWBOXZ = "LowBoxZ";
        private const string LOWBOXY = "LowBoxY";
        private const string HIGHBOXZ = "HighBoxZ";
        private const string HIGHBOXY = "HighBoxY";
        private const string PLANE = "Plane";
        private const double CONST_1 = 0.15, CONST_2 = 0.025, CONST_3 = 0.0125, CONST_4 = 0.1;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    parts.Add(new PartInfo(LOWBOXZ, "Utility_USER_FIXED_BOX_1"));
                    parts.Add(new PartInfo(HIGHBOXZ, "Utility_USER_FIXED_CYL_1"));
                    parts.Add(new PartInfo(LOWBOXY, "Utility_USER_FIXED_BOX_1"));
                    parts.Add(new PartInfo(HIGHBOXY, "Utility_USER_FIXED_CYL_1"));
                    parts.Add(new PartInfo(PLANE, "Utility_SQUARE_GROUT_1"));

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
                return 2;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                bool mirror;
                Vector globalX, globalY, globalZ, bbX, bbZ, bbY;
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                string pipeOrientation = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrBBXPipeOrietation", "PipeOrientation")).PropValue;
                PropertyValueCodelist alignmentCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrBBXAlignment", "BBXAlignment");
                mirror = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrBBXMirror", "BBXMirror")).PropValue;

                //==============================
                //Create the Bounding Box
                //==============================

                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                //Create Vectors to define the plane of the BBX

                if (pipeOrientation == "Horizontal")
                {
                    switch (int.Parse(alignmentCodelist.PropValue.ToString()))
                    {
                        case 1:
                            {
                                //Vertical Plane Normal Along Route - Z Axis Towards Structure
                                globalZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);//'Get Global Z
                                bbX = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, globalZ);
                                bbZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, bbX);// 'Project Vector From Route to Structure into the BBX Plane
                                BoundingBoxHelper.CreateBoundingBox(bbZ, bbX, "TestBBX", false, mirror, false, true);
                                break;
                            }
                        case 2:
                            {
                                //  Vertical Plane Normal Along Route - Z Axis Orthogonal to Global CS
                                globalZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);
                                bbX = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, globalZ);
                                bbZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, bbX);
                                bbY = globalZ.Cross(bbX);
                                // 'Get Orthogonal Vector in the plane of the BBX (Gz, -Gz, By, -By) depending on Angle
                                if (AngleBetweenVectors(globalZ, bbZ) < Math.Round(Math.PI, 2) / 4)

                                    bbZ = globalZ;

                                else if (AngleBetweenVectors(bbY, bbZ) > 3 * (Math.Round(Math.PI, 2) / 4))
                                {
                                    bbZ.X = -globalZ.X;
                                    bbZ.Y = -globalZ.Y;
                                    bbZ.Z = -globalZ.Z;
                                }
                                else if (AngleBetweenVectors(bbY, bbZ) < Math.Round(Math.PI, 2) / 4)
                                    bbZ = bbY;
                                else
                                    bbZ.X = -bbY.X; bbZ.Y = bbY.Y; bbZ.Z = bbY.Z;
                                BoundingBoxHelper.CreateBoundingBox(bbZ, bbX, "TestBBX", false, mirror, false, true);
                                break;
                            }
                    }
                }
                else
                {
                    switch (int.Parse(alignmentCodelist.PropValue.ToString()))
                    {
                        case 1:
                            {
                                //Horizontal Plane - Z Axis Towards Structure
                                bbX = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);
                                bbZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, bbX);// 'Project Vector From Route to Structure into the BBX Plane
                                BoundingBoxHelper.CreateBoundingBox(bbZ, bbX, "TestBBX", false, mirror, false, true);
                                break;
                            }
                        case 2:
                            {
                                // Horizontal Plane - Z Axis Othogonal to Global CS
                                bbX = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);//'Use GlobalZ as BBX X-Axis
                                bbZ = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, bbX);// 'Project vector from Route to Struct into plane of BBX
                                globalX = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalX, bbX);//'Project GlobalX into plane of BBX
                                globalY = BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalY, bbX);//'Project GlobalY into plane of BBX
                                //Get Orthogonal Vector in the plane of the BBX (Gx, -Gx, Gy or -Gy) depending on angle
                                if (AngleBetweenVectors(globalX, bbZ) < Math.Round(Math.PI, 2) / 4)
                                    bbZ = globalX;
                                else if (AngleBetweenVectors(globalX, bbZ) < Math.Round(Math.PI, 2) / 4)
                                {
                                    bbZ.X = -globalX.X; bbZ.Y = -globalX.Y;
                                    bbZ.Z = -globalX.Z;
                                }
                                else if (AngleBetweenVectors(globalY, bbZ) < Math.Round(Math.PI, 2) / 4)
                                    bbZ = globalY;
                                else
                                {
                                    bbZ.X = -globalY.X; bbZ.Y = -globalY.Y;
                                    bbZ.Z = -globalY.Z;
                                }
                                BoundingBoxHelper.CreateBoundingBox(bbZ, bbX, "TestBBX", false, mirror, false, true);
                                break;
                            }
                    }
                }

                //==============================
                //Get the Bounding Box Dimensions
                //==============================

                BoundingBox boundingBox;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                //==============================
                //Set the Attributes on the Boxes and Plate to make them a decent size (For Visual representation of BBX only)
                //==============================

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                componentDictionary[LOWBOXZ].SetPropertyValue(CONST_1, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                componentDictionary[LOWBOXZ].SetPropertyValue(CONST_2, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                componentDictionary[LOWBOXZ].SetPropertyValue(CONST_2, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                componentDictionary[HIGHBOXZ].SetPropertyValue(CONST_1, "IJOAHgrUtility_USER_FIXED_CYL", "L");
                componentDictionary[HIGHBOXZ].SetPropertyValue(CONST_3, "IJOAHgrUtility_USER_FIXED_CYL", "RADIUS");
                componentDictionary[LOWBOXY].SetPropertyValue(CONST_4, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                componentDictionary[LOWBOXY].SetPropertyValue(CONST_2, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                componentDictionary[LOWBOXY].SetPropertyValue(CONST_2, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                componentDictionary[HIGHBOXY].SetPropertyValue(CONST_4, "IJOAHgrUtility_USER_FIXED_CYL", "L");
                componentDictionary[HIGHBOXY].SetPropertyValue(CONST_3, "IJOAHgrUtility_USER_FIXED_CYL", "RADIUS");
                componentDictionary[PLANE].SetPropertyValue(0.01, "IJOAHgrUtility_SQUARE_GROUT", "L");
                componentDictionary[PLANE].SetPropertyValue(boundingBoxWidth, "IJOAHgrUtility_SQUARE_GROUT", "W");
                componentDictionary[PLANE].SetPropertyValue(boundingBoxHeight, "IJOAHgrUtility_SQUARE_GROUT", "T");

                //============
                //Create Joints
                //============
                //Create a collection to hold the joints

                JointHelper.CreateRigidJoint(LOWBOXZ, "StartOther", "-1", "TestBBX_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(LOWBOXY, "StartOther", "-1", "TestBBX_Low", Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.Y, 0, 0, 0);
                JointHelper.CreateRigidJoint(HIGHBOXZ, "StartOther", "-1", "TestBBX_High", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(HIGHBOXY, "StartOther", "-1", "TestBBX_High", Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.Y, 0, 0, 0);
                JointHelper.CreateRigidJoint(PLANE, "Other", "-1", "TestBBX_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -boundingBoxWidth / 2, 0);


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

                    routeConnections.Add(new ConnectionInfo(LOWBOXZ, 1));   //partindex, routeindex

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        private double AngleBetweenVectors(Vector vector1, Vector vector2)
        {
            return Math.Acos(vector1.Dot(vector2) / ((vector1.Length * vector2.Length)));
        }
    }
}
