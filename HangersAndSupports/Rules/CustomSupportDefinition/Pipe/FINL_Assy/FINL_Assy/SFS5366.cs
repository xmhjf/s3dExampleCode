//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5366.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5366
//   Author       :  Vijay
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Vijay   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5366 : CustomSupportDefinition
    {

        private const string LEFTGUIDE = "LeftGuide_5366";
        private const string RIGHTGUIDE = "RightGuide_5366";
        private const string ROUTECONNOBJECT = "RouteConnObject_5366";
        public Boolean Override { get; set; }
        public double ShoeHeight { get; set; }
        public double ShoeWidth { get; set; }
        public int GuideSide { get; set; }
        public double Gap { get; set; }
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    if (!Override)     //by default it is false
                    {
                        ShoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeHeight", "ShoeHeight")).PropValue;
                        ShoeWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeWidth", "ShoeWidth")).PropValue;
                        GuideSide = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLGuideSide", "GuideSide")).PropValue;
                        Gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLGap2", "Gap2")).PropValue;
                    }

                    if (GuideSide == 1)        //Both
                    {
                        parts.Add(new PartInfo(LEFTGUIDE, "Util_NotchPl_Metric_1"));
                        parts.Add(new PartInfo(RIGHTGUIDE, "Util_NotchPl_Metric_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT, "Log_Conn_Part_1"));
                    }
                    else if (GuideSide == 2)       //Right
                    {
                        parts.Add(new PartInfo(RIGHTGUIDE, "Util_NotchPl_Metric_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT, "Log_Conn_Part_1"));
                    }
                    else if (GuideSide == 3)       //Left
                    {
                        parts.Add(new PartInfo(LEFTGUIDE, "Util_NotchPl_Metric_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT, "Log_Conn_Part_1"));
                    }
                    return parts;       //Get the collection of parts
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //To get Pipe Nom Dia
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                NominalDiameter minNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                NominalDiameter maxNominalDiameter = new NominalDiameter();
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5366.cs", 120);
                    return;
                }
                string guideBom = "10 Steel Plate 50x30, Notch 25x13";
                double width = 50.00 / 1000.00;
                double depth = 30.00 / 1000.00;
                double X = 25.00 / 1000.00;
                double Z = 13.00 / 1000.00;
                double T = 10.00 / 1000.00;

                //this will be overriden by super assembly
                if (!Override)     //by default it is false
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                if (GuideSide == 1)    //Both
                {
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(width, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(depth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(X, "IJOAHgrUtilMetricX", "X");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(Z, "IJOAHgrUtilMetricZ", "Z");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(guideBom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    componentDictionary[LEFTGUIDE].SetPropertyValue(width, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(depth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(X, "IJOAHgrUtilMetricX", "X");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(Z, "IJOAHgrUtilMetricZ", "Z");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(guideBom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT, "Connection", RIGHTGUIDE, "CornerStructure", Plane.XY, Plane.NegativeZX, Axis.X, Axis.Z, -depth + pipeInfo.OutsideDiameter / 2 + ShoeHeight, X - ShoeWidth / 2 - Gap, 0);
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT, "Connection", LEFTGUIDE, "CornerStructure", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeZ, -depth + pipeInfo.OutsideDiameter / 2 + ShoeHeight, -(X - ShoeWidth / 2 - Gap), 0);
                }
                else if (GuideSide == 2)      //Right
                {
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(width, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(depth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(X, "IJOAHgrUtilMetricX", "X");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(Z, "IJOAHgrUtilMetricZ", "Z");
                    componentDictionary[RIGHTGUIDE].SetPropertyValue(guideBom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT, "Connection", RIGHTGUIDE, "CornerStructure", Plane.XY, Plane.NegativeZX, Axis.X, Axis.Z, -depth + pipeInfo.OutsideDiameter / 2 + ShoeHeight, X - ShoeWidth / 2 - Gap, 0);
                }
                else if (GuideSide == 3)      //Left
                {
                    componentDictionary[LEFTGUIDE].SetPropertyValue(width, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(depth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(T, "IJOAHgrUtilMetricT", "T");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(X, "IJOAHgrUtilMetricX", "X");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(Z, "IJOAHgrUtilMetricZ", "Z");
                    componentDictionary[LEFTGUIDE].SetPropertyValue(guideBom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT, "Connection", LEFTGUIDE, "CornerStructure", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeZ, -depth + pipeInfo.OutsideDiameter / 2 + ShoeHeight, -(X - ShoeWidth / 2 - Gap), 0);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Finl_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                    if (Override == true)
                        routeConnections.Add(new ConnectionInfo(ROUTECONNOBJECT, 1));      //partindex, routeindex
                    else
                    {
                        if (GuideSide == 1)  //Both
                        {
                            routeConnections.Add(new ConnectionInfo(LEFTGUIDE, 1));      //partindex, routeindex
                            routeConnections.Add(new ConnectionInfo(RIGHTGUIDE, 1));      //partindex, routeindex
                        }
                        else if (GuideSide == 2)     //Right
                            routeConnections.Add(new ConnectionInfo(RIGHTGUIDE, 1));       //partindex, routeindex
                        else if (GuideSide == 3)     //Left
                            routeConnections.Add(new ConnectionInfo(LEFTGUIDE, 1));       //partindex, routeindex
                    }

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

                    if (GuideSide == 1)  //Both
                    {
                        structConnections.Add(new ConnectionInfo(LEFTGUIDE, 1));      //partindex, routeindex
                        structConnections.Add(new ConnectionInfo(RIGHTGUIDE, 1));      //partindex, routeindex
                    }
                    else if (GuideSide == 2)     //Right
                        structConnections.Add(new ConnectionInfo(RIGHTGUIDE, 1));       //partindex, routeindex
                    else if (GuideSide == 3)     //Left
                        structConnections.Add(new ConnectionInfo(LEFTGUIDE, 1));       //partindex, routeindex

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

    }
}
