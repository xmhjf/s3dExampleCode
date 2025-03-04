//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RefPorts.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RefPorts
//   Author       :Vijaya
//   Creation Date:08.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  08.Apr.2013     Vijaya   CR-CP-224484-Initial Creation
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;

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

    public class Assy_RefPorts : CustomSupportDefinition
    {
        //Constants
        private string[] routePartkeys;
        private string[] structPartkeys;

        int numOfRoutes = 0,numOfStructs = 0,index = 0;
      
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    numOfRoutes = SupportHelper.SupportedObjects.Count;
                    numOfStructs = SupportHelper.SupportingObjects.Count;
                    routePartkeys = new string[numOfRoutes];
                    structPartkeys = new string[numOfStructs];

                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    for (index = 1; index <= numOfRoutes; index++)
                    {
                        routePartkeys[index - 1] = "UtilityPart" + index;
                        parts.Add(new PartInfo(routePartkeys[index - 1], "Utility_GROUT_1"));
                    }
                    for (index = 1; index <= numOfStructs; index++)
                    {
                        structPartkeys[index - 1] = "UtilitySquarePart" + index;
                        parts.Add(new PartInfo(structPartkeys[index - 1], "Utility_SQUARE_GROUT_1"));
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
              
                for (index = 0; index < numOfRoutes; index++)
                {                   
                    componentDictionary[routePartkeys[index]].SetPropertyValue(0.15, "IJOAHgrUtility_GROUT", "T");
                    componentDictionary[routePartkeys[index]].SetPropertyValue(0.025, "IJOAHgrUtility_GROUT", "DIA");
                }

                for (index = 0; index < numOfStructs; index++)
                {                    
                    componentDictionary[structPartkeys[index]].SetPropertyValue(0.15, "IJOAHgrUtility_SQUARE_GROUT", "T");
                    componentDictionary[structPartkeys[index]].SetPropertyValue(0.025, "IJOAHgrUtility_SQUARE_GROUT", "W");
                    componentDictionary[structPartkeys[index]].SetPropertyValue(0.025, "IJOAHgrUtility_SQUARE_GROUT", "L");
                }

                //Create Joints
                string routePort = string.Empty,structPort=string.Empty;

                int setRefPort = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrRefPorts", "SetRefPorts")).PropValue;

                if (setRefPort == 1) //Default
                {
                    for (index = 1; index <= numOfRoutes; index++)
                    {
                        if (index == 1)
                            routePort = "Route";
                        else
                            routePort = "Route_" + index;
                        JointHelper.CreateRigidJoint(routePartkeys[index - 1], "Other", "-1", routePort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);              

                    }
                    for (index = 1; index <= numOfStructs; index++)
                    {
                        if (index == 1)
                            structPort = "Structure";
                        else
                            structPort = "Struct_" + index;

                        JointHelper.CreateRigidJoint(structPartkeys[index - 1], "Other", "-1", structPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);          

                    }
                }
                else //Alternate
                {
                    for (index = 1; index <= numOfRoutes; index++)
                    {
                        if (index == 1)
                            routePort = "RouteAlt";
                        else
                            routePort = "RouteAlt_" + index;

                        JointHelper.CreateRigidJoint(routePartkeys[index - 1], "Other", "-1", routePort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                    for (index = 1; index < numOfStructs; index++)
                    {
                        if (index == 1)
                            structPort = "StructAlt";
                        else
                            structPort = "StructAlt_" + index;

                        JointHelper.CreateRigidJoint(structPartkeys[index - 1], "Other", "-1", structPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
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
                    routeConnections.Add(new ConnectionInfo(routePartkeys[0], 1)); // partindex, routeindex

                    //Return the collection of Route connection information
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
                    if (numOfStructs>0)
                        structConnections.Add(new ConnectionInfo(structPartkeys[0], 1)); // partindex, routeindex

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

