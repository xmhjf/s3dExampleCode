//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Utility_L_Wrapper.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.Utility_L_Wrapper
//   Author       :  Manikanth
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Manikanth   CR-CP-224492- Convert HS_FINL_Assy to C# .Net 
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
    public class Utility_L_Wrapper : CustomSupportDefinition
    {
        string L="L";
        private const string ROUTECONNOBJECT = "RouteConnObj_5378";
        public string SectionSize { get; set; }
        public double SectionL { get; set; }
        public Boolean Override { get; set; }
        public int Index { get; set; }
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //only if it is not by super assy
                    if (!Override)           //by default it is false
                    {
                        SectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSectionSize2", "SectionSize2")).PropValue;
                        SectionL = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLSectionL", "SectionL")).PropValue;
                    }

                    parts.Add(new PartInfo(L + "_" + Index, "Utility_GENERIC_L_1"));
                    parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));

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
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {

                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                double tWidth,tFlangeThickness,tDepth,webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "L", SectionSize, out tWidth, out tFlangeThickness, out webThickness, out tDepth);

                if (Override)
                {
                    string bomDesc = SectionSize + " , Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, SectionL, UnitName.DISTANCE_MILLIMETER);
                    (componentDictionary[L + "_" + Index]).SetPropertyValue(tWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    (componentDictionary[L + "_" + Index]).SetPropertyValue(tDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    (componentDictionary[L + "_" + Index]).SetPropertyValue(SectionL, "IJOAHgrUtility_GENERIC_L", "L");
                    (componentDictionary[L + "_" + Index]).SetPropertyValue(tFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    (componentDictionary[L + "_" + Index]).SetPropertyValue(bomDesc, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");
                }

                if (!Override)
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", L + "_" + Index, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

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

                    routeConnections.Add(new ConnectionInfo(ROUTECONNOBJECT + "_" + Index, 1)); // partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(ROUTECONNOBJECT + "_" + Index, 1)); // partindex, routeindex

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