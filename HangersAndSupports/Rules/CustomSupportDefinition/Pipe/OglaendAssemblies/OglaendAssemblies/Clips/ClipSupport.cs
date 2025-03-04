//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ClipSupport.cs
//   OglaendAssemblies,Ingr.SP3D.Content.Support.Rules.ClipSupport
//   Author       :  PVK
//   Creation Date:  23-01-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET  

//
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

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
    public class ClipSupport : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string CLIP_1 = "CLIP_1";
        private const string CLIP_2 = "CLIP_2";

        Boolean isClip1, isClip2;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    string strRule = "";
                    BusinessObject ruleBO = null;

                    // Add Clip part
                    string Clip1Part = (string)((PropertyValueString)part.GetPropertyValue("IJOAhsOGClipPart1", "ClipPart1")).PropValue;
                    isClip1 = OglaendAssemblyServices.AddPart(this, CLIP_1, Clip1Part, strRule, parts, ruleBO);


                    string Clip22Part = (string)((PropertyValueString)part.GetPropertyValue("IJOAhsOGClipPart2", "ClipPart2")).PropValue;
                    isClip2 = OglaendAssemblyServices.AddPart(this, CLIP_2, Clip22Part, strRule, parts, ruleBO);

                    // Return the collection of Catalog Parts
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
        //Get Assembly Implied Parts
        //-----------------------------------------------------------------------------------
        public override ReadOnlyCollection<PartInfo> ImpliedParts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> impliedParts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    string supportpartnumber = support.SupportDefinition.PartNumber;
                    OglaendAssemblyServices.AddImpliedPartbyInterface(this, impliedParts, "IJUAhsFrImpPart", supportpartnumber, 10005);
                    ReadOnlyCollection<PartInfo> rImpliedParts = new ReadOnlyCollection<PartInfo>(impliedParts);
                    return rImpliedParts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Implied Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                SP3D.Support.Middle.Support support = SupportHelper.Support;
                double XOffset = 0, YOffset = 0, ClipGap = 0;
                double ClipAngle = 0;
                double PortAngle = 0;
                PortAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                PortAngle = 180 * PortAngle / Math.PI;
                ClipGap = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsOGClipGap", "ClipGap")).PropValue;

                if (HgrCompareDoubleService.cmpdbl(Math.Round(PortAngle), 90)==true || HgrCompareDoubleService.cmpdbl(Math.Round(PortAngle),270)==true)
                {
                    ClipAngle = 0;
                    YOffset = 0;
                    XOffset = ClipGap / 2;
                }
                else
                {
                    ClipAngle = Math.PI / 2;
                    YOffset = ClipGap / 2;
                    XOffset = 0;
                }                

                JointHelper.CreateAngularRigidJoint(CLIP_1, "Steel", "-1", "Structure", new Vector(XOffset, YOffset, 0), new Vector(0, 0, ClipAngle + Math.PI / 2));
                JointHelper.CreateAngularRigidJoint(CLIP_2, "Steel", "-1", "Structure", new Vector(-XOffset, -YOffset, 0), new Vector(0, 0, ClipAngle + 3 * (Math.PI / 2)));


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
                    ////Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    //int numberOfRoutes = SupportHelper.SupportedObjects.Count;
                    //for (int routeIndex = 1; routeIndex <= numberOfRoutes; routeIndex++)
                    //{
                    //    routeConnections.Add(new ConnectionInfo(SECTION, routeIndex)); // partindex, routeindex     
                    //}

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

                    structConnections.Add(new ConnectionInfo(CLIP_1, 1)); // partindex, routeindex


                    structConnections.Add(new ConnectionInfo(CLIP_2, 1)); // partindex, routeindex

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

        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDesciption = "";
            try
            {
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject catalogpart = SupportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                string supportNumber = (string)((PropertyValueString)catalogpart.GetPropertyValue("IJUAhsSupportNumber", "SupportNumber")).PropValue;
                bomDesciption = supportNumber;

                return bomDesciption;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
