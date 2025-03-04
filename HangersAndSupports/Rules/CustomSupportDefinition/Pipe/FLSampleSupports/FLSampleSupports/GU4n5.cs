//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   GU4n5.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.GU4n5
//   Author       : Hema
//   Creation Date: 29-Aug-013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    29-Aug-2013  Hema CR-CP-224478 Convert FlSample_Supports to C# .Net 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;

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

    public class GU4n5 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string GUIDE1 = "GUIDE1";
        private const string GUIDE2 = "GUIDE2";
        private const string LOGOBJ1 = "LOGOBJ1";
        private const string LOGOBJ11 = "LOGOBJ11";
        private const string LOGOBJ2 = "LOGOBJ2";
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

                    string guide1 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Guide", "Guide")).PropValue;
                    string guide2 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Guide", "Guide")).PropValue;

                    parts.Add(new PartInfo(GUIDE1, guide1));
                    parts.Add(new PartInfo(GUIDE2, guide2));
                    parts.Add(new PartInfo(LOGOBJ1, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(LOGOBJ2, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(LOGOBJ11, "Log_Conn_Part_1"));

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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                Double guideLength, guideWidth, guideDepth, guideThickness, shoeHeight;

                guideLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Guide", "GuideL")).PropValue;
                guideWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Guide", "GuideW")).PropValue;
                guideDepth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Guide", "GuideD")).PropValue;
                guideThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Guide", "GuideT")).PropValue;

                componentDictionary[GUIDE1].SetPropertyValue(guideLength, "IJOAHgrGuide", "L");
                componentDictionary[GUIDE1].SetPropertyValue(guideWidth, "IJOAHgrGuide", "W");
                componentDictionary[GUIDE1].SetPropertyValue(guideDepth, "IJOAHgrGuide", "Depth");
                componentDictionary[GUIDE1].SetPropertyValue(guideThickness, "IJOAHgrGuide", "T");
                componentDictionary[GUIDE1].SetPropertyValue("280x100x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                componentDictionary[GUIDE2].SetPropertyValue(guideLength, "IJOAHgrGuide", "L");
                componentDictionary[GUIDE2].SetPropertyValue(guideWidth, "IJOAHgrGuide", "W");
                componentDictionary[GUIDE2].SetPropertyValue(guideDepth, "IJOAHgrGuide", "Depth");
                componentDictionary[GUIDE2].SetPropertyValue(guideThickness, "IJOAHgrGuide", "T");
                componentDictionary[GUIDE2].SetPropertyValue("280x100x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrFl_ShoeH", "ShoeH")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Double ndpPipeDiameterInch = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);
                if (ndpPipeDiameterInch < 20 || ndpPipeDiameterInch > 36)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: GU4n5 can only be placed from 20in to 36in pipe.", "", "GU4n5", 107);

                Double ndpPipeDiameter = pipeInfo.OutsideDiameter;
                ndpPipeDiameter = ndpPipeDiameter / 2;
                if (pipeInfo.InsulationThickness > 0)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: Not placable for Insulated pipe.", "", "GU4n5", 112);

                Double verticalDIS = ndpPipeDiameter;
                Double x = ((ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 + (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 + (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 + (verticalDIS - (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 - guideWidth * Math.Sqrt(2) / 2) * (verticalDIS - (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 - guideWidth * Math.Sqrt(2) / 2) - ((guideWidth + guideThickness) * Math.Sqrt(2) / 2 - verticalDIS) * ((guideWidth + guideThickness) * Math.Sqrt(2) / 2 - verticalDIS)) / (2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2);
                
                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (supportingType == "Slab")
                {
                    JointHelper.CreateRigidJoint("-1", "Structure", GUIDE1, "Neutral", Plane.XY, Plane.ZX, Axis.X, Axis.X, (guideWidth + guideThickness) * (Math.Sqrt(2) / 2), x, 0);
                    JointHelper.CreateRigidJoint("-1", "Structure", GUIDE2, "Neutral", Plane.XY, Plane.ZX, Axis.X, Axis.X, (guideWidth + guideThickness) * (Math.Sqrt(2) / 2), -x, 0);
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "Structure", GUIDE1, "Neutral", Plane.XY, Plane.ZX, Axis.X, Axis.Z, (guideWidth + guideThickness) * (Math.Sqrt(2) / 2), 0, x);
                    JointHelper.CreateRigidJoint("-1", "Structure", GUIDE2, "Neutral", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, (guideWidth + guideThickness) * (Math.Sqrt(2) / 2), 0, -x);
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

                    routeConnections.Add(new ConnectionInfo(GUIDE1, 1));// partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(GUIDE2, 1));// partindex, Structureindex    

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
        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = string.Empty;
            try
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Double pipeDiameterInch = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);
                if (pipeDiameterInch >= 20 && pipeDiameterInch <= 30)
                    bomDescription = "GU4";
                else
                    bomDescription = "GU5";
                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
