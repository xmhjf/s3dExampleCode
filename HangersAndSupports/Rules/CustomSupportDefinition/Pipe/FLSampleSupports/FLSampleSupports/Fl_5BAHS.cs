//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Fl_5BSHS.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.Fl_5BSHS
//   Author       : Hema
//   Creation Date: 29-Aug-013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29-Aug-2013  Hema CR-CP-224478 Convert FlSample_Supports to C# .Net   
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Exceptions;

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

    public class Fl_5BAHS : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PLATE1 = "PLATE1";
        private const string PLATE2 = "PLATE2";
        private const string PLATEBASE = "PLATEBASE";
        private const string FOOTING = "FOOTING";
        private const string PLATE3 = "PLATE3";
        Double pipeDiameterInch;
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
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    pipeDiameterInch = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                    string plate1 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate_1")).PropValue;
                    string plate2 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate_2")).PropValue;

                    if (3 <= pipeDiameterInch && pipeDiameterInch <= 6)
                    {
                        parts.Add(new PartInfo(PLATE1, plate1));
                        parts.Add(new PartInfo(PLATE2, plate1));
                        parts.Add(new PartInfo(PLATEBASE, plate2));
                    }
                    else if (8 <= pipeDiameterInch && pipeDiameterInch <= 20)
                    {
                        parts.Add(new PartInfo(PLATE1, plate1));
                        parts.Add(new PartInfo(PLATE2, plate1));
                        parts.Add(new PartInfo(PLATEBASE, plate2));
                        parts.Add(new PartInfo(PLATE3, plate1));
                    }
                    else
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR:" + "5BAHS can only be placed on 3in to 20in pipes.", "", "Fl_5BAHS", 74);
                        return null;
                    }
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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double ndpPipeDiameter = pipeInfo.OutsideDiameter;

                Double plate1L, plate1H, plate1T, plate2L, plate2H, plate2T, height;

                plate1L = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_1")).PropValue;
                plate1H = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_1")).PropValue;
                plate1T = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_1")).PropValue;
                plate2L = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_2")).PropValue;
                plate2H = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_2")).PropValue;
                plate2T = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_2")).PropValue;
                height = (Double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrPipeHieght", "PipeHieght")).PropValue;

                int pipeMaterialCodelist = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrPipeMaterial", "PipeMaterial")).PropValue;
                string plate1BOM = string.Empty, plate2BOM = string.Empty;

                PropertyValueCodelist bomList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrPipeMatPlateXY_5BAHS", "PlateXY_5BAHS");
                plate1BOM = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).DisplayName;
                plate2BOM = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).ShortDisplayName;

                Double longElbow = pipeDiameterInch * 1.5 * 0.0254;
                Double plate12L = 0;
                try
                {
                    if (RefPortHelper.PortLCS("TurnRef") != null)
                    {
                        plate12L = longElbow + ndpPipeDiameter / 2 - 0.04;
                        plate12L = plate12L + height + pipeInfo.InsulationThickness;
                        plate12L = plate12L - 0.0254 / 2;
                    }
                }
                catch (SupportInvalidArgumentException)
                {
                    plate12L = ndpPipeDiameter / 2 - 0.04; //the 0.045 is to adjust the plates so they dont dig into the pipe so much
                    plate12L = plate12L + height + pipeInfo.InsulationThickness;
                    plate12L = plate12L - 0.0254 / 2;
                }

                IPart part = support.SupportDefinition;
                string partNumber = part.PartNumber;
                if (pipeDiameterInch <= 6 && partNumber == "5BAHS_2")
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "For 5BAHS_2, Pipe Size should be at least 8in.", "", "Fl_5BAHS", 140);
                    return;
                }
                if (pipeDiameterInch >= 8 && partNumber == "5BAHS_1")
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "For 5BAHS_1, Pipe Size should be at least 6in.", "", "Fl_5BAHS", 145);
                    return;
                }
                if (3 <= pipeDiameterInch && pipeDiameterInch <= 6)
                {
                    componentDictionary[PLATE1].SetPropertyValue(plate1T, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE1].SetPropertyValue(ndpPipeDiameter, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE1].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE1].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATE2].SetPropertyValue(plate1T, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE2].SetPropertyValue(ndpPipeDiameter, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE2].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE2].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATEBASE].SetPropertyValue(plate2L, "IJOAHgrUtility_FourHolePl", "Depth");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2H, "IJOAHgrUtility_FourHolePl", "Width");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2T, "IJOAHgrUtility_FourHolePl", "T");
                    componentDictionary[PLATEBASE].SetPropertyValue(0.03, "IJOAHgrUtility_FourHolePl", "C");

                    componentDictionary[PLATEBASE].SetPropertyValue(0.024, "IJOAHgrUtility_FourHolePl", "HoleSize");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate1BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    double temp = Math.Sqrt((longElbow + ndpPipeDiameter / 2) * (longElbow + ndpPipeDiameter / 2) - (0.04) * (0.04));
                    try
                    {
                        if (RefPortHelper.PortLCS("TurnRef") != null)
                        {
                            JointHelper.CreateRigidJoint("-1", "TurnRef", PLATE1, "TopStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, longElbow - 0.04, 0, (temp - longElbow - ndpPipeDiameter / 2));
                            JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                            JointHelper.CreateRigidJoint(PLATE1, "BotStructure", PLATEBASE, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }
                    }
                    catch (SupportInvalidArgumentException)
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0.04, 0, 0);
                        JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                        JointHelper.CreateRigidJoint(PLATE1, "BotStructure", PLATEBASE, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                }
                else if (8 <= pipeDiameterInch && pipeDiameterInch <= 20)
                {
                    componentDictionary[PLATE1].SetPropertyValue(plate1T, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE1].SetPropertyValue(ndpPipeDiameter / 2, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE1].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE1].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATE2].SetPropertyValue(plate1T, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE2].SetPropertyValue(ndpPipeDiameter / 2, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE2].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE2].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATEBASE].SetPropertyValue(plate2L, "IJOAHgrUtility_FourHolePl", "Depth");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2H, "IJOAHgrUtility_FourHolePl", "Width");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2T, "IJOAHgrUtility_FourHolePl", "T");
                    componentDictionary[PLATEBASE].SetPropertyValue(0.03, "IJOAHgrUtility_FourHolePl", "C");

                    componentDictionary[PLATEBASE].SetPropertyValue(0.024, "IJOAHgrUtility_FourHolePl", "HoleSize");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate1BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    double temp = Math.Sqrt((longElbow + ndpPipeDiameter / 2) * (longElbow + ndpPipeDiameter / 2) - (ndpPipeDiameter / 4 + longElbow) * (ndpPipeDiameter / 4 + longElbow));

                    double plate3L = 0;
                    try
                    {
                        if (RefPortHelper.PortLCS("TurnRef") != null)
                        {
                            plate3L = longElbow + ndpPipeDiameter / 2 - Math.Sqrt((longElbow + ndpPipeDiameter / 2) * (longElbow + ndpPipeDiameter / 2) - (longElbow + ndpPipeDiameter / 2 - ndpPipeDiameter / 4) * (longElbow + ndpPipeDiameter / 2 - ndpPipeDiameter / 4));
                            plate3L = plate3L + height + pipeInfo.InsulationThickness;
                            plate3L = plate3L - 0.0254 / 2;
                        }
                    }
                    catch (SupportInvalidArgumentException)
                    {
                        plate3L = plate12L - 0.045;
                        componentDictionary[PLATE2].SetPropertyValue(plate12L - 0.045, "IJOAHgrUtility_PLATE", "THICKNESS");
                        componentDictionary[PLATE1].SetPropertyValue(plate12L - 0.045, "IJOAHgrUtility_PLATE", "THICKNESS");
                    }
                    try
                    {
                        if (RefPortHelper.PortLCS("TurnRef") != null)
                        {
                            JointHelper.CreateRigidJoint("-1", "TurnRef", PLATE1, "TopStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, ((longElbow - temp) + (plate12L - plate3L)), -(plate1T + 0.13 - plate1T * 2) / 2, ndpPipeDiameter / 4);
                            JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -(0.13 - plate1T));
                            JointHelper.CreateRigidJoint(PLATE1, "BotStructure", PLATEBASE, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -(0.13 - plate1T) / 2);
                            JointHelper.CreateRigidJoint(PLATEBASE, "TopStructure", PLATE3, "BotStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                        }
                    }
                    catch (SupportInvalidArgumentException)
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0.04 + 0.045, (plate1T + 0.13 - plate1T * 2) / 2, 0);
                        JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -(0.13 - plate1T));
                        JointHelper.CreateRigidJoint(PLATE1, "BotStructure", PLATEBASE, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -(0.13 - plate1T) / 2);
                        JointHelper.CreateRigidJoint(PLATEBASE, "TopStructure", PLATE3, "BotStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                    }

                    componentDictionary[PLATE3].SetPropertyValue(plate1T, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE3].SetPropertyValue(0.13 - plate1T * 2, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE3].SetPropertyValue(plate3L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE3].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");
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

                    routeConnections.Add(new ConnectionInfo(PLATE2, 1)); // partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(PLATE2, 1)); // partindex, Structureindex    

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
                if (3 <= pipeDiameterInch && pipeDiameterInch <= 6)
                    bomDescription = "5BAHS1";
                else
                    bomDescription = "5BAHS2";
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
