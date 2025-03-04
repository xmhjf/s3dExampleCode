//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeE.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeE
//   Author       : Hema
//   Creation Date: 17-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Sep-2013    Hema     CR-CP-224494 Converted Generic_Assy to C# .Net 
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

    public class UFrameTypeE : CustomSupportDefinition
    {
        private const string CURVEDPLATE = "CURVEDPLATE";
        private const string PLATE = "PLATE";

        int supportLocation;
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

                    // Get the attributes from assembly
                    supportLocation = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySupLoc", "SupLocation")).PropValue;

                    parts.Add(new PartInfo(CURVEDPLATE, "HgrGen_CurvedPlate_1"));
                    parts.Add(new PartInfo(PLATE, "HgrGen_TwoHolePlate_1"));

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
                //Get Pipe OD without insulation
                double pipeDiameter = pipeInfo.OutsideDiameter;

                //Get the Nominal Diameter of the Pipe
                NominalDiameter nominalDiameter = new NominalDiameter();
                nominalDiameter = pipeInfo.NominalDiameter;

                //Get bracket dimensions
                double W, L, L1, H, T, T1, D, H1;
                W = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "W", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                L = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "L", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                L1 = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "L1", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                H = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "H", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                H1 = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "H1", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                T = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "T", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                T1 = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "T1", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                D = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeE_Dim", "IJUAHgrGenSrvTypeEDim", "D", "IJUAHgrGenSrvTypeEDim", "PipeNPD", nominalDiameter.Size - 0.001, nominalDiameter.Size + 0.001);
                // '==================================
                //Set properties on part occurrences
                //==================================

                Double angleInDegrees, angleInRadians;

                angleInDegrees = 90;
                angleInRadians = angleInDegrees / 180 * Math.PI;

                string plate1Bom = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, T1, UnitName.DISTANCE_MILLIMETER) + "Plate Steel," + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER) + "X" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, H, UnitName.DISTANCE_MILLIMETER);

                componentDictionary[CURVEDPLATE].SetPropertyValue(angleInRadians, "IJOAHgrGenCurvedPlate", "Angle");
                componentDictionary[CURVEDPLATE].SetPropertyValue(T, "IJOAHgrGenCurvedPlate", "T");
                componentDictionary[CURVEDPLATE].SetPropertyValue(pipeDiameter / 2, "IJOAHgrGenCurvedPlate", "Radius");
                componentDictionary[CURVEDPLATE].SetPropertyValue(W, "IJOAHgrGenCurvedPlate", "W");

                componentDictionary[PLATE].SetPropertyValue(D, "IJOAHgrGenTwoHolePlate", "D");
                componentDictionary[PLATE].SetPropertyValue(L, "IJOAHgrGenTwoHolePlate", "L");
                componentDictionary[PLATE].SetPropertyValue(H, "IJOAHgrGenTwoHolePlate", "H");
                componentDictionary[PLATE].SetPropertyValue(H1, "IJOAHgrGenTwoHolePlate", "H1");
                componentDictionary[PLATE].SetPropertyValue(L1, "IJOAHgrGenTwoHolePlate", "L1");
                componentDictionary[PLATE].SetPropertyValue(T1, "IJOAHgrGenTwoHolePlate", "T1");
                componentDictionary[PLATE].SetPropertyValue(plate1Bom, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");

                //  '=============
                //Create Joints
                //=============
                if ((SupportHelper.SupportingObjects.Count == 0))
                    //Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    //Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                //Add Joint Between the Curved  Plate and Plate
                JointHelper.CreateRigidJoint(CURVEDPLATE, "Route", PLATE, "TopStructure", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -pipeDiameter / 2 - H / 2 - T, T1 / 2, 0);
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

                    routeConnections.Add(new ConnectionInfo(CURVEDPLATE, 1)); //partindex, routeindex

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

                    //We are not connecting to any structure so we have nothing to return

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
