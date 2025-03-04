//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PSL_VS_TS3.cs
//   PSL,Ingr.SP3D.Content.Support.Symbols.PSL_VS_TS3
//   Author       :BS   
//   Creation Date:21.08.2013  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21.08.2013      BS      CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net  
//   30-Dec-2014    PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{

    /// <summary>
    /// Implementation of the PSL_VS_TS3 class.
    /// </summary>
    /// <remarks></remarks>
    [CacheOption(CacheOptionType.Cached)]
    public class PSL_VS_TS3 : HangerComponentSymbolDefinition
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_VS_TS3"

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "TOTAL_TRAV1", "TOTAL_TRAV1", 0.999999)]
        public InputDouble TOTAL_TRAV1;
        [InputDouble(3, "WORKING_TRAV_DOWN", "WORKING_TRAV_DOWN", 0.999999)]
        public InputDouble WORKING_TRAV_DOWN;
        [InputDouble(4, "WORKING_TRAV_UP", "WORKING_TRAV_UP", 0.999999)]
        public InputDouble WORKING_TRAV_UP;
        [InputString(5, "SIZE", "SIZE", "No Value")]
        public InputString SIZE;
        [InputDouble(6, "MIN_TRAVEL", "MIN_TRAVEL", 0.999999)]
        public InputDouble MIN_TRAVEL;
        [InputDouble(7, "MAX_TRAVEL", "MAX_TRAVEL", 0.999999)]
        public InputDouble MAX_TRAVEL;
        [InputDouble(8, "OPER_LOAD", "OPER_LOAD", 0.999999)]
        public InputDouble OPER_LOAD;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("TURNB", "TURNB")]
        [SymbolOutput("ROD", "ROD")]
        [SymbolOutput("ARM", "ARM")]
        [SymbolOutput("BASE", "BASE")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("LUG1", "LUG1")]
        [SymbolOutput("LUG2", "LUG2")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                string size = SIZE.Value;
                double minTravel = MIN_TRAVEL.Value;
                double maxTravel = MAX_TRAVEL.Value;
                double operLoad = OPER_LOAD.Value;

                double workingTraveldown = WORKING_TRAV_DOWN.Value;
                double workingTravelup = WORKING_TRAV_UP.Value;

                double maxOperatingLoad = PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "MAX_LOAD", "IJUAHgrPSL_CON_LOAD", "SIZE", size);
                double minOperatingLoad = PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "MIN_LOAD", "IJUAHgrPSL_CON_LOAD", "SIZE", size);

                maxOperatingLoad = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Force, maxOperatingLoad, UnitName.FORCE_NEWTON);
                minOperatingLoad = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Force, minOperatingLoad, UnitName.FORCE_NEWTON);

                if (operLoad < minOperatingLoad - 0.0001 || operLoad > maxOperatingLoad + 0.0001)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Operating Load should be between " + minOperatingLoad.ToString() + "N and " + maxOperatingLoad.ToString() + "N.");
                    operLoad = minOperatingLoad;
                }
                minTravel = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, minTravel, UnitName.DISTANCE_MILLIMETER);
                maxTravel = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, maxTravel, UnitName.DISTANCE_MILLIMETER);
                operLoad = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Force, operLoad, UnitName.FORCE_NEWTON);

                double moment = PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "WR_MAX", "IJUAHgrPSL_CON_LOAD", "SIZE", size);
                moment = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, moment, UnitName.DISTANCE_MILLIMETER);

                double workingTravelDownConv = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, workingTraveldown, UnitName.DISTANCE_MILLIMETER);
                double workingTravelUpConv = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, workingTravelup, UnitName.DISTANCE_MILLIMETER);

                double calculatedTotalTravel = 10 * (int)(moment * 100 / operLoad + 0.5);
                //--- make 301.xxx to 300 for example; otherwise it generates error
                calculatedTotalTravel = Math.Round(calculatedTotalTravel / 1000.0, 2) * 1000;

                double totalActualTravel = workingTravelDownConv + workingTravelUpConv;
                if ((totalActualTravel + 25 <= calculatedTotalTravel) && (calculatedTotalTravel >= minTravel) && (calculatedTotalTravel <= maxTravel))
                { }
                else
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidSpringSize, "Selected spring size is not suitable for load and travel specified."));
                    return;
                }
                double totalTravel1 = calculatedTotalTravel;
                double totalTravel = totalTravel1;
                const double PI = 3.14159265359;
                double setPosition = (totalTravel + workingTravelUpConv - workingTravelDownConv) / 2;
                double leverLength = totalTravel / (Math.Sin(35 * PI / 180) + Math.Sin(10 * PI / 180));
                double leverAngleOffset = (35 * PI / 180 - Math.Asin((leverLength * Math.Sin(35 * PI / 180) - setPosition) / leverLength)) * 180 / PI;
                setPosition = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, setPosition, UnitName.DISTANCE_MILLIMETER);
                leverLength = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, leverLength, UnitName.DISTANCE_MILLIMETER);

                if (HgrCompareDoubleService.cmpdbl((int)(totalTravel1 / 10) * 10.0, totalTravel1) == false)
                    totalTravel = (int)(totalTravel1 / 10) * 10.0 + 10.0;

                totalTravel = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, totalTravel, UnitName.DISTANCE_MILLIMETER);

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_CONSTANTS", "SIZE"), size.Trim());
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_CONSTANTS", "TOTAL_TRAV"), totalTravel);

                double E = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "E", parameter);
                double rodDiameter = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "ROD_DIA", parameter);
                double FA = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FA", parameter);
                double FE = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FE", parameter);
                double FB = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FB", parameter);
                double FC = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FC", parameter);
                double TA = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TA", parameter);
                double J = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "A2", parameter);
                double SE = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "SE", parameter);
                double SA = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "SA", parameter);
                double fd = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FD", parameter);
                double TF = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TF", parameter);
                double TE = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TE", parameter);
                double AA = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "AA", parameter);
                double FG = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FG", parameter);
                double U = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "U", parameter);
                double V2 = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "V2", parameter);
                double XN = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "XN", parameter);
                double EE = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "EE", parameter);
                double BB = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "BB", parameter);
                double CC = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "CC", parameter);
                double DD = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "DD", parameter);
                double GG = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "GG", parameter);
                double FK = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FK", parameter);
                double FJ = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FJ", parameter);

                double A2 = J + setPosition;
                double D = leverLength * Math.Cos((35 - leverAngleOffset) * PI / 180);
                double C = leverLength;
                double angle = 35 - leverAngleOffset;

                if (angle > 35)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidLeverAngle, "Lever angle cannot exceed 35 degrees. Please reset Actual Travel Up and Actual Travel Down with proper values."));

                double FM = 0;
                double hypo = Math.Sqrt(fd * fd + FC * FC);
                double angle2 = (85) - Math.Atan(fd / FC) * 180.0 / Math.PI;
                double baseX = Math.Cos(angle2 * Math.PI / 180.0) * hypo;
                double baseY = Math.Sin(angle2 * Math.PI / 180.0) * hypo;

                double cylinderLength = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 250, UnitName.DISTANCE_MILLIMETER);
                double locationZ = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 100, UnitName.DISTANCE_MILLIMETER);

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, D - V2, A2 + AA), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(cylinderLength , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCylinderLength, "CYL_LEN cannot be zero"));
                    return;
                }
                if (TA <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTAGTZ, "TA value should be greater than zero"));
                    return;
                }
                if (CC <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCCGTZ, "CC value should be greater than zero"));
                    return;
                }
                if (EE <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEEGTZ, "EE value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(SA , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidSANEZ, "SA value cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(FB , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFBNEZ, "FB value cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -locationZ);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d turnB = symbolGeometryHelper.CreateCylinder(null, rodDiameter, cylinderLength);
                m_Symbolic.Outputs["TURNB"] = turnB;

                locationZ = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 50, UnitName.DISTANCE_MILLIMETER);
                double tempNumber = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 50, UnitName.DISTANCE_MILLIMETER);

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, locationZ);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rod = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2, J - TE + E - tempNumber);
                m_Symbolic.Outputs["ROD"] = rod;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d arm = symbolGeometryHelper.CreateBox(null, TA, rodDiameter * 2.0, rodDiameter * 4, 9);
                Matrix4X4 matrix = new Matrix4X4();
                matrix.Translate(new Vector(-TA, 0, -rodDiameter * 2));
                matrix.Rotate(angle * Math.PI / 180.0, new Vector(0, 1, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(D, -rodDiameter, A2 - TE));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arm.Transform(matrix);
                m_Symbolic.Outputs["ARM"] = arm;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, D + baseX + FE / 2.0, A2 - TE - FC));
                pointCollection.Add(new Position(-FA / 2.0, D + baseX + FE * 0.125, A2 - TE - FC));
                pointCollection.Add(new Position(-FA / 2.0, D + baseX - FE * 0.125, A2 - TE - FC));
                pointCollection.Add(new Position(0, D + baseX - FE / 2.0, A2 - TE - FC));
                pointCollection.Add(new Position(FA / 2.0, D + baseX - FE * 0.125, A2 - TE - FC));
                pointCollection.Add(new Position(FA / 2.0, D + baseX + FE * 0.125, A2 - TE - FC));
                pointCollection.Add(new Position(0, D + baseX + FE / 2.0, A2 - TE - FC));
                Projection3d baseProjection = new Projection3d(new LineString3d(pointCollection), new Vector(0, 0, -1), SA, true);
                m_Symbolic.Outputs["BASE"] = baseProjection;

                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-FB / 2.0, D - TF, A2 - FM));
                pointCollection.Add(new Position(-FB / 2, D + FC, A2 - FM));
                pointCollection.Add(new Position(-FB / 2, D + FE / 2 + fd, A2 - TE));
                pointCollection.Add(new Position(-FB / 2, D + FE / 2 + fd, A2 - TE - FC));
                double locationY = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 10, UnitName.DISTANCE_MILLIMETER);
                locationZ = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 10, UnitName.DISTANCE_MILLIMETER);

                pointCollection.Add(new Position(-FB / 2.0, D - locationY, A2 - TE - FC));
                pointCollection.Add(new Position(-FB / 2.0, D - locationY, A2 - TE + locationZ));
                pointCollection.Add(new Position(-FB / 2.0, D - TF, A2 - rodDiameter * 2));
                pointCollection.Add(new Position(-FB / 2.0, D - TF, A2 - FM));
                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), FB, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d lug1 = symbolGeometryHelper.CreateBox(null, CC, EE, AA + BB, 9);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(D - V2 - CC / 2.0, -GG / 2.0 - EE, A2));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                lug1.Transform(matrix);
                m_Symbolic.Outputs["LUG1"] = lug1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d lug2 = symbolGeometryHelper.CreateBox(null, CC, EE, AA + BB, 9);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(D - V2 - CC / 2.0, GG / 2.0, A2));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                lug2.Transform(matrix);
                m_Symbolic.Outputs["LUG2"] = lug2;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_VS_TS3"));
                    return;
                }
            }
        }

        #endregion
    }
}
