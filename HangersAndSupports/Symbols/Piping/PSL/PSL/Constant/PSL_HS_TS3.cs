//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_HS_TS3.cs
//   PROGID : PSL,Ingr.SP3D.Content.Support.Symbols.PSL_HS_TS3
//   Author       :  Mahanth
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Mahanth    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net   
//   30-Dec-2014    PVK          TR-CP-264951	Resolve P3 coverity issues found in November 2014 report 
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
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_HS_TS3 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_HD_TS5"
        //----------------------------------------------------------------------------------
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
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                Double workingTravDown = WORKING_TRAV_DOWN.Value;
                Double workingTravUp = WORKING_TRAV_UP.Value;
                String size = SIZE.Value;
                Double minTravel = MIN_TRAVEL.Value;
                Double maxTravel = MAX_TRAVEL.Value;
                Double operatingLoad = OPER_LOAD.Value;
                
                double maxOperatingLoad = PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "MAX_LOAD", "IJUAHgrPSL_CON_LOAD", "SIZE", size);
                double minOperatingLoad = PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "MIN_LOAD", "IJUAHgrPSL_CON_LOAD", "SIZE", size);

                maxOperatingLoad = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Force, maxOperatingLoad, UnitName.FORCE_NEWTON);
                minOperatingLoad = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Force, minOperatingLoad, UnitName.FORCE_NEWTON);

                if ((operatingLoad < minOperatingLoad - 0.0001) || (operatingLoad > maxOperatingLoad + 0.0001))
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Operating load " + operatingLoad + " should be between " + minOperatingLoad + "  and  " + maxOperatingLoad);
                    operatingLoad = minOperatingLoad;
                }

                maxTravel = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, maxTravel, UnitName.DISTANCE_MILLIMETER);
                minTravel = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, minTravel, UnitName.DISTANCE_MILLIMETER);
                operatingLoad = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Force, operatingLoad, UnitName.FORCE_NEWTON);

                double moment = PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "WR_MAX", "IJUAHgrPSL_CON_LOAD", "SIZE", size);
                moment = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, moment, UnitName.DISTANCE_MILLIMETER);

                workingTravDown = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, workingTravDown, UnitName.DISTANCE_MILLIMETER);
                workingTravUp = MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, workingTravUp, UnitName.DISTANCE_MILLIMETER);

                double calculatedTotalTravel = 10 * (int)(moment * 100 / operatingLoad + 0.5);
                calculatedTotalTravel = Math.Round(calculatedTotalTravel / 1000, 2) * 1000;

                double totalActualTravel = workingTravDown + workingTravUp;
                if ((totalActualTravel + 25 <= calculatedTotalTravel) && (calculatedTotalTravel >= minTravel) && (calculatedTotalTravel <= maxTravel))
                {}
                else
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidSpringSize, "Selected spring size is not suitable for load and travel specified."));
                    return;
                }
                double totalTravel = calculatedTotalTravel;
                double setPosition = (totalTravel + workingTravUp - workingTravDown) / 2;
                double leverLength = totalTravel / (Math.Sin(35 * (Math.PI / 180)) + Math.Sin(10 * (Math.PI / 180)));
                double leverAngleOffset = (35 * (Math.PI / 180) - Math.Asin((leverLength * Math.Sin(35 * (Math.PI / 180)) - setPosition) / leverLength)) * 180 / Math.PI;
                setPosition = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, setPosition, UnitName.DISTANCE_MILLIMETER);
                leverLength = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, leverLength, UnitName.DISTANCE_MILLIMETER);
                if (HgrCompareDoubleService.cmpdbl(Convert.ToInt32(totalTravel / 10) * 10.00, totalTravel) == false)
                    totalTravel = Convert.ToInt32((totalTravel / 10) * 10.00 + 10.00);

                totalTravel = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, totalTravel, UnitName.DISTANCE_MILLIMETER);

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_CONSTANTS", "SIZE"), size.Trim());
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_CONSTANTS", "TOTAL_TRAV"), totalTravel);

                double rodDiameter = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "ROD_DIA", parameter);
                double sa = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "SA", parameter);
                double ta = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TA", parameter);
                double fa = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FA", parameter);
                double fe = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FE", parameter);
                double tc = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TC", parameter);
                double e = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "E", parameter);
                double fb = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FB", parameter);
                double tb = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TB", parameter);
                double aa = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "AA", parameter);
                double fc = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FC", parameter);
                double fp = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FP", parameter);
                double v1 = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "V1", parameter);
                double xn = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "XN", parameter);
                double ee = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "EE", parameter);
                double bb = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "BB", parameter);
                double cc = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "CC", parameter);
                double gg = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "GG", parameter);
                double j = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "A1", parameter);
                double a1 = j + setPosition;
                double d = leverLength * Math.Cos((35 - leverAngleOffset) * Math.PI / 180);
                double angle = 35 - leverAngleOffset;
                if (angle > 35)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidLeverAngle, "Lever angle cannot exceed 35 degrees. Please reset Actual Travel Up and Actual Travel Down with proper values."));
                }
                double fm = 0;
                double cylinderLength = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 250, UnitName.DISTANCE_MILLIMETER);
                double locZ = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 100, UnitName.DISTANCE_MILLIMETER);

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, d - v1, a1 + aa), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

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
                if (ta <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTAGTZ, "TA value should be greater than zero"));
                    return;
                }
                if (cc <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCCGTZ, "CC value should be greater than zero"));
                    return;
                }
                if (ee <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEEGTZ, "EE value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(sa , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidSANEZ, "SA value cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(fb , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFBNEZ, "FB value cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d turnB = symbolGeometryHelper.CreateCylinder(null, rodDiameter, cylinderLength);
                Matrix4X4 rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(0, 0, -locZ));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                turnB.Transform(rotateMatrix);
                m_Symbolic.Outputs["TURNB"] = turnB;

                locZ = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 50, UnitName.DISTANCE_MILLIMETER);
                double tempNumber = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 50, UnitName.DISTANCE_MILLIMETER);

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d rod = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, j - tb + e - tempNumber);
                rotateMatrix.SetIdentity();
                rotateMatrix.Translate(new Vector(0, 0, locZ));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                rod.Transform(rotateMatrix);
                m_Symbolic.Outputs["ROD"] = rod;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-ta, 0, -rodDiameter * 2.0);
                Projection3d arm = symbolGeometryHelper.CreateBox(null, ta, rodDiameter * 2.0, rodDiameter * 4.0, 9);
                rotateMatrix.SetIdentity();
                rotateMatrix.Rotate(angle * Math.PI / 180, new Vector(0, 1, 0));
                rotateMatrix.Translate(new Vector(d, -rodDiameter, a1 - tb));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arm.Transform(rotateMatrix);
                m_Symbolic.Outputs["ARM"] = arm;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, d + fc, a1 - xn));
                pointCollection.Add(new Position(-fa / 2.0, d + fc, a1 - xn - fe * 0.375));
                pointCollection.Add(new Position(-fa / 2.0, d + fc, a1 - xn - fe * 0.625));
                pointCollection.Add(new Position(0, d + fc, a1 - xn - fe));
                pointCollection.Add(new Position(fa / 2.0, d + fc, a1 - xn - fe * 0.625));
                pointCollection.Add(new Position(fa / 2.0, d + fc, a1 - xn - fe * 0.375));
                pointCollection.Add(new Position(0, d + fc, a1 - xn));
                Projection3d base1 = new Projection3d(new LineString3d(pointCollection), new Vector(0, 1, 0), sa, true);
                m_Symbolic.Outputs["BASE"] = base1;

                double location = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 10, UnitName.DISTANCE_MILLIMETER);
                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-fb / 2.0, d - tc, a1 - fm));
                pointCollection.Add(new Position(-fb / 2.0, d + fc, a1 - fm));
                pointCollection.Add(new Position(-fb / 2.0, d + fc, a1 - tb - fp));
                pointCollection.Add(new Position(-fb / 2.0, d - location, a1 - tb - fp));
                pointCollection.Add(new Position(-fb / 2.0, d - location, a1 - tb + location));
                pointCollection.Add(new Position(-fb / 2.0, d - tc, a1 - rodDiameter * 2.0));
                pointCollection.Add(new Position(-fb / 2.0, d - tc, a1 - fm));
                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), fb, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(d - v1 - cc / 2.0, -gg / 2.0 - ee, a1);
                Projection3d lug1 = symbolGeometryHelper.CreateBox(null, cc, ee, aa + bb, 9);
                rotateMatrix.SetIdentity();
                rotateMatrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                lug1.Transform(rotateMatrix);
                m_Symbolic.Outputs["LUG1"] = lug1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(d - v1 - cc / 2.0, gg / 2.0, a1);
                Projection3d lug2 = symbolGeometryHelper.CreateBox(null, cc, ee, aa + bb, 9);
                rotateMatrix.SetIdentity();
                rotateMatrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                lug2.Transform(rotateMatrix);
                m_Symbolic.Outputs["LUG2"] = lug2;
            }
            catch (Exception)
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_HS_TS3.cs"));
                return;
            }
        }

        #endregion
    }
}
