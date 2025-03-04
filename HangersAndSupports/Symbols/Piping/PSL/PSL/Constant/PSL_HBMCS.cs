//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_HBMCS.cs
//   PROGID : PSL,Ingr.SP3D.Content.Support.Symbols.PSL_HBMCS
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
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.Generic;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_HBMCS : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_HBMCS"
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
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("FLANGE", "FLANGE")]
        [SymbolOutput("BAR", "BAR")]
        [SymbolOutput("ARM1", "ARM1")]
        [SymbolOutput("ARM2", "ARM2")]
        [SymbolOutput("BASE", "BASE")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("PLATE", "PLATE")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                Double workingTravUp = WORKING_TRAV_UP.Value;
                Double workingTravDown = WORKING_TRAV_DOWN.Value;
                String size = SIZE.Value;
                Double maxTravel = MAX_TRAVEL.Value;
                Double minTravel = MIN_TRAVEL.Value;
                Double operatingLoad = OPER_LOAD.Value;
                double maxOperatingLoad = (double)PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "MAX_LOAD", "IJUAHgrPSL_CON_LOAD", "SIZE", size.Trim());
                double minOperatingLoad = (double)PSLSymbolServices.GetDataByCondition("PSL_CON_LOAD", "IJUAHgrPSL_CON_LOAD", "MIN_LOAD", "IJUAHgrPSL_CON_LOAD", "SIZE", size.Trim());

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
                calculatedTotalTravel = Math.Round((calculatedTotalTravel / 1000) * 1000, 2);
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

                double fs = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FS", parameter);
                double fc = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FC", parameter);
                double c = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "C", parameter);
                double j = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "J4", parameter);
                double e = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "E", parameter);
                double fa = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FA", parameter);
                double fe = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FE", parameter);
                double fd = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FD", parameter);
                double td = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TD", parameter);
                double fb = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FB", parameter);
                double ta = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "TA", parameter);
                double sa = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "SA", parameter);
                double ff = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FF", parameter);
                double fq = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FQ", parameter);
                double xp = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "XP", parameter);
                double fr = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_CONSTANTS", "IJUAHgrPSL_CONSTANTS", "FR", parameter);

                double j4 = j - setPosition;
                double d = leverLength * Math.Cos((35 - leverAngleOffset) * Math.PI / 180);
                double angle = 35 - leverAngleOffset;
                if (angle > 35)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidLeverAngle, "Lever angle cannot exceed 35 degrees. Please reset Actual Travel Up and Actual Travel Down with proper values."));
                }
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, d + fc - fs / 2.0, -j4), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (fa <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFAGTZ, "FA value should be greater than zero"));
                    return;
                }
                if (fq <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFQGTZ, "FQ value should be greater than zero"));
                    return;
                }
                if (fr <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFRGTZ, "FR value should be greater than zero"));
                    return;
                }
                if (fs <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFSGTZ, "FS value should be greater than zero"));
                    return;
                }
                if (ff <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFFGTZ, "FF value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(sa , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidSANEZ, " SA value cannot be  zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(fb , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFBNEZ, " FB value cannot be  Zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-fq / 2.0, -fq / 2.0, -ff);
                Projection3d flange = symbolGeometryHelper.CreateBox(null, fq, fq, ff, 9);
                Matrix4X4 rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                flange.Transform(rotateMatrix);
                m_Symbolic.Outputs["FLANGE"] = flange;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-fq / 4.0, -fq / 4.0, -ff - (j - td - e - ff));
                Projection3d bar = symbolGeometryHelper.CreateBox(null, fq / 2.0, fq / 2.0, j - td - e - ff, 9);
                rotateMatrix.SetIdentity();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                bar.Transform(rotateMatrix);
                m_Symbolic.Outputs["BAR"] = bar;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-c, 0, -fq / 8);
                Projection3d arm1 = symbolGeometryHelper.CreateBox(null, c, fq / 4.0, fq / 4.0, 9);
                rotateMatrix.SetIdentity();
                rotateMatrix.Rotate(angle * Math.PI / 180, new Vector(0, 1, 0));
                rotateMatrix.Translate(new Vector(d, -fq / 8, -(j4 - td - xp) - ff));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arm1.Transform(rotateMatrix);
                m_Symbolic.Outputs["ARM1"] = arm1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-ta, 0, -fq / 4.0);
                Projection3d arm2 = symbolGeometryHelper.CreateBox(null, ta, fq / 2.0, fq / 2.0, 9);
                rotateMatrix.SetIdentity();
                rotateMatrix.Rotate(angle * Math.PI / 180, new Vector(0, 1, 0));
                rotateMatrix.Translate(new Vector(d, -fq / 4.0, -(j4 - td) - ff));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arm2.Transform(rotateMatrix);
                m_Symbolic.Outputs["ARM2"] = arm2;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, d + fc, -(j4 - fd - td - fe / 2.0) + ff));
                pointCollection.Add(new Position(fa / 2.0, d + fc, -(j4 - td - fd - fe * 0.125) + ff));
                pointCollection.Add(new Position(fa / 2.0, d + fc, -(j4 - td - fd + fe * 0.125) + ff));
                pointCollection.Add(new Position(0, d + fc, -(j4 - td - fd + fe / 2.0) + ff));
                pointCollection.Add(new Position(-fa / 2.0, d + fc, -(j4 - td - fd + fe * 0.125) + ff));
                pointCollection.Add(new Position(-fa / 2.0, d + fc, -(j4 - td - fd - fe * 0.125) + ff));
                pointCollection.Add(new Position(0, d + fc, -(j4 - fd - td - fe / 2.0) + ff));
                Projection3d base1 = new Projection3d(new LineString3d(pointCollection), new Vector(0, 1, 0), sa, true);
                m_Symbolic.Outputs["BASE"] = base1;

                double location = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, 10, UnitName.DISTANCE_MILLIMETER);
                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-fb / 2.0, d + fc, -j4 + ff));
                pointCollection.Add(new Position(-fb / 2.0, d + fc - fs, -j4 + ff));
                pointCollection.Add(new Position(-fb / 2.0, d + fc - fs, -(j4 - td / 2.0) + ff));
                pointCollection.Add(new Position(-fb / 2.0, d - fc / 2.0, -(j4 - td) + ff));
                pointCollection.Add(new Position(-fb / 2.0, d - fc / 2.0, -(j4 - td - fe) + ff));
                pointCollection.Add(new Position(-fb / 2.0, d + fc, -(j4 - td - fd - fe / 2.0) + ff));
                pointCollection.Add(new Position(-fb / 2.0, d + fc, -j4 + ff));
                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), fb, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-fr / 2.0, d + fc - fs / 2.0 - fs / 2.0, -j4);
                Projection3d plate = symbolGeometryHelper.CreateBox(null, fr, fs, ff, 9);
                rotateMatrix.SetIdentity();
                plate.Transform(rotateMatrix);
                m_Symbolic.Outputs["PLATE"] = plate;
            }
            catch (Exception)
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_HBMCS.cs."));
                return;
            }
        }

        #endregion

    }
}
