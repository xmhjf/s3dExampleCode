//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_V2_TS1.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_V2_TS1
//   Author       :  Vijay
//   Creation Date:  27-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net  
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
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_V2_TS1 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_V2_TS1"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "WORKING_TRAV", "WORKING_TRAV", 0.999999)]
        public InputDouble WORKING_TRAV;
        [InputDouble(3, "DIR", "DIR", 0.999999)]
        public InputDouble DIR;
        [InputDouble(4, "HOT_LOAD", "HOT_LOAD", 0.999999)]
        public InputDouble HOT_LOAD;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(6, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(7, "L", "L", 0.999999)]
        public InputDouble L;
        [InputDouble(8, "M", "M", 0.999999)]
        public InputDouble M;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(10, "K", "K", 0.999999)]
        public InputDouble K;
        [InputDouble(11, "RTO", "RTO", 0.999999)]
        public InputDouble RTO;
        [InputString(12, "SIZE", "SIZE", "No Value")]
        public InputString SIZE;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("CYLINDER1", "CYLINDER1")]
        [SymbolOutput("CYLINDER2", "CYLINDER2")]
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
                Part part = (Part)m_PartInput.Value;
                String size = SIZE.Value;
                Double workingTravel = WORKING_TRAV.Value;
                int dir = (int)DIR.Value;
                Double hotLoad = HOT_LOAD.Value;
                Double a = A.Value;
                Double rodDiameter = ROD_DIA.Value;
                Double l = L.Value;
                Double m = M.Value;
                Double d = D.Value;
                Double k = K.Value;
                Double rto = RTO.Value;

                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrRod_Dia_mm", "ROD_DIA"), rodDiameter);
                Double b = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_235", "IJUAHgrPSL_235", "B", parameter, 0.001, PSLSymbolServices.ComparisionOperator.BETWEEN_WITHOUT_LIMITS);

                parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>("IJUAHgrPSL_SPRING_RATE", "SIZE"), size.Trim());
                Double springRate = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_SPRING_RATE", "IJUAHgrPSL_SPRING_RATE", "RATE", parameter);
                Double minLoad = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_SPRING_RATE", "IJUAHgrPSL_SPRING_RATE", "MIN_LOAD", parameter);
                Double maxLoad = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_SPRING_RATE", "IJUAHgrPSL_SPRING_RATE", "MAX_LOAD", parameter);
                Double maxOverTravelLoad = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_SPRING_RATE", "IJUAHgrPSL_SPRING_RATE", "MAX_OVERTRAVEL_LOAD", parameter);
                Double minOverTravelLoad = (double)PSLSymbolServices.GetDataByMultipleConditions("PSL_SPRING_RATE", "IJUAHgrPSL_SPRING_RATE", "MIN_OVERTRAVEL_LOAD", parameter);

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                workingTravel = Convert.ToDouble(MiddleServiceProvider.UOMMgr.ConvertDBUtoUnit(UnitType.Distance, workingTravel, UnitName.DISTANCE_MILLIMETER));
                if (dir == 2)      //to match the formular movement_up - Movement down
                    workingTravel = workingTravel * (-1);

                Double coldLoad = ((hotLoad) + ((workingTravel) * springRate));
                double minLoadCalc = 0.0, maxLoadCalc = 0.0;

                if (coldLoad < hotLoad)
                {
                    minLoadCalc = coldLoad;
                    maxLoadCalc = hotLoad;
                }
                else
                {
                    minLoadCalc = hotLoad;
                    maxLoadCalc = coldLoad;
                }

                if (minLoadCalc >= minOverTravelLoad && maxLoadCalc <= maxOverTravelLoad + 0.0001)
                {
                    if (minLoadCalc < minLoad - 0.0001 || maxLoadCalc > maxLoad)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "Cold load or hot load out of range Min:" + Convert.ToString(minLoad) + " MAX:" + Convert.ToString(maxLoad));
                    }
                }
                else    //if out of overtravel range make it fail
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,"Cold load or hot load out of range Min:" + Convert.ToString(minLoad) + " MAX:" + Convert.ToString(maxLoad));
                    return;
                }
                double preset = (coldLoad - minLoad) / springRate;
                preset = Convert.ToDouble(MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, preset, UnitName.DISTANCE_MILLIMETER, UnitName.DISTANCE_METER));

                if (HgrCompareDoubleService.cmpdbl(hotLoad, 0) == true && HgrCompareDoubleService.cmpdbl(workingTravel, 0) == true)
                {
                    coldLoad = 0;
                    preset = 0;
                }
                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, rto + preset), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero."));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(a,0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidANEZ, "A value cannot be zero"));
                    return;
                }
                double x = k - Math.Sqrt((d * d) - (m / 2.0) * (m / 2.0));

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(m / 2.0, -k / 2.0 + x, rto - l + preset));
                pointCollection.Add(new Position(m / 2.0, k / 2.0 - x, rto - l + preset));
                pointCollection.Add(new Position(0, k / 2.0, rto - l + preset));
                pointCollection.Add(new Position(-m / 2.0, k / 2.0 - x, rto - l + preset));
                pointCollection.Add(new Position(-m / 2.0, -k / 2.0 + x, rto - l + preset));
                pointCollection.Add(new Position(0, -k / 2.0, rto - l + preset));
                pointCollection.Add(new Position(m / 2.0, -k / 2.0 + x, rto - l + preset));

                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(0, 0, 1), l, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, 0, rto - l + preset));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d cylinder1 = symbolGeometryHelper.CreateCylinder(null, rodDiameter, (rto - l) + 0.077 / 2 + b + preset);
                cylinder1.Transform(matrix);
                m_Symbolic.Outputs["CYLINDER1"] = cylinder1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(0, 0, rto + preset));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d cylinder2 = symbolGeometryHelper.CreateCylinder(null, rodDiameter, a);
                cylinder2.Transform(matrix);
                m_Symbolic.Outputs["CYLINDER2"] = cylinder2;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_V2_TS1."));
                    return;
                }
            }
        }

        #endregion

    }
}
