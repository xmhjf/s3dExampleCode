//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_V2_DS.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_V2_DS
//   Author       :  Vijay
//   Creation Date:  26-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   26-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net  
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
    public class PSL_V2_DS : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_V2_DS"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "HOT_LOAD", "HOT_LOAD", 0.999999)]
        public InputDouble HOT_LOAD;
        [InputDouble(3, "SHOE_H", "SHOE_H", 0.999999)]
        public InputDouble SHOE_H;
        [InputDouble(4, "SPAN", "SPAN", 0.999999)]
        public InputDouble SPAN;
        [InputDouble(5, "WORKING_TRAV", "WORKING_TRAV", 0.999999)]
        public InputDouble WORKING_TRAV;
        [InputDouble(6, "DIR", "DIR", 0.999999)]
        public InputDouble DIR;
        [InputDouble(7, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(8, "AA", "AA", 0.999999)]
        public InputDouble AA;
        [InputDouble(9, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(10, "L", "L", 0.999999)]
        public InputDouble L;
        [InputDouble(11, "M", "M", 0.999999)]
        public InputDouble M;
        [InputDouble(12, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(13, "K", "K", 0.999999)]
        public InputDouble K;
        [InputDouble(14, "RTO", "RTO", 0.999999)]
        public InputDouble RTO;
        [InputDouble(15, "W", "W", 0.999999)]
        public InputDouble W;
        [InputDouble(16, "D1", "D1", 0.999999)]
        public InputDouble D1;
        [InputDouble(17, "BB1", "BB1", 0.999999)]
        public InputDouble BB1;
        [InputDouble(18, "BB2", "BB2", 0.999999)]
        public InputDouble BB2;
        [InputDouble(19, "D2", "D2", 0.999999)]
        public InputDouble D2;
        [InputDouble(20, "BB3", "BB3", 0.999999)]
        public InputDouble BB3;
        [InputDouble(21, "D3", "D3", 0.999999)]
        public InputDouble D3;
        [InputString(22, "SIZE", "SIZE", "No Value")]
        public InputString SIZE;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY1", "BODY1")]
        [SymbolOutput("BODY2", "BODY2")]
        [SymbolOutput("CYL1", "CYL1")]
        [SymbolOutput("CYL2", "CYL2")]
        [SymbolOutput("BEAM", "BEAM")]
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
                Double hotLoad = HOT_LOAD.Value;
                Double shoeHeight = SHOE_H.Value;
                Double span = SPAN.Value;
                Double workingTravel = WORKING_TRAV.Value;
                int dir = (int)DIR.Value;
                Double pipeDiameter = PIPE_DIA.Value;
                Double aa = AA.Value;
                Double rodDiameter = ROD_DIA.Value;
                Double l = L.Value;
                Double m = M.Value;
                Double d = D.Value;
                Double k = K.Value;
                Double rto = RTO.Value;
                Double w = W.Value;
                Double d1 = D1.Value;
                Double bb = BB1.Value;

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

                hotLoad = hotLoad / 2;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Cold load or hot load out of range Min:" + Convert.ToString(minLoad) + " MAX:" + Convert.ToString(maxLoad));
                    return;
                }

                double preset = (coldLoad - minLoad) / springRate;
                preset = Convert.ToDouble(MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, preset, UnitName.DISTANCE_MILLIMETER, UnitName.DISTANCE_METER));

                if (HgrCompareDoubleService.cmpdbl(hotLoad, 0) == true && HgrCompareDoubleService.cmpdbl(workingTravel, 0) == true)
                {
                    coldLoad = 0;
                    preset = 0;
                }

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, span / 2.0, -pipeDiameter / 2.0 + aa + rto + preset - shoeHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, -span / 2.0, -pipeDiameter / 2.0 + aa + rto + preset - shoeHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(span , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidSpanNEZ, "Center to Center Rod Dimension value cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(bb, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBBNEZ, "BB value cannot be zero"));
                    return;
                }

                double down = shoeHeight;
                if (shoeHeight < pipeDiameter / 2.0)
                    down = pipeDiameter / 2.0;

                double x = k - Math.Sqrt((d * d) - (m / 2.0) * (m / 2.0));

                if (span > 800 && span <= 1200)
                {
                    bb = BB2.Value;
                    d1 = D2.Value;
                }

                if (span > 1200 && span <= 1600)
                {
                    bb = BB3.Value;
                    d1 = D3.Value;
                }

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(m / 2.0, span / 2.0 - k / 2.0 + x, -down + aa - l));
                pointCollection.Add(new Position(m / 2.0, span / 2.0 + k / 2.0 - x, -down + aa - l));
                pointCollection.Add(new Position(0, span / 2.0 + k / 2.0, -down + aa - l));
                pointCollection.Add(new Position(-m / 2.0, span / 2.0 + k / 2.0 - x, -down + aa - l));
                pointCollection.Add(new Position(-m / 2.0, span / 2.0 - k / 2.0 + x, -down + aa - l));
                pointCollection.Add(new Position(0, span / 2.0 - k / 2.0, -down + aa - l));
                pointCollection.Add(new Position(m / 2.0, span / 2.0 - k / 2.0 + x, -down + aa - l));

                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(0, 0, 1), l, true);
                m_Symbolic.Outputs["BODY1"] = body;

                pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(m / 2.0, -span / 2.0 - k / 2.0 + x, -down + aa - l));
                pointCollection.Add(new Position(m / 2.0, -span / 2.0 + k / 2.0 - x, -down + aa - l));
                pointCollection.Add(new Position(0, -span / 2.0 + k / 2.0, -down + aa - l));
                pointCollection.Add(new Position(-m / 2.0, -span / 2.0 + k / 2.0 - x, -down + aa - l));
                pointCollection.Add(new Position(-m / 2.0, -span / 2.0 - k / 2.0 + x, -down + aa - l));
                pointCollection.Add(new Position(0, -span / 2.0 - k / 2.0, -down + aa - l));
                pointCollection.Add(new Position(m / 2.0, -span / 2.0 - k / 2.0 + x, -down + aa - l));

                body = new Projection3d(new LineString3d(pointCollection), new Vector(0, 0, 1), l, true);
                m_Symbolic.Outputs["BODY2"] = body;

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(span / 2.0, 0, -down + aa));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d cylinder1 = symbolGeometryHelper.CreateCylinder(null, rodDiameter, 0.152 / 2 + b + preset);
                cylinder1.Transform(matrix);
                m_Symbolic.Outputs["CYL1"] = cylinder1;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(-span / 2.0, 0, -down + aa));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d cylinder2 = symbolGeometryHelper.CreateCylinder(null, rodDiameter, 0.152 / 2 + b + preset);
                cylinder2.Transform(matrix);
                m_Symbolic.Outputs["CYL2"] = cylinder2;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.Translate(new Vector(-span / 2.0, -w / 2.0 - d1, -down - bb));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d beambox = symbolGeometryHelper.CreateBox(null, span, w + d1 * 2.0, bb, 9);
                beambox.Transform(matrix);
                m_Symbolic.Outputs["BEAM"] = beambox;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_V2_DS."));
                    return;
                }
            }
        }

        #endregion

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            double span=double.NaN;
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string size = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;                
                if ((size.Substring(0, 2)).Equals("V1"))
                    span = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_V1_DS", "SPAN")).PropValue;
                else if ((size.Substring(0, 2)).Equals("V2"))
                    span = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_V2_DS", "SPAN")).PropValue;
                else if ((size.Substring(0, 2)).Equals("V3"))
                    span = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_V3_DS", "SPAN")).PropValue;

                bomDescription = "PSL Size " + size + "-DS Variable Supports, Center to Center Dimension: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, span, UnitName.DISTANCE_MILLIMETER);
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_V2_DS.cs."));
                return "";
            }
        }

        #endregion
    }
}
