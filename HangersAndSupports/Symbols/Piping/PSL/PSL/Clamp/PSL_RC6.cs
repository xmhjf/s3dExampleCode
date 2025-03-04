//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_RC6.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_RC6
//   Author       :  Vijaya
//   Creation Date:  26-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   26-08-2013     Vijaya    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net 
//   30-Dec-2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report   
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_RC6 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {

        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_RC6"

        #region "Definition of Inputs"


        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "F", "F", 0.999999)]
        public InputDouble F;
        [InputString(3, "TEMP2", "TEMP2", "No Value")]
        public InputString TEMP2;
        [InputString(4, "PIPE_NOM_DIA", "PIPE_NOM_DIA", "No Value")]
        public InputString PIPE_NOM_DIA;
        [InputString(5, "LOAD_GRP", "LOAD_GRP", "No Value")]
        public InputString LOAD_GRP;
        [InputDouble(6, "MIN_F", "MIN_F", 0.999999)]
        public InputDouble MIN_F;
        [InputDouble(7, "MAX_F", "MAX_F", 0.999999)]
        public InputDouble MAX_F;
        [InputDouble(8, "MIN_E", "MIN_E", 0.999999)]
        public InputDouble MIN_E;
        [InputDouble(9, "MAX_E", "MAX_E", 0.999999)]
        public InputDouble MAX_E;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LEFT_FRONT", "LEFT_FRONT")]
        [SymbolOutput("LEFT_BACK", "LEFT_BACK")]
        [SymbolOutput("RIGHT_FRONT", "RIGHT_FRONT")]
        [SymbolOutput("RIGHT_BACK", "RIGHT_BACK")]
        [SymbolOutput("LEFT_BOLT1", "LEFT_BOLT1")]
        [SymbolOutput("RIGHT_BOLT1", "RIGHT_BOLT1")]
        [SymbolOutput("LEFT_BOLT2", "LEFT_BOLT2")]
        [SymbolOutput("RIGHT_BOLT2", "RIGHT_BOLT2")]
        [SymbolOutput("LEFT_BOLT3", "LEFT_BOLT3")]
        [SymbolOutput("RIGHT_BOLT3", "RIGHT_BOLT3")]
        [SymbolOutput("LINE", "LINE")]
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

                Double pipeDiameter, t, a, b, c, boltLength, e, pinLength, f = F.Value, minimumE = MIN_E.Value, minimumF = MIN_F.Value, maximumE = MAX_E.Value, maximumF = MAX_F.Value;
                string temp2 = TEMP2.Value;
                //check if Rod Centre (F) false within Max and Min F
                if (f < minimumF || f > maximumF)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConstructOutputs" + ": " + "ERROR: " + "Rod Centre must be between " + minimumF * 1000 + " mm and " + maximumF * 1000 + "mm.", "", "PSL_RC6.cs", 103);
                string f1 = Convert.ToString(f * 1000);

                if (HgrCompareDoubleService.cmpdbl((f * 1000 / 50) * 50 , f * 1000)==false)
                    if (HgrCompareDoubleService.cmpdbl(maximumF - minimumF ,0) ==false)
                        f1 = Convert.ToString((f * 1000 / 50) * 50 + 50);

                string partNumber = "RC6-" + PIPE_NOM_DIA.Value + "/" + f1.Trim() + "-" + LOAD_GRP.Value + "-" + temp2;
                string auxialaryTable = string.Empty, interfaceName = string.Empty;
                if (temp2.Equals("400"))
                {
                    auxialaryTable = "PSL_RC6_AUX_400";
                    interfaceName = "IJUAHgrPSL_RC6_AUX_400";
                }
                if (temp2.Equals("490"))
                {
                    auxialaryTable = "PSL_RC6_AUX_490";
                    interfaceName = "IJUAHgrPSL_RC6_AUX_490";
                }
                if (temp2.Equals("530"))
                {
                    auxialaryTable = "PSL_RC6_AUX_530";
                    interfaceName = "IJUAHgrPSL_RC6_AUX_530";
                }
                if (temp2.Equals("560"))
                {
                    auxialaryTable = "PSL_RC6_AUX_560";
                    interfaceName = "IJUAHgrPSL_RC6_AUX_560";
                }
                if (temp2.Equals("600"))
                {
                    auxialaryTable = "PSL_RC6_AUX_600";
                    interfaceName = "IJUAHgrPSL_RC6_AUX_600";
                }
                Dictionary<KeyValuePair<string, string>, Object> parameter = new Dictionary<KeyValuePair<string, string>, Object>();
                parameter.Add(new KeyValuePair<string, string>(interfaceName, "PART_NUMBER"), partNumber);
                pipeDiameter = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "PIPE_DIA", parameter);
                t = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "T", parameter);
                b = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "B", parameter);
                a = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "A", parameter);
                c = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "C", parameter);
                boltLength = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "BOLT_LEN", parameter);
                pinLength = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "PIN_LEN", parameter);

                // Calculate the value of E by interpolating MIN_E and MAX_E
                if (HgrCompareDoubleService.cmpdbl(maximumF - minimumF, 0) == false)
                    e = minimumE + (maximumE - minimumE) * (f - minimumF) / (maximumF - minimumF);
                else
                    e = (double)PSLSymbolServices.GetDataByMultipleConditions(auxialaryTable, interfaceName, "E", parameter);

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, -f / 2, -c - a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "Pin", new Position(0, f / 2, -c - a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;
                //Validating Inputs               
                if (t <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTGTZ, "T value should be greater than zero"));
                    return;
                }
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + t, e);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, -e));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-(f / 2 + a * 1.75), -b / 2 - t, -e);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d leftFront = symbolGeometryHelper.CreateBox(null, f / 2 + a * 1.75 - pipeDiameter / 2, t, e, 9);
                leftFront.Transform(matrix);
                m_Symbolic.Outputs["LEFT_FRONT"] = leftFront;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-(f / 2 + a * 1.75), b / 2, -e);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d leftBack = symbolGeometryHelper.CreateBox(null, f / 2 + a * 1.75 - pipeDiameter / 2, t, e, 9);
                leftBack.Transform(matrix);
                m_Symbolic.Outputs["LEFT_BACK"] = leftBack;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2, -b / 2 - t, -e);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d rightFront = symbolGeometryHelper.CreateBox(null, f / 2 + a * 1.75 - pipeDiameter / 2, t, e, 9);
                rightFront.Transform(matrix);
                m_Symbolic.Outputs["RIGHT_FRONT"] = rightFront;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2, b / 2, -e);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d rightBack = symbolGeometryHelper.CreateBox(null, f / 2 + a * 1.75 - pipeDiameter / 2, t, e, 9);
                rightBack.Transform(matrix);
                m_Symbolic.Outputs["RIGHT_BACK"] = rightBack;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(boltLength / 2, -f / 2, -c).Subtract(new Position(-boltLength / 2, -f / 2, -c));
                symbolGeometryHelper.ActivePosition = new Position(-boltLength / 2, -f / 2, -c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d leftBolt1 = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["LEFT_BOLT1"] = leftBolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(boltLength / 2, f / 2, -c).Subtract(new Position(-boltLength / 2, f / 2, -c));
                symbolGeometryHelper.ActivePosition = new Position(-boltLength / 2, f / 2, -c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rightBolt1 = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["RIGHT_BOLT1"] = rightBolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(pinLength / 2, -pipeDiameter / 2 - a - t * 2.5 - 0.01, -e / 2 - e / 4).Subtract(new Position(-pinLength / 2, -pipeDiameter / 2 - a - t * 2.5 - 0.01, -e / 2 - e / 4));
                symbolGeometryHelper.ActivePosition = new Position(-pinLength / 2, -pipeDiameter / 2 - a - t * 2.5 - 0.01, -e / 2 - e / 4);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d leftBolt = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["LEFT_BOLT2"] = leftBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(pinLength / 2, pipeDiameter / 2 + a + t * 2.5 + 0.01, -e / 2 - e / 4).Subtract(new Position(-pinLength / 2, pipeDiameter / 2 + a + t * 2.5 + 0.01, -e / 2 - e / 4));
                symbolGeometryHelper.ActivePosition = new Position(-pinLength / 2, pipeDiameter / 2 + a + t * 2.5 + 0.01, -e / 2 - e / 4);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rightBolt2 = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["RIGHT_BOLT2"] = rightBolt2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(pinLength / 2, -pipeDiameter / 2 - a - t * 2.5 - 0.01, -e / 2 + e / 4).Subtract(new Position(-pinLength / 2, -pipeDiameter / 2 - a - t * 2.5 - 0.01, -e / 2 + e / 4));
                symbolGeometryHelper.ActivePosition = new Position(-pinLength / 2, -pipeDiameter / 2 - a - t * 2.5 - 0.01, -e / 2 + e / 4);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d leftBolt3 = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["LEFT_BOLT3"] = leftBolt3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(pinLength / 2, pipeDiameter / 2 + a + t * 2.5 + 0.01, -e / 2 + e / 4).Subtract(new Position(-pinLength / 2, pipeDiameter / 2 + a + t * 2.5 + 0.01, -e / 2 + e / 4));
                symbolGeometryHelper.ActivePosition = new Position(-pinLength / 2, pipeDiameter / 2 + a + t * 2.5 + 0.01, -e / 2 + e / 4);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rightBolt3 = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["RIGHT_BOLT3"] = rightBolt3;

                Line3d line = new Line3d(new Position(0, 0, -e / 2), new Position(0, 0, e / 2));
                m_Symbolic.Outputs["LINE"] = line;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_RC6.cs"));
                    return;
                }
            }
        }

        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                string size = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;
                double rodCentersF = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_RC6", "F")).PropValue;

                bomDescription = "PSL " + size + "Pressed Riser Clamp Six Bolt Type, Rod Centre =" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodCentersF, UnitName.DISTANCE_MILLIMETER);

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_RC6.cs."));
                return "";
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject oSupportOrComponent)
        {
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double maximumWeight, minimumWeight, minimumF, maximumF, f;
                minimumWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_RC6", "MIN_WEIGHT")).PropValue;
                try
                {
                    maximumWeight = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_RC6", "MAX_WEIGHT")).PropValue;
                }
                catch
                {
                    maximumWeight = 0.0;
                }

                minimumF = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_RC6", "MIN_F")).PropValue;
                maximumF = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_RC6", "MAX_F")).PropValue;
                f = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_RC6", "F")).PropValue;

                double weight = 0.0, cogX, cogY, cogZ;

                if (HgrCompareDoubleService.cmpdbl(maximumF - minimumF, 0) == false)
                    weight = minimumWeight + (maximumWeight - minimumWeight) * (f - minimumF) / (maximumF - minimumF);

                cogX = 0;
                cogY = 0;
                cogZ = 0;

                SupportComponent supportComponent = (SupportComponent)oSupportOrComponent;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_RC6."));
            }
        }
        #endregion
    }
}
