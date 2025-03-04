//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_Brace_L1.cs
//    HVAC,Ingr.SP3D.Content.Support.Symbols.Gen_Brace_L1
//   Author       :  Hema
//   Creation Date:  29-11-2012
//   Description:    Converted HS_HVAC VB Project to C# .Net

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29-11-2012     Hema    CR-CP-222294 Converted HS_HVAC VB Project to C# .Net
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   03/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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
    [VariableOutputs]
    public class Gen_Brace_L2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HVAC,Ingr.SP3D.Content.Support.Symbols.Gen_Brace_L1"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(4, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputDouble(5, "FlangeT", "FlangeT", 0.999999)]
        public InputDouble m_dFlangeT;
        [InputDouble(6, "WebT", "WebT", 0.999999)]
        public InputDouble m_dWebT;
        [InputString(7, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_oInputBomDesc;
        [InputString(8, "BraceType", "BraceType", "No Value")]
        public InputString m_oBraceType;
        [InputDouble(9, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;
        [InputDouble(10, "BraceOrient", "BraceOrient", 0.999999)]
        public InputDouble m_oBraceOrient;
        [InputDouble(11, "CutBackAngle1", "CutBackAngle1", 0.999999)]
        public InputDouble m_dCutBackAngle1;
        [InputDouble(12, "CutBackAngle2", "CutBackAngle2", 0.999999)]
        public InputDouble m_dCutBackAngle2;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("BODY1", "BODY1")]
        [SymbolOutput("BODY2", "BODY2")]
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
              
                double L = m_dL.Value;
                double width = m_dW.Value;
                double depth = m_dDepth.Value;
                double flangeThickness = m_dFlangeT.Value;
                double webThickness = m_dWebT.Value;
                String braceType = m_oBraceType.Value;
                double angle = m_dAngle.Value;
                int braceOrient =(int) m_oBraceOrient.Value;
                double cutBackAngle1 = m_dCutBackAngle1.Value;
                double cutBackAngle2 = m_dCutBackAngle2.Value;
                double Z1, Z2, Z3, Z4;

                double xDirX1 = 0;
                double xDirY1 = 0;
                double xDirZ1 = 0;
                double zDirX1 = 0;
                double zDirY1 = 0;
                double zDirZ1 = 0;
                double xDirX2 = 0;
                double xDirY2 = 0;
                double xDirZ2 = 0;
                double zDirX2 = 0;
                double zDirY2 = 0;
                double zDirZ2 = 0;

                if (HgrCompareDoubleService.cmpdbl(flangeThickness, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidFlangeThickness, "Flange Thickness cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(depth - flangeThickness, 0) == true) 
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidDepthFlangeThickness, "Depth - Flange Thickness value cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(depth, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidW, "W cannot be zero"));
                    return;
                }

                Z1 = (flangeThickness) / Math.Tan((cutBackAngle1));
                Z2 = (flangeThickness) / (Math.Tan((cutBackAngle2)));
                Z3 = (width) / (Math.Tan((cutBackAngle1)));
                Z4 = (width) / Math.Tan((cutBackAngle2));


                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if (HgrCompareDoubleService.cmpdbl(angle, 0) == true)
                {
                    xDirX1 = 1;
                    xDirY1 = 0;
                    xDirZ1 = 0;
                    zDirX1 = 0;
                    zDirY1 = 0;
                    zDirZ1 = 1;
                    xDirX2 = 1;
                    xDirY2 = 0;
                    xDirZ2 = 0;
                    zDirX2 = 0;
                    zDirY2 = 0;
                    zDirZ2 = 1;
                }
                else if (braceOrient == 1)
                {
                    xDirX1 = Math.Cos(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    xDirY1 = 0;
                    xDirZ1 = Math.Sin(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    zDirX1 = Math.Cos(Math.PI * (90 / ((float)180)) + angle) / (Math.Cos(Math.PI * (90 / ((float)180)) + angle) * Math.Cos(Math.PI * (90 / ((float)180)) + angle) + Math.Sin(Math.PI * (90 / ((float)180)) + angle) * Math.Sin(Math.PI * (90 / ((float)180)) + angle));
                    zDirY1 = 0;
                    zDirZ1 = Math.Sin(Math.PI * (90 / ((float)180)) + angle) / (Math.Cos(Math.PI * (90 / ((float)180)) + angle) * Math.Cos(Math.PI * (90 / ((float)180)) + angle) + Math.Sin(Math.PI * (90 / ((float)180)) + angle) * Math.Sin(Math.PI * (90 / ((float)180)) + angle));
                    xDirX2 = Math.Cos(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    xDirY2 = 0;
                    xDirZ2 = Math.Sin(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    zDirX2 = Math.Cos(Math.PI * (90 / ((float)180)) - angle) / (Math.Cos(Math.PI * (90 / ((float)180)) - angle) * Math.Cos(Math.PI * (90 / ((float)180)) - angle) + Math.Sin(Math.PI * (90 / ((float)180)) - angle) * Math.Sin(Math.PI * (90 / ((float)180)) - angle));
                    zDirY2 = 0;
                    zDirZ2 = Math.Sin(Math.PI * (90 / ((float)180)) - angle) / (Math.Cos(Math.PI * (90 / ((float)180)) + angle) * Math.Cos(Math.PI * (90 / ((float)180)) + angle) + Math.Sin(Math.PI * (90 / ((float)180)) - angle) * Math.Sin(Math.PI * (90 / ((float)180)) - angle));
                }
                else
                {
                    xDirX1 = Math.Sin(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    xDirY1 = Math.Cos(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    xDirZ1 = 0;
                    zDirX1 = 0;
                    zDirY1 = 0;
                    zDirZ1 = 1;
                    xDirX2 = Math.Sin(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    xDirY2 = Math.Cos(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    xDirZ2 = 0;
                    zDirX2 = 0;
                    zDirY2 = 0;
                    zDirZ2 = 1;
                }

                Collection<Position> pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(0, 0, 0));
                pointCollection.Add(new Position(L, 0, 0));
                pointCollection.Add(new Position(L - Z4, 0, width));
                pointCollection.Add(new Position(Z3, 0, width));
                pointCollection.Add(new Position(0, 0, 0));
              
                Projection3d body1 = new Projection3d(new LineString3d(pointCollection), new Vector(0, 1, 0), flangeThickness, true);
                m_Symbolic.Outputs["BODY1"] = body1;

                pointCollection.Clear();

                pointCollection.Add(new Position(0, flangeThickness, 0));
                pointCollection.Add(new Position(L, flangeThickness, 0));
                pointCollection.Add(new Position(L - Z2, flangeThickness, flangeThickness));
                pointCollection.Add(new Position(Z1, flangeThickness, flangeThickness));
                pointCollection.Add(new Position(0, flangeThickness, 0));
              
                Projection3d body2 = new Projection3d(new LineString3d(pointCollection), new Vector(0, 1, 0), depth - flangeThickness, true);
                m_Symbolic.Outputs["BODY2"] = body2;

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "StartStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "EndStructure", new Position(L, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                m_Symbolic.Outputs["Port3"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, width, depth), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;
                Port port5 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                m_Symbolic.Outputs["Port5"] = port5;
            }
            catch    //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Gen_Brace_L2"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                double flangeThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACGenBrace", "FlangeT")).PropValue;
                double webThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACGenBrace", "WebT")).PropValue;
                double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACGenBrace", "W")).PropValue;
                double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACGenBrace", "Depth")).PropValue;
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACGenBrace", "L")).PropValue;
                string braceType = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACGenBrace", "BraceType")).PropValue;

                string bomDescription = (String)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACGenBrace", "InputBomDesc")).PropValue;

                string flangeThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, flangeThicknessValue, UnitName.DISTANCE_INCH);
                string webThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, webThicknessValue, UnitName.DISTANCE_INCH);
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);
                string depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_INCH);
                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_INCH);

                if (bomDescription.Trim() == "None")
                    bomString = " ";
                else if (bomDescription.Trim() == null)
                    bomString = "L" + W + " x " + depth + " x " + flangeThickness + ", Length = " + L;
                else
                    bomString = bomDescription.Trim();
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrConstructOutputs, "Error in BOMDescription of Gen_Brace_L2"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                const int getSteelDensityKGPerM = 7900;
                double weight, cogX, cogY, cogZ;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrHVACGenBrace", "L")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrHVACGenBrace", "W")).PropValue;
                double flangeThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrHVACGenBrace", "FlangeT")).PropValue;
                double webThickness = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrHVACGenBrace", "WebT")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrHVACGenBrace", "Depth")).PropValue;
                string braceType = (string)((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAHgrHVACGenBrace", "BraceType")).PropValue;

                weight = ((L * (width - flangeThickness) * flangeThickness) + (L * depth * flangeThickness)) * getSteelDensityKGPerM;
                cogX = 0;
                cogY = 0;
                cogZ = 0;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrConstructOutputs, "Error in WeightCG of Gen_Brace_L2"));
                }
            }
        }
        #endregion
    }
}
