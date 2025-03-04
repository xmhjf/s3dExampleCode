//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_Brace.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.Gen_Brace
//   Author       :  Hema
//   Creation Date:  16-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16-11-2012      Hema   CR-CP-222274  Converted HS_Generic VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
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

    [CacheOption(CacheOptionType.NonCached)] 
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Gen_Brace : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.Gen_Brace"
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
        [InputString(7, "InputBomDesc", "InputBomDesc","No Value")]
        public InputString m_oInputBomDesc;
        [InputString(8, "BraceType", "BraceType","No Value")]
        public InputString m_oBraceType;
        [InputDouble(9, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;      
        [InputDouble(10, "BraceOrient", "BraceOrient",1)]
        public InputDouble m_oBraceOrient;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY", "BODY")]
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

                Double L = m_dL.Value;
                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidL, "L cannot be zero"));
                    return;
                }
                Double width = m_dW.Value;
                Double depth = m_dDepth.Value;
                Double flangeThickness = m_dFlangeT.Value;
                Double webThickness = m_dWebT.Value;
                String braceType = m_oBraceType.Value;
                Double angle = m_dAngle.Value;
                Double braceOrient = m_oBraceOrient.Value;
               
                Double xDirX1 = 0, xDirY1 = 0,xDirZ1 = 0,zDirX1 = 0, zDirY1 = 0,zDirZ1 = 0,xDirX2 = 0,xDirY2 = 0,xDirZ2 = 0,zDirX2 = 0,zDirY2 = 0,zDirZ2 = 0;
                Collection<Position> pointCollection;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                
                if (angle == 0)
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
                    zDirX2 = Math.Cos(Math.PI * (90 / ((float)180)) - angle) / (Math.Cos(Math.PI * (90 / ((float)180)) - angle) * Math.Cos(Math.PI * (90 / ((float)180)) - angle) + Math.Sin(Math.PI * (90 / ((float)180)) - angle) * Math.Sin(Math.PI * (90 / ((float)180)) - angle) );
                    zDirY2 = 0;
                    zDirZ2 = Math.Sin(Math.PI * (90 / ((float)180)) - angle) / (Math.Cos(Math.PI * (90 / ((float)180)) + angle) * Math.Cos(Math.PI * (90 / ((float)180)) + angle)  + Math.Sin(Math.PI * (90 / ((float)180)) - angle) * Math.Sin(Math.PI * (90 / ((float)180)) - angle));
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
                 if( (braceType.ToUpper()) == "W SECTION" )
                 {
                     Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                     m_Symbolic.Outputs["Port1"] = port1;
                     Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, 0, depth / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                     m_Symbolic.Outputs["Port2"] = port2;
                     Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                     m_Symbolic.Outputs["Port3"] = port3;

                     pointCollection = new Collection<Position>();

                     pointCollection.Add(new Position(0, -width / 2, depth / 2));
                     pointCollection.Add(new Position(0, width / 2, depth / 2));
                     pointCollection.Add(new Position(0, width / 2, depth / 2 - flangeThickness));
                     pointCollection.Add(new Position(0, webThickness / 2, depth / 2 - flangeThickness));
                     pointCollection.Add(new Position(0, webThickness / 2, -depth / 2 + flangeThickness));
                     pointCollection.Add(new Position(0, width / 2, -depth / 2 + flangeThickness));
                     pointCollection.Add(new Position(0, width / 2, -depth / 2));
                     pointCollection.Add(new Position(0, -width / 2, -depth / 2));
                     pointCollection.Add(new Position(0, -width / 2, -depth / 2 + flangeThickness));
                     pointCollection.Add(new Position(0, -webThickness / 2, -depth / 2 + flangeThickness));
                     pointCollection.Add(new Position(0, -webThickness / 2, depth / 2 - flangeThickness));
                     pointCollection.Add(new Position(0, -width / 2, depth / 2 - flangeThickness));
                     pointCollection.Add(new Position(0, -width / 2, depth / 2));

                     Vector projectionVector = new Vector(L, 0, 0);
                     Projection3d body = new Projection3d( new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                     m_Symbolic.Outputs["BODY"] = body;
                 }
                 else if (braceType.ToUpper() == "C SECTION")
                 {
                     Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                     m_Symbolic.Outputs["Port1"] = port1;
                     Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, width / 2, depth / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                     m_Symbolic.Outputs["Port2"] = port2;
                     Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                     m_Symbolic.Outputs["Port3"] = port3;

                     pointCollection = new Collection<Position>();

                     pointCollection.Add(new Position(0, 0, 0));
                     pointCollection.Add(new Position(0, 0, depth));
                     pointCollection.Add(new Position(0, flangeThickness, depth));
                     pointCollection.Add(new Position(0, flangeThickness, webThickness));
                     pointCollection.Add(new Position(0, width - flangeThickness, webThickness));
                     pointCollection.Add(new Position(0, width - flangeThickness, depth));
                     pointCollection.Add(new Position(0, width, depth));
                     pointCollection.Add(new Position(0, width, 0));
                     pointCollection.Add(new Position(0, 0, 0));
                   
                     Vector projectionVector = new Vector(L, 0, 0);
                     Projection3d body = new Projection3d( new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                     m_Symbolic.Outputs["BODY"] = body;
                 }
                 else if (braceType.ToUpper() == "HSS SECTION")
                 {
                     Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                     m_Symbolic.Outputs["Port1"] = port1;
                     Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, width / 2, depth / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                     m_Symbolic.Outputs["Port2"] = port2;
                     Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                     m_Symbolic.Outputs["Port3"] = port3;

                     pointCollection = new Collection<Position>();

                     pointCollection.Add(new Position(0, 0, 0));
                     pointCollection.Add(new Position(0, 0, depth));
                     pointCollection.Add(new Position(0, width, depth));
                     pointCollection.Add(new Position(0, width, 0));
                     pointCollection.Add(new Position(0, 0, 0));

                     Vector projectionVector = new Vector(L, 0, 0);
                     Projection3d body = new Projection3d( new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                     m_Symbolic.Outputs["BODY"] = body;
                 }
                 else
                 {
                     Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                     m_Symbolic.Outputs["Port1"] = port1;
                     Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, width, depth), new Vector(1, 0, 0), new Vector(0, 0, 1));
                     m_Symbolic.Outputs["Port2"] = port2;
                     Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                     m_Symbolic.Outputs["Port3"] = port3;

                     pointCollection = new Collection<Position>();

                     pointCollection.Add(new Position(0, 0, 0));
                     pointCollection.Add(new Position(0, 0, depth));
                     pointCollection.Add(new Position(0, flangeThickness, depth));
                     pointCollection.Add(new Position(0, flangeThickness, flangeThickness));
                     pointCollection.Add(new Position(0, width, flangeThickness));
                     pointCollection.Add(new Position(0, width, 0));
                     pointCollection.Add(new Position(0,0,0));

                     Vector projectionVector = new Vector(L, 0, 0);
                     Projection3d body = new Projection3d( new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                     m_Symbolic.Outputs["BODY"] = body;
                 }
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of Gen_Brace"));
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
                double flangeThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericBrace", "FlangeT")).PropValue;
                double webThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericBrace", "WebT")).PropValue;
                double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericBrace", "W")).PropValue;
                double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericBrace", "Depth")).PropValue;
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericBrace", "L")).PropValue;
                string braceType = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericBrace", "BraceType")).PropValue;

                string bomDescription = (String)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericBrace", "InputBomDesc")).PropValue;

                string flangeThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, flangeThicknessValue, UnitName.DISTANCE_INCH);
                string webThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, webThicknessValue, UnitName.DISTANCE_INCH);
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);
                string depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_INCH);
                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_INCH);
                

                if (bomDescription.Trim() == "None")
                    bomString = " ";
                else if (bomDescription.Trim() == " ")
                    if (braceType.Trim() == "W SECTION")
                        bomString = "W " + W + " x " + depth + ", Flange =" + flangeThickness + ", Web =" + webThickness + ", Length = " + L;
                    else if (braceType.Trim() == "C SECTION")
                        bomString = "C " + W + " x " + depth + ", Flange =" + flangeThickness + ", Web =" + webThickness + ", Length = " + L;
                    else
                        bomString = "L" + W + " x " + depth + " x " + flangeThickness + ", Length = " + L;
                else
                    bomString = bomDescription;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrBOMDescription, "Error in BOMDescription of Gen_Brace"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                const int getSteelDensityKGPerM = 7900;
                Double weight, cogX=0, cogY=0, cogZ=0;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenericBrace", "L")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenericBrace", "W")).PropValue;
                double flangeThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenericBrace", "FlangeT")).PropValue;
                double webThickness = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrGenericBrace", "WebT")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrGenericBrace", "Depth")).PropValue;
                string braceType = (string)((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAHgrGenericBrace", "BraceType")).PropValue;

                if (braceType.ToUpper() == "W SECTION")
                    weight = (((L * width * flangeThickness) * 2) + (L * (depth - flangeThickness * 2) * webThickness)) * getSteelDensityKGPerM;
                else if (braceType.ToUpper() == "W SECTION")
                    weight = 2 * (flangeThickness * depth) + (webThickness * width) * L * getSteelDensityKGPerM;
                else
                    weight = ((L * (width - flangeThickness) * flangeThickness) + (L * depth * flangeThickness)) * getSteelDensityKGPerM;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrWeightCG, "Error in WeightCG of Gen_Brace"));
                }
            }
        }
        #endregion
    }
}
