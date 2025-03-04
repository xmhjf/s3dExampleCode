//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SwivelRing.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.SwivelRing
//   Author       :  Hema
//   Creation Date:  25/02/2013
//   Description:    Converted SwivelRing Smartpart VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   25/02/2013     Hema     CR-CP-222467 Converted SwivelRing Smartpart VB Project to C# .Net 
//   25/Mar/2013    Hema     DI-CP-228142  Modify the error handling for delivered H&S symbols
//   11-02-2013       Chethan DI-CP-263820  Fix priority 3 items to .net SmartParts as a result of new testing  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
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

    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]

    public class SwivelRing : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.SwivelRing"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("RodEnd", "RodEnd")]
        [SymbolOutput("Surface", "Surface")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddSwivelRingInputs(2, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddSwivelRingOutputs(additionalOutputs);
            }
            return additionalOutputs;
        }

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

                //Get the values from the database and put them into the swivel variable.
                SwivelInputs Swivel = LoadSwivelData(2);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                double calcDia, dx, angle1, angle2, angle3, lowerRightZ, lowerRightY, lowerLeftZ, lowerLeftY, dx2, angle4, tempUpperRightY, upperRightY, rotateAngle = 0;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if (Swivel.PipeOD > Swivel.Diameter1 - 0.000001)
                    calcDia = Swivel.PipeOD;
                else
                    calcDia = Swivel.Diameter1;

                // Check to see if Height3 had a value other then 0 if it does then this will be a jShape hanger and need three ports.
                if (Swivel.Height3 > 0.000001)
                {
                    dx = Math.Sqrt(((Swivel.Width2) * (Swivel.Width2)) + ((Swivel.Height1 + (Swivel.Thickness1 * 2)) * (Swivel.Height1 + (Swivel.Thickness1 * 2))));

                    // Calculations for the left side of the symbol
                    angle1 = Math.Asin(Swivel.Width2 / dx) * 180 / Math.PI;
                    angle2 = Math.Acos(((calcDia / 2) + Swivel.Thickness1) / dx) * 180 / Math.PI;
                    angle3 = 90 - angle1 - angle2;

                    // Set the values for the dimensions of the left side
                    lowerRightZ = Math.Sin(angle3 * Math.PI / 180) * (calcDia / 2);
                    lowerRightY = Math.Cos(angle3 * Math.PI / 180) * (calcDia / 2);
                    lowerLeftZ = Math.Sin(angle3 * Math.PI / 180) * ((calcDia / 2) + Swivel.Thickness1);
                    lowerLeftY = Math.Cos(angle3 * Math.PI / 180) * ((calcDia / 2) + Swivel.Thickness1);

                    dx2 = lowerLeftY - Swivel.Width2;

                    angle4 = Math.Atan(((dx2 / (Swivel.Height1 + (Swivel.Thickness1 * 2))))) * 180 / Math.PI;
                    tempUpperRightY = Math.Tan(angle4 * 180 / Math.PI) * (Swivel.Height1 + Swivel.Thickness1);

                    upperRightY = (lowerRightY - tempUpperRightY);

                    // This angle is used to rotate the swivel shape and bolt.  This will normally only be used if it is the J shape hanger.
                    rotateAngle = Math.Atan((((Swivel.Height1 + Swivel.Thickness1) - lowerRightZ) / (lowerRightY - upperRightY))) * 180 / Math.PI;
                    
                    //ports

                    Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, Math.Sin(rotateAngle * Math.PI / 180), Math.Cos(rotateAngle * Math.PI / 180)));                    
                    m_PhysicalAspect.Outputs["Route"] = route;                    
                    Port side = new Port(OccurrenceConnection, part, "RodEnd", new Position(0, -calcDia / 4, Swivel.RodTakeOut), new Vector(0, Math.Cos((rotateAngle) * Math.PI / 180), -Math.Sin((rotateAngle) * Math.PI / 180)), new Vector(0, Math.Cos((270+rotateAngle) * Math.PI / 180), -Math.Sin((270+rotateAngle) * Math.PI / 180)));                    
                    m_PhysicalAspect.Outputs["RodEnd"] = side;
                    Port surface = new Port(OccurrenceConnection, part, "Surface", new Position(0, -calcDia / 2 - 2 * Swivel.Thickness2 - Swivel.Thickness1, Swivel.Height3), new Vector(0, 0, 1), new Vector(0, 1, 0));
                    m_PhysicalAspect.Outputs["Surface"] = surface;
                }
                else
                {
                    //As this is not a JShape hanger then we only need two ports.
                    Port route = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Route"] = route;
                    Port side = new Port(OccurrenceConnection, part, "RodEnd", new Position(0, 0, Swivel.RodTakeOut), new Vector(0, 1, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["RodEnd"] = side;
                }
                //Add the swivel ring shape, this will create all three types of swivel rings based on the values added. If this is a JHanger this will also add the bolt that goes between the ears.
                AddSwivelRing(Swivel, rotateAngle, new Matrix4X4(), m_PhysicalAspect.Outputs, "SwivelRing");
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SwivelRing"));
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of SwivelRing"));
                }
            }
        }
        #endregion

    }

}
