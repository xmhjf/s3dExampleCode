////+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
////   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
////
////   Nut.cs
////    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Grout
////   Author       :  Havish
////   Creation Date:  25-03-2014
////   Description:    CR-CP-240075  Check-in new Grout Smart Part into V2015.  


////   Change History:    
////   dd.mmm.yyyy     who     change description
////   -----------     ---     ------------------
////   25-03-2014    Havish    CR-CP-240075  Check-in new Grout Smart Part into V2015.
////   12-12-2014      PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report    
////   09-02-2015     Siva     TR-CP-260035  Grout SmartPart is difficult to use           
////   28-04-2015      PVK	   Resolve Coverity issues found in April
////   30-11=2015      VDP     Integrate the newly developed SmartParts into Product(DI-CP-282644)
////+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Support.Exceptions;


namespace Ingr.SP3D.Content.Support.Symbols
{
    //        //-----------------------------------------------------------------------------------
    //    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //    //It is recommended that customers specify namespace of their symbols to be
    //    //<CompanyName>.SP3D.Content.Support.<Specialization>.
    //    //It is also recommended that if customers want to change this symbol to suit their
    //    //requirements, they should change namespace/symbol name so the identity of the modified
    //    //symbol will be different from the one delivered by Intergraph.
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Grout : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Grout"
        //        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
 
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Structure1", "Structure1")]
        [SymbolOutput("Structure2", "Structure2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion
        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddGroutInputs(2, out endIndex, additionalInputs);
                return additionalInputs;
            }
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
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;
                int startindex = 2, endIndex;
                GroutInputs grout = LoadGroutData(startindex, out endIndex);

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                Port port1 = new Port(connection, part, "Structure1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Structure1"] = port1;

                Port port2 = new Port(connection, part, "Structure2", new Position(0, 0, grout.GroutHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Structure2"] = port2;

                AddGrout(grout, m_PhysicalAspect.Outputs, "Grout");

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Grout."));
                }
            }
        }

        #endregion


        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
            
            {

            Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
            double bottomWidth1, bottomWidth2, topWidth1, topWidth2, GroutDensity, GroutHeight;
            string materialType = "", materialGrade = "";
            double materialDensity = 0;

            if (supportComponentBO.SupportsInterface("IJOAhsMaterialEx"))
            {
                try
                {
                    materialType = (((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue);
                    materialGrade = (((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue);
                }
                catch
                {
                    materialType = String.Empty;
                    materialGrade = String.Empty;
                }

            }
            else if (catalogPart.SupportsInterface("IJOAhsMaterialEx"))
            {
                try
                {
                    materialType = (((PropertyValueString)catalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialType")).PropValue);
                    materialGrade = (((PropertyValueString)catalogPart.GetPropertyValue("IJOAhsMaterialEx", "MaterialGrade")).PropValue);
                }
                catch
                {
                    materialType = String.Empty;
                    materialGrade = String.Empty;
                }

            }
            else
            {
                materialType = String.Empty;
                materialGrade = String.Empty;
            }

            //Obtaining the Material Density
            CatalogStructHelper catalogStructHelper = new CatalogStructHelper();


            Material material;
            try
            {
                material = catalogStructHelper.GetMaterial(materialType, materialGrade);
                materialDensity = material.Density;
            }
            catch
            {
                // the specified MaterialType is not available.refdata needs to be checked.
                // so assigning 0 to materialDensity.
                materialDensity = 0;
            }
          

            try
            {
                try
                {
                    bottomWidth1 = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsGroutBotWidth1", "BottomWidth1")).PropValue;
                }
                catch
                {
                    bottomWidth1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsGroutBotWidth1", "BottomWidth1")).PropValue;
                }
            }
            catch
            {
                bottomWidth1 = 0;
            }
            try
            {
                try
                {
                    bottomWidth2 = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsGroutBotWidth2", "BottomWidth2")).PropValue;
                }
                catch
                {
                    bottomWidth2 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsGroutBotWidth2", "BottomWidth2")).PropValue;
                }

                 
            }
            catch
            {
                 bottomWidth2 = 0;
            }
            try
            {
                try
                {
                    topWidth1 = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsGroutTopWidth1", "TopWidth1")).PropValue;
                }
                catch
                {
                    topWidth1 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsGroutTopWidth1", "TopWidth1")).PropValue;
                }
                 
            }
            catch
            {
                 topWidth1 = 0;
            }
            try
            {
                try
                {
                    topWidth2 = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsGroutTopWidth2", "TopWidth2")).PropValue;
                }
                catch
                {
                    topWidth2 = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsGroutTopWidth2", "TopWidth2")).PropValue;
                }
                 
            }
            catch
            {
                 topWidth2 = 0;
            }
            try
            {
                try
                {
                    GroutHeight = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAhsGroutHeight", "GroutHeight")).PropValue;
                }
                catch
                {
                    GroutHeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAhsGroutHeight", "GroutHeight")).PropValue;
                }
                 
            }
            catch
            {
                 GroutHeight = 0;
            }
            try
            {
                GroutDensity = materialDensity;
            }
            catch
            {
                GroutDensity = 0;
            }
               
                
            int ShapeType = (int)((PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAhsGroutShape", "ShapeType")).PropValue;

            if (bottomWidth2 <= 0)
                bottomWidth2 = bottomWidth1;
            else if (bottomWidth1 <= 0)
                bottomWidth1 = bottomWidth2;

            if (topWidth2 <= 0)
                topWidth2 = topWidth1;
            else if (topWidth1 <= 0)
                topWidth1 = topWidth2;


            // The below code is for calculating the Bottom and Top area's of Round Grout and Square Grout respectively
            double cGroutArea1 = Math.PI * Math.Pow(bottomWidth1, 2);
            double cGroutArea2 = Math.PI * Math.Pow(topWidth1, 2);
            double sGroutArea1 = Math.Pow(bottomWidth1, 2);
            double sGroutArea2 = Math.Pow(topWidth1, 2);

            double Weight = 0, CogX = 0, CogY = 0, CogZ = 0;
            try
            {
                ////System WCG Attributes
            
                switch (ShapeType)
                {
                    case 1:
                        try
                        {
                            Weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                        }
                        catch
                        {
                            Weight = (double)(GroutDensity * GroutHeight * (cGroutArea1 + cGroutArea2 + (Math.Pow((cGroutArea1 * cGroutArea2), 1.0 / 2.0)))) / 3;
                        }
                        //Center of Gravity
                        try
                        {
                            CogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                        }
                        catch
                        {
                            CogX = 0;
                        }
                        try
                        {
                            CogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                        }
                        catch
                        {
                            CogY = 0;
                        }
                        try
                        {
                            CogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                        }
                        catch
                        {
                            CogZ = (GroutHeight * (cGroutArea1 + 2 * Math.Pow((cGroutArea1 * cGroutArea2), 1.0 / 2.0) + (3 * cGroutArea2)) / (4 * (cGroutArea1 + cGroutArea2 + (Math.Pow((cGroutArea1 * cGroutArea2), 1.0 / 2.0)))));
                        }
                        break;

                    case 2:

                        try
                        {
                            Weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                        }
                        catch
                        {
                            Weight = (1.0 / 3.0) * (GroutDensity * GroutHeight * (sGroutArea1 + sGroutArea2 + (Math.Pow((sGroutArea1 * sGroutArea2), 1.0 / 2.0)))); 
                        }
                        //Center of Gravity
                        try
                        {
                            CogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                        }
                        catch
                        {
                            CogX = 0;
                        }
                        try
                        {
                            CogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                        }
                        catch
                        {
                            CogY = 0;
                        }
                        try
                        {
                            CogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                        }
                        catch
                        {
                            CogZ = (double)((GroutHeight * (sGroutArea1 + 2 * Math.Pow((sGroutArea1 * sGroutArea2), 1.0 / 2.0) + (3 * sGroutArea2)) / (4 * (sGroutArea1 + sGroutArea2 + (Math.Pow((sGroutArea1 * sGroutArea2), 1.0 / 2.0))))));
                        }

                        break;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(Weight, CogX, CogY, CogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Grout."));
                }
            }
        }
        #endregion

    }
}