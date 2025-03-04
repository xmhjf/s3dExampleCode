//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2015, Intergraph Corporation. All rights reserved.
//
//   SpreaderBeam.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.SpreaderBeam
//   Author       :  Vinay/Devender
//   Creation Date:  11-08-2014
//   Description:CR-CP-179357       Initial Creation
//
//   Change History:
//   dd.mm.yyyy     who                change description
//   12-11-2014     Chethan             DI-CP-263659  Improvise to add End plates to Spreader beam smart part
//   30-11=2015      VDP     Integrate the newly developed SmartParts into Product(DI-CP-282644)
  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Collections;

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
    public class SpreaderBeam : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.SpreaderBeam"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #endregion
       
        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Topport1", "Topport1")]
        [SymbolOutput("Topport2", "Topport2")]
        [SymbolOutput("Botport1", "Botport1")]
        [SymbolOutput("Botport2", "Botport2")]
        [SymbolOutput("BotMiddleport", "BotMiddleport")]
        [SymbolOutput("SpreaderBeam", "SpreaderBeam")]
        [SymbolOutput("EndPlatePort1", "EndPlatePort1")]
        [SymbolOutput("EndPlatePort2", "EndPlatePort2")]
        public AspectDefinition m_oSymbolic;
        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddSpreaderBeamInputs(2, out endIndex, additionalInputs);
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
                #region Declaration of variables and their initialisation
                Part part = (Part)m_PartInput.Value;
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                //Retrieve XLS data through array of inputs
                SpreaderBeamInputs SpreaderBeam = LoadSpreaderBeamData(2);

                //SpreaderBeamInputs SpreaderBeam;
                PlateInputs stiffenerInputs, topplateInputs, bottomplateinputs, middleplateInputs, toplugplateInputs, bottomlugplateinputs, middlelugplateInputs, endplateInputs;
                UBoltInputs topuboltInputs, middleuboltInputs, bottomuboltInputs;
                WBABoltInputs topattWBABoltInputs, middleattWBABoltInputs, bottomattWBABoltInputs, topconWBABoltInputs, middleconWBABoltInputs, bottomconWBABoltInputs;
                WBAHoleInputs topattWBAHoleInputs, middleattWBAHoleInputs, bottomattWBAHoleInputs;
                SwivelInputs topswivelInputs, middleswivelInputs, bottomswivelInputs;
                ClevisHangerInputs topclevishangerInputs, middleclevishangerInputs, bottomclevishangerInputs;
                EyeNutInputs topeyenutInputs, middleeyenutInputs, bottomeyenutInputs;
                ClevisInputs topclevisInputs, middleclevisInputs, bottomclevisInputs;

                string topUbolt = "", middleUbolt = "", bottomUbolt = "", topPlate = "", middlePlate = "", bottomPlate = "", topLugplate = "", middleLugplate = "", bottomLugplate = "";
                string topattWBABolt = "", middleattWBABolt = "", bottomattWBABolt = "", topattWBAHole = "", middleattWBAHole = "", bottomattWBAHole = "", topconWBABolt = "", middleconWBABolt = "", bottomconWBABolt = "";
                string topSwivel = "", middleSwivel = "", bottomSwivel = "", topClevishanger = "", middleClevishanger = "", bottomClevishanger = "", topClevis = "", middleClevis = "", bottomClevis = "";
                string topEyenut = "", middleEyenut = "", bottomEyenut = "",stiffener = "", endplate="";

                Boolean btopUbolt = false, bmiddleUbolt = false, bbottomUbolt = false, btopPlate = false, bmiddlePlate = false, bbottomPlate = false, btopLugPlate = false, bmiddleLugPlate = false, bbottomLugPlate = false;
                Boolean btopattWBABolt = false, bmiddleattWBABolt = false, bbottomattWBABolt = false, btopattWBAHole = false, bmiddleattWBAHole = false, bbottomattWBAHole = false, btopconWBABolt = false, bmiddleconWBABolt = false, bbottomconWBABolt = false;
                Boolean btopSwivel = false, bmiddleSwivel = false, bbottomSwivel = false, btopClevis = false, bmiddleClevis = false, bbottomClevis = false;
                Boolean btopEyenut = false, bmiddleEyenut = false, bbottomEyenut = false, bstiffener = false, bendplate=false;
                Boolean bTopAttachment = true, bMiddleAttachment = true, bTopMiddleAttachment = false, bBotMiddleAttachment = false, bBottomAttachment = true, bTopConnection = true, bMiddleConnection = true, bBottomConnection = true;

                double multi1Qty, multi1LocateBy, multi1Location, EndPlateVerticalOffset;
                Matrix4X4 matrix = new Matrix4X4();



                //=============================='
                //'Load The Shape Data Structures'
                //'=============================='

                //Top Attachment
                if (!string.IsNullOrEmpty(SpreaderBeam.TopShape) & SpreaderBeam.TopShape != "No Value")
                {
                    switch (SpreaderBeam.TopAtt)
                    {
                        case 1:
                            topLugplate = SpreaderBeam.TopShape;
                            btopLugPlate = true;
                            break;
                        case 2:
                            topattWBABolt = SpreaderBeam.TopShape;
                            btopattWBABolt = true;
                            break;
                        case 3:
                            topUbolt = SpreaderBeam.TopShape;
                            btopUbolt = true;
                            break;
                        case 4:
                            topPlate = SpreaderBeam.TopShape;
                            btopPlate = true;
                            break;
                        case 5:
                            topattWBAHole = SpreaderBeam.TopShape;
                            btopattWBAHole = true;
                            break;
                        case -1:
                            bTopAttachment = false;
                            break;
                    }
                }
                else
                    bTopAttachment = false;

                // Middle Attachment
                if (!string.IsNullOrEmpty(SpreaderBeam.MiddleShape) & SpreaderBeam.MiddleShape != "No Value")
                {
                    switch (SpreaderBeam.MiddleAtt)
                    {
                        case 1:
                            middleLugplate = SpreaderBeam.MiddleShape;
                            bmiddleLugPlate = true;
                            break;
                        case 2:
                            middleattWBABolt = SpreaderBeam.MiddleShape;
                            bmiddleattWBABolt = true;
                            break;
                        case 3:
                            middleUbolt = SpreaderBeam.MiddleShape;
                            bmiddleUbolt = true;
                            break;
                        case 4:
                            middlePlate = SpreaderBeam.MiddleShape;
                            bmiddlePlate = true;
                            break;
                        case 5:
                            middleattWBAHole = SpreaderBeam.MiddleShape;
                            bmiddleattWBAHole = true;
                            break;
                        case -1:
                            bMiddleAttachment = false;
                            break;
                    }
                }
                else
                    bMiddleAttachment = false;

                // Bottom Attachment
                if (!string.IsNullOrEmpty(SpreaderBeam.BottomShape) & SpreaderBeam.BottomShape != "No Value")
                {
                    switch (SpreaderBeam.BottomAtt)
                    {
                        case 1:
                            bottomLugplate = SpreaderBeam.BottomShape;
                            bbottomLugPlate = true;
                            break;
                        case 2:
                            bottomattWBABolt = SpreaderBeam.BottomShape;
                            bbottomattWBABolt = true;
                            break;
                        case 3:
                            bottomUbolt = SpreaderBeam.BottomShape;
                            bbottomUbolt = true;
                            break;
                        case 4:
                            bottomPlate = SpreaderBeam.BottomShape;
                            bbottomPlate = true;
                            break;
                        case 5:
                            bottomattWBAHole = SpreaderBeam.BottomShape;
                            bbottomattWBAHole = true;
                            break;
                        case -1:
                            bBottomAttachment = false;
                            break;
                    }
                }
                else
                    bBottomAttachment = false;

                // Top Connection
                if (!string.IsNullOrEmpty(SpreaderBeam.TopConShape) & SpreaderBeam.TopConShape != "No Value")
                {
                    switch (SpreaderBeam.TopCon)
                    {
                        case 1:
                            topClevis = SpreaderBeam.TopConShape;
                            btopClevis = true;
                            break;
                        case 2:
                            topEyenut = SpreaderBeam.TopConShape;
                            btopEyenut = true;
                            break;
                        case 3:
                            topSwivel = SpreaderBeam.TopConShape;
                            btopSwivel = true;
                            break;
                        case 4:
                            topconWBABolt = SpreaderBeam.TopConShape;
                            btopconWBABolt = true;
                            break;
                        case -1:
                            bTopConnection = false;
                            break;
                    }
                }
                else
                    bTopConnection = false;

                // Middle Connection
                if (!string.IsNullOrEmpty(SpreaderBeam.MidConShape) & SpreaderBeam.MidConShape != "No Value")
                {
                    switch (SpreaderBeam.MiddleCon)
                    {
                        case 1:
                            middleClevis = SpreaderBeam.MidConShape;
                            bmiddleClevis = true;
                            break;
                        case 2:
                            middleEyenut = SpreaderBeam.MidConShape;
                            bmiddleEyenut = true;
                            break;
                        case 3:
                            middleSwivel = SpreaderBeam.MidConShape;
                            bmiddleSwivel = true;
                            break;
                        case 4:
                            middleconWBABolt = SpreaderBeam.MidConShape;
                            bmiddleconWBABolt = true;
                            break;
                        case -1:
                            bMiddleConnection = false;
                            break;
                    }
                }
                else
                    bMiddleConnection = false;

                // Bottom Connection
                if (!string.IsNullOrEmpty(SpreaderBeam.BotConShape) & SpreaderBeam.BotConShape != "No Value")
                {
                    switch (SpreaderBeam.BottomCon)
                    {
                        case 1:
                            bottomClevis = SpreaderBeam.BotConShape;
                            bbottomClevis = true;
                            break;
                        case 2:
                            bottomEyenut = SpreaderBeam.BotConShape;
                            bbottomEyenut = true;
                            break;
                        case 3:
                            bottomSwivel = SpreaderBeam.BotConShape;
                            bbottomSwivel = true;
                            break;
                        case 4:
                            bottomconWBABolt = SpreaderBeam.BotConShape;
                            bbottomconWBABolt = true;
                            break;
                        case -1:
                            bBottomConnection = false;
                            break;
                    }
                }
                else
                    bBottomConnection = false;
                // Stiffener
                stiffener = SpreaderBeam.Stiffener;
                if (!string.IsNullOrEmpty(stiffener) & stiffener != "No Value")
                {
                    bstiffener = true;
                }
                // Endplate 
                endplate = SpreaderBeam.SpreaderBeamEndPlate;

                if (!string.IsNullOrEmpty(endplate) & endplate != "No Value")
                {
                    bendplate = true;
                }
                
                EndPlateVerticalOffset = SpreaderBeam.EndPlateVerticalOffset;

            #endregion

            #region Load shape data
                //Load Data for all Components

                //Load Ubolt data by query
                topuboltInputs = LoadUBoltDataByQuery(topUbolt);
                middleuboltInputs = LoadUBoltDataByQuery(middleUbolt);
                bottomuboltInputs = LoadUBoltDataByQuery(bottomUbolt);
                //Load Plate data by query
                topplateInputs = LoadPlateDataByQuery(topPlate);
                middleplateInputs = LoadPlateDataByQuery(middlePlate);
                bottomplateinputs = LoadPlateDataByQuery(bottomPlate);
                //Load Lug Plate data by query
                toplugplateInputs = LoadPlateDataByQuery(topLugplate);
                middlelugplateInputs = LoadPlateDataByQuery(middleLugplate);
                bottomlugplateinputs = LoadPlateDataByQuery(bottomLugplate);
                //Load Top WBA Bolt data by query
                topattWBABoltInputs = LoadWBABoltDataByQuery(topattWBABolt);
                middleattWBABoltInputs = LoadWBABoltDataByQuery(middleattWBABolt);
                bottomattWBABoltInputs = LoadWBABoltDataByQuery(bottomattWBABolt);
                //Load Top WBA Hole data by query
                topattWBAHoleInputs = LoadWBAHoleDataByQuery(topattWBAHole);
                middleattWBAHoleInputs = LoadWBAHoleDataByQuery(middleattWBAHole);
                bottomattWBAHoleInputs = LoadWBAHoleDataByQuery(bottomattWBAHole);
                //Load Bottom WBA data by query
                topconWBABoltInputs = LoadWBABoltDataByQuery(topconWBABolt);
                middleconWBABoltInputs = LoadWBABoltDataByQuery(middleconWBABolt);
                bottomconWBABoltInputs = LoadWBABoltDataByQuery(bottomconWBABolt);
                //Load Swivel data by query
                topswivelInputs = LoadSwivelDataByQuery(topSwivel);
                middleswivelInputs = LoadSwivelDataByQuery(middleSwivel);
                bottomswivelInputs = LoadSwivelDataByQuery(bottomSwivel);
                //Load Clevis data by query
                topclevishangerInputs = LoadClevisHangerDataByQuery(topClevishanger);
                middleclevishangerInputs = LoadClevisHangerDataByQuery(middleClevishanger);
                bottomclevishangerInputs = LoadClevisHangerDataByQuery(bottomClevishanger);
                //Load Clevis Hanger data by query
                topclevisInputs = LoadClevisDataByQuery(topClevis);
                middleclevisInputs = LoadClevisDataByQuery(middleClevis);
                bottomclevisInputs = LoadClevisDataByQuery(bottomClevis);
                //Load EyeNut data by query
                topeyenutInputs = LoadEyeNutDataByQuery(topEyenut);
                middleeyenutInputs = LoadEyeNutDataByQuery(middleEyenut);
                bottomeyenutInputs = LoadEyeNutDataByQuery(bottomEyenut);
                //Load Stiffener data by query
                stiffenerInputs = LoadPlateDataByQuery(stiffener);
                //Load endPlates data by query
                endplateInputs = LoadPlateDataByQuery(endplate);
                #endregion

                StringBuilder error = new StringBuilder();
                ////Warn for Common Errors
                if (SpreaderBeam.Length <= 0)
                {
                    error.Append(SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidSteelLength, "Length of the section cannot be zero"));
                }
                if (SpreaderBeam.SteelName =="")
                {
                    error.Append(SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidSteelName, "SteelName can not be Empty"));
                }
                if (SpreaderBeam.SteelStandard == "")
                {
                    error.Append(SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidSteelStandard, "SteelStandard can not be Empty"));
                }
                if (SpreaderBeam.SteelType  == "")
                {
                    error.Append(SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrInvalidSteelType, "SteelTye can not be Empty"));
                }


                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                #region Check for placement type
                // Check for Placement Type
                int iTopPlacementType = 0;
                iTopPlacementType = SpreaderBeam.TopPlacementType;
                // Centre
                if (iTopPlacementType == 1)
                {
                    bTopAttachment = false;
                    bTopMiddleAttachment = true;
                }
                //Edge
                else if (iTopPlacementType == 2)
                {
                    bTopAttachment = true;
                    bTopMiddleAttachment = false;
                }
                //Center and Edge
                else if (iTopPlacementType == 3)
                {
                    bTopAttachment = true;
                    bTopMiddleAttachment = true;
                }

                int iBotPlacementType = 0;
                iBotPlacementType = SpreaderBeam.BotPlacementType;
                //Center
                if (iBotPlacementType == 1)
                {
                    bBotMiddleAttachment = true;
                    bBottomAttachment = false;
                }
                //Edge
                else if (iBotPlacementType == 2)
                {
                    bBotMiddleAttachment = false;
                    bBottomAttachment = true;
                }
                //Center and Edge
                else if (iBotPlacementType == 3)
                {
                    bBotMiddleAttachment = true;
                    bBottomAttachment = true;
                }

                //Check for no Middle Attachments Case
                if (bTopMiddleAttachment == false & bBotMiddleAttachment == false)
                {
                    bMiddleAttachment = false;
                }
                else
                {
                    bMiddleAttachment = true;
                }

                // Offsets
                int iTopOffsetDef = 0;
                int iBottomOffsetDef = 0;
                int iMiddleOffsetDef = 0;
                double TopOffset1 = 0;
                double TopOffset2 = 0;
                double BottomOffset1 = 0;
                double BottomOffset2 = 0;
                double MiddleOffset = 0;

                iTopOffsetDef = SpreaderBeam.TopOffsetDef;
                iBottomOffsetDef = SpreaderBeam.MiddleOffsetDef;
                iMiddleOffsetDef = SpreaderBeam.BotOffsetDef;

                // Offsets for Top Attachments                
                if (iTopOffsetDef == 1)  // centre
                {
                    TopOffset1 = SpreaderBeam.TopOff1;
                    TopOffset2 = SpreaderBeam.TopOff2;                   
                }
                else     //edge
                {
                    TopOffset1 = SpreaderBeam.Length / 2 - SpreaderBeam.TopOff1;
                    TopOffset2 = SpreaderBeam.Length / 2 - SpreaderBeam.TopOff2;
                }

                // Offsets for Bottom Attachments                
                if (iBottomOffsetDef == 1)       //center
                {
                    BottomOffset1 = SpreaderBeam.BottomOff1;
                    BottomOffset2 = SpreaderBeam.BottomOff2;                    
                }
                else      //Edge
                {
                    BottomOffset1 = SpreaderBeam.Length / 2 - SpreaderBeam.BottomOff1;
                    BottomOffset2 = SpreaderBeam.Length / 2 - SpreaderBeam.BottomOff2;
                }


                // Offsets for Middle Attachment                
                if (iMiddleOffsetDef == 1)        //center
                {
                    MiddleOffset = SpreaderBeam.MiddleOff;
                }
                else
                {
                    MiddleOffset = SpreaderBeam.MiddleOff - SpreaderBeam.Length / 2;
                }
                #endregion

                // Shoe Height
                double TotalHeight = 0;
                TotalHeight = SpreaderBeam.Diameter1 / 2 + SpreaderBeam.ShoeHeight / 2;

                // Add Beam                
                SteelMember tSteel = new SteelMember();
                double dSectionDepth = 0;
                double dWebThickness = 0;
                double dSectionLength = 0;
                double dFlangeThickness = 0;

                tSteel = GetSectionDataFromSection(SpreaderBeam.SteelStandard, SpreaderBeam.SteelType, SpreaderBeam.SteelName);
                dSectionDepth = tSteel.depth;
                dWebThickness = tSteel.webThickness;
                dSectionLength = SpreaderBeam.Length;
                dFlangeThickness = tSteel.flangeThickness;

                if (SpreaderBeam.BBGap > 0)
                {
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    matrix.Translate(new Vector(0, SpreaderBeam.BBGap / 2 + tSteel.width / 2, -TotalHeight));
                    DrawStandardSection(SpreaderBeam.SteelStandard, SpreaderBeam.SteelType, SpreaderBeam.SteelName, SpreaderBeam.Length, SpreaderBeam.SteelAngle, SpreaderBeam.SteelCpoint, m_oSymbolic.Outputs, matrix, "Beam1");
                    matrix.SetIdentity();
                    matrix.Rotate(-Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    matrix.Translate(new Vector(0, -SpreaderBeam.BBGap / 2 - tSteel.width / 2, -TotalHeight));
                    DrawStandardSection(SpreaderBeam.SteelStandard, SpreaderBeam.SteelType, SpreaderBeam.SteelName, SpreaderBeam.Length, SpreaderBeam.SteelAngle, SpreaderBeam.SteelCpoint, m_oSymbolic.Outputs, matrix, "Beam2");
                }
                else
                {
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    matrix.Translate(new Vector(0, 0, -TotalHeight));
                    DrawStandardSection(SpreaderBeam.SteelStandard, SpreaderBeam.SteelType, SpreaderBeam.SteelName, SpreaderBeam.Length, SpreaderBeam.SteelAngle, SpreaderBeam.SteelCpoint, m_oSymbolic.Outputs, matrix, "Beam1");
                }

                //Orientation for Attachnments
                long topAttachOrient1 = 0;
                long topAttachOrient2 = 0;
                long middleAttachOrient1 = 0;
                long middleAttachOrient2 = 0;
                long botAttachOrient1 = 0;
                long botAttachOrient2 = 0;

                //Rotations for Connections(Swivel Hanger)
                long topConRot1 = 1;
                long topConRot2 = 1;
                long middleConRot1 = 1;
                long middleConRot2 = 1;
                long botConRot1 = 1;
                long botConRot2 =1;

                topAttachOrient1 = SpreaderBeam.TopAtt1Orient;
                topAttachOrient2 = SpreaderBeam.TopAtt2Orient;
                middleAttachOrient1 = SpreaderBeam.MidAtt1Orient;
                middleAttachOrient2 = SpreaderBeam.MidAtt2Orient;
                botAttachOrient1 = SpreaderBeam.BotAtt1Orient;
                botAttachOrient2 = SpreaderBeam.BotAtt2Orient;

                topConRot1 = SpreaderBeam.TopCon1Rot;
                topConRot2 = SpreaderBeam.TopCon2Rot;
                middleConRot1 = SpreaderBeam.MidCon1Rot;
                middleConRot2 = SpreaderBeam.MidCon2Rot;
                botConRot1 = SpreaderBeam.BotCon1Rot;
                botConRot2 = SpreaderBeam.BotCon2Rot;

                #region Add top attachment
                //Add Top Attachment
                if (bTopAttachment == true)
                {
                    if (btopLugPlate == true) //adding lug plate
                    {
                        if (topAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-TopOffset1  - toplugplateInputs.width1 / 2, -TotalHeight, -toplugplateInputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddPlate(toplugplateInputs, matrix, m_oSymbolic.Outputs, "TopLug1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-toplugplateInputs.width1 / 2, -TotalHeight, -TopOffset1  - toplugplateInputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(toplugplateInputs, matrix, m_oSymbolic.Outputs, "TopLug1");
                        }
                        if (topAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(TopOffset2 - toplugplateInputs.width1 / 2, -TotalHeight, -toplugplateInputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddPlate(toplugplateInputs, matrix, m_oSymbolic.Outputs, "TopLug2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-toplugplateInputs.width1 / 2, -TotalHeight, TopOffset2 - toplugplateInputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(toplugplateInputs, matrix, m_oSymbolic.Outputs, "TopLug2");
                        }
                    }
                    else if (btopattWBABolt == true) //adding WBA bolt
                    {
                        if (topAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-TopOffset1 , 0, TotalHeight));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddWBABolt(topattWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltAtt1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, -TopOffset1 , TotalHeight));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBABolt(topattWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltAtt1");
                        }
                        if (topAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(TopOffset2, 0, TotalHeight));
                            matrix.Rotate(Math.PI , new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddWBABolt(topattWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltAtt2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, TopOffset2, TotalHeight));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBABolt(topattWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltAtt2");
                        }
                    }
                    else if (btopUbolt == true) //adding ubolt
                    {
                        if (topAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, TopOffset1 , -TotalHeight + (topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddUBolt(topuboltInputs, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "TopUbolt1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-TopOffset1 , 0, -TotalHeight + (topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) / 2));
                            AddUBolt(topuboltInputs, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "TopUbolt1");
                        }
                        if (topAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, -TopOffset2, -TotalHeight + (topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddUBolt(topuboltInputs, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "TopUbolt2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(TopOffset2, 0, -TotalHeight + (topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) / 2));
                            AddUBolt(topuboltInputs, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "TopUbolt2");
                        }
                    }
                    else if (btopPlate == true) //adding plate
                    {
                        if (topAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-TopOffset1  - topplateInputs.width1 / 2, -topplateInputs.length1 / 2, -TotalHeight));
                            AddPlate(topplateInputs, matrix, m_oSymbolic.Outputs, "TopPlate1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-topplateInputs.width1 / 2, TopOffset1  - topplateInputs.length1 / 2, -TotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(topplateInputs, matrix, m_oSymbolic.Outputs, "TopPlate1");
                        }
                        if (topAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(TopOffset2 - topplateInputs.width1 / 2, -topplateInputs.length1 / 2, -TotalHeight));
                            //matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(topplateInputs, matrix, m_oSymbolic.Outputs, "TopPlate2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-topplateInputs.width1 / 2, -TopOffset2 - topplateInputs.length1 / 2, -TotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(topplateInputs, matrix, m_oSymbolic.Outputs, "TopPlate2");
                        }
                    }
                    else if (btopattWBAHole == true) //adding WBA Hole
                    {
                        if (topAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-TopOffset1 , 0, TotalHeight));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddWBAHole(topattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "TopWBAHoleAtt1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, -TopOffset1 , TotalHeight));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBAHole(topattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "TopWBAHoleAtt1");
                        }
                        if (topAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(TopOffset2, 0, TotalHeight));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddWBAHole(topattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "TopWBAHoleAtt2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, TopOffset2, TotalHeight));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBAHole(topattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "TopWBAHoleAtt2");
                        }
                    }
                }
                #endregion

                #region Add middle attachment
                //Add middle Attachment
                if (bMiddleAttachment == true)
                {
                    if (bmiddleLugPlate == true) //adding lug plate
                    {
                        if (bTopMiddleAttachment == true)
                        {
                            if (middleAttachOrient1 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset  - middlelugplateInputs.width1 / 2, -TotalHeight, -middlelugplateInputs.thickness1 / 2));
                                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                                AddPlate(middlelugplateInputs, matrix, m_oSymbolic.Outputs, "MiddleLug1");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(-middlelugplateInputs.width1 / 2, -TotalHeight, MiddleOffset  - middlelugplateInputs.thickness1 / 2));
                                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddPlate(middlelugplateInputs, matrix, m_oSymbolic.Outputs, "MiddleLug1");
                            }
                        }
                        if (bBotMiddleAttachment == true)
                        {
                            if (middleAttachOrient2 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(-MiddleOffset  - middlelugplateInputs.width1 / 2, dSectionDepth + TotalHeight, -middlelugplateInputs.thickness1 / 2));
                                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                AddPlate(middlelugplateInputs, matrix, m_oSymbolic.Outputs, "MiddleLug2");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(-middlelugplateInputs.width1 / 2, dSectionDepth + TotalHeight, MiddleOffset  - middlelugplateInputs.thickness1 / 2));
                                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddPlate(middlelugplateInputs, matrix, m_oSymbolic.Outputs, "MiddleLug2");
                            }
                        }
                    }
                    else if (bmiddleattWBABolt == true) //adding WBA bolt
                    {
                        if (bTopMiddleAttachment == true)
                        {
                            if (middleAttachOrient1 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset , 0, TotalHeight));
                                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                AddWBABolt(middleattWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltAtt1");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, MiddleOffset , TotalHeight));
                                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddWBABolt(middleattWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltAtt1");
                            }
                        }
                        if (bBotMiddleAttachment == true)
                        {
                            if (middleAttachOrient2 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset , 0, -dSectionDepth - TotalHeight));
                                AddWBABolt(middleattWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltAtt2");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, -MiddleOffset , -dSectionDepth - TotalHeight));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddWBABolt(middleattWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltAtt2");
                            }
                        }
                    }
                    else if (bmiddleUbolt == true) //adding ubolt
                    {
                        if (bTopMiddleAttachment == true)
                        {
                            if (middleAttachOrient1 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, -MiddleOffset , -TotalHeight + (middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) / 2));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddUBolt(middleuboltInputs, middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "MiddleUbolt1");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset , 0, -TotalHeight + (middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) / 2));
                                AddUBolt(middleuboltInputs, middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "MiddleUbolt1");
                            }
                        }
                        if (bBotMiddleAttachment == true)
                        {
                            if (middleAttachOrient2 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, MiddleOffset , dSectionDepth + TotalHeight + (middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) / 2));
                                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddUBolt(middleuboltInputs, middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "MiddleUbolt2");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset , 0, dSectionDepth + TotalHeight + (middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) / 2));
                                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                AddUBolt(middleuboltInputs, middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "MiddleUbolt2");
                            }
                        }
                    }
                    else if (bmiddlePlate == true) //adding plate
                    {
                        if (bTopMiddleAttachment == true)
                        {
                            if (middleAttachOrient1 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset  - middleplateInputs.width1 / 2, -middleplateInputs.length1 / 2, -TotalHeight));
                                AddPlate(middleplateInputs, matrix, m_oSymbolic.Outputs, "MiddlePlate1");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(-middleplateInputs.width1 / 2, -MiddleOffset  - middleplateInputs.length1 / 2, -TotalHeight));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddPlate(middleplateInputs, matrix, m_oSymbolic.Outputs, "MiddlePlate1");
                            }
                        }
                        if (bBotMiddleAttachment == true)
                        {
                            if (middleAttachOrient2== 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset - middleplateInputs.width1 / 2, -middleplateInputs.length1 / 2, -dSectionDepth - TotalHeight - middleplateInputs.thickness1));
                                AddPlate(middleplateInputs, matrix, m_oSymbolic.Outputs, "MiddlePlate2");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(-middleplateInputs.width1 / 2, -MiddleOffset - middleplateInputs.length1 / 2, -dSectionDepth - TotalHeight - middleplateInputs.thickness1));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddPlate(middleplateInputs, matrix, m_oSymbolic.Outputs, "MiddlePlate2");
                            }
                        }
                    }
                    else if (bmiddleattWBAHole == true) //adding WBA Hole
                    {
                        if (bTopMiddleAttachment == true)
                        {
                            if (middleAttachOrient1 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset , 0, TotalHeight));
                                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                AddWBAHole(middleattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "MiddleWBAHoleAtt1");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, MiddleOffset , TotalHeight));
                                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddWBAHole(middleattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "MiddleWBAHoleAtt1");
                            }
                        }
                        if (bBotMiddleAttachment == true)
                        {
                            if (middleAttachOrient2 == 2)
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(MiddleOffset , 0, -dSectionDepth - TotalHeight));
                                //matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                AddWBAHole(middleattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "MiddleWBAHoleAtt2");
                            }
                            else
                            {
                                matrix = new Matrix4X4();
                                matrix.SetIdentity();
                                matrix.Translate(new Vector(0, -MiddleOffset , -dSectionDepth - TotalHeight));
                                //matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                AddWBAHole(middleattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "MiddleWBAHoleAtt2");
                            }
                        }
                    }
                }
                #endregion

                #region Add bottom attachment
                //Add Bottom Attachment
                if (bBottomAttachment == true)
                {                    
                    if (bbottomLugPlate == true) //Lug Plate
                    {
                        if (botAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(BottomOffset1 - bottomlugplateinputs.width1 / 2, dSectionDepth + TotalHeight, -bottomlugplateinputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(bottomlugplateinputs, matrix, m_oSymbolic.Outputs, "BottomLug1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-bottomlugplateinputs.width1 / 2, dSectionDepth + TotalHeight, -BottomOffset1 - bottomlugplateinputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI/2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(bottomlugplateinputs, matrix, m_oSymbolic.Outputs, "BottomLug1");
                        }
                        if (botAttachOrient2== 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-BottomOffset2 - bottomlugplateinputs.width1 / 2, dSectionDepth + TotalHeight, -bottomlugplateinputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            AddPlate(bottomlugplateinputs, matrix, m_oSymbolic.Outputs, "BottomLug2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-bottomlugplateinputs.width1 / 2, dSectionDepth + TotalHeight, BottomOffset2 - bottomlugplateinputs.thickness1 / 2));
                            matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(bottomlugplateinputs, matrix, m_oSymbolic.Outputs, "BottomLug2");
                        }

                    }
                    else if (bbottomattWBABolt == true) //WBA Bolt
                    {
                        if (botAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-BottomOffset1, 0, -dSectionDepth - TotalHeight));
                            AddWBABolt(bottomattWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltAtt1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, BottomOffset1, -dSectionDepth - TotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBABolt(bottomattWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltAtt1");
                        }

                        if (botAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(BottomOffset2, 0, -dSectionDepth - TotalHeight));
                            AddWBABolt(bottomattWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltAtt2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, -BottomOffset2, -dSectionDepth - TotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBABolt(bottomattWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltAtt2");
                        }
                        
                    }
                    else if (bbottomUbolt == true)  //U Bolt
                    {
                        if (botAttachOrient1== 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, -BottomOffset1, dSectionDepth + TotalHeight + (bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia) / 2));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddUBolt(bottomuboltInputs, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "BottomUbolt1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-BottomOffset1, 0, dSectionDepth + TotalHeight + (bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia) / 2));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddUBolt(bottomuboltInputs, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "BottomUbolt1");
                        }

                        if (botAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, BottomOffset2, dSectionDepth + TotalHeight + (bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia) / 2));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddUBolt(bottomuboltInputs, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "BottomUbolt2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(BottomOffset2, 0, dSectionDepth + TotalHeight + (bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia) / 2));
                            matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                            AddUBolt(bottomuboltInputs, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia, matrix, m_oSymbolic.Outputs, "BottomUbolt2");
                        }                        
                    }
                    else if (bbottomPlate == true)   //Plate
                    {
                        if (botAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-BottomOffset1 - bottomplateinputs.width1 / 2, -bottomplateinputs.length1 / 2, -dSectionDepth - TotalHeight - bottomplateinputs.thickness1));
                            AddPlate(bottomplateinputs, matrix, m_oSymbolic.Outputs, "BottomPlate1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-bottomplateinputs.width1 / 2, BottomOffset1 - bottomplateinputs.length1 / 2, -dSectionDepth - TotalHeight - bottomplateinputs.thickness1));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(bottomplateinputs, matrix, m_oSymbolic.Outputs, "BottomPlate1");
                        }

                        if (botAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(BottomOffset2 - bottomplateinputs.width1 / 2, -bottomplateinputs.length1 / 2, -dSectionDepth - TotalHeight - bottomplateinputs.thickness1));
                            AddPlate(bottomplateinputs, matrix, m_oSymbolic.Outputs, "BottomPlate2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-bottomplateinputs.width1 / 2, -BottomOffset2 - bottomplateinputs.length1 / 2, -dSectionDepth - TotalHeight - bottomplateinputs.thickness1));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddPlate(bottomplateinputs, matrix, m_oSymbolic.Outputs, "BottomPlate2");
                        }                        
                    }
                    else if (bbottomattWBAHole == true)   //WBA Hole
                    {
                        if (botAttachOrient1 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(-BottomOffset1, 0, -dSectionDepth - TotalHeight));
                            AddWBAHole(bottomattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "BottomWBAHoleAtt1");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, BottomOffset1, -dSectionDepth - TotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBAHole(bottomattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "BottomWBAHoleAtt1");
                        }

                        if (botAttachOrient2 == 2)
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(BottomOffset2, 0, -dSectionDepth - TotalHeight));
                            AddWBAHole(bottomattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "BottomWBAHoleAtt2");
                        }
                        else
                        {
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();
                            matrix.Translate(new Vector(0, -BottomOffset2, -dSectionDepth - TotalHeight));
                            matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                            AddWBAHole(bottomattWBAHoleInputs, matrix, m_oSymbolic.Outputs, "BottomWBAHoleAtt2");
                        }
                    }
                }
                #endregion

                #region Add top connection
                //Add Top Connection
                if (bTopAttachment == true)
                {
                    if (bTopConnection == true)
                    {
                        if (btopLugPlate == true)
                        {                            
                            if (btopClevis == true)     // Clevis
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, toplugplateInputs.length1 / 2 - TotalHeight));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , toplugplateInputs.length1 / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, toplugplateInputs.length1 / 2 - TotalHeight));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, toplugplateInputs.length1 / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                            } 
                            else if (btopEyenut == true)        //EyeNut
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + toplugplateInputs.length1 / 2 + topeyenutInputs.Thickness1 / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + toplugplateInputs.length1 / 2 + topeyenutInputs.Thickness1 / 2 - TotalHeight));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection, matrix);
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + toplugplateInputs.length1 / 2 + topeyenutInputs.Thickness1 / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection, matrix);

                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + toplugplateInputs.length1 / 2 + topeyenutInputs.Thickness1 / 2 - TotalHeight));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection, matrix);
                                }
                            }                           
                            else if (btopSwivel == true)        // Swivel
                            {                                
                                if (topConRot1 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset1 , 0, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                    
                                }
                                else if (topConRot1 == 2)      // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");
                                }
                                else if (topConRot1 == 3)        // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                    
                                }
                                else if (topConRot1 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset1, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset1, 0, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");
                                }
                                
                                
                                if (topConRot2 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset2, 0, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                   
                                }
                                else if (topConRot2 == 4)        // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset2, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset2, 0, -(topswivelInputs.Height1 + topswivelInputs.Thickness1 + toplugplateInputs.length1 / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");
                                }
                            }
                            
                            else if (btopconWBABolt == true)        //WBA Bolt
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, topconWBABoltInputs.Height2 + toplugplateInputs.length1 / 2 - TotalHeight));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , topconWBABoltInputs.Height2 + toplugplateInputs.length1 / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon1");
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, topconWBABoltInputs.Height2 + toplugplateInputs.length1 / 2 - TotalHeight));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, topconWBABoltInputs.Height2 + toplugplateInputs.length1 / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon2");
                                }
                            }
                        }
                        else if (btopattWBABolt == true)
                        {                           
                            if (btopClevis == true)      // Clevis
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - topclevisInputs.Pin1Diameter / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - topclevisInputs.Pin1Diameter / 2 - TotalHeight));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - topclevisInputs.Pin1Diameter / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - topclevisInputs.Pin1Diameter / 2 - TotalHeight));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                            }
                            
                            else if (btopEyenut == true)        //EyeNut
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection,matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, TopOffset1 , topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection,matrix);
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(TopOffset2, 0, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection,matrix);

                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, -TopOffset2, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection,matrix);
                                }
                            }
                            // Swivel
                            else if (bTopConnection == true)
                            {                                
                                if (topConRot1 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset1 , 0, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                    
                                }
                                else if (topConRot1 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                  
                                }
                                else if (topConRot1 == 3)         // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                    
                                }
                                else if (topConRot1 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset1 , -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset1 , 0, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1"); 
                                 }
                                if (topConRot2 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset2, 0, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 2)      // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset2, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset2, 0, -(topswivelInputs.Height1 + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");
                                }

                            }
                        }
                        else if (btopUbolt == true)
                        {                            
                            if (btopClevis == true)     // Clevis
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight - topclevisInputs.Pin1Diameter / 2));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight - topclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight - topclevisInputs.Pin1Diameter / 2));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight - topclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                            }
                            
                            else if (btopEyenut == true)        //EyeNut
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, TopOffset1 , topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection,matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection, matrix);
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, -TopOffset2, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection,matrix);

                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(TopOffset2, 0, topeyenutInputs.InnerLength2 - topeyenutInputs.OverLength1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - TotalHeight));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection, matrix);
                                }
                            }
                            
                            else if (btopSwivel == true)        // Swivel
                            {                                
                                if (topConRot1 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                    
                                }
                                else if (topConRot1 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset1 , -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                    
                                }
                                else if (topConRot1 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset1 , 0, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");                                    
                                }
                                else if (topConRot1 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1, 0, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel1");
                                }
                               
                                if (topConRot2 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset2, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset2, 0, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");                                    
                                }
                                else if (topConRot2 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, -(topswivelInputs.Height1 + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddSwivelRing(topswivelInputs, 0, matrix, m_oSymbolic.Outputs, "TopSwivel2");
                                }
                            }
                            else if (btopconWBABolt == true)        //WBA Bolt
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topconWBABoltInputs.Height2 - topconWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topconWBABoltInputs.Height2 - topconWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon1");
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topconWBABoltInputs.Height2 - topconWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topconWBABoltInputs.Height2 -topconWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon2");
                                }
                            }
                        }
                        else if (btopPlate == true)
                        {                            
                            if (btopClevis == true)     // Clevis
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset1 , 0, -(topclevisInputs.Opening1 + topclevisInputs.Pin1Diameter + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset1 , -(topclevisInputs.Opening1 + topclevisInputs.Pin1Diameter + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis1");
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset2, 0, -(topclevisInputs.Opening1 + topclevisInputs.Pin1Diameter + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset2, -(topclevisInputs.Opening1 + topclevisInputs.Pin1Diameter + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(topclevisInputs, matrix, m_oSymbolic.Outputs, "TopClevis2");
                                }
                            }
                            
                            else if (btopEyenut == true)        //EyeNut
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(TopOffset1 , 0, -(topeyenutInputs.Nut.ShapeLength + topeyenutInputs.OverLength1 + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, TopOffset1 , -(topeyenutInputs.Nut.ShapeLength + topeyenutInputs.OverLength1 + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut1", out eyeNutObjectCollection, matrix);
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(-TopOffset2, 0, -(topeyenutInputs.Nut.ShapeLength + topeyenutInputs.OverLength1 + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection, matrix);

                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, -TopOffset2, -(topeyenutInputs.Nut.ShapeLength + topeyenutInputs.OverLength1 + topplateInputs.thickness1) + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(topeyenutInputs, m_oSymbolic.Outputs, "TopEyeNut2", out eyeNutObjectCollection, matrix);
                                }
                                
                            }
                            else if (btopconWBABolt == true)        //WBA Bolt
                            {
                                if (topAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-TopOffset1 , 0, -topplateInputs.thickness1 + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -TopOffset1 , -topplateInputs.thickness1 + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon1");
                                }

                                if (topAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(TopOffset2, 0, -topplateInputs.thickness1 + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, TopOffset2, -topplateInputs.thickness1 + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(topconWBABoltInputs, matrix, m_oSymbolic.Outputs, "TopWBABoltCon2");
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Add middle connection
                //Add Middle Connection
                // Connections will be added only if attachments are present
                if (bMiddleAttachment == true)
                {
                    if (bMiddleConnection == true)
                    {
                        if (bmiddleLugPlate == true)
                        {                            
                            if (bmiddleClevis == true)      // Clevis
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middlelugplateInputs.length1 / 2 - TotalHeight));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middlelugplateInputs.length1 / 2 - TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, dSectionDepth + TotalHeight + middlelugplateInputs.length1 / 2 ));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector( 0,-MiddleOffset , dSectionDepth + TotalHeight + middlelugplateInputs.length1 / 2 ));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                }                                
                            }
                            else if (bmiddleEyenut == true)     //EyeNut
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middlelugplateInputs.length1 / 2 + middleeyenutInputs.Thickness1 / 2 - TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middlelugplateInputs.length1 / 2 + middleeyenutInputs.Thickness1 / 2 - TotalHeight));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);

                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middlelugplateInputs.length1 / 2 + middleeyenutInputs.Thickness1 / 2 + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middlelugplateInputs.length1 / 2 + middleeyenutInputs.Thickness1 / 2 + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                }
                            }                            
                            else if (bmiddleSwivel == true)     // Swivel
                            {
                                if (bTopMiddleAttachment == true)
                                {                                    
                                    if (middleConRot1 == 1)     // 0 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 - TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 2)        // 90 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 - TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 3)        // 180 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 - TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 4)        // 270 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 - TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset, 0, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 - TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {                                    
                                    if (middleConRot2 == 1)     // 0 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 2)        // 90 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 3)        // 180 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 4)        // 270 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset, 0, -(middleswivelInputs.Height1 + middleswivelInputs.Thickness1 + middlelugplateInputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");
                                    }
                                }
                            }                            
                            else if (bmiddleconWBABolt == true)     //WBA Bolt
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleconWBABoltInputs.Height2 + middlelugplateInputs.length1 / 2 - TotalHeight));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleconWBABoltInputs.Height2 + middlelugplateInputs.length1 / 2 - TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, dSectionDepth + TotalHeight + middleconWBABoltInputs.Height2 + middlelugplateInputs.length1 / 2));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset , dSectionDepth + TotalHeight + middleconWBABoltInputs.Height2 + middlelugplateInputs.length1 / 2));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon2");
                                    }
                                }
                            }
                        }
                        else if (bmiddleattWBABolt == true)
                        {                            
                            if (bmiddleClevis == true)      // Clevis
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 -middleclevisInputs.Pin1Diameter/2 - TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 -middleclevisInputs.Pin1Diameter/2  - TotalHeight));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , dSectionDepth + TotalHeight + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2-middleclevisInputs.Pin1Diameter/2));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, dSectionDepth + TotalHeight + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2-middleclevisInputs.Pin1Diameter/2));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                }                                
                            }
                            else if (bmiddleEyenut == true)     //EyeNut
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);

                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                }                                
                            }
                            else if (bmiddleSwivel == true)     // Swivel
                            {
                                if (bTopMiddleAttachment == true)
                                {                                    
                                    if (middleConRot1 == 1)     // 0 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 2)        // 90 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 3)        // 180 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 4)        // 270 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset, 0, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {                                    
                                    if (middleConRot2 == 1)     // 0 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 2)        // 90 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 3)        // 180 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 4)        // 270 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset, 0, -(middleswivelInputs.Height1 + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");
                                    }
                                }
                            }
                        }
                        else if (bmiddleUbolt == true)
                        {                            
                            if (bmiddleClevis == true)              // Clevis
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia - TotalHeight - middleclevisInputs.Pin1Diameter / 2));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia - TotalHeight - middleclevisInputs.Pin1Diameter / 2));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, dSectionDepth + TotalHeight + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia - middleclevisInputs.Pin1Diameter / 2));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , dSectionDepth + TotalHeight + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia - middleclevisInputs.Pin1Diameter / 2));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                }                                
                            }
                            else if (bmiddleEyenut == true)             //EyeNut
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, MiddleOffset , middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia - TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia - TotalHeight));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);

                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleeyenutInputs.InnerLength2 - middleeyenutInputs.OverLength1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                }
                            }                                
                            else if (bmiddleSwivel == true)             // Swivel           
                            {
                                if (bTopMiddleAttachment == true)
                                {                                    
                                    if (middleConRot1 == 1)         // 0 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 2)        // 90 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset , -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 3)        // 180 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");                                        
                                    }
                                    else if (middleConRot1 == 4)           // 270 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset, 0, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {                                    
                                    if (middleConRot2 == 1)         // 0 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 2)        // 90 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset , -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 3)        // 180 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");                                        
                                    }
                                    else if (middleConRot2 == 4)        // 270 Degrees
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset, 0, -(middleswivelInputs.Height1 + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddSwivelRing(middleswivelInputs, 0, matrix, m_oSymbolic.Outputs, "MiddleSwivel2");
                                    }
                                }
                            }                            
                            else if (bmiddleconWBABolt == true)             //WBA Bolt
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + middleconWBABoltInputs.Height2 - middleconWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + middleconWBABoltInputs.Height2 - middleconWBABoltInputs.Pin1Diameter / 2 - TotalHeight));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon1");
                                    }

                                    if (middleAttachOrient2 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + middleconWBABoltInputs.Height2 - middleconWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset , middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + middleconWBABoltInputs.Height2 - middleconWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon2");
                                    }
                                }
                            }
                        }
                        else if (bmiddlePlate == true)
                        {                            
                            if (bmiddleClevis == true)          // Clevis
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleclevisInputs.Opening1 + middleclevisInputs.Pin1Diameter + middleplateInputs.thickness1) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleclevisInputs.Opening1 + middleclevisInputs.Pin1Diameter + middleplateInputs.thickness1) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleclevisInputs.Opening1 + middleclevisInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight + middleplateInputs.thickness1)));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleclevisInputs.Opening1 + middleclevisInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight + middleplateInputs.thickness1)));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddClevis(middleclevisInputs, matrix, m_oSymbolic.Outputs, "MiddleClevis2");
                                    }
                                }                                
                            }
                            else if (bmiddleEyenut == true)         //EyeNut
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(-MiddleOffset , 0, -(middleeyenutInputs.Nut.ShapeLength + middleeyenutInputs.OverLength1 + middleplateInputs.thickness1) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleeyenutInputs.Nut.ShapeLength + middleeyenutInputs.OverLength1 + middleplateInputs.thickness1) + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut1", out eyeNutObjectCollection, matrix);

                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(middleeyenutInputs.Nut.ShapeLength + middleeyenutInputs.OverLength1 + middleplateInputs.thickness1 + dSectionDepth + TotalHeight)));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                    else
                                    {
                                        ArrayList eyeNutObjectCollection;
                                        matrix = new Matrix4X4();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(middleeyenutInputs.Nut.ShapeLength + middleeyenutInputs.OverLength1 + middleplateInputs.thickness1 + dSectionDepth + TotalHeight)));
                                        matrix.Rotate(Math.PI/2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddEyeNutwithLocation(middleeyenutInputs, m_oSymbolic.Outputs, "MiddleEyeNut2", out eyeNutObjectCollection, matrix);
                                    }
                                }                                
                            }
                            else if (bmiddleconWBABolt == true)                 //WBA Bolt
                            {
                                if (bTopMiddleAttachment == true)
                                {
                                    if (middleAttachOrient1 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -middleplateInputs.thickness1 + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon1");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, MiddleOffset , -middleplateInputs.thickness1 + TotalHeight));
                                        matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon1");
                                    }
                                }
                                if (bBotMiddleAttachment == true)
                                {
                                    if (middleAttachOrient2 == 2)
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleplateInputs.thickness1)));
                                        //matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon2");
                                    }
                                    else
                                    {
                                        matrix = new Matrix4X4();
                                        matrix.SetIdentity();
                                        matrix.Translate(new Vector(0, -MiddleOffset , -(dSectionDepth + TotalHeight + middleplateInputs.thickness1)));
                                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                        AddWBABolt(middleconWBABoltInputs, matrix, m_oSymbolic.Outputs, "MiddleWBABoltCon2");
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Add bottom connection
                // Add Bottom Connection
                if (bBottomAttachment == true)
                {
                    if (bBottomConnection == true)
                    {
                        if (bbottomLugPlate == true)
                        {                            
                            if (bbottomClevis == true)          // Clevis
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset1, dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset2, dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }                                
                            }
                            else if (bbottomEyenut == true)             //EyeNut
                            {

                                if (botAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, BottomOffset1, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomlugplateinputs.length1 / 2 + bottomeyenutInputs.Thickness1 / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomlugplateinputs.length1 / 2 + bottomeyenutInputs.Thickness1 / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);

                                }
                                if (botAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, -BottomOffset2, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomlugplateinputs.length1 / 2 + bottomeyenutInputs.Thickness1 / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection, matrix);

                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(BottomOffset2, 0, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomlugplateinputs.length1 / 2 + bottomeyenutInputs.Thickness1 / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection, matrix);
                                }                                
                            }
                            else if (bbottomSwivel == true)                 // Swivel
                            {                                
                                if (botConRot1 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset1, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset1, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");
                                }
                                
                                if (botConRot2 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset2, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset2, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1 + bottomlugplateinputs.length1 / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");
                                }
                            }
                            else if (bbottomconWBABolt == true)             //WBA Bolt
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, dSectionDepth + TotalHeight + bottomconWBABoltInputs.Height2 + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset1, dSectionDepth + TotalHeight + bottomconWBABoltInputs.Height2 + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon1");
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, dSectionDepth + TotalHeight + bottomconWBABoltInputs.Height2 + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset2, dSectionDepth + TotalHeight + bottomconWBABoltInputs.Height2 + bottomlugplateinputs.length1 / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon2");
                                }
                            }
                        }
                        else if (bbottomattWBABolt == true)
                        {                            
                            if (bbottomClevis == true)              // Clevis
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset1, dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset2, dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }                                
                            }
                            else if (bbottomEyenut == true)         //EyeNut
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(BottomOffset1, 0, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, BottomOffset1, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, -BottomOffset2, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection, matrix);
                                }                                
                            }
                            else if (bbottomSwivel == true)         // Swivel
                            {                                
                                if (botConRot1 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset1, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset1, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");
                                }

                                if (botConRot2 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset2, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset2, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");
                                }
                            }
                        }
                        else if (bbottomUbolt == true)
                        {                            
                            if (bbottomClevis == true)          // Clevis
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset1, dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset2, dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia - bottomclevisInputs.Pin1Diameter / 2));
                                    matrix.Rotate(Math.PI, new Vector(0, 1, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }                                
                            }
                            else if (bbottomEyenut == true)             //EyeNut
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, BottomOffset1, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, -BottomOffset2, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(BottomOffset2, 0, bottomeyenutInputs.InnerLength2 - bottomeyenutInputs.OverLength1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection, matrix);
                                }                                
                            }
                            else if (bbottomSwivel == true)             // Swivel
                            {                                
                                if (botConRot1 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset1, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");                                    
                                }
                                else if (botConRot1 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset1, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset1, 0, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel1");
                                }
                                
                                if (botConRot2 == 1)        // 0 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 2)       // 90 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset2, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                    
                                }
                                else if (botConRot2 == 3)       // 180 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");                                   
                                }
                                else if (botConRot2 == 4)       // 270 Degrees
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset2, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset2, 0, -(bottomswivelInputs.Height1 + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + dSectionDepth + TotalHeight)));
                                    matrix.Rotate(Math.PI, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddSwivelRing(bottomswivelInputs, 0, matrix, m_oSymbolic.Outputs, "BottomSwivel2");
                                }
                            }
                            else if (bbottomconWBABolt == true)      //WBA Bolt
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomconWBABoltInputs.Height2 - bottomconWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset1, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomconWBABoltInputs.Height2 - bottomconWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon1");
                                }
                                if (botAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomconWBABoltInputs.Height2 - bottomconWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset2, bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomconWBABoltInputs.Height2 - bottomconWBABoltInputs.Pin1Diameter / 2 + dSectionDepth + TotalHeight));
                                    matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon2");
                                }
                            }
                        }
                        else if (bbottomPlate == true)
                        {                            
                            if (bbottomClevis == true)              // Clevis
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(bottomclevisInputs.Opening1 + bottomclevisInputs.Pin1Diameter + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset1, -(bottomclevisInputs.Opening1 + bottomclevisInputs.Pin1Diameter + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis1");
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(bottomclevisInputs.Opening1 + bottomclevisInputs.Pin1Diameter + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset2, -(bottomclevisInputs.Opening1 + bottomclevisInputs.Pin1Diameter + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddClevis(bottomclevisInputs, matrix, m_oSymbolic.Outputs, "BottomClevis2");
                                }                                
                            }
                            else if (bbottomEyenut == true)     //EyeNut
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(bottomeyenutInputs.Nut.ShapeLength + bottomeyenutInputs.OverLength1 + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, BottomOffset1, -(bottomeyenutInputs.Nut.ShapeLength + bottomeyenutInputs.OverLength1 + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut1", out eyeNutObjectCollection, matrix);
                                }

                                if (botAttachOrient2 == 2)
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(bottomeyenutInputs.Nut.ShapeLength + bottomeyenutInputs.OverLength1 + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection, matrix);
                                }
                                else
                                {
                                    ArrayList eyeNutObjectCollection;
                                    matrix = new Matrix4X4();
                                    matrix.Translate(new Vector(0, -BottomOffset2, -(bottomeyenutInputs.Nut.ShapeLength + bottomeyenutInputs.OverLength1 + dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddEyeNutwithLocation(bottomeyenutInputs, m_oSymbolic.Outputs, "BottomEyeNut2", out eyeNutObjectCollection,matrix);
                                }                                
                            }
                            else if (bbottomconWBABolt == true)         //WBA Bolt
                            {
                                if (botAttachOrient1 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon1");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, BottomOffset1, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon1");
                                }
                                if (botAttachOrient2 == 2)
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon2");
                                }
                                else
                                {
                                    matrix = new Matrix4X4();
                                    matrix.SetIdentity();
                                    matrix.Translate(new Vector(0, -BottomOffset2, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1)));
                                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                                    AddWBABolt(bottomconWBABoltInputs, matrix, m_oSymbolic.Outputs, "BottomWBABoltCon2");
                                }
                            }
                        }
                    }
                }
                #endregion

                #region Add stiffeners
                // Add Right Stiffeners
                multi1Qty = SpreaderBeam.Multi1Qty;
                multi1LocateBy = SpreaderBeam.Multi1LocateBy;
                multi1Location = SpreaderBeam.Multi1Location;


                int I = 0;
                double[] dXLocation = new double[Convert.ToInt16(multi1Qty)];
                if (bstiffener == true)
                {
                    
                    for (I = 0; I <= multi1Qty - 1; I++)
                    {
                        dXLocation[I] = MultiPosition(dSectionLength, multi1Qty, multi1LocateBy, multi1Location, stiffenerInputs.thickness1)[I];
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(dWebThickness / 2, -dSectionDepth - TotalHeight + dFlangeThickness, -dSectionLength / 2 - stiffenerInputs.thickness1 / 2 + dXLocation[I]));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        AddPlate(stiffenerInputs, matrix, m_oSymbolic.Outputs, "Stiffener1" + Convert.ToString(I));
                    }
                }

                // Add Left stiffners
                if (bstiffener == true)
                {
                    //double[] dXLocation = new double[Convert.ToInt16(multi1Qty)];
                    for (I = 0; I <= multi1Qty - 1; I++)
                    {
                        dXLocation[I] = MultiPosition(dSectionLength, multi1Qty, multi1LocateBy, multi1Location, stiffenerInputs.thickness1)[I];
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(-stiffenerInputs.width1 - dWebThickness / 2, -dSectionDepth - TotalHeight + dFlangeThickness, -dSectionLength / 2 - stiffenerInputs.thickness1 / 2 + dXLocation[I]));
                        matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                        matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                        AddPlate(stiffenerInputs, matrix, m_oSymbolic.Outputs, "Stiffener2" + Convert.ToString(I));
                    }
                }
                #endregion

                #region Add Endplates
                if (bendplate ==true)
                {
                    //Add Left EndPlate
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-(endplateInputs.width1 / 2), -dSectionDepth - TotalHeight-EndPlateVerticalOffset, -dSectionLength / 2 - endplateInputs.thickness1));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    AddPlate(endplateInputs, matrix, m_oSymbolic.Outputs, "EndPLate1");
                    //Add Right EndPlate
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(-endplateInputs.width1 / 2, -dSectionDepth - TotalHeight-EndPlateVerticalOffset, dSectionLength / 2));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    AddPlate(endplateInputs, matrix, m_oSymbolic.Outputs, "EndPlate2");
                }
                #endregion

                #region AddPorts
                //   Add the Ports
                // Add "Route" port and "BotMiddlePort" port
                Port port1 = new Port(connection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Route"] = port1;
                if (bBotMiddleAttachment == false) //No Middle Attachment and No Middle Connection
                {
                    Port port2 = new Port(connection, part, "BotMiddleport", new Position(0, 0, -dSectionDepth - TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oSymbolic.Outputs["BotMiddleport"] = port2;
                }                
                else if (bMiddleConnection == false)    // With only Middle Attachment
                {
                    if (bBotMiddleAttachment == true)
                    {
                        if (bmiddleLugPlate == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -dSectionDepth - TotalHeight - middlelugplateInputs.length1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleattWBABolt == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -dSectionDepth - TotalHeight - middleattWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleUbolt == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -dSectionDepth - TotalHeight - (middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia) / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddlePlate == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -dSectionDepth - TotalHeight - middleplateInputs.thickness1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleattWBAHole)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -dSectionDepth - TotalHeight - middleattWBAHoleInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                    }
                    else
                    {
                        Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset, 0, -dSectionDepth - TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["BotMiddleport"] = port2;
                    }                    
                }
                else            // With Middle Attachment and Middle Connection
                {               
                    if (bmiddleLugPlate == true)
                    {
                        if (bmiddleClevis == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + middlelugplateInputs.length1 / 2 + TotalHeight + middleclevisInputs.Opening1 + middleclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleEyenut == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleeyenutInputs.InnerLength2 + middleeyenutInputs.Nut.ShapeLength + middleeyenutInputs.Thickness1 / 2 + middlelugplateInputs.length1 / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleSwivel == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middlelugplateInputs.length1 / 2 + middleswivelInputs.Height1 + middleswivelInputs.Thickness1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleconWBABolt == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middlelugplateInputs.length1 / 2 + middleconWBABoltInputs.Height2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                    }
                    else if (bmiddleattWBABolt == true)
                    {
                        if (bmiddleClevis == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 - middleclevisInputs.Pin1Diameter / 2 + middleclevisInputs.Opening1 + middleclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleEyenut == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + middleeyenutInputs.InnerLength2 + middleeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleSwivel == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleattWBABoltInputs.Height2 - middleattWBABoltInputs.Pin1Diameter / 2 + middleswivelInputs.Height1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                    }
                    else if (bmiddleUbolt == true)
                    {
                        if (bmiddleClevis == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia - middleclevisInputs.Pin1Diameter / 2 + middleclevisInputs.Opening1 + middleclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleEyenut == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + middleeyenutInputs.InnerLength2 + middleeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleSwivel == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + middleswivelInputs.Height1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleconWBABolt == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleuboltInputs.UBoltWidth - middleuboltInputs.UBoltRodDia + middleconWBABoltInputs.Height2 - middleconWBABoltInputs.Pin1Diameter / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                    }
                    else if (bmiddlePlate == true)
                    {
                        if (bmiddleClevis == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleplateInputs.thickness1 + middleclevisInputs.Opening1 + middleclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleEyenut == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleplateInputs.thickness1 + middleeyenutInputs.InnerLength2 + middleeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                        else if (bmiddleconWBABolt == true)
                        {
                            Port port2 = new Port(connection, part, "BotMiddleport", new Position(MiddleOffset , 0, -(dSectionDepth + TotalHeight + middleplateInputs.thickness1 + middleconWBABoltInputs.Height2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["BotMiddleport"] = port2;
                        }
                    }
                }
                // Add "Topport1" port and "Topport2" port


                if (bTopAttachment == false) //No Top Attachment and No Top Connection
                {
                    Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oSymbolic.Outputs["Topport1"] = port3;
                    m_oSymbolic.Outputs["Topport2"] = port4;
                }
                else if (bTopConnection == false) // With only Top Attachment
                {
                    if (btopLugPlate == true)
                    {
                        Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + toplugplateInputs.length1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + toplugplateInputs.length1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Topport1"] = port3;
                        m_oSymbolic.Outputs["Topport2"] = port4;
                    }
                    else if (btopattWBABolt == true)
                    {
                        Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + topattWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topattWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Topport1"] = port3;
                        m_oSymbolic.Outputs["Topport2"] = port4;
                    }
                    else if (btopUbolt == true)
                    {
                        Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + (topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia) / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + SpreaderBeam.Diameter1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Topport1"] = port3;
                        m_oSymbolic.Outputs["Topport2"] = port4;
                    }
                    else if (btopPlate == true)
                    {
                        Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + topplateInputs.thickness1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topplateInputs.thickness1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Topport1"] = port3;
                        m_oSymbolic.Outputs["Topport2"] = port4;
                    }
                    else if (btopattWBAHole)
                    {
                        Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + topattWBAHoleInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topattWBAHoleInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Topport1"] = port3;
                        m_oSymbolic.Outputs["Topport2"] = port4;
                    }
                }
                else // With Top Attachment and Top Connection
                {
                    if (btopLugPlate == true)
                    {
                        if (btopClevis == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + toplugplateInputs.length1 / 2 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + toplugplateInputs.length1 / 2 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1)); 
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopEyenut == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength + topeyenutInputs.Thickness1 / 2 + toplugplateInputs.length1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength + topeyenutInputs.Thickness1 / 2 + toplugplateInputs.length1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopSwivel == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + toplugplateInputs.length1 / 2 + topswivelInputs.Height1 + topswivelInputs.Thickness1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + toplugplateInputs.length1 / 2 + topswivelInputs.Height1 + topswivelInputs.Thickness1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopconWBABolt == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + toplugplateInputs.length1 / 2 + topconWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + toplugplateInputs.length1 / 2 + topconWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                    }
                    else if (btopattWBABolt == true)
                    {
                        if (btopClevis == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - topclevisInputs.Pin1Diameter / 2 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 - topclevisInputs.Pin1Diameter / 2 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopEyenut == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopSwivel == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1 , 0, -TotalHeight + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 + topswivelInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topattWBABoltInputs.Height2 - topattWBABoltInputs.Pin1Diameter / 2 + topswivelInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopconWBABolt == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                    }
                    else if (btopUbolt == true)
                    {
                        if (btopClevis == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - topclevisInputs.Pin1Diameter / 2 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia - topclevisInputs.Pin1Diameter / 2 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopEyenut == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopSwivel == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topswivelInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topswivelInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopconWBABolt == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topconWBABoltInputs.Height2 - topconWBABoltInputs.Pin1Diameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topuboltInputs.UBoltWidth - topuboltInputs.UBoltRodDia + topconWBABoltInputs.Height2 - topconWBABoltInputs.Pin1Diameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                    }
                    else if (btopPlate == true)
                    {
                        if (btopClevis == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topplateInputs.thickness1 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topplateInputs.thickness1 + topclevisInputs.Opening1 + topclevisInputs.nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopEyenut == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topplateInputs.thickness1 + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topplateInputs.thickness1 + topeyenutInputs.InnerLength2 + topeyenutInputs.Nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                        else if (btopconWBABolt == true)
                        {
                            Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topplateInputs.thickness1 + topconWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topplateInputs.thickness1 + topconWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Topport1"] = port3;
                            m_oSymbolic.Outputs["Topport2"] = port4;
                        }
                    }
                    else if (btopattWBAHole)
                    {
                        Port port3 = new Port(connection, part, "Topport1", new Position(-TopOffset1, 0, -TotalHeight + topplateInputs.thickness1 + topattWBAHoleInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port4 = new Port(connection, part, "Topport2", new Position(TopOffset2, 0, -TotalHeight + topplateInputs.thickness1 + topattWBAHoleInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Topport1"] = port3;
                        m_oSymbolic.Outputs["Topport2"] = port4;
                    }
                }

                // Add "Botport1" port and "Botport2" port      
                if (bBottomAttachment == false)//No Bottom Attachment and No Bottom Connection
                {
                    Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -dSectionDepth - TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -dSectionDepth - TotalHeight), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oSymbolic.Outputs["Botport1"] = port5;
                    m_oSymbolic.Outputs["Botport2"] = port6;
                }
                else if (bBottomConnection == false)// With only Bottom Attachment 
                {
                    if (bbottomLugPlate == true)
                    {
                        Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -dSectionDepth - TotalHeight - bottomlugplateinputs.length1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -dSectionDepth - TotalHeight - bottomlugplateinputs.length1 / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Botport1"] = port5;
                        m_oSymbolic.Outputs["Botport2"] = port6;
                    }
                    else if (bbottomattWBABolt == true)
                    {
                        Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -dSectionDepth - TotalHeight - bottomattWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -dSectionDepth - TotalHeight - bottomattWBABoltInputs.Height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Botport1"] = port5;
                        m_oSymbolic.Outputs["Botport2"] = port6;
                    }
                    else if (bbottomUbolt == true)
                    {
                        Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -dSectionDepth - TotalHeight - (bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia) / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -dSectionDepth - TotalHeight - (bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia) / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Botport1"] = port5;
                        m_oSymbolic.Outputs["Botport2"] = port6;
                    }
                    else if (bbottomPlate == true)
                    {
                        Port port5 = new Port(connection, part,  "Botport1", new Position(-BottomOffset1, 0, -dSectionDepth - TotalHeight - bottomplateinputs.thickness1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -dSectionDepth - TotalHeight - bottomplateinputs.thickness1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Botport1"] = port5;
                        m_oSymbolic.Outputs["Botport2"] = port6;
                    }
                    else if (bbottomattWBAHole == true)
                    {
                        Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -dSectionDepth - TotalHeight - bottomattWBAHoleInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -dSectionDepth - TotalHeight - bottomattWBAHoleInputs.Height1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                        m_oSymbolic.Outputs["Botport1"] = port5;
                        m_oSymbolic.Outputs["Botport2"] = port6;
                    }

                }
                else // With  Bottom Attachment and Bottom Connection
                {
                    if (bbottomLugPlate == true)
                    {
                        if (bbottomClevis == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomEyenut == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength + bottomeyenutInputs.Thickness1 / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength + bottomeyenutInputs.Thickness1 / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomSwivel == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomswivelInputs.Height1 + bottomswivelInputs.Thickness1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomconWBABolt == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomconWBABoltInputs.Height2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomlugplateinputs.length1 / 2 + bottomconWBABoltInputs.Height2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                    }
                    else if (bbottomattWBABolt == true)
                    {
                        if (bbottomClevis == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 - bottomclevisInputs.Pin1Diameter / 2 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 - bottomclevisInputs.Pin1Diameter / 2 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomEyenut == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomSwivel == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + bottomswivelInputs.Height1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomattWBABoltInputs.Height2 - bottomattWBABoltInputs.Pin1Diameter / 2 + bottomswivelInputs.Height1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                    }
                    else if (bbottomUbolt == true)
                    {
                        if (bbottomClevis == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia - bottomclevisInputs.Pin1Diameter / 2 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia - bottomclevisInputs.Pin1Diameter / 2 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomEyenut == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomSwivel == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomswivelInputs.Height1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomswivelInputs.Height1)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomconWBABolt == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomconWBABoltInputs.Height2 - bottomconWBABoltInputs.Pin1Diameter / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomuboltInputs.UBoltWidth - bottomuboltInputs.UBoltRodDia + bottomconWBABoltInputs.Height2 - bottomconWBABoltInputs.Pin1Diameter / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                    }
                    else if (bbottomPlate == true)
                    {
                        if (bbottomClevis == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1 + bottomclevisInputs.Opening1 + bottomclevisInputs.nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomEyenut == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1 + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1 + bottomeyenutInputs.InnerLength2 + bottomeyenutInputs.Nut.ShapeLength)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                        else if (bbottomconWBABolt == true)
                        {
                            Port port5 = new Port(connection, part, "Botport1", new Position(-BottomOffset1, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1 + bottomconWBABoltInputs.Height2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            Port port6 = new Port(connection, part, "Botport2", new Position(BottomOffset2, 0, -(dSectionDepth + TotalHeight + bottomplateinputs.thickness1 + bottomconWBABoltInputs.Height2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                            m_oSymbolic.Outputs["Botport1"] = port5;
                            m_oSymbolic.Outputs["Botport2"] = port6;
                        }
                    }
                }
                if (bendplate == true)
                {
                    Port port7 = new Port(connection, part, "EndPlatePort1", new Position(-dSectionLength / 2 - endplateInputs.thickness1 / 2, 0, -TotalHeight + (endplateInputs.length1 - EndPlateVerticalOffset - dSectionDepth) / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    Port port8 = new Port(connection, part, "EndPlatePort2", new Position(dSectionLength / 2 + endplateInputs.thickness1 / 2, 0, -TotalHeight + (endplateInputs.length1 - EndPlateVerticalOffset - dSectionDepth) / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));                   
                    m_oSymbolic.Outputs["EndPlatePort1"] = port7;
                    m_oSymbolic.Outputs["EndPlatePort2"] = port8;
                }
                #endregion

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SpreaderBeam.cs"));
                }
            }
        }
        #endregion

        #region "Caluculating WeightCG"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                string materialType, materialGrade;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                VolumeCG volumeCG =  supportComponent.GetVolumeAndCOG();

                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();

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


                Material material;
                double materialDensity;
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
                double weight,cogx,cogy,cogz;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = volumeCG.Volume * materialDensity;
                }

                try
                {
                    cogx = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogx = volumeCG.COGX;
                }

                try
                {
                    cogy = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogy = volumeCG.COGY;
                }

                try
                {
                    cogz = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogz = volumeCG.COGZ;
                }
                supportComponent.SetWeightAndCOG(weight, cogx, cogy, cogz);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of SpreaderBeam"));
                }
            }
        }
        #endregion

    }

}

