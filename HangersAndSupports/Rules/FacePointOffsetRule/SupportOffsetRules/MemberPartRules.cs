//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   MemberPartRules.cs
//   SupportOffsetRules,Ingr.SP3D.Content.Support.Rules.MemberPartGageRule
//   SupportOffsetRules,Ingr.SP3D.Content.Support.Rules.MemberPartGageRule2
//   SupportOffsetRules,Ingr.SP3D.Content.Support.Rules.MemberPartRatioRule
//   Author       :  BS
//   Creation Date:  26.Sept.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01.Nov.2013     BS      CR233078  Convert HgrSupOffsetRules to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using System;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class MemberPartGageRule : IHgrOffsetRule
    {
        public double GetFacePositionOffset(CrossSection crossSection, IPort port, double portLength, int cardinalPoint)
        {
            
            double flangeGage = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedBoltGage", "gf")).PropValue;

            return 0.5 * flangeGage;
        }
    }
    public class MemberPartGageRule2 : IHgrOffsetRule
    {
        public double GetFacePositionOffset(CrossSection crossSection, IPort port, double portLength, int cardinalPoint)
        {
            double longSideGage = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructAngleBoltGage", "lsg")).PropValue;

            return longSideGage;
        }
    }
    public class MemberPartRatioRule : IHgrOffsetRule
    {
        public double GetFacePositionOffset(CrossSection crossSection, IPort port, double portLength, int cardinalPoint)
        {
            double ratio = 0.25,offset;
            offset = ratio * Math.Round(portLength, 5);
            if (cardinalPoint == 3 || cardinalPoint == 9)
                offset = offset * -1;
            return offset;
        }
    }    
}



