//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSLServiceClass.cs
//   PSL,Ingr.SP3D.Content.Support.Symbols.PSLServiceClass
//   Author       :  Hema
//   Creation Date:  21.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-Aug-2013     Hema      CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public static class PTPBOMServices
    {
        /// <summary>
        /// This method will be called to get UOMFormat of given Unittype and PrecisionType
        /// </summary>
        /// <param name="unit">type of Unit</param>
        /// <param name="precisionType">type of precision</param>
        /// <returns></returns>
        ///<code>
        ///GetUOMFormat(UnitType.Force, (PrecisionType)collection[0])
        ///</code>
        public static UOMFormat GetUOMFormat(UnitType unit, PrecisionType precisionType)
        {
            UOMFormat uomFormat = MiddleServiceProvider.UOMMgr.GetDefaultUnitFormat(unit);
            uomFormat.PrecisionType = precisionType;
            if (uomFormat.PrecisionType == PrecisionType.PRECISIONTYPE_FRACTIONAL)
                uomFormat.FractionalPrecision = uomFormat.FractionalPrecision;
            if (uomFormat.PrecisionType == PrecisionType.PRECISIONTYPE_DECIMAL)
                uomFormat.DecimalPrecision = uomFormat.DecimalPrecision;
            return uomFormat;
        }
    }
}