//***************************************************************************
//   Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
//   
//   Project: EquipNamingRules
//
//   Abstract: The file contains Localizer Id's for EquipNamingRules
//
//   History:
//       Madhuri                      Mar/17/2015             CR-CP-269668 : Re-writing the existing EquipNamingRules name rule from vb6 to 3dapi
//***************************************************************************

using System;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Equipment.Rules
{
    public static  class LocalizerResourceIDS
    {
        public const string DEFAULT_ASSEMBL = "EquipmentNamingRules";
        public const string DEFAULT_RESOURC = "Ingr.SP3D.Content.Equipment.Rules.EquipmentResource";
        public const int IDS_STRING_SHAPE = 1;
        public const int IDS_STRING_DATUMSHAPE = 2;



        /// <summary>
        /// Retrieves the string from the resource.
        /// </summary>
        /// <param name="resourceId">Message ID.</param>
        /// <param name="defaultMessage">Default string to return.</param>
        public static string GetString(int resourceId, string defaultMessage)
        {
            return CmnLocalizer.GetString(resourceId, defaultMessage, LocalizerResourceIDS.DEFAULT_RESOURC, LocalizerResourceIDS.DEFAULT_ASSEMBL);
        }
    }
}
