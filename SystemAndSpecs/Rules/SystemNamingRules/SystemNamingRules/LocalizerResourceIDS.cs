/*******************************************************************
  Copyright (C) 1999, Intergraph Corporation.  All rights reserved.
'
'  Project: SystemsAndSpecs SystemNameRules
'  Class:   LocalizerResourceIDS
'
'  Abstract: The file contains Loaclizer ID's for System and Specs Name Rule
'
             04 june 2015 Mishra,Anamay  TR-CP-273693	HierarchyRoot Class doesn't support INamedItem Interface
'
'******************************************************************/
using Ingr.SP3D.Common.Middle.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ingr.SP3D.Content.System.Rules
{
    public static class LocalizerResourceIDS
    {

        public const string DEFAULT_ASSEMBL = "SystemNamingRules";
        public const string DEFAULT_RESOURC = "Ingr.SP3D.Content.System.Rules.SystemResource";
        public const int IDS_DUPLICATENAME = 1;
        public const int IDS_HASCHANGED = 2;
        public const int IDS_BLANKNAME = 3;
        public const int IDS_SYSTEM = 4;
        public const int IDS_NEW = 5;
        public const int IDS_GENERIC = 6;
        public const int IDS_CONDUIT = 7;
        public const int IDS_DUCTING = 8;
        public const int IDS_MACHINERY = 9;
        public const int IDS_PIPELINE = 10;
        public const int IDS_PIPING = 11;
        public const int IDS_STRUCTURAL = 12;
        public const int IDS_UNIT = 13;
        public const int IDS_AREA = 14;



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
