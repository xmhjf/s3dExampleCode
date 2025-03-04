﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// SeqBasePlateNoHull rule class Implements PlnAssemblyRangeRuleBase
    /// </summary>
    public class SeqBasePlateNoHull : PlnAssemblySequenceRuleBase
    {
        /// <summary>
        /// Sequence parts with out hull
        /// </summary>
        /// <param name="sequenceData"></param>
        public override void Evaluate(SequenceData sequenceData)
        {
            #region Input Error Handling

            if (sequenceData == null)
            {
                throw new ArgumentNullException("sequenceData Information is null");
            }

            #endregion //Input Error Handling

            try
            {
                int basePlateCount = 0;
                Matrix4X4 viewMatrix = null;
                object[] unrelatedPart = null;

                BusinessObject assemblyBase = sequenceData.Assembly;
                Collection<BusinessObject> assyChildren = new Collection<BusinessObject>();
                Collection<PLATE_INFO> basePlateInfoColl = new Collection<PLATE_INFO>();
                foreach (BusinessObject children in sequenceData.AssemblyChildren)
                {
                    assyChildren.Add(children);
                }

                // Specify parts that will be sequenced
                int sequenceOption = SEQUENCE_FILTER_BASEPLATE_NON_HULL | SEQUENCE_FILTER_STIFFENING_PROFILE | SEQUENCE_FILTER_CONNECTED_PLATE_NON_HULL | SEQUENCE_FILTER_CONNECTED_PROFILE;
                int unrelatedPartCount = 0;
                CollectBasePlateInfo(assemblyBase, assyChildren, sequenceOption, basePlateInfoColl, ref basePlateCount, ref viewMatrix, ref unrelatedPartCount, ref unrelatedPart);

                // Sequence parts
                for (int basePlateIndex = 0; basePlateIndex <= basePlateCount - 1; basePlateIndex++)
                {
                    SequenceParts(assemblyBase, basePlateInfoColl[basePlateIndex], viewMatrix, SequenceType.SeqBasePlateNoHull);
                }

            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }
   
    }
}
