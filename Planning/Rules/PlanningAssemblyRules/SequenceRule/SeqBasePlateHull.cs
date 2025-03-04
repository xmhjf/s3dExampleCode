using System;
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
    /// SeqBasePlateHull rule class Implements PlnAssemblyRangeRuleBase
    /// </summary>
    public class SeqBasePlateHull : PlnAssemblySequenceRuleBase
    {
        /// <summary>
        /// Sequence parts with hull
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
                Collection<BusinessObject> assemblyChildren = new Collection<BusinessObject>();
                foreach (BusinessObject child in sequenceData.AssemblyChildren)
                {
                    assemblyChildren.Add(child);
                }

                // Specify parts that will be sequenced
                int sequenceFilter = (SEQUENCE_FILTER_BASEPLATE_NON_HULL | SEQUENCE_FILTER_BASEPLATE_HULL | SEQUENCE_FILTER_STIFFENING_PROFILE | SEQUENCE_FILTER_CONNECTED_PLATE_NON_HULL | SEQUENCE_FILTER_CONNECTED_PLATE_HULL | SEQUENCE_FILTER_CONNECTED_PROFILE);

                int unrelatedPartCount = 0;
                int basePlateCount = 0;
                object[] unrelatedPart = null;
                Collection<PLATE_INFO> basePlateInfoColl = new Collection<PLATE_INFO>();
                Matrix4X4 viewMatrix = null;

                CollectBasePlateInfo(sequenceData.Assembly, assemblyChildren, sequenceFilter, basePlateInfoColl, ref basePlateCount, ref viewMatrix, ref unrelatedPartCount, ref unrelatedPart);

                // Sequence all plate parts and profile parts for each deck
                for (int basePlateIndex = 0; basePlateIndex <= basePlateCount - 1; basePlateIndex++)
                {
                    SequenceParts(sequenceData.Assembly, basePlateInfoColl[basePlateIndex], viewMatrix, SequenceType.SeqBasePlateHull);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }
  
    }
}
