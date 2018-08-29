using System;
using System.Collections.Generic;

using ExtendOffice.Office.Summary;

namespace ExtendOffice.Office
{
    public interface IOfficeFile
    {
        Dictionary<String, String> DocumentSummaryInformation { get; }

        Dictionary<String, String> SummaryInformation { get; }
    }
}