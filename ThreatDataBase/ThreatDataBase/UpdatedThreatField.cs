using System.Collections.Generic;
namespace ThreatDataBase
{
    class UpdatedThreatField
    {
        public string  Id{get;set;}
        public ThreatInfo FieldName { get; set; } = new ThreatInfo(0, "", "", "", "", "", "", "");
        public ThreatInfo Field { get; set; } = new ThreatInfo(0, "", "", "", "", "", "", "");
        public ThreatInfo UpdatedField { get; set; } = new ThreatInfo(0, "", "", "", "", "", "", "");
    }

}
