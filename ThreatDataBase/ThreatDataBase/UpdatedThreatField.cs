using System.Collections.Generic;
namespace ThreatDataBase
{
    class UpdatedThreatField
    {
        public string  Id{get;set;}
     
        public ThreatInfo Fields { get; set; } = new ThreatInfo(0, "", "", "", "", "", "", "");
        public ThreatInfo UpdatedFields { get; set; } = new ThreatInfo(0, "", "", "", "", "", "", "");
    }

}
