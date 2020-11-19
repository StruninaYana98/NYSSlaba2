namespace ThreatDataBase
{
    internal class ThreatShortInfo
    {
        public string Idshort {get;set;}
        public string Nameshort { get; set; }

        public ThreatShortInfo(string id, string name)
        {
            Idshort = id;
            Nameshort = name;
        }
    }

}
