namespace ThreatDataBase
{
    internal class ThreatInfo
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public string Target { get; set; }
        public string BreachOfConfid { get; set; }
        public string IntegrityViolation { get; set; } 
        public string AccessibilityViolation { get; set; }


        public ThreatInfo(int id, string name, string description, string source, string target, string breachOfConfid, string integrityViolation, string accessibilityViolation)
        {
            this.Id = id;
            this.Name = name;
            this.Description = description;
            this.Source = source;
            this.Target = target;
            this.BreachOfConfid = breachOfConfid;
            this.IntegrityViolation = integrityViolation;
            this.AccessibilityViolation = accessibilityViolation;
        }

        public override bool Equals(object obj)
        {
            return obj is ThreatInfo info &&
                   Id == info.Id &&
                   Name == info.Name &&
                   Description == info.Description &&
                   Source == info.Source &&
                   Target == info.Target &&
                   BreachOfConfid == info.BreachOfConfid &&
                   IntegrityViolation == info.IntegrityViolation &&
                   AccessibilityViolation == info.AccessibilityViolation;
        }
    }
}
