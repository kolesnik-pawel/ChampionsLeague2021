
using System;

namespace ChampionsLeague2021.Models
{
    public class SimplyObject : IEquatable<SimplyObject>
    {
        public string Name { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            SimplyObject simplyObject = obj as SimplyObject;
            if (simplyObject == null) return false;
            else
            return base.Equals(simplyObject);
        }
        public bool Equals(SimplyObject other)
        {
            if (other == null) return false;
            return (this.Name.Equals(other.Name));
        }

        public SimplyObject ShallowCopy()
        {
            return (SimplyObject)this.MemberwiseClone();
        }
    }
}
