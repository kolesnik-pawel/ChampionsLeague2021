namespace ChampionsLeague2021.Models
{
    public class GroupInformation
    {

        public string competitionId { get; set; }
        public string id { get; set; }
        public MetaData metaData { get; set; }
        public int order { get; set; }
        public string phase { get; set; }
        public int roundId { get; set; }
        public int seasonYear { get; set; }
        public Translations translations { get; set; }
        public string type { get; set; }

    }

    public class MetaData
    {
        public string groupName { get; set; }
        public string groupShortName { get; set; }
    }

    public class Translations
    {
        public enTrnas name { get; set; }
        public enTrnas shortName { get; set; }
    }
}
