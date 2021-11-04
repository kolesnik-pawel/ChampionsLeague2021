namespace ChampionsLeague2021.Models
{
    public class TranslationsTeam
    {
        public enTrnas countryName { get; set; }
        public enTrnas displayName { get; set; }
        public enTrnas displayOfficialName { get; set; }
        public enTrnas displayTeamCode { get; set; }
    }

    public class enTrnas
    {
        public string EN { get; set; }
    }
}
