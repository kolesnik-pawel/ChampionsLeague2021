namespace ChampionsLeague2021.Models
{
    public class Team
    {
        public string countryCode { get; set; }
        public string associationLogoUrl { get; set; }
        public string bigLogoUrl { get; set; }
        public int id { get; set; }
        public string idProvider { get; set; }
        public string internationalName { get; set; }
        public string isPlaceHolder { get; set; }
        public string logoUrl { get; set; }
        public string mediumLogoUrl { get; set; }
        public string teamCode { get; set; }
        public string teamTypeDetail { get; set; }
        public TranslationsTeam translations { get; set; }

    }
}
