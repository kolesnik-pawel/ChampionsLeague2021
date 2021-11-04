
namespace ChampionsLeague2021.Models
{
    public class TourmentClassification
    {
        public string Winner { get; set; }

        public string SecondPlace { get; set; }

        public string ThirdPlace { get; set; }

        public Bets Bet { get; set; }

        public BetsResult Points { get; set; }
    }
}
