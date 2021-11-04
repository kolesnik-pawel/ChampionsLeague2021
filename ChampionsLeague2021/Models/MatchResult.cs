using ChampionsLeague2021.Enums;

namespace ChampionsLeague2021.Models
{
    public class MatchResult
    {
        public string homeTeam { get; set; }

        public string awayTeam { get; set; }

        public MatchResultsEnum winTeam { get; set; }

        public int ReadCards { get; set; }

        public int YellowCards { get; set; }

        public Score score { get; set; }

        public int penalty { get; set; }

        public Bets bets { get; set; }

        public BetsResult Points { get; set; }

    }
}
