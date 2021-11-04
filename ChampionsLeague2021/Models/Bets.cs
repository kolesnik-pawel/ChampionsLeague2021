
using ChampionsLeague2021.Enums;

namespace ChampionsLeague2021.Models
{
    public class Bets
    {
        public TeamsValues RegularTimeWin { get; set; }

        public MatchResultsEnum winTeam { get; set; }

        public int PenatlyWin { get; set; }

        public bool? RedCard { get; set; }

        public int YellowCards { get; set; }

        public string GroupQualification { get; set; }

        public string Winner { get; set; }

        public string SecondPlace { get; set; }

        public string ThirdPlace { get; set; }

    }
}
