using System;

namespace ChampionsLeague2021.Models
{
    class MatchInfo
    {
        public int id { get; set; }
        public Team awayTeam { get; set; }

        public Team homeTeam { get; set; }

        public MatchKickOff kickOffTime { get; set; }

        public GroupInformation group { get; set; }

        public Matchday matchday { get; set; }

        public string status { get; set; }

        public string type { get; set; }

        public Score score { get; set; }

        public string phase { get; set; }

        public TeamsValues fouls { get; set; }

        public Winner winner { get; set; }

    }

    public class MatchKickOff
    {
        public string date { get; set; }
        public DateTime dateTime { get; set; }
        public int utcOffsetInHours { get; set; }
    }

    public class TeamsValues
    {
        public int away { get; set; }

        public int home { get; set; }
    }

    public class Score
    {
        public TeamsValues total { get; set; }

        public TeamsValues regular { get; set; }

        public TeamsValues aggregate { get; set; }

        public TeamsValues penalty { get; set; }
    }

    public class Winner
    {
        public Match match { get; set; }
    }

    public class Match
    {
        public string reason { get; set; }

        public Team team { get; set; }
    }

}
