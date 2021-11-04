using System.Collections.Generic;
using ChampionsLeague2021.Enums;


namespace ChampionsLeague2021.Models
{
    public class  Player
    {
            public string Name { get; set; }

            public int ColumnWynikLocation { get; set; }

            public int ColumnRedCardLocation { get; set; }

            public int ColumnGroupQualificationLocation { get; set; }

            public int ColumnTourmentClassification { get; set; }

            public List<MatchResult> Matches { get; set; }

           // public List<QualificationResult> Qualifications { get; set; }

            public int WinTeamPoints { get; set; }

            public int CorrectScorePoints { get; set; }

            public int RedCardPoints { get; set; }

            public int GroupQualificationPoints { get; set; }

            public int TotalMachWinPoints { get; set; }

            public int TotalPoints { get; set; }

            public TourmentClassification TourmentClassification { get; set; }
                   
            public int TourmentClassificationPoints { get; set; }
        
    }
}
