using System.Collections.Generic;
using Google.Apis.Sheets.v4.Data;
using ChampionsLeague2021.Enums;
using ChampionsLeague2021.Models;
using Newtonsoft.Json;
using System.Linq;
using System;
using System.IO;

namespace ChampionsLeague2021
{
   
    class AddPlayers
    {
        private bool playersUpdate = false;
        private ValueRange ReadPlayersFromSheets(SheetsEnum sheetName, string spreadsheetId)
        {
            SheetsHelper Sheets = new SheetsHelper();
            return Sheets.Read(sheetName, "A1", "A1000", spreadsheetId);
        }

        private List<SimplyObject> ReadPlayersFromJson(string path)
        {
            Console.WriteLine(Directory.GetCurrentDirectory());
            var file = File.ReadAllText(path);

            return  JsonConvert.DeserializeObject<List<SimplyObject>>(file);
        }

        private List<SimplyObject> ConvertReadedPlayersFromSheetToListOfPlyers(ValueRange valueRange)
        {
            List<SimplyObject> returned = new List<SimplyObject>();
            foreach(var value in valueRange.Values)
            {
               
                returned.Add(new SimplyObject() { Name = value[0].ToString() });
               
            }

            return returned;
        }

        public List<SimplyObject> CompareAndUpdatePlayersName(List<SimplyObject> readFromJson, List<SimplyObject> readFromSheets)
        {
            List<SimplyObject> tmpPlayerFromJson = new List<SimplyObject>();
            List<SimplyObject> tmpPlayerFromSheets = new List<SimplyObject>();

            foreach (var player in readFromJson)
            {
                tmpPlayerFromJson.Add(player);
            }
            foreach (var player in readFromSheets)
            {
                tmpPlayerFromSheets.Add(player);
            }

            // var tmpPlayerFromSheets = ref readFromSheets;
            foreach (var playerFromJson in tmpPlayerFromJson)
            {
                if (!readFromSheets.Contains(playerFromJson))
                {
                    readFromJson.Remove(playerFromJson);
                    playersUpdate = true;
                }
            }

            foreach (var playerFromSheets in tmpPlayerFromSheets)
            {
               if( !readFromJson.Contains(playerFromSheets))
                {
                    readFromJson.Add(playerFromSheets);
                    playersUpdate = true;
                }
            }

            return readFromJson;
        }

        public void SaveUpdatesPlayersToFile(string path, List<SimplyObject> updatedPlayers)
        {
            try
            {
                File.WriteAllText(path, JsonConvert.SerializeObject(updatedPlayers));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void ChackAndUpdatePlayers(string pathJsonPlayersFile, SheetsEnum sheetName, string spreadsheetId)
        {
            var readedPlayersFormJson = ReadPlayersFromJson(pathJsonPlayersFile);
            var readedPlayersFromSheets = ConvertReadedPlayersFromSheetToListOfPlyers(ReadPlayersFromSheets(sheetName, spreadsheetId));
            var comparedPlayers = CompareAndUpdatePlayersName(readedPlayersFormJson, readedPlayersFromSheets);

            if (playersUpdate)
            {
                Console.WriteLine(Directory.GetCurrentDirectory());
                SaveUpdatesPlayersToFile(pathJsonPlayersFile, comparedPlayers);
            }

        }

        public ValueRange AddPlayer(string path)
        {
             string yellowCardImg = "https://img.uefa.com/imgml/uefacom/euro2020/liveblog/icons/code_1.svg";
            var file = File.ReadAllText(path);

            var players = JsonConvert.DeserializeObject<List<Player>>(file);
            ValueRange valueRange = new ValueRange();
            var tmp = new List<IList<object>>();
            IList<object> contentList = new List<object>();
            IList<object> gameList = new List<object>();
            // dodaje userów
            foreach (var player in players)
            {
                contentList.Add(player.Name);
                contentList.Add("");
                contentList.Add($@"=IMAGE(""{yellowCardImg}"")");
                contentList.Add("");
            }
            tmp.Add(contentList);

            // dodaje typy zakładów
            foreach (var player in players)
            {
                gameList.Add("Wynik");
                gameList.Add("0:0");
                gameList.Add("");
                gameList.Add("2<5");
            }
            tmp.Add(gameList);
            valueRange.Values = tmp;

            return valueRange;
        }

        public void AddDataValidation(SheetsEnum sheet, int SheetGid)
        {
            string startCell = "B1";
            string endCell = "B70";
            var sh = new SheetsHelper();
            int rownumber = 0;
            var readed = sh.Read(sheet, startCell, endCell);
            var readPlayers = sh.Read(sheet, "L3", "Z3");
            var increment = 0;
            var startColum = Enums.ColumnEnum.L;

            foreach (var row in readed.Values)
            {
                if (row.Count != 0)
                {
                    var startColumLoop = startColum;

                    foreach (var player in readPlayers.Values.FirstOrDefault())
                    {
                        if (player.ToString() == "Czerwona")
                        {
                            System.Threading.Thread.Sleep(1010);
                           // sh.DataValidation(SheetGid, startColumLoop, startColumLoop + 1, rownumber, rownumber + 1);
                            startColumLoop++;
                            Console.WriteLine(increment++);
                        }
                        else
                        {
                            startColumLoop++;
                        }
                    }
                }
                rownumber++;
            }
        }

        private List<Player> ReadPlayers(ValueRange read, GamesEnum gameType, List<Player> players)
        {
            //znajdź i określ pozycję graczy
            //var columnCount = 0;
            foreach (var rows in read.Values)
            {
                var columnCount = 0;
                foreach (var col in rows)
                {
                    foreach (var player in players)
                    {
                        if (col.ToString() == player.Name)
                        {
                            switch (gameType)
                            {
                                case GamesEnum.Mach:
                                    player.ColumnWynikLocation = columnCount;
                                    player.ColumnRedCardLocation = columnCount + 1;
                                    if (player.Matches == null)
                                    {
                                        player.Matches = new List<MatchResult>();
                                    }
                                    break;
                                //case GamesEnum.GroupQualification:
                                //    player.ColumnGroupQualificationLocation = columnCount;
                                //    player.Qualifications = new List<QualificationResult>();
                                //    break;
                                //case GamesEnum.TourmentWinner:
                                //    player.ColumnTourmentClassification = columnCount;
                                //    player.TourmentClassification = new TourmentClassification();
                                //    break;
                            }
                        }
                    }
                    columnCount++;

                }

            }

            return players;
        }

        private List<Player> MatchBestResults(List<Player> players, ValueRange read, bool? finaleFase = false)
        {
            int increment = 0;
            foreach (var wynik in read.Values)
            {
                if (wynik.Count < 8)
                {
                    continue;
                }

                if (wynik[0].ToString() != "0"
                    && wynik[1].ToString() != ""
                    && wynik[1].ToString() != "Gospodarz"
                    && wynik[5].ToString() != "")
                {
                    foreach (var player in players)
                    {

                        var mach = new MatchResult();
                        mach.homeTeam = wynik[1].ToString();
                        mach.awayTeam = wynik[5].ToString();

                        mach.score = new Score();
                        mach.score.regular = SplitResult(wynik[6].ToString());
                        mach.score.penalty = SplitResult(wynik[7].ToString());
                        mach.winTeam = SetWinner(mach.score.regular.home, mach.score.regular.away);
                        //mach.penalty = wynik[8].ToString() >0 ? true : false;
                        mach.ReadCards = SetValue(wynik[9].ToString());

                        mach.bets = new Bets();
                        mach.bets.RegularTimeWin = SplitResult(wynik[player.ColumnWynikLocation].ToString());
                        mach.bets.RedCard = RedCardSet(wynik[player.ColumnRedCardLocation].ToString());
                        mach.bets.winTeam = SetWinner(mach.bets.RegularTimeWin.home, mach.bets.RegularTimeWin.away);
                        mach.Points = new BetsResult();
                        mach.Points = CountPoints(mach, finaleFase);
                        //mach.bets.TourmentClassification.
                        player.WinTeamPoints += mach.Points.WinTeamBets;
                        player.CorrectScorePoints += mach.Points.MachScoreBets;
                        player.RedCardPoints += mach.Points.RedCardBets;
                        player.TotalMachWinPoints = player.WinTeamPoints + player.CorrectScorePoints;
                        player.TotalPoints = player.TotalMachWinPoints + player.RedCardPoints + player.GroupQualificationPoints;
                        player.Matches.Add(mach);

                    }
                }
                increment++;
                Console.WriteLine($"Dodnao mecz nr {increment} z {read.Values.Count}");
            }

            return players;
        }

        //private BetsResult SetQualiPoint(bool qualification, string playerQualiSet)
        //{
        //    BetsResult result = new BetsResult();

        //    if (qualification == true)
        //    {
        //        switch (playerQualiSet)
        //        {
        //            case "Tak":
        //                result.GroupQualification = (int)PiontsEnum.GroupQualification;
        //                break;
        //            case "1st":
        //                result.GroupQualification = (int)PiontsEnum.GroupPositionQualification;
        //                break;
        //            case "2nd":
        //                result.GroupQualification = (int)PiontsEnum.GroupPositionQualification;
        //                break;

        //            default:
        //                result.GroupQualification = 0;
        //                break;
        //        }
        //    }
        //    return result;
        //}
        //private List<Player> QualificationBstResult(List<Player> players, ValueRange read)
        //{
        //    int increment = 0;
        //    foreach (var item in read.Values)
        //    {
        //        if (item.Count == 0)
        //        {
        //            continue;
        //        }
        //        if (item[0].ToString() != "")
        //        {
        //            var Team = new TeamAtGroup();
        //            foreach (var player in players)
        //            {
        //                var group = new QualificationResult();
        //                group.GroupName = item[0].ToString();
        //                group.TeamName = item[2].ToString();
        //                group.Qualified = item[11].ToString() == "FALSE" ? false : true;
        //                group.TeamPoints = int.Parse(item[10].ToString());
        //                group.FirstQuali = int.Parse(item[12].ToString()) == 1 ? true : false;
        //                group.SecondQuali = int.Parse(item[12].ToString()) == 2 ? true : false;
        //                group.Bet = new Bets();

        //                group.Bet.GroupQualification = item[player.ColumnGroupQualificationLocation].ToString();
        //                group.Points = new BetsResult();
        //                group.Points = SetQualiPoint(group.Qualified, group.Bet.GroupQualification);
        //                player.Qualifications.Add(group);

        //            }
        //        }
        //        increment++;
        //        Console.WriteLine($"Dodnao Qualification nr {increment} z {read.Values.Count}");
        //    }
        //    return players;
        //}

        public List<Player> TourmentBestResult(List<Player> players, ValueRange read)
        {
            foreach (var player in players)
            {
                player.TourmentClassification.Bet = new Bets();
            }
            int increment = 0;
            foreach (var item in read.Values)
            {
                if (item.Count == 0)
                {
                    continue;
                }
                if (item[0].ToString() == "")
                {
                    continue;
                }

                foreach (var player in players)
                {
                    //var tourmentResult = new TourmentClassification();


                    Enum.TryParse(typeof(TourmentClassificationEnum), item[0].ToString(), out var place);

                    switch (place)
                    {
                        case TourmentClassificationEnum.Winner:
                            //tourmentResult.Winner = item[player.ColumnTourmentClassification].ToString();
                            player.TourmentClassification.Bet.Winner = item[player.ColumnTourmentClassification].ToString();
                            break;
                        case TourmentClassificationEnum.Second:
                            player.TourmentClassification.Bet.SecondPlace = item[player.ColumnTourmentClassification].ToString();
                            break;
                        case TourmentClassificationEnum.Third:
                            player.TourmentClassification.Bet.ThirdPlace = item[player.ColumnTourmentClassification].ToString();
                            break;
                    }


                }
                increment++;
                Console.WriteLine($"Dodnao TourmentClassifivation nr {increment} z {read.Values.Count}");
            }
            return players;

        }
        public void AddPoints(SheetsEnum MatchTable, SheetsEnum QualificationTable, SheetsEnum TourmentTable, SheetsEnum FinaleFase)
        {
            string startCell = "A1";
            string endCell = "Z69";
            var sh = new SheetsHelper();
            var readed = sh.Read(MatchTable, startCell, endCell);
            var readedFinaleFaze = sh.Read(FinaleFase, startCell, "Z43");
            var readQuali = sh.Read(QualificationTable, "A1", "V43");
            var readTourment = sh.Read(TourmentTable, "A1", "I5");

            //List<Player> players = new List<Player>();
            var file = System.IO.File.ReadAllText("Json/players.json");
            List<Player> players = JsonConvert.DeserializeObject<List<Player>>(file);

            //var mach = new MatchResult();

            players = ReadPlayers(readed, GamesEnum.Mach, players);
            players = MatchBestResults(players, readed);

            players = ReadPlayers(readed, GamesEnum.Mach, players);
            players = MatchBestResults(players, readedFinaleFaze, true);

            players = ReadPlayers(readQuali, GamesEnum.GroupQualification, players);
            //players = QualificationBstResult(players, readQuali);

            players = ReadPlayers(readTourment, GamesEnum.TourmentWinner, players);
            players = TourmentBestResult(players, readTourment);


            UpdatePoints(players);
            System.Threading.Thread.Sleep(3000);
            Console.WriteLine("Zakończono updatowanie Punków");
            Console.WriteLine("Rozpoczęto zapisywanie danych");
            SaveToJson(JsonConvert.SerializeObject(players));
            System.Threading.Thread.Sleep(3000);

        }

        private TeamsValues SplitResult(string result)
        {
            result.Trim();
            TeamsValues tValues = new TeamsValues();
            if (result.Contains(':'))
            {
                var splitted = result.Split(':');
                int.TryParse(splitted[0], out int home);
                tValues.home = home;
                int.TryParse(splitted[1], out int away);
                tValues.away = away;
            }
            else
            {
                tValues.away = -1;
                tValues.home = -1;
            }
            return tValues;
        }
        private MatchResultsEnum SetWinner(int home, int away)
        {
            if (home == -1 && away == -1)
            {
                return MatchResultsEnum.notSett;
            }
            else if (home == away)
            {
                return MatchResultsEnum.draw;
            }
            else if (home > away)
            {
                return MatchResultsEnum.home;
            }
            else if (home < away)
            {
                return MatchResultsEnum.away;
            }
            else
            {
                return MatchResultsEnum.notSett;
            }
        }
        private BetsResult CountPoints(MatchResult match, bool? finaleFase)
        {

            BetsResult betsResult = new BetsResult();
            /// brak podanych wartości
            if ((int)match.winTeam == -1 && match.score.regular.home == -1)
            {
                betsResult.MachScoreBets = 0;
                betsResult.WinTeamBets = 0;
            }
            else
            {

                // trafiony wynik
                if (match.bets.RegularTimeWin.home == match.score.regular.home
                    && match.bets.RegularTimeWin.away == match.score.regular.away)
                {
                    if (finaleFase == false)
                    {
                        betsResult.MachScoreBets = (int)PiontsEnum.CorrectScore;
                    }
                    else
                    {
                        betsResult.MachScoreBets = (int)PiontsEnum.CorrectScoreFinle;
                    }
                }
                // wytypowana drużyna
                if (match.bets.winTeam == match.winTeam)
                {
                    if (finaleFase == false)
                    {
                        betsResult.WinTeamBets = (int)PiontsEnum.WinTeamFinale;
                    }
                    else
                    {
                        betsResult.WinTeamBets = (int)PiontsEnum.WinTeam;
                    }
                }
            }
            // czerwone kartki
            if (match.ReadCards > 0)
            {
                if (match.bets.RedCard == true)
                {
                    betsResult.RedCardBets = (int)PiontsEnum.ReadCard;
                }
            }
            else
            {
                if (match.bets.RedCard == false)
                {
                    betsResult.RedCardBets = (int)PiontsEnum.ReadCard;
                }
            }

            return betsResult;
        }

        private int SetValue(string value)
        {
            if (value.Trim() != "")
            {
                int.TryParse(value, out int result);
                return result;
            }
            else
            {
                return 0;
            }

        }

        public void SaveToJson(string data)
        {
            Console.WriteLine("Tworzenie pliku Json");

            string path = Directory.GetCurrentDirectory();
            string fullPath = Path.Combine(path, "CurrentBets.json");

            //string fullPathOld = Path.Combine(path,@$"CurrentBets_{DateTime.Now.ToString("dd_MM_yyyy_hh_mm")}.json");

            if (File.Exists(fullPath))
            {
                var crationTime = File.GetCreationTime(fullPath);

                string fullPathOld = Path.Combine(path, $"HistoryBets");
                string fullPathOldData = Path.Combine(fullPathOld, $"CurrentBets_{crationTime.ToString("dd_MM_yyyy_HH_mm_ss")}.json");
                int loopCount = 0;
                while (File.Exists(fullPathOldData))
                {
                    loopCount++;
                    fullPathOldData = Path.Combine(fullPathOld, $"CurrentBets_{crationTime.ToString("dd_MM_yyyy_HH_mm_ss")}_{loopCount}.json");
                }
                // if(File.Exists(fullPathOld))
                // {
                //     fullPathOld += "_1";
                // }
                Console.WriteLine($"Rozpoczęcie prznoszenia pliku {fullPath} od {fullPathOldData}");
                File.Move(fullPath, fullPathOldData);
                Console.WriteLine($"Przeniesiono  {fullPath} od {fullPathOldData}");

                Console.WriteLine($"Plik usuwany: {fullPath} ");

                File.Delete(fullPath);

                while (File.Exists(fullPath))
                {
                    Console.WriteLine($"Plik jest usuwany: {fullPath} ");
                    System.Threading.Thread.Sleep(100);
                }
                System.Threading.Thread.Sleep(1000);
                Console.WriteLine();
                Console.WriteLine($"Plik został usunięty : {File.Exists(fullPath)}");
            }


            Console.WriteLine($"Tworzenie  {fullPath} ");
            using (StreamWriter sw = File.AppendText(fullPath))
            {
                foreach (var item in data)
                {
                    sw.Write(item);
                }
                //data.ForEach(x => sw.WriteLine( x.convertToCsvRow()));
                Console.WriteLine($"Utworzono {fullPath} ");
            }
            System.Threading.Thread.Sleep(3000);
        }

        private void UpdatePoints(List<Player> players)
        {
            var sh = new SheetsHelper();

            foreach (var player in players)
            {
                ValueRange update = new ValueRange();
                List<IList<object>> values = new List<IList<object>>();
                values.Add(new List<object>() { player.TotalMachWinPoints.ToString(), player.RedCardPoints.ToString() });
                update.Values = values;
                //player.ColumnWynikLocation
                Enum.TryParse(typeof(ColumnEnum), player.ColumnWynikLocation.ToString(), out object column);

                sh.UpdateCell(SheetsEnum.Grupy, $"{column}4", update);
                Console.WriteLine($"Zakualizowano punkty dla {player.Name}");
            }
            Console.WriteLine($"Update Points - zakończono");
            // sh.UpdateCell(SheetsEnum.Grupy, string cell, ValueRange value)
        }

        private bool? RedCardSet(string value)
        {

            if (value == "Tak")
            {
                return true;
            }
            else if (value == "Nie")
            {
                return false;
            }
            else
            {
                return null;
            }
        }
    }
}
