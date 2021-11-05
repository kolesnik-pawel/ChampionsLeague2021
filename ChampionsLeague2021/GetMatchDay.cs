using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4.Data;
using ChampionsLeague2021.Enums;
using ChampionsLeague2021.Models;
using System.Configuration;

namespace ChampionsLeague2021
{
    class GetMatchDay
    {
        static string path = ConfigurationManager.AppSettings.Get("GetMatchesEndpoint");// $"https://match.uefa.com/v2/matches?matchId=";
        static string pathEvents = $"https://match.uefa.com/v2/matches/";


       // static string eventsPath = "/events?filter=LINEUP&offset=0&limit=100";
        static string pathEvent = ConfigurationManager.AppSettings.Get("GetEventsEndpoint");

        public List<MatchInfo> GetDataFromEndpoint(string matchIds)
        {
            RestHelper restHelper = new RestHelper();

            var client = restHelper.SetRestClient(string.Concat(path,matchIds));
            var request = restHelper.CreateGetRequest(RestSharp.Method.GET);

            var response = restHelper.GetResponse(client, request);
            var content = restHelper.Content2<MatchInfo>(response);

            return content.OrderBy(x => x.kickOffTime.dateTime).ToList();

        }

        public MatchInfo GetSingleMatch(int MachId)
        {
            var url = path + MachId.ToString();

            RestHelper restHelper = new RestHelper();

            var client = restHelper.SetRestClient(url);
            var request = restHelper.CreateGetRequest(RestSharp.Method.GET);

            var response = restHelper.GetResponse(client, request);
            var content = restHelper.Content<MatchInfo>(response);

            return content;
        }

        public List<MatchEvents> GetSingleMatchEvents(int MachId)
        {
            //var url = pathEvents + MachId.ToString() + eventsPath;
            
            var url = string.Format(pathEvent, MachId);


            RestHelper restHelper = new RestHelper();

            var client = restHelper.SetRestClient(url);
            var request = restHelper.CreateGetRequest(RestSharp.Method.GET);

            var response = restHelper.GetResponse(client, request);
            if (response.Content.Count() == 0)
            {
                return null;
            }
            var content = restHelper.Content2<MatchEvents>(response);

            return content;
        }

        public int CountEvents(List<MatchEvents> content, EventsEnum events)
        {
            int count = 0;
            foreach (var item in content)
            {
                if (item.type.ToString() == events.ToString())
                {
                    count++;
                }
            }
            return count;
        }

        public ValueRange PrepareSheetsEntries(List<MatchInfo> content)
        {
            ValueRange valueRange = new ValueRange();
            var tmp = new List<IList<object>>();
            string matchday = "";
            int rowNumber = 1;

            tmp.Add(new List<object>{"Grupa","Gospodarz","","Godzina","Goście","","id meczu","Data meczu",
            "czas aktualny","róznica w czasie","przeliczona na dni", "Status meczu", "Typ rozgrywki", "Wynik", "Zwycięzca", 
                "Żółte kartki", "Czerwone kartki", "Karne", "Data Aktualizacji"});

            foreach (var item in content)
            {
                rowNumber++;
                IList<object> contentList = new List<object>();
                
                //Sets at column 3 (D) Match date
                if (item.kickOffTime.date != matchday)
                {
                    matchday = item.kickOffTime.date;
                    //contentList.Add(groupTmp);
                    tmp.Add(new List<object>());
                    tmp.Add(new List<object>() { "", "", "", matchday });
                    rowNumber = rowNumber + 2;
                }

                
                string imageHome = $@"=IMAGE(""{item.homeTeam.logoUrl}"")";
                string imageAway = $@"=IMAGE(""{item.awayTeam.logoUrl}"")";
                var matchHour = item.kickOffTime.dateTime.ToLocalTime().ToString("t");
                var ifGroupExist = item.group != null ? item.group.metaData.groupName : "";
                var groupLink = ifGroupExist != "" ? $@"=HYPERLINK(""https://www.uefa.com/uefaeuro-2020/match/{item.id}/standings"";""{ifGroupExist}"")" : "";

                // fill row cells...
                contentList.Add(groupLink);
                //item.kickOffTime.dateTime.AddHours(item.kickOffTime.utcOffsetInHours).ToLocalTime();
                contentList.Add(item.homeTeam.internationalName);
                contentList.Add(imageHome);
                contentList.Add($@"=HYPERLINK(""https://www.uefa.com/uefaeuro-2020/match/{item.id}"";""{matchHour}"")");
                contentList.Add(item.awayTeam.internationalName);
                contentList.Add(imageAway);
                contentList.Add(item.id);
                // contentList.Add(item.kickOffTime.dateTime.AddHours(item.kickOffTime.utcOffsetInHours).ToString("yyyy-MM-dd HH:mm:ss"));
                contentList.Add(item.kickOffTime.dateTime.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"));
                contentList.Add("=Now()");
                contentList.Add($@"=if(I{rowNumber} > H{rowNumber}; ""00:00""; H{rowNumber}-I{rowNumber})");
                contentList.Add($@"=IFS(INT(J{rowNumber})<0; ""00:00""; INT(J{rowNumber})>=1;CONCATENATE(INT(J{rowNumber});"" dni ""; HOUR(J{rowNumber}); "" g "");  INT(J{rowNumber})=0; CONCATENATE(HOUR(J{rowNumber}); "" g: "";MINUTE(J{rowNumber}); "" min""))");
                contentList.Add(item.status);
                contentList.Add(item.type);
                if (item.score != null)
                {
                    contentList.Add(item.score.regular.home + ":" + item.score.regular.away);
                    //contentList.Add(item.score.penalty);
                    if (item.winner.match.reason == "DRAW")
                    {
                        contentList.Add(item.winner.match.reason);
                    }
                    else
                    {
                        contentList.Add(item.winner.match.team.internationalName);
                    }
                }

                // add rows to list of rows 
                tmp.Add(contentList);
            }
            valueRange.Values = tmp;

            return valueRange;
        }

        public void UpdateStatus()
        {
            string SpreadsheetId2 = ConfigurationManager.AppSettings.Get("DataSheets");
            var sheetName = SheetsEnum.Dane;
            var sh = new SheetsHelper();

            var readed = sh.Read(sheetName, "A1", "Q96", SpreadsheetId2);


            ValueRange update = new ValueRange();
            List<IList<object>> values = new List<IList<object>>();
            //List<object> rowValue = new List<object>();

            int rowCount = 0;

            foreach (var match in readed.Values)
            {
                rowCount++;

                if (match.Count < 5)
                {
                    values.Add(new List<object>());
                    continue;
                }
                if (match[1].ToString() != "" && match[1].ToString() != "Gospodarz")
                {
                    if (match[(int)DataColumnName.roznica_w_czasie].ToString() != "00:00")
                    {
                        double divInTime = double.Parse(match[(int)DataColumnName.roznica_w_czasie].ToString());
                        if (divInTime > 1)
                        {
                            values.Add(new List<object>());
                            continue;
                        }
                    }

                    if (match[11].ToString() == StatusEnum.UPCOMING.ToString() || match[11].ToString() == StatusEnum.LIVE.ToString()
                    || match[11].ToString() == StatusEnum.FINISHED.ToString())
                    {

                        MatchInfo matchRead = new MatchInfo();
                        List<MatchEvents> marchReadEvents = new List<MatchEvents>();
                        
                        matchRead = GetSingleMatch(int.Parse(match[6].ToString()));
                        marchReadEvents = GetSingleMatchEvents(int.Parse(match[6].ToString()));

                        int yellowCards = CountEvents(marchReadEvents, EventsEnum.YELLOW_CARD);
                        int redCards = CountEvents(marchReadEvents, EventsEnum.RED_CARD);
                        int redYellowCards = CountEvents(marchReadEvents, EventsEnum.RED_YELLOW_CARD);

                        List<object> rowValue = new List<object>();

                        rowValue.Add(matchRead.status);
                        rowValue.Add(matchRead.type);
                        rowValue.Add($"{matchRead.score.total.home}:{matchRead.score.total.away}");

                        if (matchRead.status == StatusEnum.FINISHED.ToString())
                        {                                              
                            if (matchRead.winner.match.reason == "WIN_REGULAR" ||
                                matchRead.winner.match.reason == "WIN_ON_EXTRA_TIME" ||
                                matchRead.winner.match.reason == "WIN_ON_PENALTIES")
                            {
                                rowValue.Add(matchRead.winner.match.team.internationalName);
                            }
                            else if (matchRead.winner.match.reason == "DRAW")
                            {
                                rowValue.Add(matchRead.winner.match.reason);
                            }
                        }
                        else if (matchRead.status == StatusEnum.LIVE.ToString())
                        {
                            rowValue.Add("Live");
                        }

                        rowValue.Add(yellowCards);
                        rowValue.Add(redYellowCards + redCards);
                        //karne
                        if (matchRead.score.penalty != null)
                        {
                            rowValue.Add($"{matchRead.score.penalty.home}:{matchRead.score.penalty.away}");
                        }
                        else
                        {
                            rowValue.Add("brak");
                        }

                        rowValue.Add(DateTime.Now.ToString("dd-MM-yy HH:mm:ss"));

                        values.Add(rowValue);

                        Console.WriteLine($"update match {matchRead.homeTeam.internationalName} vs {matchRead.awayTeam.internationalName}");

                    }
                }
            }
            
            string colStatusName = ((ColumnEnum)DataColumnName.Status_meczu).ToString();
            update.Values = values;
            sh.UpdateCell(sheetName, $"{colStatusName}{2}", update, SpreadsheetId2);

        }

        public void AddProtectedRange()
        {
            var sheetName = SheetsEnum.Dane;
            var sh = new SheetsHelper();
            var readed = sh.Read(sheetName, "A1", "Q63");

            int rowCount = 0;
            foreach (var match in readed.Values)
            {
                rowCount++;

                if (match.Count < 5)
                {
                    continue;
                }
                // if(match[1].ToString() != "" && match[1].ToString() != "Gospodarz")
                // {
                //       if (match[11].ToString() == StatusEnum.LIVE.ToString() || match[11].ToString() == StatusEnum.FINISHED.ToString())
                //       {

                //       }
                // }
                if (match[11].ToString() == StatusEnum.UPCOMING.ToString())
                {
                    break;
                }
            }
            sh.ProtectedRange(43786814, $"L8:Z{rowCount + 6}", 790);
            if (rowCount > 65)
            {
                sh.ProtectedRange(48770365, $"L8:Z{rowCount + 6 - 66}", 791);
            }
        }

    }
}
