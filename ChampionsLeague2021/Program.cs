using System;
using Google.Apis.Sheets.v4.Data;
using ChampionsLeague2021.Enums;
using System.Linq;
using System.Configuration;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Drawing;

namespace ChampionsLeague2021
{
    class Program
    {
     
        //"1AdCREzt0sBqlICp-hooH0wNU9YT9MaFtZ8CGOviDT6o";
        
        static void Main(string[] args)
        {
            string SpreadsheetId2 = ConfigurationManager.AppSettings.Get("DataSheets");

            Console.WriteLine("Hello World!");
            SheetsHelper sheetsHelper = new SheetsHelper(SpreadsheetId2);
            AddPlayers addPlyers = new AddPlayers();
            addPlyers.ChackAndUpdatePlayers("Json/players.json", SheetsEnum.Players, SpreadsheetId2);

            //sheetsHelper.UpdateCell(SheetsEnum.Arkusz1, "A1", addPlyers.AddPlayer("Json/players.json"));
            //sheetsHelper.UnmargeCellRequest(0, ColumnEnum.A, ColumnEnum.Z, 0, 100);
            //sheetsHelper.BorderCell(0, ColumnEnum.D, ColumnEnum.E, 1, 100);
            //sheetsHelper.MargeCellRequest(0, ColumnEnum.A, ColumnEnum.C, 0, 2);
            //sheetsHelper.MargeCellRequest(0, ColumnEnum.C, ColumnEnum.E, 0, 2);
            //sheetsHelper.CellBackgroundCollor(0, ColumnEnum.C, ColumnEnum.D, 1, 3, sheetsHelper.ColorConvert(System.Drawing.ColorTranslator.FromHtml("#aaea9999")));

            sheetsHelper.ConvertToGridRange("Z11:AA20");
            //sheetsHelper.AutoResizeCell(1414534158, ColumnEnum.A, ColumnEnum.Z);
            //sheetsHelper.AutoResizeCell(0, ColumnEnum.A, ColumnEnum.Z);

            GetMatchDay Matches = new GetMatchDay();
            List<Models.MatchInfo> result = Matches.GetDataFromEndpoint(GetMatchsIdAndSetupTechnicalData(SheetsEnum.Arkusz3,
                                                                                                         "A2",
                                                                                                         "A40",
                                                                                                         SpreadsheetId2));


               
            ValueRange valueRange = Matches.PrepareSheetsEntries(result);

            //SheetsHelper sheetsHelper = new SheetsHelper(SpreadsheetId2);
           // sheetsHelper.Write2(SheetsEnum.Dane, "A1", valueRange);
            sheetsHelper.UpdateCell(SheetsEnum.Dane, "A1", valueRange);
            Matches.UpdateStatus();


        }

       static private string GetMatchsIdAndSetupTechnicalData(SheetsEnum arkusz, string startCell, string endCell, string spreadsheetId)
        {
            SheetsHelper Sheets = new SheetsHelper();
            ValueRange tmp = Sheets.Read(arkusz, startCell, endCell, spreadsheetId);

            string matchIds = "";

            foreach (var row in tmp.Values)
            {
              if (row[0] == tmp.Values.Last()[0] )
                {
                    matchIds += row[0];
                }
              else
                    matchIds += row[0] + ",";
            }

            return matchIds;            
        }
    }
}
