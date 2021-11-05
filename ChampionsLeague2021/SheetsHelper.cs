using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Requests;
using ChampionsLeague2021.Enums;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using System.Threading;
using Google.Apis.Util.Store;
using System.Configuration;

namespace ChampionsLeague2021
{
    public class SheetsHelper
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };

        static readonly string ApplicationName = "Legislators";

        string SpreadsheetId = "1o_DOHZdmXc8y2IUaPwBTJMFjWfSZ0M9xhRXFJxhg8dY";

        //static readonly string SpreadsheetId2 = "1AdCREzt0sBqlICp-hooH0wNU9YT9MaFtZ8CGOviDT6o";

        //static readonly string sheet = "Arkusz1";

        static string SpreadsheetId2 = ConfigurationManager.AppSettings.Get("DataSheets");

        static SheetsService service;

        public SheetsHelper()
        {
            SetConnection();

        }

        public SheetsHelper(string spreadsheetId)
        {
            SpreadsheetId = spreadsheetId;
            SetConnection();
        }

        private void SetConnection()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

            }

            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        public ValueRange Read(SheetsEnum sheet, string startCell, string endCell)
        {
           return Read(sheet, startCell, endCell, SpreadsheetId);

        }

        public ValueRange Read(SheetsEnum sheet, string startCell, string endCell, string spreadsheetId)
        {
            var range = $"{sheet}!{startCell}:{endCell}";

            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;

            return response;
        }

        public void Write(SheetsEnum sheet, string startCell, IList<object> values)
        {
            var range = $"{sheet}!{startCell}";
            var valueRange = new ValueRange();
            var tmp = new List<IList<object>>();
            tmp.Add(values);

            valueRange.Values = tmp;

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            var appendResponse = appendRequest.Execute();
        }

        public void Write2(SheetsEnum sheet, string startCell, ValueRange valueRange)
        {
            var range = $"{sheet}!{startCell}";
            // var valueRange = new ValueRange();
            // var tmp = new List<IList<object>>();
            // tmp.Add(values);

            //valueRange.Values = values;

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            var appendResponse = appendRequest.Execute();
        }
        public void ClearSheet(SheetsEnum sheet)
        {
            var range = $"{sheet}!A:ZZ";
            var request = service.Spreadsheets.Values.Clear(new ClearValuesRequest(), SpreadsheetId, range);

            var appendResponse = request.Execute();
        }

        public void ClearSheet(SheetsEnum sheet, string startCell, string endCell)
        {
            var range = $"{sheet}!{startCell}:{endCell}";
            var request = service.Spreadsheets.Values.Clear(new ClearValuesRequest(), SpreadsheetId, range);

            var appendResponse = request.Execute();
        }

        public void UpdateCell(SheetsEnum sheet, string cell, ValueRange value)
        {
            this.UpdateCell(sheet, cell, value, SpreadsheetId);

        }
        public void UpdateCell(SheetsEnum sheet, string cell, ValueRange value, string spreadsheetId)
        {
            var range = $"{sheet}!{cell}";

            var UpdateRequest = service.Spreadsheets.Values.Update(value, spreadsheetId, range);

            UpdateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            var appendResponse = UpdateRequest.Execute();

        }

        public List<ConditionValue> ConditionValidationList(List<string> conditionList)
        {
            List<ConditionValue> conditionValueList = new List<ConditionValue>();
            foreach (var condition in conditionList)
            {
                ConditionValue conditionValue = new ConditionValue() { UserEnteredValue = condition };
                conditionValueList.Add(conditionValue);
            }

            return conditionValueList;

        }       
        public void AutoResizeCell(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn)
        {            
            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        AutoResizeDimensions = new AutoResizeDimensionsRequest()
                        {
                            Dimensions = new DimensionRange()
                            {
                                Dimension = "COLUMNS",
                                SheetId = sheetId,
                                StartIndex = (int?)StartColum,
                                EndIndex = (int?)EndColumn
                            }
                        }
                    }
                }
            };
            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId2);

            set.Execute();
        }
        public void UnmargeCellRequest(int sheetId, string rangeCells)
        {
            GridRange range = ConvertToGridRange(rangeCells);
            range.SheetId = sheetId;
            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        UnmergeCells = new UnmergeCellsRequest()
                        {
                            Range = range
                        }
                    }
                }
            };

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId2);

            set.Execute();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetId">GId of google sheets (find at end of url address) </param>
        /// <param name="rangeCells"></param>
        /// <param name="mergeType">MERGE_ALL MERGE_COLUMNS MERGE_ROWS</param>
        public void MargeCellRequest(int sheetId, string rangeCells, MergeCellTypesEnum mergeType)
        {
            GridRange range = ConvertToGridRange(rangeCells);
            range.SheetId = sheetId;

            BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        MergeCells = new MergeCellsRequest()
                        {
                            MergeType = mergeType.ToString(),
                            Range = range
                        }                        
                                                    
                    }
                }
            };

            var set = service.Spreadsheets.BatchUpdate(requestBody, SpreadsheetId2);

            set.Execute();
        }
        public void CellBackgroundCollor(int sheetId, string rangeCells, Color bacgroundColor)
        {
            GridRange range = ConvertToGridRange(rangeCells);
            range.SheetId = sheetId;

            BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        RepeatCell = new RepeatCellRequest()
                        {
                            Range = range,
                            Cell = new CellData()
                            {
                                UserEnteredFormat = new CellFormat()
                                {
                                    BackgroundColor = bacgroundColor
                                }
                            },
                            Fields = "userEnteredFormat(backgroundColor)"
                        }
                    }

                }
            };

            var set = service.Spreadsheets.BatchUpdate(requestBody, SpreadsheetId2);

            set.Execute();
        }

        public void BorderCell(int sheetId, string rangeCells, int RowEnd)
        {
            Request borderCellRequest = new Request();
            GridRange gridRange = ConvertToGridRange(rangeCells);
            gridRange.SheetId = sheetId;

            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        UpdateBorders = new UpdateBordersRequest()
                        {
                            Range = gridRange,
                            Right = new Border()
                            {
                                Style = "SOLID",
                                Color = new Color()
                                {
                                    Alpha = 0.99f,
                                    Blue = 0.0f,
                                    Green = 0.1f,
                                    Red = 0.1f
                                }
                            }
                        }
                    }
                }
            };
            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId2);

            set.Execute();
        }
        public void DataValidation(int sheetId, string rangeCells, List<ConditionValue> conditionValueList)
        {
            // List<Request> body = new List<Request>();

            GridRange range = ConvertToGridRange(rangeCells);

            range.SheetId = sheetId;

            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>() 
                { 
                    new Request()
                    {
                        SetDataValidation = new SetDataValidationRequest()
                        {
                            Range = range,
                            Rule = new DataValidationRule()
                            {
                                Condition = new BooleanCondition()
                                {
                                    Type = "ONE_OF_LIST",
                                    Values = conditionValueList
                                }
                            }
                        }
                    }
                }                
            };

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId);

            set.Execute();

        }
        public void ProtectedRange(int sheetId, string rangeCells, int ProtectedRangeId)
        {
            GridRange range = ConvertToGridRange(rangeCells);

            range.SheetId = sheetId;

            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                       
                        UpdateProtectedRange = new UpdateProtectedRangeRequest()
                        {
                            ProtectedRange = new ProtectedRange()
                            {
                                Editors = new Editors()
                                {
                                    Users = new List<string>()
                                    {
                                        ConfigurationManager.AppSettings.Get("GoogleSheetsEmailAddress"),
                                        ConfigurationManager.AppSettings.Get("SheetsOwnerEmailAddress")
                                    }
                                },
                               ProtectedRangeId = ProtectedRangeId,
                               Range = range,
                               RequestingUserCanEdit = true,
                               Description = "test"                             
                            },
                            Fields = "*"
                        }
                    }
                }
            };

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId);
            var s = service.Spreadsheets.Sheets;
            //service.Spreadsheets.Get();

            set.Execute();
        }
        public void AddNewProtectedRange(int sheetId, string rangeCells, int ProtectedRangeId)
        {
            GridRange range = ConvertToGridRange(rangeCells);

            range.SheetId = sheetId;

            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {

                        AddProtectedRange = new AddProtectedRangeRequest()
                        {
                            ProtectedRange = new ProtectedRange()
                            {
                                Editors = new Editors()
                                {
                                    Users = new List<string>()
                                    {
                                        ConfigurationManager.AppSettings.Get("GoogleSheetsEmailAddress"),
                                        ConfigurationManager.AppSettings.Get("SheetsOwnerEmailAddress")
                                    }
                                },
                               ProtectedRangeId = ProtectedRangeId,
                               Range = range,
                               RequestingUserCanEdit = true,
                               Description = "test1",
                            }
                        }
                    }
                }
            };

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId);

            set.Execute();
        }
        public void SetDataValidationPost()
        {
            var url = $"https://sheets.googleapis.com/v4/spreadsheets/{SpreadsheetId}:batchUpdate";
            RestHelper restHelper = new RestHelper();



            var client = restHelper.SetRestClient(url);
            var request = restHelper.CreateGetRequest(RestSharp.Method.POST);
            // request.AddJsonBody(JsonConvert.SerializeObject(new FileStream("Json//SetDataValidation.json",FileMode.Open, FileAccess.Read)));
            request.AddJsonBody("Json//SetDataValidation.json");
            var response = restHelper.GetResponse(client, request);
            //var content = restHelper.Content2<Match>(response);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellRange"> example: "A1:B2"</param>
        /// <returns></returns>
        public GridRange ConvertToGridRange(string cellRange)
        {
            GridRange gridRange = new GridRange();

            
            List<string> listColumn = new List<string>();
            List<string> listRow = new List<string>();


            for (int j = 0; j < cellRange.Split(':').Length; j++)
            {
                var address = cellRange.Split(':')[j].Trim(':');
                string column = null;
                string row = null;

                for (int i = 0; i < address.Length; i++)
                {
                    if (int.TryParse(address[i].ToString(), out _))
                    {
                        row += address[i].ToString();
                    }

                    else
                    {
                        column += address[i].ToString();
                    }
                }
                listColumn.Add(column);
                listRow.Add(row);
            }

            gridRange.StartColumnIndex = ConwertColumnStringAddressToIntIndex(listColumn[0]);
            gridRange.StartRowIndex = int.Parse(listRow[0])-1;
            gridRange.EndColumnIndex = ConwertColumnStringAddressToIntIndex(listColumn[1])+1;
            gridRange.EndRowIndex = int.Parse(listRow[1]);

            return gridRange;
        }

        private int ConwertColumnStringAddressToIntIndex(string address)
        {
            if (string.IsNullOrEmpty(address))
                throw new System.ArgumentNullException("address");

           address = address.ToUpper();
            int sum = 0;

            for (int i = 0; i < address.Length; i++)
            {
                sum *= 26;
                sum += (address[i] - 'A' );
            }

            return sum;
        }

    }
}