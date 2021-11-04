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
            var range = $"{sheet}!{startCell}:{endCell}";

            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;

            return response;
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
        public void DataValidation(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd)
        {
            List<Request> body = new List<Request>();
            Request dataValidation = new Request();
            GridRange range = new GridRange();

            range.SheetId = sheetId;
            range.StartColumnIndex = (int?)StartColum;
            range.EndColumnIndex = (int?)EndColumn;
            range.StartRowIndex = RowStart;
            range.EndRowIndex = RowEnd;

            DataValidationRule rule = new DataValidationRule();

            BooleanCondition condition = new BooleanCondition();
            condition.Type = "ONE_OF_LIST";
            ConditionValue conditionValue1 = new ConditionValue() { UserEnteredValue = "Tak" };
            ConditionValue conditionValue2 = new ConditionValue() { UserEnteredValue = "Nie" };

            List<ConditionValue> conditionValueList = new List<ConditionValue>();
            conditionValueList.Add(conditionValue1);
            conditionValueList.Add(conditionValue2);
            condition.Values = conditionValueList;

            rule.Condition = condition;

            SetDataValidationRequest dataValidationRequest = new SetDataValidationRequest();

            dataValidationRequest.Range = range;
            dataValidationRequest.Rule = rule;

            dataValidation.SetDataValidation = dataValidationRequest;

            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest();

            request.Requests = new List<Request>() { dataValidation };

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId2);

            set.Execute();
        }

        public void UnmargeCellRequest(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd)
        {
            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        UnmergeCells = new UnmergeCellsRequest()
                        {
                            Range = new GridRange()
                            {
                                SheetId = sheetId,
                                StartColumnIndex = (int?)StartColum,
                                EndColumnIndex = (int?)EndColumn,
                                StartRowIndex = RowStart,
                                EndRowIndex = RowEnd
                            }
                        }
                    }
                }
            };

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId2);

            set.Execute();
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
        public void MargeCellRequest(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd)
        {
            BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        MergeCells = new MergeCellsRequest()
                        {
                            MergeType = "MERGE_ROWS",
                            Range = new GridRange()
                            {
                                SheetId = sheetId,
                                StartColumnIndex = (int?)StartColum,
                                EndColumnIndex = (int?)EndColumn,
                                StartRowIndex = RowStart,
                                EndRowIndex = RowEnd
                            }
                        }                        
                                                    
                    }
                }
            };

            var set = service.Spreadsheets.BatchUpdate(requestBody, SpreadsheetId2);

            set.Execute();
        }
        public void CellBackgroundCollor(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd, Color bacgroundColor)
        {
            BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        RepeatCell = new RepeatCellRequest()
                        {
                            Range = new GridRange()
                            {
                                SheetId = sheetId,
                                StartColumnIndex = (int?)StartColum,
                                EndColumnIndex = (int?)EndColumn,
                                StartRowIndex = RowStart,
                                EndRowIndex = RowEnd
                            },
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

        public void BorderCell(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd)
        {
            Request borderCellRequest = new Request();
            GridRange gridRange = new GridRange();
            
            gridRange.SheetId = sheetId;
            gridRange.StartColumnIndex = (int?)StartColum;
            gridRange.EndColumnIndex = (int?)EndColumn;
            gridRange.StartRowIndex = RowStart;
            gridRange.EndRowIndex = RowEnd;

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
        public void DataValidation(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd, List<ConditionValue> conditionValueList)
        {
           // List<Request> body = new List<Request>();
            Request dataValidation = new Request();
            GridRange range = new GridRange();

            range.SheetId = sheetId;
            range.StartColumnIndex = (int?)StartColum;
            range.EndColumnIndex = (int?)EndColumn;
            range.StartRowIndex = RowStart;
            range.EndRowIndex = RowEnd;

            DataValidationRule rule = new DataValidationRule();

            BooleanCondition condition = new BooleanCondition();
            condition.Type = "ONE_OF_LIST";

            condition.Values = conditionValueList;

            rule.Condition = condition;

            SetDataValidationRequest dataValidationRequest = new SetDataValidationRequest();

            dataValidationRequest.Range = range;
            dataValidationRequest.Rule = rule;

            dataValidation.SetDataValidation = dataValidationRequest;

            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest();

            request.Requests = new List<Request>() { dataValidation };

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId);

            set.Execute();


        }
        public void DataValidation2(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd, List<ConditionValue> conditionValueList)
        {
            // List<Request> body = new List<Request>();
            Request dataValidation = new Request();
            GridRange range = new GridRange();

            range.SheetId = sheetId;
            range.StartColumnIndex = (int?)StartColum;
            range.EndColumnIndex = (int?)EndColumn;
            range.StartRowIndex = RowStart;
            range.EndRowIndex = RowEnd;

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
        public void ProtectedRange(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd, int ProtectedRangeId)
        {
            //var tmp = service.Spreadsheets.Values.Get(SpreadsheetId, "A1:Z20").Service.;
            List<Request> body = new List<Request>();
            Request dataProtected = new Request();
            GridRange range = new GridRange();
            ProtectedRange protectedRange = new ProtectedRange();
            protectedRange.ProtectedRangeId = ProtectedRangeId;

            range.SheetId = sheetId;
            range.StartColumnIndex = (int?)StartColum;
            range.EndColumnIndex = (int?)EndColumn;
            range.StartRowIndex = RowStart;
            range.EndRowIndex = RowEnd;

            protectedRange.Range = range;

            // protectedRange.NamedRangeId = "Nie ma obstawiania";

            Editors usersCanEdit = new Editors();
            usersCanEdit.Users = new List<string>();
            usersCanEdit.Users.Add("sancho0510@gmail.com");
            usersCanEdit.Users.Add("sheets@fluid-isotope-311615.iam.gserviceaccount.com");
            protectedRange.Editors = usersCanEdit;

            // dataProtected.AddProtectedRange =  new AddProtectedRangeRequest();
            // dataProtected.AddProtectedRange.ProtectedRange = protectedRange;

            dataProtected.UpdateProtectedRange = new UpdateProtectedRangeRequest();
            dataProtected.UpdateProtectedRange.ProtectedRange = protectedRange;
            dataProtected.UpdateProtectedRange.Fields = "*";

            BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest();

            request.Requests = new List<Request>() { dataProtected };
            //request.Requests[0].

            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId);
            //service.Spreadsheets.Get();

            set.Execute();

        }

        public void ProtectedRange2(int sheetId, ColumnEnum StartColum, ColumnEnum EndColumn, int RowStart, int RowEnd, int ProtectedRangeId)
        {
            //var tmp = service.Spreadsheets.Values.Get(SpreadsheetId, "A1:Z20").Service.;
            List<Request> body = new List<Request>();
            Request dataProtected = new Request();
            GridRange range = new GridRange();
            ProtectedRange protectedRange = new ProtectedRange();
            protectedRange.ProtectedRangeId = ProtectedRangeId;

            range.SheetId = sheetId;
            range.StartColumnIndex = (int?)StartColum;
            range.EndColumnIndex = (int?)EndColumn;
            range.StartRowIndex = RowStart;
            range.EndRowIndex = RowEnd;

            protectedRange.Range = range;

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
                                        "sancho0510@gmail.com",
                                        "sheets@fluid-isotope-311615.iam.gserviceaccount.com"
                                    }
                                },
                               ProtectedRangeId = ProtectedRangeId
                            },
                            Fields = "*"
                        }
                    }
                }
            };

            // protectedRange.NamedRangeId = "Nie ma obstawiania";

            Editors usersCanEdit = new Editors();
            usersCanEdit.Users = new List<string>();
            usersCanEdit.Users.Add("sancho0510@gmail.com");
            usersCanEdit.Users.Add("sheets@fluid-isotope-311615.iam.gserviceaccount.com");
            protectedRange.Editors = usersCanEdit;

            // dataProtected.AddProtectedRange =  new AddProtectedRangeRequest();
            // dataProtected.AddProtectedRange.ProtectedRange = protectedRange;

            dataProtected.UpdateProtectedRange = new UpdateProtectedRangeRequest();
            dataProtected.UpdateProtectedRange.ProtectedRange = protectedRange;
            dataProtected.UpdateProtectedRange.Fields = "*";

           // BatchUpdateSpreadsheetRequest request = new BatchUpdateSpreadsheetRequest();

            request.Requests = new List<Request>() { dataProtected };
            //request.Requests[0].


            var set = service.Spreadsheets.BatchUpdate(request, SpreadsheetId);
            //service.Spreadsheets.Get();

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
            gridRange.StartRowIndex = int.Parse(listRow[0]);
            gridRange.EndColumnIndex = ConwertColumnStringAddressToIntIndex(listColumn[1]);
            gridRange.EndRowIndex = int.Parse(listRow[1]);

            return gridRange;
        }

        private int ConwertColumnStringAddressToIntIndex(string address)
        {
            if (string.IsNullOrEmpty(address))
                throw new System.ArgumentNullException("address");

            address.ToUpper();
            int sum = 0;

            for (int i = 0; i < address.Length; i++)
            {
                sum *= 26;
                sum += (address[i] - 'A' + 1);
            }

            return sum;
        }

    }
}