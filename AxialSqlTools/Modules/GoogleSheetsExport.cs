using Microsoft.VisualStudio.Shell;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace AxialSqlTools
{
    public class GoogleSheetsExportResult
    {
        public string SpreadsheetId { get; set; }

        public string SpreadsheetUrl { get; set; }

        public string SpreadsheetTitle { get; set; }
    }

    public class GoogleSheetsAuthorizationResult
    {
        public string AccessToken { get; set; }

        public string RefreshToken { get; set; }
    }

    public static class GoogleSheetsExport
    {
        private const string SheetsScope = "https://www.googleapis.com/auth/spreadsheets";
        private const string AuthorizationEndpoint = "https://accounts.google.com/o/oauth2/v2/auth";
        private const string TokenEndpoint = "https://oauth2.googleapis.com/token";
        private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

        public static string ExpandDateWildcards(string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern))
            {
                throw new ArgumentException("Pattern must not be empty", nameof(pattern));
            }

            string result = System.Text.RegularExpressions.Regex.Replace(pattern, @"\{([^}]+)\}", match =>
            {
                string formatInside = match.Groups[1].Value;
                try
                {
                    return DateTime.Now.ToString(formatInside);
                }
                catch (FormatException)
                {
                    return match.Value;
                }
            });

            return result;
        }

        public static string BuildAuthorizationUrl(SettingsManager.GoogleSheetsSettings settings)
        {
            var query = HttpUtility.ParseQueryString(string.Empty);
            query["client_id"] = settings.clientId;
            query["redirect_uri"] = RedirectUri;
            query["response_type"] = "code";
            query["scope"] = SheetsScope;
            query["access_type"] = "offline";
            query["prompt"] = "consent";

            return new UriBuilder(AuthorizationEndpoint)
            {
                Query = query.ToString()
            }.ToString();
        }

        public static async Task<GoogleSheetsAuthorizationResult> ExchangeAuthorizationCodeAsync(SettingsManager.GoogleSheetsSettings settings, string authorizationCode, CancellationToken cancellationToken)
        {
            using (var httpClient = new HttpClient())
            {
                var content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    { "grant_type", "authorization_code" },
                    { "code", authorizationCode },
                    { "redirect_uri", RedirectUri },
                    { "client_id", settings.clientId },
                    { "client_secret", settings.clientSecret }
                });

                var response = await httpClient.PostAsync(TokenEndpoint, content, cancellationToken);
                response.EnsureSuccessStatusCode();

                string json = await response.Content.ReadAsStringAsync();
                JObject token = JObject.Parse(json);

                return new GoogleSheetsAuthorizationResult
                {
                    AccessToken = token.Value<string>("access_token"),
                    RefreshToken = token.Value<string>("refresh_token")
                };
            }
        }

        private static async Task<string> GetAccessTokenAsync(SettingsManager.GoogleSheetsSettings settings, CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(settings.refreshToken))
            {
                throw new InvalidOperationException("Google Sheets is not authorized. Please authorize first.");
            }

            using (var httpClient = new HttpClient())
            {
                var content = new FormUrlEncodedContent(new Dictionary<string, string>
                {
                    { "grant_type", "refresh_token" },
                    { "refresh_token", settings.refreshToken },
                    { "client_id", settings.clientId },
                    { "client_secret", settings.clientSecret }
                });

                var response = await httpClient.PostAsync(TokenEndpoint, content, cancellationToken);

                if (!response.IsSuccessStatusCode)
                {
                    string errorBody = await response.Content.ReadAsStringAsync();
                    throw new InvalidOperationException($"Failed to obtain access token: {response.ReasonPhrase}. {errorBody}");
                }

                string json = await response.Content.ReadAsStringAsync();
                JObject token = JObject.Parse(json);

                string newRefreshToken = token.Value<string>("refresh_token");
                if (!string.IsNullOrWhiteSpace(newRefreshToken) && !string.Equals(newRefreshToken, settings.refreshToken, StringComparison.Ordinal))
                {
                    settings.refreshToken = newRefreshToken;
                    SettingsManager.SaveGoogleSheetsSettings(settings);
                }

                return token.Value<string>("access_token");
            }
        }

        public static async Task<GoogleSheetsExportResult> ExportToNewSpreadsheetAsync(List<DataTable> dataTables, SettingsManager.GoogleSheetsSettings settings, bool isShiftPressed, CancellationToken cancellationToken)
        {
            if (dataTables == null || dataTables.Count == 0)
            {
                throw new ArgumentException("No data tables provided for export", nameof(dataTables));
            }

            bool includeSource = settings.includeSourceQuery ^ isShiftPressed;
            string sourceQuery = includeSource ? GetSourceQueryText() : string.Empty;
            string spreadsheetTitle = ExpandDateWildcards(settings.GetSpreadsheetTitle());

            string accessToken = await GetAccessTokenAsync(settings, cancellationToken);

            var sheetNames = BuildSheetNames(dataTables.Count, includeSource);
            var createPayload = new
            {
                properties = new { title = spreadsheetTitle },
                sheets = sheetNames.Select(name => new { properties = new { title = name } }).ToList()
            };

            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var createContent = new StringContent(JsonConvert.SerializeObject(createPayload), Encoding.UTF8, "application/json");
                var createResponse = await httpClient.PostAsync(
                    "https://sheets.googleapis.com/v4/spreadsheets",
                    createContent,
                    cancellationToken);

                if (!createResponse.IsSuccessStatusCode)
                {
                    var errorBody = await createResponse.Content.ReadAsStringAsync();
                    throw new InvalidOperationException(
                        $"Sheets create failed: {(int)createResponse.StatusCode} {createResponse.ReasonPhrase}\n{errorBody}");
                }

                string createJson = await createResponse.Content.ReadAsStringAsync();
                JObject spreadsheet = JObject.Parse(createJson);
                string spreadsheetId = spreadsheet.Value<string>("spreadsheetId");
                string spreadsheetUrl = spreadsheet.Value<string>("spreadsheetUrl");

                Dictionary<string, int> sheetNameToId = spreadsheet["sheets"]?
                    .OfType<JObject>()
                    .Select(s => s["properties"] as JObject)
                    .Where(p => p != null)
                    .ToDictionary(
                        p => p.Value<string>("title"),
                        p => p.Value<int>("sheetId"),
                        StringComparer.OrdinalIgnoreCase)
                    ?? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                List<object> batchRequests = new List<object>();

                for (int i = 0; i < dataTables.Count; i++)
                {
                    IList<IList<object>> values = BuildValues(dataTables[i], settings.exportBoolsAsNumbers);
                    await AppendValuesAsync(httpClient, spreadsheetId, sheetNames[i], values, cancellationToken);

                    if (sheetNameToId.TryGetValue(sheetNames[i], out int sheetId))
                    {
                        int columnCount = dataTables[i].Columns.Count;
                        batchRequests.Add(CreateFilterRequest(sheetId, columnCount));
                        batchRequests.Add(CreateAutoResizeRequest(sheetId, columnCount));
                    }
                }

                if (includeSource)
                {
                    IList<IList<object>> sourceValues = BuildSourceQueryValues(sourceQuery);
                    string sourceSheetName = sheetNames.Last();
                    await AppendValuesAsync(httpClient, spreadsheetId, sourceSheetName, sourceValues, cancellationToken);

                    if (sheetNameToId.TryGetValue(sourceSheetName, out int sourceSheetId))
                    {
                        const int sourceColumnCount = 1;
                        batchRequests.Add(CreateFilterRequest(sourceSheetId, sourceColumnCount));
                        batchRequests.Add(CreateAutoResizeRequest(sourceSheetId, sourceColumnCount));
                    }
                }

                if (batchRequests.Count > 0)
                {
                    await BatchUpdateAsync(httpClient, spreadsheetId, batchRequests, cancellationToken);
                }

                return new GoogleSheetsExportResult
                {
                    SpreadsheetId = spreadsheetId,
                    SpreadsheetUrl = spreadsheetUrl,
                    SpreadsheetTitle = spreadsheetTitle
                };
            }
        }

        private static async Task AppendValuesAsync(HttpClient httpClient, string spreadsheetId, string sheetName, IList<IList<object>> values, CancellationToken cancellationToken)
        {
            var payload = new
            {
                range = $"{sheetName}!A1",
                majorDimension = "ROWS",
                values = values
            };

            string range = $"{sheetName}!A1";

            string requestUri =
                $"https://sheets.googleapis.com/v4/spreadsheets/" +
                $"{Uri.EscapeDataString(spreadsheetId)}/values/" +
                $"{Uri.EscapeDataString(range)}:append?valueInputOption=USER_ENTERED";

            var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
            var response = await httpClient.PostAsync(requestUri, content, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                var responseBody = await response.Content.ReadAsStringAsync();
                throw new Exception($"Append failure: {(int)response.StatusCode} {response.ReasonPhrase}\n{responseBody}");
            }
        }

        private static IList<IList<object>> BuildValues(DataTable dataTable, bool exportBoolsAsNumbers)
        {
            var values = new List<IList<object>>();

            var header = dataTable.Columns
                .Cast<DataColumn>()
                .Select(col => col.ExtendedProperties.ContainsKey("columnName") ? col.ExtendedProperties["columnName"]?.ToString() : col.ColumnName)
                .Cast<object>()
                .ToList();

            values.Add(header);

            foreach (DataRow row in dataTable.Rows)
            {
                List<object> rowValues = new List<object>();
                foreach (DataColumn column in dataTable.Columns)
                {
                    object value = row[column];
                    if (value is DBNull)
                    {
                        rowValues.Add(string.Empty);
                    }
                    else if (column.DataType == typeof(bool) && exportBoolsAsNumbers)
                    {
                        rowValues.Add((bool)value ? 1 : 0);
                    }
                    else
                    {
                        rowValues.Add(value);
                    }
                }

                values.Add(rowValues);
            }

            return values;
        }

        private static IList<IList<object>> BuildSourceQueryValues(string sourceQuery)
        {
            var values = new List<IList<object>>
            {
                new List<object> { "Query Text" }
            };

            string[] lines = sourceQuery?.Replace("\r\n", "\n").Split(new[] { '\n' }, StringSplitOptions.None) ?? Array.Empty<string>();
            foreach (string line in lines)
            {
                values.Add(new List<object> { line });
            }

            return values;
        }

        private static List<string> BuildSheetNames(int dataTableCount, bool includeSource)
        {
            List<string> sheetNames = new List<string>();
            for (int i = 1; i <= dataTableCount; i++)
            {
                sheetNames.Add($"QueryResult_{i}");
            }

            if (includeSource)
            {
                sheetNames.Add("SourceQuery");
            }

            return sheetNames;
        }

        private static string GetSourceQueryText()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return ""; // ScriptFactoryAccess.GetActiveQueryWindowText();
        }

        private static object CreateFilterRequest(int sheetId, int columnCount)
        {
            return new
            {
                setBasicFilter = new
                {
                    filter = new
                    {
                        range = new
                        {
                            sheetId = sheetId,
                            startRowIndex = 0,
                            endRowIndex = 1,
                            startColumnIndex = 0,
                            endColumnIndex = columnCount
                        }
                    }
                }
            };
        }

        private static object CreateAutoResizeRequest(int sheetId, int columnCount)
        {
            return new
            {
                autoResizeDimensions = new
                {
                    dimensions = new
                    {
                        sheetId = sheetId,
                        dimension = "COLUMNS",
                        startIndex = 0,
                        endIndex = columnCount
                    }
                }
            };
        }

        private static async Task BatchUpdateAsync(HttpClient httpClient, string spreadsheetId, List<object> requests, CancellationToken cancellationToken)
        {
            var payload = new
            {
                requests = requests
            };

            string requestUri = $"https://sheets.googleapis.com/v4/spreadsheets/{Uri.EscapeDataString(spreadsheetId)}:batchUpdate";

            var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
            var response = await httpClient.PostAsync(requestUri, content, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                var responseBody = await response.Content.ReadAsStringAsync();
                throw new Exception($"BatchUpdate failure: {(int)response.StatusCode} {response.ReasonPhrase}\n{responseBody}");
            }
        }
    }
}
