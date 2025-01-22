namespace AxialSqlTools
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Net.Http;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for AskChatGptWindowControl.
    /// </summary>
    public partial class AskChatGptWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AskChatGptWindowControl"/> class.
        /// </summary>
        public AskChatGptWindowControl()
        {
            this.InitializeComponent();

            PromptTextBox.Text = "What is the capital of France?";
        }

        /// <summary>
        /// Handles click on the button by displaying a message box.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Justification = "Sample code")]
        [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Default event handler naming pattern")]
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                string.Format(System.Globalization.CultureInfo.CurrentUICulture, "Invoked '{0}'", this.ToString()),
                "AskChatGptWindow");
        }

        async private void buttonAskChatGpt_Click(object sender, RoutedEventArgs e)
        {

            buttonSend.IsEnabled = false;
            ResponseResult.Text = "";

            string apiKey = SettingsManager.GetOpenAiApiKey();

            // The API URL
            string url = "https://api.openai.com/v1/chat/completions";

            // Create the HTTP client
            using (HttpClient client = new HttpClient())
            {
                // Set the headers
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                
                var jsonObject = new
                {
                    model = "gpt-4o",
                    messages = new List<object>
                    {
                        new { role = "system", content = "You are a helpful assistant." },
                        new { role = "user", content = PromptTextBox.Text }
                    }
                };

                string json = JsonConvert.SerializeObject(jsonObject, Formatting.Indented);

                // Send the request
                StringContent content = new StringContent(json, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(url, content);

                // Get the response as JSON
                string result = await response.Content.ReadAsStringAsync();

                var jsonResult = JObject.Parse(result);

                string contentResult = jsonResult["choices"]?[0]?["message"]?["content"]?.ToString();

                ResponseResult.Text = contentResult;

                buttonSend.IsEnabled = true;

            }



        }
    }
}