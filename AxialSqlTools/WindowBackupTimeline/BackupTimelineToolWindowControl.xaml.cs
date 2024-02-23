namespace AxialSqlTools
{
    using System.Data.SqlClient;
    using System.Diagnostics.CodeAnalysis;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for BackupTimelineToolWindowControl.
    /// </summary>
    public partial class BackupTimelineToolWindowControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="BackupTimelineToolWindowControl"/> class.
        /// </summary>
        public BackupTimelineToolWindowControl()
        {
            this.InitializeComponent();
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


            var ci = ScriptFactoryAccess.GetCurrentConnectionInfo();

            string sourceQuery = @"SELECT 
                database_name, 
                backup_start_date, 
                backup_finish_date, 
                CASE type
                    WHEN 'D' THEN 'Data'
                    WHEN 'L' THEN 'Log'
                END AS BackupType, 
                ROW_NUMBER() OVER (ORDER BY database_name, type) AS RN
            FROM
                msdb.dbo.backupset
            WHERE
                type IN('D', 'L') AND backup_start_date > GETDATE() - 7
            ORDER BY
                database_name, backup_start_date;
";


            //var model = new PlotModel { Title = "SQL Server Backup History" };
            //var series = new RectangleBarSeries();



            using (SqlConnection sourceConn = new SqlConnection(ci.FullConnectionString))
            {
                sourceConn.Open();
                using (SqlCommand cmd = new SqlCommand(sourceQuery, sourceConn))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            //var categoryIndex = reader.GetInt64(4);
                            //var startDate = DateTimeAxis.ToDouble(reader.GetDateTime(1));
                            //var finishDate = DateTimeAxis.ToDouble(reader.GetDateTime(2));

                            //series.Items.Add(new RectangleBarItem(startDate, categoryIndex, finishDate, categoryIndex + 0.8));
                        }
                    }
                }
            }


            //model.Series.Add(series);
            //this.TimelineModel.Model = model;


        }



    }
}