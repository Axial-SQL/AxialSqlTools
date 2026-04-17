namespace AxialSqlTools
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Diagnostics.CodeAnalysis;
    using System.Windows;
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for HealthDashboard_ServersControl.
    /// </summary>
    public partial class HealthDashboard_ServersControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HealthDashboard_ServersControl"/> class.
        /// </summary>
        /// 
        public ObservableCollection<MyRowModel> Items { get; set; }


        public HealthDashboard_ServersControl()
        {
            this.InitializeComponent();

            Items = new ObservableCollection<MyRowModel>();
            // Bind the collection to the DataGrid
            MyDataGrid.ItemsSource = Items;

            // Example: Adding rows to the DataGrid
            Items.Add(new MyRowModel { Property1 = "Value1", Property2 = "Value2" });
            Items.Add(new MyRowModel { Property1 = "Value3", Property2 = "Value4" });
            // Add more items as needed


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
                "HealthDashboard_Servers");
        }

    }


    public class MyRowModel
    {
        public string Property1 { get; set; }
        public string Property2 { get; set; }
        // Add additional properties as needed
    }
}