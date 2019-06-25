using System.Windows;
using System.Windows.Controls;
using InsuranceCompareTool.ViewModels;
namespace InsuranceCompareTool.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public MainWindowViewModel ViewModel
        {
            get
            {
                return this.DataContext as MainWindowViewModel; 
            }
        }
    

    private void Selector_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
         ViewModel.SelectTabCommand.Execute(e);
    }
 
    }
}
