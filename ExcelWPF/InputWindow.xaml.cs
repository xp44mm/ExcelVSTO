using System.Windows;

namespace ExcelWPF
{
    /// <summary>
    /// ExecutionWindow.xaml 的交互逻辑
    /// </summary>
    public partial class InputWindow
    {
        public string Input { get; private set; }
        public InputWindow(string title, string input)
        {
            this.InitializeComponent();
            this.Title = title;
            this.tbInput.Text = input;

        }

        private void Ok_Click(System.Object sender, RoutedEventArgs e)
        {
            this.Input = this.tbInput.Text;
            this.DialogResult = true;
            this.Close();

        }
    }
}
