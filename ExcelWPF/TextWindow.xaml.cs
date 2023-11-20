using System.Windows;

namespace ExcelWPF
{
    public partial class TextWindow
    {
        public TextWindow(string title, string textContent)
        {
            this.InitializeComponent();
            this.Title = title;
            this.TextContent.Text = textContent;
        }

    }
}
