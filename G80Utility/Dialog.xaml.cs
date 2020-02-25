using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;


namespace G80Utility
{
    /// <summary>
    /// AdminDialog.xaml 的互動邏輯
    /// </summary>
    public partial class Dialog : Window
    {
        public Dialog()
        {
            InitializeComponent();
        }

        public string PwdText
        {
            get { return PwdTextBox.Text; }
            set { PwdTextBox.Text = value; }
        }

        //確認
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        //取消
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
