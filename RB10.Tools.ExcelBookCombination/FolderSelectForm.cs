using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RB10.Tools.ExcelBookCombination
{
    public partial class FolderSelectForm : Form
    {
        public FolderSelectForm()
        {
            InitializeComponent();
        }

        private void FolderSelectButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.Cancel) return;

            TargetFolderTextBox.Text = dlg.SelectedPath;
        }

        private void ExecuteButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (TargetFolderTextBox.Text == "")
                {
                    MessageBox.Show("読み込み対象のファイルが格納されているフォルダーを選択して下さい。", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }              

                FolderBrowserDialog dlg = new FolderBrowserDialog();
                dlg.Description = "出力先のフォルダーを選択してください。";
                if (dlg.ShowDialog() == DialogResult.Cancel) return;

                var bc = new BookCombination(TargetFolderTextBox.Text, dlg.SelectedPath);
                bc.Run();

                MessageBox.Show("正常終了", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
