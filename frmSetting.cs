using System;
using System.Drawing;
using System.Windows.Forms;

namespace ComList
{
    public partial class frmSetting :Form
    {
        private const int programNum    = 10;
        private const int padding       = 8;

        public frmSetting()
        {
            // デザイナーサポートに必要なメソッド(自動生成)
            InitializeComponent();

            // タイトル
            Text = Application.ProductName + "設定";

            // 実行するプログラムの入力欄を生成
            for ( int i = 0; i < programNum; i++ ) {
                var lblProgram          = new Label();
                var txtProgramName      = new TextBox();
                var txtProgramPath      = new TextBox();
                var txtProgramArguments = new TextBox();

                lblProgram.Name = "lblProgram" + i;
                lblProgram.Text = "プログラム" + (i + 1);

                txtProgramName.Name = "txtProgramName" + i;
                if ( i < Program.Settings.ProgramName.Count ) {
                    txtProgramName.Text = Program.Settings.ProgramName[i];
                }

                txtProgramPath.Name = "txtProgramPath" + i;
                if ( i < Program.Settings.ProgramPath.Count ) {
                    txtProgramPath.Text = Program.Settings.ProgramPath[i];
                }

                txtProgramArguments.Name = "txtProgramArguments" + i;
                if ( i < Program.Settings.ProgramArguments.Count ) {
                    txtProgramArguments.Text = Program.Settings.ProgramArguments[i];
                }

                lblProgram.Left = padding;
                lblProgram.TextAlign = ContentAlignment.MiddleLeft;

                lblProgram.Height = 25;
                lblProgram.Top = (lblProgram.Height * i) + padding;
                txtProgramName.Top = (lblProgram.Height * i) + padding;
                txtProgramPath.Top = (lblProgram.Height * i) + padding;
                txtProgramArguments.Top = (lblProgram.Height * i) + padding;

                Controls.Add( lblProgram );
                Controls.Add( txtProgramName );
                Controls.Add( txtProgramPath );
                Controls.Add( txtProgramArguments );
            }

            // 幅調整
            frmSetting_Resize( null, null );
        }

        //
        // フォームリサイズ
        //
        private void frmSetting_Resize( object sender, EventArgs e )
        {
            // テキストボックスの幅をフォーム幅に応じて変更する ... 実装がいまいち
            for ( int i = 0; i < programNum; i++ ) {
                var lblProgram          = Controls["lblProgram" + i];
                var txtProgramName      = Controls["txtProgramName" + i];
                var txtProgramPath      = Controls["txtProgramPath" + i];
                var txtProgramArguments = Controls["txtProgramArguments" + i];

                txtProgramName.Width = (int)((Width - (padding * 3) - (lblProgram.Bounds.X + lblProgram.Width)) * 0.25);
                txtProgramPath.Width = (int)((Width - (padding * 3) - (lblProgram.Bounds.X + lblProgram.Width)) * 0.55);
                txtProgramArguments.Width = (int)((Width - (padding * 3) - (lblProgram.Bounds.X + lblProgram.Width)) * 0.20);

                txtProgramName.Left = lblProgram.Bounds.X + lblProgram.Width;
                txtProgramPath.Left = txtProgramName.Bounds.X + txtProgramName.Width;
                txtProgramArguments.Left = txtProgramPath.Bounds.X + txtProgramPath.Width;
            }
        }

        //
        // OKボタン押下
        //
        private void btnOk_Click( object sender, EventArgs e )
        {
            // リストをクリア
            Program.Settings.ProgramName.Clear();
            Program.Settings.ProgramPath.Clear();
            Program.Settings.ProgramArguments.Clear();

            // 表示内容をリスト化
            for ( int i = 0; i < programNum; i++ ) {
                if ( Controls["txtProgramName" + i] != null ) {
                    Program.Settings.ProgramName.Add( Controls["txtProgramName" + i].Text );
                    Program.Settings.ProgramPath.Add( Controls["txtProgramPath" + i].Text );
                    Program.Settings.ProgramArguments.Add( Controls["txtProgramArguments" + i].Text );
                }
            }

            // フォーム閉じる
            Close();
        }

        //
        // キャンセル押下
        //
        private void btnCancel_Click( object sender, EventArgs e )
        {
            // フォーム閉じる
            Close();
        }
    }
}
