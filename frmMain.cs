using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;
using System.Management;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ComList
{
    public partial class frmMain :Form
    {
        //
        // メンバ変数
        //
        private Timer       DeviceChangeDelay       = new Timer();  // デバイス変更時のコンテキストメニュー更新の遅延実行用タイマ
        private const int   DeviceChangeDelayTime   = 5000;         // デバイス変更時のコンテキストメニュー更新の遅延実行用タイマの時間
        private const int   ComDescriptionMaxLength = 40;           // COMポートのDescriptionの最大文字数(超えた場合は省略)

        //
        // コンストラクタ
        //
        public frmMain()
        {
            // デザイナーサポートに必要なメソッド(自動生成)
            InitializeComponent();

            // タイトル
            Text = Application.ProductName;
            notifyIcon.Text = Application.ProductName;

            // タスクバーに表示しない
            ShowInTaskbar = false;

            // 最小化で表示
            WindowState = FormWindowState.Minimized;

            // コンテキストメニューを登録
            createContextMenu();

            // デバイス変更時のコンテキストメニュー更新の遅延実行用タイマの初期化
            DeviceChangeDelay.Interval = DeviceChangeDelayTime;
            DeviceChangeDelay.Tick += new EventHandler( DeviceChangeDelay_Tick );
        }

        //
        // フォーム表示時
        //
        private void frmMain_Shown( object sender, EventArgs e )
        {
            // 非表示に
            Hide();
        }

        //
        // アプリの切り替え(Alt+Tab)に表示しない処理
        //
        const int WS_EX_TOOLWINDOW = 0x00000080;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle = cp.ExStyle | WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        //
        // ウィンドウプロシージャのオーバーライド
        //
        protected override void WndProc( ref Message m )
        {
            switch ( m.Msg ) {
            case 0x0219:    // 0x0219 WM_DEVICECHANGE ... デバイスの挿抜を検知して、メニューを更新する
                Debug.WriteLine( "WM_DEVICECHANGE" );
                DeviceChangeDelay.Start();  // 挿抜時は、バタつくので遅延実行する
                break;
            }
            base.WndProc( ref m );
        }

        //
        // デバイス変更時のコンテキストメニュー更新の遅延実行用タイマのハンドラ
        //
        private void DeviceChangeDelay_Tick( object sender, EventArgs e )
        {
            Debug.WriteLine( "DeviceChangeDelay_Tick" );
            DeviceChangeDelay.Stop();
            createContextMenu();
        }

        //
        // コンテキストメニューの生成
        //
        private void createContextMenu()
        {
            // 生成前にコンテキストメニューを一旦クリア
            contextMenuStrip.Items.Clear();

            // シリアルポート名のリストから生成
            var portList = SerialPort.GetPortNames();
            foreach ( string portName in portList ) {
                // コンテキストメニューに追加
                var tsiCom = new ToolStripMenuItem();
                contextMenuStrip.Items.Add( tsiCom );

                // COMポート文字列とDescriptionを表示
                tsiCom.Text = portName + Environment.NewLine + " " + getPortProperty( portName, "Description" );
                if ( tsiCom.Text.Length > ComDescriptionMaxLength ) {
                    tsiCom.Text = tsiCom.Text.Substring( 0, ComDescriptionMaxLength ) + " ...";
                }

                // プログラム実行のサブメニューの追加
                var tsiOpenWith = new List<ToolStripMenuItem>();
                for ( int i = 0; i < Program.Settings.ProgramName.Count; i++ ) {
                    if ( Program.Settings.ProgramName[i].Length == 0 ) {
                        continue;
                    }
                    var tsi = new ToolStripMenuItem();
                    tsi.Text = Program.Settings.ProgramName[i];
                    tsi.Click += ToolStripMenuItemOpenWith_Click;
                    tsiOpenWith.Add( tsi );
                }
                tsiCom.DropDownItems.AddRange( tsiOpenWith.ToArray() );
            }

            // セパレータの追加
            contextMenuStrip.Items.Add( new ToolStripSeparator() );

            // 設定メニューの追加
            var tsiSetting = new ToolStripMenuItem();
            tsiSetting.Text = "設定";
            tsiSetting.Click += ToolStripMenuItemSetting_Click;
            contextMenuStrip.Items.Add( tsiSetting );

            // 終了メニューの追加
            var tsiExit = new ToolStripMenuItem();
            tsiExit.Text = "終了";
            tsiExit.Click += ToolStripMenuItemExit_Click;
            contextMenuStrip.Items.Add( tsiExit );
        }

        //
        // COMポートの情報を取得する
        //
        private string getPortProperty( string portName, string propertyName )
        {
            try {
                // WMIを使用しデバイスの情報を取得し、COMポートの一覧を抽出する
                //      * WMI ... Windows Management Instrumentation、ハードウェアやソフトウェアの情報取得するインターフェース。
                //      * Win32_PnPEntity ... プラグアンドプレイデバイスに関する情報を取得する。プロパティ一覧は、http://www.wmifun.net/library/win32_pnpentity.html を参照。
                //      * 設定で、参照 -> 参照の追加、で System.Management にチェックが必要?
                //      * PowershellでGet-WmiObject -Query "SELECT * FROM Win32_PnPEntity"を実行するのと同様?
                //
                ManagementObjectSearcher mos = new ManagementObjectSearcher();
                mos.Query.QueryString = "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(" + portName + ")%'";
                var moc = mos.Get();
                foreach ( var m in moc ) {
                    Debug.WriteLine( "Name         :" + m.GetPropertyValue( "Name" ) );
                    Debug.WriteLine( "Caption      :" + m.GetPropertyValue( "Caption" ) );
                    Debug.WriteLine( "Description  :" + m.GetPropertyValue( "Description" ) );
                    Debug.WriteLine( "Manufacturer :" + m.GetPropertyValue( "Manufacturer" ) );
                    Debug.WriteLine( "Service      :" + m.GetPropertyValue( "Service" ) );

                    if ( m.GetPropertyValue( propertyName ) == null ) {
                        return "";
                    }
                    return m.GetPropertyValue( propertyName ).ToString();
                }
            }
            catch ( Exception ex ) {
                Debug.WriteLine( ex.Message );
            }
            return "";
        }

        //
        // メニューのプログラム実行をクリック
        //
        private void ToolStripMenuItemOpenWith_Click( object sender, EventArgs e )
        {
            // 親アイテムのテキストからCOMポート番号を抽出
            Match m = Regex.Match( ((ToolStripMenuItem)sender).OwnerItem.Text, @"COM([1-9][0-9]?[0-9]?)" );
            if ( !m.Success ) {
                return;
            }
            var comPort = m.Groups[1].ToString();

            // 設定から実行するプログラムの情報を取得
            var programName = ((ToolStripMenuItem)sender).Text;
            var programPath = Program.Settings.ProgramPath[Program.Settings.ProgramName.IndexOf( programName )];
            var programArguments = Program.Settings.ProgramArguments[Program.Settings.ProgramName.IndexOf( programName )];

            // プログラムのパスの環境変数を展開
            programPath = Environment.ExpandEnvironmentVariables( programPath );
            if ( !File.Exists( programPath ) ) {    // 存在チェック
                MessageBox.Show( "指定されたプログラムが存在しません。", "プログラム実行エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
                return;
            }

            // 引数の %% をCOMポート番号に置換
            programArguments = programArguments.Replace( "%%", comPort );

            // プログラムを実行
            Debug.WriteLine( programName + ", " + programPath + ", " + programArguments );
            Process.Start( programPath, programArguments );
        }

        //
        // メニューの設定をクリック
        //
        private void ToolStripMenuItemSetting_Click( object sender, EventArgs e )
        {
            // 設定フォームを表示
            var frmSetting = new frmSetting();
            frmSetting.Show();

            // コンテキストメニューを更新
            createContextMenu();
        }

        //
        // メニューの終了をクリック
        //
        private void ToolStripMenuItemExit_Click( object sender, EventArgs e )
        {
            Close();
        }
    }
}
