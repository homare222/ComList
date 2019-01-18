using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace ComList
{
    static class Program
    {
        public static Settings Settings;

        //
        // アプリケーションのメインエントリポイント
        //
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault( false );

            // 設定を読込
            Settings.Load( ref Settings );

            // 実行
            Application.Run( new frmMain() );

            // 設定を保存
            Settings.Save( ref Settings );
        }
    }

    //
    // 設定クラス
    //
    public class Settings
    {
        public List<string> ProgramName;            // 実行するプログラム名
        public List<string> ProgramPath;            // 実行するプログラムのパス
        public List<string> ProgramArguments;       // 実行するプログラムに渡す引数

        //
        // 設定ファイルのパス
        //
        private static string FilePath()
        {
            // exeと同じディレクトリの"(アプリケーション名)Settings.xml"
            return Path.GetDirectoryName( Assembly.GetExecutingAssembly().Location ) + @"\" + Application.ProductName + "Settings.xml";
        }

        //
        // 設定を保存
        //
        public static void Load( ref Settings Settings )
        {
            if ( !File.Exists( FilePath() ) ) {
                Settings = new Settings();  // ファイルがない場合は、空のインスタンスを生成する
                Settings.ProgramName = new List<string>{ "Tera Term 9600bps", "Tera Term 115200bps", };
                Settings.ProgramPath = new List<string>{ @"C:\Program Files (x86)\teraterm\ttermpro.exe", @"C:\Program Files (x86)\teraterm\ttermpro.exe", };
                Settings.ProgramArguments = new List<string>{ "/C=%% /BAUD=9600", "/C=%% /BAUD=115200" };
                return;
            }

            XmlSerializer   slzr    = new XmlSerializer( typeof( Settings ) );
            StreamReader    sr      = new StreamReader( FilePath(), new UTF8Encoding( false ) );
            Settings = (Settings)slzr.Deserialize( sr );
            sr.Close();
        }

        //
        // 設定を読込
        //
        public static void Save( ref Settings Settings )
        {
            XmlSerializer   slzr    = new XmlSerializer( typeof( Settings ) );
            StreamWriter    sw      = new StreamWriter( FilePath(), false, new UTF8Encoding( false ) );
            slzr.Serialize( sw, Settings );
            sw.Close();
        }
    }
}
