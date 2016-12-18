using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using C1.WPF.Word;
using System.IO;

namespace C1WordSample
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // ドキュメントクラスを作成
            C1WordDocument word = new C1WordDocument();
            word.Clear();

            // ドキュメント情報を設定します
            word.Info.Author = "CodeZine";
            word.Info.Subject = "C1Wordサンプル";
            word.Info.Title = "サンプル";

            /**
             * テキストの表示処理
             */
            this.drawText(word);

            /**
             * 画像の表示処理
             */
            this.drawImage(word);

            /**
             * 図形の描画
             */
            this.drawCircle(word);

            // ドキュメントフォルダーのパス
            string document_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // ファイル名
            string file_name = "sample.docx";

            // 保存処理
            word.Save(document_path + System.IO.Path.DirectorySeparatorChar + file_name);
        }

        private void drawText(C1WordDocument word)
        {
            // テキストとフォントの設定 
            string text = "C1WordでWord形式のファイルを保存するサンプル";
            Font font = new Font("Segoe UI Light", 20, RtfFontStyle.Bold);

            // パラグラフを追加
            word.AddParagraph(text, font, Colors.Blue, RtfHorizontalAlignment.Justify);
        }

        private void drawImage(C1WordDocument word)
        {
            // ドキュメントフォルダーのパス
            string document_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            /**
             * 画像の表示処理
             */

            BitmapImage bitmap = new BitmapImage();
            FileStream stream = File.OpenRead(document_path + System.IO.Path.DirectorySeparatorChar + "sample.png");

            // BitmapImageの初期化
            bitmap.BeginInit();

            bitmap.StreamSource = stream;

            bitmap.EndInit();

            var wb = new WriteableBitmap(bitmap);

            // 描画領域の指定
            Rect rect = new Rect(30, 30, 120, 120);

            // 画像の描画
            word.DrawImage(wb, rect);
        }

        private void drawCircle(C1WordDocument word)
        {
            // 図形を描く範囲を指定
            Rect rect = new Rect(30, 180, 120, 120);

            // 円を描く
            word.FillPie(Colors.Red, rect, 0, 360);
        }
    }
}
