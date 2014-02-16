using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net.Sockets;

namespace MailLib
{
    public class Pop : IDisposable
    {
        /// <summary>TCP 接続</summary>
        private TcpClient tcp = null;

        /// <summary>TCP 接続からのリーダー</summary>
        private StreamReader reader = null;

        /// <summary>
        /// コンストラクタです。POPサーバと接続します。
        /// </summary>
        /// <param name="hostname">POPサーバのホスト名。</param>
        /// <param name="port">POPサーバのポート番号（通常は110）。</param>
        public Pop(string hostname, int port)
        {
            // サーバと接続
            this.tcp = new TcpClient(hostname, port);
            this.reader = new StreamReader(this.tcp.GetStream(), Encoding.ASCII);

            // オープニング受信
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("接続時に POP サーバが \"" + s + "\" を返しました。");
            }
        }

        /// <summary>
        /// 解放処理を行います。
        /// </summary>
        public void Dispose()
        {
            if (this.reader != null) {
                ((IDisposable)this.reader).Dispose();
                this.reader = null;
            }
            if (this.tcp != null) {
                ((IDisposable)this.tcp).Dispose();
                this.tcp = null;
            }
        }

        /// <summary>
        /// POP サーバにログインします。
        /// </summary>
        /// <param name="username">ユーザ名。</param>
        /// <param name="password">パスワード。</param>
        public void Login(string username, string password)
        {
            // ユーザ名送信
            SendLine("USER " + username);
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("USER 送信時に POP サーバが \"" + s + "\" を返しました。");
            }

            // パスワード送信
            SendLine("PASS " + password);
            s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("PASS 送信時に POP サーバが \"" + s + "\" を返しました。");
            }
        }

        /// <summary>
        /// POP サーバに溜まっているメールのリストを取得します。
        /// </summary>
        /// <returns>System.String を格納した ArrayList。</returns>
        public ArrayList GetList()
        {
            // LIST 送信
            SendLine("LIST");
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("LIST 送信時に POP サーバが \"" + s + "\" を返しました。");
            }

            // サーバにたまっているメールの数を取得
            ArrayList list = new ArrayList();

            while (true) {
                s = ReadLine();
                if (s == ".") {
                    // 終端に到達
                    break;
                }
                // メール番号部分のみを取り出し格納
                int p = s.IndexOf(' ');
                if (p > 0) {
                    s = s.Substring(0, p);
                }
                list.Add(s);
            }
            return list;
        }

        /// <summary>
        /// POP サーバに溜まっているメールのサイズリストを取得します。
        /// </summary>
        /// <returns>System.String を格納した ArrayList。</returns>
        public ArrayList GetSizeList()
        {
            // LIST 送信
            SendLine("LIST");
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("LIST 送信時に POP サーバが \"" + s + "\" を返しました。");
            }

            // サーバにたまっているメールの数を取得
            ArrayList list = new ArrayList();

            while (true) {
                s = ReadLine();
                if (s == ".") {
                    // 終端に到達
                    break;
                }
                // メールサイズ部分のみを取り出し格納
                int p = s.IndexOf(' ');
                if (p > 0) {
                    s = s.Substring((p + 1), s.Length - (p + 1));
                }
                list.Add(s);
            }
            return list;
        }

        /// <summary>
        /// POP サーバに溜まっているメールの  UIDLリストを取得します。
        /// </summary>
        /// <returns>System.String を格納した ArrayList。</returns>
        public ArrayList GetUidlList()
        {
            // LIST 送信
            SendLine("UIDL");
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("UIDL 送信時に POP サーバが \"" + s + "\" を返しました。");
            }

            // サーバにたまっているメールの数を取得
            ArrayList list = new ArrayList();

            while (true) {
                s = ReadLine();
                if (s == ".") {
                    // 終端に到達
                    break;
                }
                // UIDL部分のみを取り出し格納
                int p = s.IndexOf(' ');
                if (p > 0) {
                    s = s.Substring((p + 1), s.Length - (p + 1));
                }
                list.Add(s);
            }
            return list;
        }

        /// <summary>
        /// POP サーバからメールを 1つ取得します。
        /// </summary>
        /// <param name="num">GetList() メソッドで取得したメールの番号。</param>
        /// <returns>メールの本体。</returns>
        public string GetMail(string num)
        {
            // RETR 送信
            SendLine("RETR " + num);
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("RETR 送信時に POP サーバが \"" + s + "\" を返しました。");
            }

            // メール取得
            StringBuilder sb = new StringBuilder();
            while (true) {
                s = ReadLine();
                if (s == ".") {
                    // "." のみの場合はメールの終端を表す
                    break;
                }
                sb.Append(s);
                sb.Append("\r\n");
            }
            return sb.ToString();
        }

        /// <summary>
        /// POP サーバのメールを 1つ削除します。
        /// </summary>
        /// <param name="num">GetList() メソッドで取得したメールの番号。</param>
        public void Delete(string num)
        {
            // DELE 送信
            SendLine("DELE " + num);
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("DELE 送信時に POP サーバが \"" + s + "\" を返しました。");
            }
        }

        /// <summary>
        /// POP サーバと切断します。
        /// </summary>
        public void Close()
        {
            // QUIT 送信
            SendLine("QUIT");
            string s = ReadLine();
            if (!s.StartsWith("+OK")) {
                throw new PopException("QUIT 送信時に POP サーバが \"" + s + "\" を返しました。");
            }

            ((IDisposable)this.reader).Dispose();
            this.reader = null;
            ((IDisposable)this.tcp).Dispose();
            this.tcp = null;
        }

        /// <summary>
        /// POP サーバにコマンドを送信します。
        /// </summary>
        /// <param name="s">送信する文字列。</param>
        private void Send(string s)
        {
            Print("送信: " + s);
            byte[] b = Encoding.ASCII.GetBytes(s);
            this.tcp.GetStream().Write(b, 0, b.Length);
        }

        /// <summary>
        /// POP サーバにコマンドを送信します。末尾に改行を付加します。
        /// </summary>
        /// <param name="s">送信する文字列。</param>
        private void SendLine(string s)
        {
            Print("送信: " + s + "\\r\\n");
            byte[] b = Encoding.ASCII.GetBytes(s + "\r\n");
            this.tcp.GetStream().Write(b, 0, b.Length);
        }

        /// <summary>
        /// POP サーバから 1行読み込みます。
        /// </summary>
        /// <returns>読み込んだ文字列。</returns>
        private string ReadLine()
        {
            string s = this.reader.ReadLine();
            Print("受信: " + s + "\\r\\n");
            return s;
        }

        /// <summary>
        /// チェック用にコンソールに出力します。
        /// </summary>
        /// <param name="msg">出力する文字列。</param>
        private void Print(string msg)
        {
            Console.WriteLine(msg);
        }
    }
}
