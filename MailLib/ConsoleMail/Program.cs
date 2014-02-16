using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using MailLib;

namespace ConsoleMail
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 4) {
                // メール受信のテスト(サーバ、ユーザ名、パスは引数で指定)
                GetMailTest(args[0], args[1], args[2], int.Parse(args[3]));
            }
            else if (args.Length == 3) {
                // メール受信のテスト(サーバ、ユーザ名、パスは引数で指定)
                GetMailTest(args[0], args[1], args[2]);
            }
            else{
                Console.WriteLine("引数が少ないまたは多いため受信処理をスキップします。");
            }

            // メール表示のテスト(通常メール)
            ShowMailTest1();

            // メール表示のテスト(添付メール)
            ShowMailTest2();
        }

        /// <summary>
        /// POP に接続しメールを取得するテストを実行します。
        /// </summary>
        public static void GetMailTest(string hostname, string username, string password, int portnumber = 110)
        {
            try {
                // POP サーバに接続します。
                using (Pop pop = new Pop(hostname, portnumber)) {
                    // ログインします。
                    pop.Login(username, password);

                    // POP サーバに溜まっているメールのリストを取得します。
                    ArrayList list = pop.GetList();
                    ArrayList size_list = pop.GetSizeList();
                    ArrayList uidl_list = pop.GetUidlList();

                    for (int i = 0; i < list.Count; ++i) {
                        // メール本体を取得します。
                        string mail = pop.GetMail((string)list[i]);
                        string size = (string)size_list[i];
                        string uidl = (string)uidl_list[i];

                        // 確認用に取得したメールをそのままカレントディレクトリに書き出します。
                        using (StreamWriter sw = new StreamWriter(DateTime.Now.ToString("yyyyMMddHHmmssfff") + i.ToString("0000") + ".txt")) {
                            sw.Write(mail);
                        }

                        // メールを POP サーバから取得します。
                        // ★注意★
                        // 削除したメールを元に戻すことはできません。
                        // 本当に削除していい場合は以下のコメントをはずしてください。
                        //pop.Delete((string)list[i]);
                    }

                    // 切断します。
                    pop.Close();
                }
            }
            catch (PopException ex) {
                Console.WriteLine(ex.Message);
            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// ファイル化したメールを画面に表示します
        /// </summary>
        private static void ShowMailTest1()
        {
            Console.WriteLine("sample1.txt のテスト");

            // テストのために POP から取得する代わりにファイルから読み込みます。
            string mailtext = null;

            using (StreamReader sr = new StreamReader("sample1.txt")) {
                mailtext = sr.ReadToEnd();
            }

            // Mail クラスを作成します。
            Mail mail = new Mail(mailtext);

            // From、To、Subject を表示します。
            Console.WriteLine(mail.Header["From"][0]);
            Console.WriteLine(mail.Header["To"][0]);
            Console.WriteLine(mail.Header["Subject"][0]);
            Console.WriteLine("--");

            // デコードしたFrom、To、Subject を表示します。
            Console.WriteLine(MailHeader.Decode(mail.Header["From"][0]));
            Console.WriteLine(MailHeader.Decode(mail.Header["To"][0]));
            Console.WriteLine(MailHeader.Decode(mail.Header["Subject"][0]));
            Console.WriteLine("--");

            // Content-Type を表示します。
            Console.WriteLine(mail.Header["Content-Type"][0]);
            Console.WriteLine("--");

            // メール本文を表示します。
            // 本来は Content-Type の charset を参照してデコードすべきですが
            // ここではサンプルとして iso-2022-jp 固定でデコードします。
            byte[] bytes = Encoding.ASCII.GetBytes(mail.Body.Text);
            string mailbody = Encoding.GetEncoding("iso-2022-jp").GetString(bytes);

            Console.WriteLine("↓本文↓");
            Console.WriteLine(mailbody);
        }

        /// <summary>
        /// sample2.txt ファイルを使って Mail クラスのテストを行います。
        /// </summary>
        private static void ShowMailTest2()
        {
            Console.WriteLine("sample2.txt のテスト");

            // テストのために POP から取得する代わりにファイルから読み込みます。
            string mailtext = null;
            
            using (StreamReader sr = new StreamReader("sample2.txt")) {
                mailtext = sr.ReadToEnd();
            }

            // Mail クラスを作成します。
            Mail mail = new Mail(mailtext);

            // From、To、Subject を表示します。
            Console.WriteLine(mail.Header["From"][0]);
            Console.WriteLine(mail.Header["To"][0]);
            Console.WriteLine(mail.Header["Subject"][0]);
            Console.WriteLine("--");

            // デコードしたFrom、To、Subject を表示します。
            Console.WriteLine(MailHeader.Decode(mail.Header["From"][0]));
            Console.WriteLine(MailHeader.Decode(mail.Header["To"][0]));
            Console.WriteLine(MailHeader.Decode(mail.Header["Subject"][0]));
            Console.WriteLine("--");

            // Content-Type を表示します。
            Console.WriteLine(mail.Header["Content-Type"][0]);
            Console.WriteLine("--");

            // 1つ目のパートの Content-Type、メール本文を表示します。
            // 本来は Content-Type の charset を参照してデコードすべきですが
            // ここではサンプルとして iso-2022-jp 固定でデコードします。
            MailMultipart part1 = mail.Body.Multiparts[0];
            Console.WriteLine("パート1");
            Console.WriteLine(part1.Header["Content-Type"][0]);
            Console.WriteLine("--");

            byte[] bytes = Encoding.ASCII.GetBytes(part1.Body.Text);
            string mailbody = Encoding.GetEncoding("iso-2022-jp").GetString(bytes);

            Console.WriteLine("↓本文↓");
            Console.WriteLine(mailbody);
            Console.WriteLine("--");

            // 2つ目のパートの Content-Type を表示し、BASE64 をデコードしてファイルとして保存します。
            // 本来は Content-Transfer-Encoding が base64 であることを確認したり、
            // Content-Type の name を参照してファイル名を決めたりすべきですが、ここでは省略しています。
            MailMultipart part2 = mail.Body.Multiparts[1];
            Console.WriteLine("パート2");
            Console.WriteLine(part2.Header["Content-Type"][0]);
            Console.WriteLine("--");

            bytes = Convert.FromBase64String(part2.Body.Text);
            using (Stream stm = File.Open("sample2.gif", FileMode.Create))
            
            using (BinaryWriter bw = new BinaryWriter(stm)) {
                bw.Write(bytes);
            }

            Console.WriteLine("添付ファイルを保存しました。ファイル名は sample2.gif です。");
        }
    }
}
