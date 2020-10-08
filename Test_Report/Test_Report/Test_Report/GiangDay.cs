using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test_Report
{
    class GiangDay
    {
        public static string[] dulieu = new string[10];
        public static int thuc_giang_A=0;
        public static int phai_giang_A=0;
        public static int thuc_giang_B=0;
        public static string khoa;
        public static string boMon;
        public static string Day;
        public static string Month;
        public static string Year;
        public static string namHoc;
        public static string HoTen;
        public static string namSinh;
        public static string chucVu;
        public static string luong;
        public static string hocHam;
        public void ReplaceWordStub(string stubtoReplace, string text, Microsoft.Office.Interop.Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubtoReplace, ReplaceWith: text);
        }

    }
}
