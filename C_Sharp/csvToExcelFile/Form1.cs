using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;

namespace csvToExcelFile
{
    public partial class Form1 : Form
    {
        private static readonly int baseDataIndex = 27;
        private static readonly int changeDateIndex = 33;
        private static readonly int dayCellNum = 2;

        private string[] fileNames;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = System.Environment.CurrentDirectory + @"\csv";
            ofd.Filter = "CSVファイル(*.csv)|*.csv";
            ofd.Title = "読み込むファイル";
            ofd.RestoreDirectory = true;
            ofd.Multiselect = true;

            if(ofd.ShowDialog() == DialogResult.OK)
            {
                fileNames = ofd.FileNames;
            }

            string pathText = "読み込むファイル：" + System.Environment.NewLine;
            foreach (string fileName in fileNames)
            {
                pathText += fileName + System.Environment.NewLine;
            }
            textBox.Text = pathText;
        }
        
        /// <summary>
        /// CSVファイルの読み込み
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private List<string[]> ReadCSVFile(string fileName)
        {
            List<string[]> readCSVList = new List<string[]>();
            using (StreamReader sr = new StreamReader(fileName, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                while (!sr.EndOfStream)
                {
                    string[] words = sr.ReadLine().Split(',');
                    readCSVList.Add(words);
                }
            }
            return readCSVList;
        }

        /// <summary>
        /// 指定されたExcelファイルにリストを保存
        /// </summary>
        /// <param name="readCSVList"></param>
        /// <param name="dir"></param>
        /// <param name="name"></param>
        private void SaveXlsxFile(List<string[]> readCSVList, string dir, string name)
        {
            using (var book = new XLWorkbook(System.Environment.CurrentDirectory + @"\xlsx\template.xlsx"))
            {
                var sheet = book.Worksheet("sheet1");

                for (int i = 0; i < readCSVList.Count; i++)
                {
                    for (int j = 0; j < readCSVList[i].Length; j++)
                    {
                        sheet.Cell(i + 1, j + 1).Value = readCSVList[i][j];
                    }
                }

                book.SaveAs(dir + "/xlsx/" + name + ".xlsx");
            }
        }

        /// <summary>
        /// 変換ボタンを押したらCSVファイルを変換してxlsxに保存する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConvertButton_Click(object sender, EventArgs e)
        {
            foreach (string fileName in fileNames)
            {
                List<string[]> readCSVList = ReadCSVFile(fileName);

                DateTime basedTime = DateTime.ParseExact(readCSVList[baseDataIndex][dayCellNum], "yyyy-MM-dd", null);
                readCSVList[changeDateIndex][dayCellNum] = (basedTime.Year + 1).ToString() + "-" + basedTime.Month.ToString() + "-" + basedTime.Day.ToString();

                string dir = Path.GetDirectoryName(fileName);
                string name = Path.GetFileNameWithoutExtension(fileName);

                SaveXlsxFile(readCSVList, dir, name);
            }
        }
    }
}
