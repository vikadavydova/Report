using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Report
{
    public partial class MainForm : Form
    {
        public enum TypeOfAnalysis
        {
            Statistics,
            Ttest,
            MannWitney,
            Hi,
            ANOVA,
            RegAnalysis,
            Clustering,
            CorAnalysis,
            CorPleiade,
            None
        }

        public List<string> listOfAnalyses = new List<string>() 
        {
            "Описательные статистики",
            "T-тест",
            "Тест Манна-Уитни",
            "Хи-Квадрат",
            "ANOVA",
            "Регрессионный анализ",
            "Кластеризация",
            "Корреляционный анализ"
        };

        private Word.Application wordApp;
        private Word.Document wordDoc;
        private Word.Range wordRange;

        private Excel.Application exApp;
        private Excel.Workbook exBook;

        List<TypeOfAnalysis> Types = new List<TypeOfAnalysis>();
        public int indexOfParagraphs;
        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            chlbxAnalysis.Items.AddRange(listOfAnalyses.ToArray());
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            string name = tbxName.Text;
            string label = tbxLabel.Text;
            DateTime date = dtPicker.Value;
            foreach(string item in chlbxAnalysis.CheckedItems)
            {
                switch(item)
                {
                    case "Описательные статистики": { Types.Add(TypeOfAnalysis.Statistics); break; };
                    case "T-тест": { Types.Add(TypeOfAnalysis.Ttest); break; };
                    case "Тест Манна-Уитни": { Types.Add(TypeOfAnalysis.MannWitney); break; };
                    case "Хи-Квадрат": { Types.Add(TypeOfAnalysis.Hi); break; };
                    case "ANOVA": { Types.Add(TypeOfAnalysis.ANOVA); break; };
                    case "Корреляционный анализ": { Types.Add(TypeOfAnalysis.CorAnalysis); break; };
                    case "Корреляционная плеяда": { Types.Add(TypeOfAnalysis.CorPleiade); break; };
                    case "Регрессионный анализ": { Types.Add(TypeOfAnalysis.RegAnalysis); break; };
                    case "Кластеризация": { Types.Add(TypeOfAnalysis.Clustering); break; };
                    default: { MessageBox.Show("Как вы вообще здесь оказались?!"); break; }
                }
            }
 

            try
            {
                //Создаем объект Word - равносильно запуску Word 
                wordApp = new Word.Application();
                //Делаем его видимым 
                wordApp.Visible = true;
                // открываем документ, соответствующий шаблону
                String templatePath = Environment.CurrentDirectory + "\\TemplateWord.dotx";
                wordDoc = wordApp.Documents.Add(templatePath);

                // вставка текста по закладке (закладки можно создавать в Word: Вставка - Ссылки - Закладка)
                wordRange = wordDoc.Bookmarks["label"].Range;
                wordRange.Text = label;

                wordRange = wordDoc.Bookmarks["gender"].Range;
                if(chbxFemale.Checked)
                {
                    wordRange.Text = "а";
                }
                wordRange = wordDoc.Bookmarks["name"].Range;
                wordRange.Text = name;

                wordRange = wordDoc.Bookmarks["date"].Range;
                wordRange.Text = date.ToString("D");

                wordDoc.Paragraphs.Add();

                wordDoc.Paragraphs.Add();
                wordRange = wordDoc.Paragraphs[7].Range;
                wordRange.Text = "Здесь будет описание данных... ";
                wordRange.Font.Size = 12;
                wordRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                wordDoc.Paragraphs.Add();
                indexOfParagraphs = 9;
                Types.Sort();

                foreach(var type in Types)
                {
                    switch (type)
                    {
                        case TypeOfAnalysis.Statistics:
                            {
                                indexOfParagraphs = addSomeInfo("Описательные статистики", indexOfParagraphs);
                                break;
                            };
                        case TypeOfAnalysis.Ttest:
                            {
                                indexOfParagraphs = addSomeInfo("Т Тест", indexOfParagraphs);
                                break;
                            };
                        case TypeOfAnalysis.MannWitney:
                            {
                                indexOfParagraphs = addSomeInfo("Тест Манна-Уитни", indexOfParagraphs);
                                break;
                            };
                        case TypeOfAnalysis.Hi:
                            {
                                indexOfParagraphs = addSomeInfo("Хи-Квадрат", indexOfParagraphs);
                                break;
                            };
                        case TypeOfAnalysis.ANOVA:
                            {
                                indexOfParagraphs = addSomeInfo("АНОВА", indexOfParagraphs);
                                break;
                            };
                        case TypeOfAnalysis.CorAnalysis:
                            {
                                indexOfParagraphs = addSomeInfo("Корреляционный анализ", indexOfParagraphs);

                                //вставка таблицы
                                Word.Table wordtable = wordDoc.Tables.Add(wordRange, 3, 3);

                                // прорисовка границ -- по умолчанию границы таблицы не рисуются
                                wordtable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                wordtable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDashSmallGap;

                                for(int i = 1; i < 4; ++i)
                                {
                                    for(int j = 1; j < 4; ++j)
                                    {
                                        if(i == 1 && j == 1)
                                        {
                                            continue;
                                        }

                                        if(i == 1)
                                        {
                                            wordRange = wordtable.Cell(i, j).Range;
                                            wordRange.Text = "X" + (j - 1);
                                            continue;
                                        }

                                        if (j == 1)
                                        {
                                            wordRange = wordtable.Cell(i, j).Range;
                                            wordRange.Text = "Y" + (i - 1);
                                            continue;
                                        }

                                        wordRange = wordtable.Cell(i, j).Range;
                                    }
                                }

                                wordRange = wordtable.Cell(3, 3).Range;
                                wordDoc.Paragraphs.Add();
                                indexOfParagraphs = wordDoc.Paragraphs.Count;
                                wordRange = wordDoc.Paragraphs[indexOfParagraphs].Range;
                                break;
                            };
                        case TypeOfAnalysis.CorPleiade:
                            {
                                indexOfParagraphs = addSomeInfo("Корреляционная плеяда", indexOfParagraphs);
                                break;
                            };
                        case TypeOfAnalysis.RegAnalysis:
                            {
                                indexOfParagraphs = addSomeInfo("Регрессионный анализ", indexOfParagraphs);
                                break;
                            };
                        case TypeOfAnalysis.Clustering:
                            {
                                indexOfParagraphs = addSomeInfo("Кластеризация", indexOfParagraphs);
                                break;
                            };
                        default:
                            {
                                MessageBox.Show("Как вы вообще здесь оказались?!");
                                break;
                            }
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Упс.. " + ex.Message);
                wordApp.Quit();
                wordDoc = null;
                wordApp = null;

            }
        }

        private int addSomeInfo(string labelOfMethod, int index)
        {
            wordDoc.Paragraphs.Add();
            wordRange = wordDoc.Paragraphs[index].Range;
            wordRange.Text = labelOfMethod;
            wordRange.Font.Size = 14;
            wordRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wordRange.Bold = 1;
            index++;

            wordDoc.Paragraphs.Add();
            wordRange = wordDoc.Paragraphs[index].Range;
            wordRange.Text = "Здесь будет текст...:)";
            wordRange.Font.Size = 12;
            wordRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wordRange.Bold = 0;
            index++;

            wordDoc.Paragraphs.Add();
            wordRange = wordDoc.Paragraphs[index].Range;
            index++;

            return index;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if (wordApp != null)
                {
                    wordApp.Quit();
                    wordDoc = null;
                    wordApp = null;
                    MessageBox.Show("Закрыто");
                }

                if (exApp != null)
                {
                    exApp.Quit();
                    exApp = null;
                    MessageBox.Show("Закрыто");
                }

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("Закрыто ранее");
            }
        }

        private void chbxMale_CheckedChanged(object sender, EventArgs e)
        {
            if(chbxMale.Checked)
            {
                chbxFemale.Enabled = false;
            }
            else
            {
                chbxFemale.Enabled = true;
            }
        }

        private void chbxFemale_CheckedChanged(object sender, EventArgs e)
        {
            if (chbxFemale.Checked)
            {
                chbxMale.Enabled = false;
            }
            else
            {
                chbxMale.Enabled = true;
            }
        }

        private void chlbxAnalysis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (chlbxAnalysis.GetItemCheckState(5) == CheckState.Checked || chlbxAnalysis.GetItemCheckState(7) == CheckState.Checked)
            {
                excelButton.Visible = true;
            }
            else
            {
                excelButton.Visible = false;
            }

            if (chlbxAnalysis.GetItemCheckState(7) == CheckState.Checked)
            {
                if (!chlbxAnalysis.Items.Contains("Корреляционная плеяда"))
                {
                    chlbxAnalysis.Items.Add("Корреляционная плеяда");
                }
            }
            else
            {
                chlbxAnalysis.Items.Remove("Корреляционная плеяда");
            }
        }

        private void excelButton_Click(object sender, EventArgs e)
        {
            // создание нового файла 
            try
            {
                //Создаем объект Excel - равносильно запуску Excel 
                exApp = new Excel.Application();
                //Делаем его видимым 
                exApp.Visible = true;
                foreach (string item in chlbxAnalysis.CheckedItems)
                {
                    switch (item)
                    {
                        case "Корреляционный анализ": { Types.Add(TypeOfAnalysis.CorAnalysis); break; };
                        case "Регрессионный анализ": { Types.Add(TypeOfAnalysis.RegAnalysis); break; };
                        default: { break; }
                    }
                }

                Excel.Worksheet excelWS;

                if (Types.Contains(TypeOfAnalysis.CorAnalysis) && Types.Contains(TypeOfAnalysis.RegAnalysis))
                {
                    exApp.SheetsInNewWorkbook = 2;
                    exBook = exApp.Workbooks.Add();

                    makeCorSheet(1);

                    makeRegSheet(2);
                }
                else
                {
                    if (Types.Contains(TypeOfAnalysis.CorAnalysis))
                    {
                        exBook = exApp.Workbooks.Add();
                        makeCorSheet(1);
                    }
                    if (Types.Contains(TypeOfAnalysis.RegAnalysis))
                    {
                        exBook = exApp.Workbooks.Add();
                        makeRegSheet(1);
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Упс.. " + ex.Message);
                exApp.Quit();
                exApp = null;

            }
            Types.Clear();
        }

        private void makeCorSheet(int number)
        {
            Excel.Worksheet sheet = exBook.Worksheets[number];
            sheet.Name = "Корреляционный анализ";
            Excel.Range range;

            for (int i = 1; i < 4; ++i)
            {
                for (int j = 1; j < 4; ++j)
                {
                    if (i == 1 && j == 1)
                    {
                        range = sheet.Cells[i, j];
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        continue;
                    }

                    if (i == 1)
                    {
                        range = sheet.Cells[i, j];
                        range.Value = "X" + (j - 1);
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        continue;
                    }

                    if (j == 1)
                    {
                        range = sheet.Cells[i, j];
                        range.Value = "Y" + (i - 1);
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        continue;
                    }

                    range = sheet.Cells[i, j];
                }
            }
        }

        private void makeRegSheet(int number)
        {
            var sheet = exBook.Worksheets[number];
            sheet.Name = "Регрессионный анализ";
            Excel.Range range;

            range = sheet.Cells[1, 1];
            range.Value = "Коэффициент";
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.EntireColumn.AutoFit();
            range.EntireRow.AutoFit();

            range = sheet.Cells[1, 2];
            range.Value = "Значение";
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            range = sheet.Cells[1, 3];
            range.Value = "P-Value";
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            range = sheet.Cells[2, 1];
            range.Value = "Коэффициент 1";
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.EntireColumn.AutoFit();
            range.EntireRow.AutoFit();

            range = sheet.Cells[3, 1];
            range.Value = "Коэффициент 2";
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.EntireColumn.AutoFit();
            range.EntireRow.AutoFit();
        }
    }
}
