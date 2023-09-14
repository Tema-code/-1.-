using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using W = Microsoft.Office.Interop.Word;
using System.Reflection;
using Microsoft.Office.Interop.Word;

namespace Работа1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int yearPostup = 2021; 
        int month = 8; 
        int year = 2023; 
        int kurs = -1;

        private void button1_Click(object sender, EventArgs e)
        {
            {
                if (month >= 9)
                {
                    kurs += 1;
                }
                for (int i = year - yearPostup; i >= 0; i--)
                {
                    kurs += 1;
                }
                
                object EndOfDoc = "\\endofdoc";
                W.Application oWord = new W.Application();
                W.Document oDoc = oWord.Documents.Add();

                W.Paragraph oPrg = oDoc.Paragraphs.Add();
                oWord.Visible = true;

                oPrg.Range.Text = "Федеральное государственное бюджетное образовательное учреждение высшего образования";

                oPrg.Format.LeftIndent = 0; 
                oPrg.Range.Font.Bold = 0; 
                oPrg.Range.Font.AllCaps = 0; 
                oPrg.Range.Font.Name = "Times New Roman"; 
                oPrg.Range.Font.Size = 14f; 

                oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter; 
                oPrg.Range.InsertParagraphAfter(); 

                oPrg.Range.Text = "«Российский университет транспорта» (РУТ (МИИТ))";
                oPrg.Range.InsertParagraphAfter(); 

                oPrg.Range.Text = "Институт транспортной техники и систем управления";
                oPrg.Range.Font.Size = 16f;
                oPrg.Format.SpaceAfter = 30f; 
                oPrg.Range.InsertParagraphAfter(); 

                oPrg.Range.Text = "Кафедра Управление и защита информации";
                oPrg.Range.Font.Bold = 1;

                W.InlineShape line = oDoc.Paragraphs.Last.Range.InlineShapes.AddHorizontalLineStandard();
                line.Height = 1;
                line.Fill.Solid();
                line.HorizontalLineFormat.NoShade = true;

                line.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                line.HorizontalLineFormat.PercentWidth = 90;
                line.HorizontalLineFormat.Alignment = WdHorizontalLineAlignment.wdHorizontalLineAlignCenter;
                oPrg.Range.InsertParagraphAfter();

                oPrg.Format.SpaceAfter = 30f;
                oPrg.Range.InsertParagraphAfter();

                oPrg.Range.Text = "Отчёт";
                oPrg.Range.Font.Bold = 1;
                oPrg.Range.Font.AllCaps = 1;
                oPrg.Range.Font.Size = 26f;
                oPrg.Format.SpaceAfter = 0;
                oPrg.Range.InsertParagraphAfter();

                oPrg.Range.Text = "по учебной практике";
                oPrg.Range.Font.Bold = 0;
                oPrg.Range.Font.AllCaps = 0;
                oPrg.Range.Font.Size = 14f;
                oPrg.Format.SpaceAfter = 80f;
                oPrg.Range.InsertParagraphAfter();

                oPrg.Format.SpaceAfter = 0;
                oPrg.Range.Text = "Выполнил: ст. гр. ТУУ-" + kurs + "11";
                oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphRight;
                oPrg.Range.Font.Size = 12f;
                oPrg.Range.InsertParagraphAfter();

                oPrg.Range.Text = "Маликов А. Д.";
                oPrg.Range.InsertParagraphAfter();

                oPrg.Range.Text = "Вариант №12";
                oPrg.Range.InsertParagraphAfter();

                oPrg.Range.Text = "Проверил: доц. Сафронов А. И.";
                oPrg.Format.SpaceAfter = 200f;
                oPrg.Range.InsertParagraphAfter();

                oPrg.Range.Text = "Москва - " + year + " г.";
                oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }
    }
}
