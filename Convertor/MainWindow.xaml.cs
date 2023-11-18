using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Printing;
using System.Reflection.Metadata;
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
using System.Xml.Linq;
using org.apache.pdfbox.pdfviewer;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using Xceed.Words.NET;
using Spire.Xls;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Document = Spire.Doc.Document;
using javax.xml.parsers;
using Aspose.Words;

namespace Convertor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog() { Filter = "PDF files |*.pdf" };

            Nullable<bool> result = openFileDlg.ShowDialog();

            if (result == true)
            {
                txt1.Text = openFileDlg.FileName;
            }
        }

        private void btn2_Click(object sender, RoutedEventArgs e)
        {
            PDDocument doc = PDDocument.load(txt1.Text);
            PDFTextStripper stripper = new PDFTextStripper();

            var word = (stripper.getText(doc));
            var docName = System.IO.Path.GetFileNameWithoutExtension(txt1.Text) + ".docx";
            var worddoc = DocX.Create(docName);
            worddoc.InsertParagraph(word);
            worddoc.Save();

            var p = new Process();
            p.StartInfo = new ProcessStartInfo(docName.ToString())
            {
                UseShellExecute = true
            };
            p.Start();

            txt1.Text = "";
        }

        private void btn3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog() { Filter = "Excel files  (*.xlsx,*.xls)|*.xlsx;*.xls" };

            Nullable<bool> result = openFileDlg.ShowDialog();

            if (result == true)
            {
                txt2.Text = openFileDlg.FileName;
            }
        }

        private void btn4_Click(object sender, RoutedEventArgs e)
        {

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(txt2.Text);

            Worksheet sheet = workbook.Worksheets[0];

            Document doc = new Document();
            Spire.Doc.Section section = doc.AddSection();
            section.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Landscape;

            Spire.Doc.Table table = section.AddTable(true);
            table.ResetCells(sheet.LastRow, sheet.LastColumn);

            MergeCells(sheet, table);

            for (int r = 1; r <= sheet.LastRow; r++)
            {
                table.Rows[r - 1].Height = (float)sheet.Rows[r - 1].RowHeight;

                for (int c = 1; c <= sheet.LastColumn; c++)
                {

                    CellRange xCell = sheet.Range[r, c];

                    Spire.Doc.TableCell wCell = table.Rows[r - 1].Cells[c - 1];

                    Spire.Doc.Fields.TextRange textRange = wCell.AddParagraph().AppendText(xCell.NumberText);
                    CopyStyle(textRange, xCell, wCell);
                }
            }

            doc.SaveToFile(System.IO.Path.GetFileNameWithoutExtension(txt2.Text) + ".docx", Spire.Doc.FileFormat.Docx);

            var docName = System.IO.Path.GetFileNameWithoutExtension(txt2.Text) + ".docx";

            var p = new Process();
            p.StartInfo = new ProcessStartInfo(docName.ToString())
            {
                UseShellExecute = true
            };
            p.Start();

            txt2.Text = "";
        }           

        private void btn5_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog() { Filter = "JPG files |*.jpg" };

            Nullable<bool> result = openFileDlg.ShowDialog();

            if (result == true)
            {
                txt3.Text = openFileDlg.FileName;
            }
        }

        private void btn6_Click(object sender, RoutedEventArgs e)
        {
            var doc = new Aspose.Words.Document();
            var builder = new Aspose.Words.DocumentBuilder(doc);

            builder.InsertImage(txt3.Text);
            doc.Save((txt3.Text) + ".docx");

            var p = new Process();
            p.StartInfo = new ProcessStartInfo((txt3.Text) + ".docx".ToString())
            {
                UseShellExecute = true
            };
            p.Start();

            txt3.Text = "";
        }

        private static void MergeCells(Worksheet sheet, Spire.Doc.Table table)
        {
            if (sheet.HasMergedCells)
            {

                CellRange[] ranges = sheet.MergedCells;

                for (int i = 0; i < ranges.Length; i++)
                {

                    int startRow = ranges[i].Row;
                    int startColumn = ranges[i].Column;
                    int rowCount = ranges[i].RowCount;

                    int columnCount = ranges[i].ColumnCount;
                    if (rowCount > 1 && columnCount > 1)
                    {
                        for (int j = startRow; j <= startRow + rowCount; j++)
                        {
                            table.ApplyHorizontalMerge(j - 1, startColumn - 1, startColumn - 1 + columnCount - 1);
                        }

                        table.ApplyVerticalMerge(startColumn - 1, startRow - 1, startRow - 1 + rowCount - 1);
                    }

                    if (rowCount > 1 && columnCount == 1)
                    {
                        table.ApplyVerticalMerge(startColumn - 1, startRow - 1, startRow - 1 + rowCount - 1);
                    }

                    if (columnCount > 1 && rowCount == 1)
                    {
                        table.ApplyHorizontalMerge(startRow - 1, startColumn - 1, startColumn - 1 + columnCount - 1);
                    }
                }
            }
        }

        private static void CopyStyle(Spire.Doc.Fields.TextRange wTextRange, CellRange xCell, Spire.Doc.TableCell wCell)
        {

            wTextRange.CharacterFormat.TextColor = xCell.Style.Font.Color;
            wTextRange.CharacterFormat.FontSize = (float)xCell.Style.Font.Size;
            wTextRange.CharacterFormat.FontName = xCell.Style.Font.FontName;
            wTextRange.CharacterFormat.Bold = xCell.Style.Font.IsBold;
            wTextRange.CharacterFormat.Italic = xCell.Style.Font.IsItalic;
            wCell.CellFormat.BackColor = xCell.Style.Color;

            switch (xCell.HorizontalAlignment)
            {
                case HorizontalAlignType.Left:
                    wTextRange.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                    break;

                case HorizontalAlignType.Center:
                    wTextRange.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                    break;

                case HorizontalAlignType.Right:
                    wTextRange.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

                    break;

            }

            //Copy vertical alignment

            switch (xCell.VerticalAlignment)

            {
                case VerticalAlignType.Bottom:
                    wCell.CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Bottom;

                    break;

                case VerticalAlignType.Center:
                    wCell.CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    break;

                case VerticalAlignType.Top:
                    wCell.CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Top;

                    break;
            }
        }
    }
}
