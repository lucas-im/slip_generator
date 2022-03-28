using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.IO.Enumeration;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.IO;
using Spire.Xls;
using static PdfSharp.Drawing.XStringFormat;

namespace SlipGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly Settings _settings = Settings.Default;

        //Directory of slip template pdf files.
        private readonly string _pdfDir = "C:\\Users\\lucas\\Desktop\\ExcellTool";

        private class ExlData
        {
            private string _insurance;
            private string _recipient;
            private string _streetPoBox;
            private string _cityStateZip;
            private List<string> _patients;

            public ExlData(string insurance, string recipient, string streetPoBox, string cityStateZip,
                List<string> patients)
            {
                _insurance = insurance;
                _recipient = recipient;
                _streetPoBox = streetPoBox;
                _cityStateZip = cityStateZip;
                _patients = patients;
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            TxtOpenExp.Text = _settings.LastExpDir;
            TxtAdr.Text = _settings.SelAdr;
        }

        private void BtnOpenExl_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                var fileName = openFileDialog.FileName;
                if (!(fileName.EndsWith("xlsx") || fileName.EndsWith("xlsm") || fileName.EndsWith("xlsb") ||
                      fileName.EndsWith("xltx")))
                {
                    TxtOpenExl.Text = openFileDialog.FileName;
                    TxtOpenExl.CaretIndex = TxtOpenExp.GetLineLength(0);
                    ResultLabel.Content = "Invalid file type!";
                    ResultLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                ResultLabel.Content = "";
                TxtOpenExl.Text = openFileDialog.FileName;
                TxtOpenExl.CaretIndex = TxtOpenExp.GetLineLength(0);
            }
        }

        private void BtnOpenExp_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                InitialDirectory = TxtOpenExp.Text,
                Filter = "Directory|*.this.directory",
                FileName = "select"
            };

            if (dialog.ShowDialog() == true)
            {
                string path = dialog.FileName;
                TxtOpenExp.Text = path;
                _settings.LastExpDir = path;
                _settings.Save();
            }
        }

        private void ValColToRead(object sender, RoutedEventArgs e)
        {
            var subStr = TxtColToRead.Text.Replace(" ", "").Split(",");
            foreach (var str in subStr)
            {
                if (Regex.Matches(str, "[^A-Za-z]").Count > 0 || str.Length > 1)
                {
                    ResultLabel.Content = "Invalid column tag!";
                    ResultLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                ResultLabel.Content = "";
            }
        }

        private void PrevAdr_Click(object sender, RoutedEventArgs e)
        {
            PrevAdrWindow prevAdrWindow = new PrevAdrWindow();
            prevAdrWindow.Show();
            prevAdrWindow!.Closing += PrevAdrWindowClosed;
        }

        private void PrevAdrWindowClosed(object sender, CancelEventArgs e)
        {
            TxtAdr.Text = _settings.SelAdr;
        }

        private void BtnGenSlip_Click(object sender, RoutedEventArgs e)
        {
            var arr = _settings.Adr.Split(',').ToList();
            if (arr.Count == 0)
            {
                _settings.Adr = TxtAdr.Text;
                _settings.Save();
                return;
            }

            arr.Add(TxtAdr.Text);
            var value = string.Join(",", arr?.Select(i => i.ToString()).ToArray());
            _settings.Adr = value;
            _settings.Save();

            var workbook = new Workbook();
            workbook.LoadFromFile(TxtOpenExl.Text);
            var sheet = workbook.Worksheets[0];
            if (sheet == null)
            {
                ResultLabel.Content = "Invalid excel file.";
                ResultLabel.Foreground = new SolidColorBrush(Colors.Red);
                return;
            }

            ResultLabel.Content = "";
            for (var i = 1; i < sheet.Rows.Length; i++)
            {
           
                try
                {
                    //SlipType is A or Both.
                    if (SlipTypeCb.SelectedIndex == 0 || SlipTypeCb.SelectedIndex == 2)
                    {
                        var pdfDoc = PdfReader.Open(_pdfDir + "\\SlipTypeA.pdf");

                    }

                    //SlipType is B or Both.
                    if (SlipTypeCb.SelectedIndex == 1 || SlipTypeCb.SelectedIndex == 2)
                    {
                        var pdfDoc = PdfReader.Open(_pdfDir + "\\SlipTypeB.pdf");
                        var providerName = pdfDoc.AcroForm.Fields["PROVIDER NAME"];
                        var streetPoBox = pdfDoc.AcroForm.Fields["STREET / PO BOX"];
                        var cityStateZip = pdfDoc.AcroForm.Fields["CITY / STATE / ZIP"];
                        var recipientName = pdfDoc.AcroForm.Fields["RECIPIENT NAME"];
                        var insuranceName = pdfDoc.AcroForm.Fields["INSURANCE NAME"];
                        for(var j = 0; i < sheet.Rows[i].Columns.Length; j++)
                        {
                            var row = sheet.Rows[i].Columns[j].Text;
                            
                        }
                        // PdfString caseNamePdfStr = new PdfString("12345");
                        // currentField.Value = caseNamePdfStr;
                        pdfDoc.Save(TxtOpenExp.Text.Replace("select.this.directory.this.directory", "") + i + ".pdf");
                    }
                }
                catch (Exception exception)
                {
                    ResultLabel.Content = "Failed to generate.";
                    ResultLabel.Foreground = new SolidColorBrush(Colors.Red);
                    MessageBox.Show("Failed to generate slips. " + exception);
                    return;
                }
            }
        }
    }
}