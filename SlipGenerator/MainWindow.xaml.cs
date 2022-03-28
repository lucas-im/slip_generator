using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.IO.Enumeration;
using System.Linq;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
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
            var subStr = TxtColToRead.Text.ToLower().Replace(" ", "").Split(",");
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
            if (TxtAdr.Text.Length > 0)
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
            }

            var workbook = new Workbook();
            workbook.LoadFromFile(TxtOpenExl.Text);
            var sheets = workbook.Worksheets;
            if (sheets == null)
            {
                ResultLabel.Content = "Invalid excel file.";
                ResultLabel.Foreground = new SolidColorBrush(Colors.Red);
                return;
            }

            ResultLabel.Content = "";

            var colToRead = TxtColToRead.Text.Split(',').ToList();
            var isExclude = ColToReadCB.SelectedIndex == 0;

            for (var k = 0; k < sheets.Count; k++)
            {
                var sheet = sheets[k];
                for (var i = 1; i < sheet.Rows.Length; i++)
                {
                    //SlipType is A or Both.
                    if (SlipTypeCb.SelectedIndex == 0 || SlipTypeCb.SelectedIndex == 2)
                    {
                        var pdfDoc = PdfReader.Open(_pdfDir + "\\SlipTypeA.pdf");
                        var providerName = pdfDoc.AcroForm.Fields["PROVIDER NAME"];
                        var streetPoBox = pdfDoc.AcroForm.Fields["Street and Apt No or PO Box No"];
                        var cityStateZip = pdfDoc.AcroForm.Fields["City State ZIP4"];
                        var recipientName = pdfDoc.AcroForm.Fields["RECIPIENT"];
                        var insuranceName = pdfDoc.AcroForm.Fields["INSURANCE NAME"];
                        var patient1 = pdfDoc.AcroForm.Fields["PATIENT_1"];
                        var patient2 = pdfDoc.AcroForm.Fields["PATIENT_2"];
                        var patient3 = pdfDoc.AcroForm.Fields["PATIENT_3"];
                        var patient4 = pdfDoc.AcroForm.Fields["PATIENT_4"];
                        var patient5 = pdfDoc.AcroForm.Fields["PATIENT_5"];
                        var patient6 = pdfDoc.AcroForm.Fields["PATIENT_6"];
                        var patient7 = pdfDoc.AcroForm.Fields["PATIENT_7"];
                        var patient8 = pdfDoc.AcroForm.Fields["PATIENT_8"];
                        var patient9 = pdfDoc.AcroForm.Fields["PATIENT_9"];
                        var patient10 = pdfDoc.AcroForm.Fields["PATIENT_10"];
                        for (var j = 0; j < sheet.Rows[i].Columns.Length; j++)
                        {
                            var adr = sheet.Rows[i].Columns[j].RangeAddressLocal[0].ToString().ToLower();
                            var row = sheet.Rows[i].Columns[j].Text;
                            if (!colToRead.Contains(adr))
                            {
                                switch (sheet.Rows[0].Columns[j].Text)
                                {
                                    case "INSURANCE":
                                        insuranceName.Value = new PdfString(row);
                                        break;
                                    case "RECIPIENT":
                                        recipientName.Value = new PdfString(row);
                                        break;
                                    case "STREET / P O Box":
                                        streetPoBox.Value = new PdfString(row);
                                        break;
                                    case "CITY_STATE_ZIP":
                                        cityStateZip.Value = new PdfString(row);
                                        break;
                                    case "Patient-1":
                                        patient1.Value = new PdfString(row);
                                        break;
                                    case "Patient-2":
                                        patient2.Value = new PdfString(row);
                                        break;
                                    case "Patient-3":
                                        patient3.Value = new PdfString(row);
                                        break;
                                    case "Patient-4":
                                        patient4.Value = new PdfString(row);
                                        break;
                                    case "Patient-5":
                                        patient5.Value = new PdfString(row);
                                        break;
                                    case "Patient-6":
                                        patient6.Value = new PdfString(row);
                                        break;
                                    case "Patient-7":
                                        patient7.Value = new PdfString(row);
                                        break;
                                    case "Patient-8":
                                        patient8.Value = new PdfString(row);
                                        break;
                                    case "Patient-9":
                                        patient9.Value = new PdfString(row);
                                        break;
                                    case "Patient-10":
                                        patient10.Value = new PdfString(row);
                                        break;
                                }
                            }
                        }

                        pdfDoc.Save(TxtOpenExp.Text.Replace("select.this.directory.this.directory", "") + "SlipTypeA" +
                                    k + i + ".pdf");
                    }

                    //SlipType is B or Both.
                    if (SlipTypeCb.SelectedIndex == 1 || SlipTypeCb.SelectedIndex == 2)
                    {
                        var pdfDoc = PdfReader.Open(_pdfDir + "\\SlipTypeB.pdf");
                        var providerName = pdfDoc.AcroForm.Fields[0];
                        var streetPoBox = pdfDoc.AcroForm.Fields[1];
                        var cityStateZip = pdfDoc.AcroForm.Fields[2];
                        var recipientName = pdfDoc.AcroForm.Fields[3];
                        var insuranceName = pdfDoc.AcroForm.Fields[4];
                        for (var j = 0; j < sheet.Rows[i].Columns.Length; j++)
                        {
                            var adr = sheet.Rows[i]?.Columns[j]?.RangeAddressLocal[0].ToString().ToLower();
                            var row = sheet.Rows[i].Columns[j].Text;
                            if (!colToRead.Contains(adr))
                            {
                                switch (sheet.Rows[0].Columns[j].Text)
                                {
                                    case "INSURANCE":
                                        insuranceName.Value = new PdfString(row);
                                        break;
                                    case "RECIPIENT":
                                        recipientName.Value = new PdfString(row);
                                        break;
                                    case "STREET / P O Box":
                                        streetPoBox.Value = new PdfString(row);
                                        break;
                                    case "CITY_STATE_ZIP":
                                        cityStateZip.Value = new PdfString(row);
                                        break;
                                }
                            }
                        }

                        try
                        {
                            providerName.Value = new PdfString("Jhonny  Wellness Center");
                            using (XGraphics gfx = XGraphics.FromPdfPage(pdfDoc.Pages[0]))
                            {
                                gfx.DrawRectangle(XBrushes.Black, new XRect(0, 0, 1, 1));
                            }
                            pdfDoc.Save(TxtOpenExp.Text.Replace("select.this.directory.this.directory", "") +
                                        "SlipTypeB" +
                                        k + i + ".pdf");
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
    }
}