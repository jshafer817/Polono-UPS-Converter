using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ImageMagick;
using PdfiumViewer;
using System.Drawing.Printing;

namespace Polono_UPS
{
    public partial class Form1 : Form
    {
        bool FileSelected;

        public Form1(string[] args)
        {
            InitializeComponent();
            FileSelected = false;
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);

            this.Show();
            File.Delete(Application.LocalUserAppDataPath + "\\" + "temp.pdf");

            if (args.Length == 1)
            {
                ConvertImage(args[0]);
                PrintPDF();
            }
            else 
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    if (Properties.Settings.Default.DirectoryLastUsed == "")
                    {
                        openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    }
                    else
                    {
                        openFileDialog.InitialDirectory = Properties.Settings.Default.DirectoryLastUsed;
                    }

                    openFileDialog.Filter = "Gif Files|*.gif";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        Properties.Settings.Default.DirectoryLastUsed = openFileDialog.FileName;
                        Properties.Settings.Default.Save();

                        FileSelected = true;

                        //Get the path of specified file
                        ConvertImage(openFileDialog.FileName);
                        PrintPDF();
                    }
                }
               
            }            
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                if (Properties.Settings.Default.DirectoryLastUsed == "")
                {
                    openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
                else
                {
                    openFileDialog.InitialDirectory = Properties.Settings.Default.DirectoryLastUsed;
                }
                openFileDialog.Filter = "Gif Files|*.gif";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.DirectoryLastUsed = openFileDialog.FileName;
                    Properties.Settings.Default.Save();

                    FileSelected = true;

                    //Get the path of specified file
                    ConvertImage(openFileDialog.FileName);
                }
            }
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            PrintPDF();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            File.Delete(Application.LocalUserAppDataPath + "\\" + "temp.pdf");
            Dispose();
        }

        public void ConvertImage(string filename)
        {
            // Read from file
            using (var image = new MagickImage(filename))
            {
                var chop = new MagickGeometry(0, 200);
                image.Rotate(270);
                image.Chop(chop);
                image.Rotate(180);
                var scale = new MagickGeometry(760, 1140);
                image.Scale(scale);
                var extent = new MagickGeometry(800, 1200);
                var backgroundColor = new MagickColor("white");
                image.Extent(extent, Gravity.Center, backgroundColor);
                image.Density = new Density(266.5, DensityUnit.PixelsPerInch); //3.90
                image.Write(Application.LocalUserAppDataPath + "\\" + "temp.pdf");

                Bitmap bmp = null;
                using (MemoryStream ms = new MemoryStream())
                {
                    var fmt = MagickFormat.Png24;
                    image.Write(ms, fmt);
                    ms.Position = 0;
                    bmp = (Bitmap)Bitmap.FromStream(ms);
                }
                
                pictureBox1.Image = bmp;                
            }
        }

        public bool PrintPDF()
        {
            bool result = false;

            if (FileSelected == false)
                return false;
            PrintDialog printDlg = new PrintDialog();
            if (printDlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // Create the printer settings for our printer
                    var printerSettings = new PrinterSettings
                    {
                        PrinterName = printDlg.PrinterSettings.PrinterName//,
                                                                         //Copies = (short)copies,
                    };

                    // Create our page settings for the paper size selected
                    var pageSettings = new PageSettings(printerSettings)
                    {
                        Margins = new Margins(0, 0, 0, 0),
                    };

                    //foreach (PaperSize paperSize in printerSettings.PaperSizes)
                    //{
                    //  if (paperSize.PaperName == paperName)
                    //{
                    //  pageSettings.PaperSize = paperSize;
                    // break;
                    //}
                    //}

                    // Now print the PDF document
                    using (var document = PdfDocument.Load(Application.LocalUserAppDataPath + "\\" + "temp.pdf"))
                    {
                        using (var printDocument = document.CreatePrintDocument(PdfiumViewer.PdfPrintMode.CutMargin))
                        {
                            printDocument.PrinterSettings = printerSettings;
                            printDocument.DefaultPageSettings = pageSettings;
                            printDocument.PrintController = new StandardPrintController();
                            printDocument.Print();
                        }
                    }
                    result = true;
                }
                catch
                {
                    result = false;
                }
            }
            return result;
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            FileSelected = true;
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                ConvertImage(file);
                PrintPDF();            
            }
        }
    }
}
