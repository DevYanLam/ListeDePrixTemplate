using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ListeDePrixNovago.PDFTemplate;
using ListeDePrixNovago.Utility;
using Microsoft.Win32;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;

namespace ListeDePrixNovago
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private PriceListConfig config;
        private string pdfFileName;
       
        public MainWindow()
        {
            InitializeComponent();
            ShowConfig();
        }

        private bool showPDF(TableType type)
        {
            try
            {
                //Create a PDF document
                Document doc = new Document();
                

                //Create a section inside the document
                Section template = doc.AddSection();
                
                template.PageSetup.RightMargin = 30;
                template.PageSetup.LeftMargin = 30;
                template.PageSetup.FooterDistance = new Unit(0, UnitType.Point);
                template.PageSetup.DifferentFirstPageHeaderFooter = true;

                //Headers
                MigraDoc.DocumentObjectModel.Shapes.Image logo = template.Headers.FirstPage.AddImage(config.LogoPath);
                logo.ScaleWidth = 0.5;
                logo.ScaleHeight = 0.5;

                template.AddParagraph();
                template.AddParagraph();
                template.AddParagraph();

                template.Headers.FirstPage.AddParagraph();
                Paragraph titre = template.Headers.FirstPage.AddParagraph(TitleSet.Text);
                titre.Format.Font.Bold = true;
                titre.Format.Alignment = ParagraphAlignment.Center;
                titre.Format.Font.Size = new Unit(16,UnitType.Point);

                template.AddParagraph();
                template.AddParagraph();
                template.AddParagraph();

                //Excel Table
                CreateTable(ExcelFilePath.Text, template, type, GetCheckedItems());
                
                //Footers
                DateTime input = DateTime.Today;
                int deltaMonday = DayOfWeek.Monday - input.DayOfWeek;
                DateTime monday = input.AddDays(deltaMonday);
                int deltaSunday = DayOfWeek.Sunday - input.DayOfWeek;
                DateTime sunday = input.AddDays(deltaSunday);
                string validText = "";
                if (config.IsValidityDateInFooter)
                    validText = "Valide du " + monday.ToShortDateString() + " au " + sunday.ToShortDateString() + "\n\n";
                string contactText = validText + config.Footer;
                template.AddParagraph();
                template.AddParagraph();
                template.AddParagraph(contactText);

                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false, PdfFontEmbedding.Always);
                pdfRenderer.Document = doc;
                pdfRenderer.RenderDocument();
                pdfFileName = Environment.CurrentDirectory + "\\" + TitleSet.Text + ".pdf";
                pdfRenderer.PdfDocument.Save(pdfFileName);
                var p = Process.Start(pdfFileName);
                p.WaitForExit();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Un problème est survenu durant la création du fichier PDF\n" + ex.Message);
                return false;
            }
            return true;
        }

        private void CreateTable(string excelFilePath, Section section, TableType type, List<Price> priceList)
        {
            ExcelReader r = new ExcelReader(excelFilePath);
            if (type == TableType.PriceList)
            {
                Table t = section.AddTable();
                t.KeepTogether = true;
                if (DropDownPriceList.SelectedItem.ToString() != null)
                    r.AddListPrice(t, DropDownPriceList.SelectedItem.ToString(), priceList);
                else
                    MessageBox.Show("Veuillez sélectionner une liste de prix","Aucune liste de prix sélectionné", MessageBoxButton.OK);
            }
            else if(type == TableType.CatalogList)
            {
                r.AddPriceCatalogTables(section, priceList);
            }
        }

        private void ShowLogo()
        {
            try
            {
                BitmapImage b = new BitmapImage();
                b.BeginInit();
                b.UriSource = new Uri(this.LogoPath.Text);
                b.EndInit();
                this.LogoPreview.Source = b;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Impossible de présenter le logo");
            }
        }

       

        private void LogoButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog fileChooser = new OpenFileDialog();
                fileChooser.Filter = "PNG File (*.png)|*.png|JPG File (*.jpg)|*.jpg";
                fileChooser.Title = "Sélectionnez un logo";

                if (fileChooser.ShowDialog() == true)
                {
                    this.LogoPath.Text = fileChooser.FileName;
                    File.Copy(fileChooser.FileName, Environment.CurrentDirectory + "/" + fileChooser.SafeFileName, true);
                    ShowLogo();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void ExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog fileChooser = new OpenFileDialog();
                fileChooser.Filter = "Excel File (*.xls)|*.xls|Excel File (*.xlsx)|*.xlsx";
                fileChooser.Title = "Sélectionnez un logo";

                if (fileChooser.ShowDialog() == true)
                {
                    string newPath = Environment.CurrentDirectory + "/" + fileChooser.SafeFileName;
                    this.ExcelFilePath.Text = newPath;
                    File.Copy(fileChooser.FileName, newPath, true);

                    ExcelReader re = new ExcelReader(newPath);
                    DropDownPriceList.ItemsSource = re.GetListTypeList();
                    ListBoxPrices.ItemsSource = re.GetPriceColumns();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ApplySettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PriceListConfig config = new PriceListConfig();
                config.LogoPath = LogoPath.Text;
                config.Footer = FooterSet.Text;
                if (IsValidityFooter.IsChecked == null)
                    config.IsValidityDateInFooter = false;
                else
                {
                    config.IsValidityDateInFooter = (bool)IsValidityFooter.IsChecked;
                }
                config.SmtpServer = SmtpServerSet.Text;
                config.SmtpPort = Int32.Parse(SmtpServerPort.Text);
                config.SmtpUsername = SmtpUsernameSet.Text;
                config.SmtpPassword = SmtpPasswordSet.Password;

                if (MessageBox.Show("Voulez-vous enregistrer la configuration?", "Enregistrement des paramètres", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                    SaveXml.SaveData(config, Environment.CurrentDirectory + "/config.xml");

                ShowConfig();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Impossible de sauvegarder la configuration\n" + ex.Message);
            }
        }

        private void ShowConfig()
        {
            try
            {
                PriceListConfig config = SaveXml.GetData(Environment.CurrentDirectory + "/config.xml");
                this.config = config;
                LogoPath.Text = config.LogoPath;
                FooterSet.Text = config.Footer;
                IsValidityFooter.IsChecked = config.IsValidityDateInFooter;
                SmtpServerSet.Text = config.SmtpServer;
                SmtpUsernameSet.Text = config.SmtpUsername;
                SmtpPasswordSet.Password = config.SmtpPassword;
                SmtpServerPort.Text = config.SmtpPort.ToString();
                ShowLogo();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Impossible de lire le fichier de configuration. Il est peut-être vide.");
            }
            
        }

        private void SendEmail()
        {
            if (MessageBox.Show("Voulez-vous envoyer ce document à " + RecipientsEmail.Text + "?", "Confirmation d'envoie", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
            {
                try
                {
                    SendEmail sm = new SendEmail(SmtpServerSet.Text, Int32.Parse(SmtpServerPort.Text), SmtpUsernameSet.Text, SmtpPasswordSet.Password);
                    sm.SendPriceList(SmtpUsernameSet.Text, RecipientsEmail.Text.Split(';'), TitleSet.Text, pdfFileName);
                    MessageBox.Show("Le courriel a bien été envoyé");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private List<Price> GetCheckedItems()
        {
            List<Price> priceList = new List<Price>();
            foreach(var item in ListBoxPrices.Items)
            {
                var i = item as Price;
                if(i.IsChecked)
                {
                    priceList.Add(i);
                }
            }
            return priceList;
        }

        private void SendEmailButton_Click(object sender, RoutedEventArgs e)
        {
            if(showPDF(TableType.PriceList))
                SendEmail();
        }

        private void SendCatalog_Click(object sender, RoutedEventArgs e)
        {
            if(showPDF(TableType.CatalogList))
                SendEmail();
        }

        public void RemoveText(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            if(tb.Text.Equals("Séparez les adresses courriels par des points-virgules."))
            tb.Text = "";
        }

        public void AddText(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            if (string.IsNullOrWhiteSpace(tb.Text))
                tb.Text = "Séparez les adresses courriels par des points-virgules.";
        }

        private void Gabarit_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
