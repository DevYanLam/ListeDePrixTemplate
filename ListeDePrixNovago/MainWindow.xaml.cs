using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using ListeDePrixNovago.PDFTemplate;
using ListeDePrixNovago.Utility;
using ListeDePrixNovago.Utility.TeamsAuthHelper;
using Microsoft.Graph;
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
        private GraphServiceClient graphClient = null;
        private List<NovagoSite> teamsName;
        private string logoPath;

        public MainWindow()
        {
            try
            {
                InitializeComponent();
                ShowConfig();
            }
            catch (Exception)
            {
                MessageBox.Show("Un problème est survenu");
            }
        }

        private bool showPDF()
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
                titre.Format.Font.Size = new Unit(16, UnitType.Point);

                template.AddParagraph();
                template.AddParagraph();
                template.AddParagraph();

                //Excel Table
                CreateTable(ExcelFilePath.Text, template, GetCheckedItems());

                //Footers
                DateTime input = DateTime.Today;
                int deltaMonday = System.DayOfWeek.Monday - input.DayOfWeek;
                DateTime monday = input.AddDays(deltaMonday);
                int deltaSunday = System.DayOfWeek.Sunday - input.DayOfWeek;
                DateTime sunday = input.AddDays(deltaSunday);
                string validText = "";
                if (config.IsValidityDateInFooter)
                    validText = "Valide du " + monday.ToShortDateString() + " au " + sunday.ToShortDateString() + "\n\n";
                string contactText = validText + config.Footer;
                template.AddParagraph();
                template.AddParagraph();
                template.AddParagraph(contactText);

                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false);
                pdfRenderer.Document = doc;
                pdfRenderer.RenderDocument();
                pdfFileName = Environment.CurrentDirectory + "\\" + TitleSet.Text + ".pdf";
                pdfRenderer.PdfDocument.Save(pdfFileName);
                var p = Process.Start(pdfFileName);
                p.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Un problème est survenu durant la création du fichier PDF\n" + ex.Message);
                return false;
            }
            return true;
        }

        private void CreateTable(string excelFilePath, Section section, List<Price> priceList)
        {
            ExcelReader r = new ExcelReader(excelFilePath);
            DataType type = r.TemplateType;
            if (type == DataType.PriceList)
            {
                Table t = section.AddTable();
                t.KeepTogether = true;
                if (DropDownPriceList.SelectedItem.ToString() != null)
                    r.AddListPrice(section, priceList, DropDownPriceList.SelectedItem.ToString());
                else
                    MessageBox.Show("Veuillez sélectionner une liste de prix", "Aucune liste de prix sélectionné", MessageBoxButton.OK);
            }
            else if (type == DataType.CatalogList)
            {
                r.AddPriceCatalogTables(section, priceList);
            }
        }

        private void ShowLogo(string path)
        {
            try
            {
                BitmapImage b = new BitmapImage();
                b.BeginInit();
                b.UriSource = new Uri(path);
                b.EndInit();
                this.LogoPreview.Source = b;
            }
            catch (Exception ex)
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
                    System.IO.File.Copy(fileChooser.FileName, Environment.CurrentDirectory + "/" + fileChooser.SafeFileName, true);
                    logoPath = Environment.CurrentDirectory + "/" + fileChooser.SafeFileName;
                    ShowLogo(logoPath);
                    LogoPath.Text = logoPath.Split('/')[logoPath.Split('/').Length - 1];
                }
            }
            catch (Exception ex)
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
                fileChooser.Title = "Open";

                if (fileChooser.ShowDialog() == true)
                {
                    this.ExcelFilePath.Text = fileChooser.FileName;

                    ExcelReader re = new ExcelReader(fileChooser.FileName);
                    var listOfPriceList = re.GetListTypeList();
                    if (((List<string>)listOfPriceList).Count > 0)
                    {
                        DropDownPriceList.Visibility = Visibility.Visible;
                        ListeDePrixLabel.Visibility = Visibility.Visible;
                        DropDownPriceList.ItemsSource = listOfPriceList;
                    }
                    else
                    {
                        DropDownPriceList.Visibility = Visibility.Hidden;
                        ListeDePrixLabel.Visibility = Visibility.Hidden;
                    }

                    ListBoxPrices.ItemsSource = re.GetPriceColumns();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ApplySettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PriceListConfig config = new PriceListConfig();
                config.LogoPath = logoPath;
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
                if (DropDownChannel.SelectedValue != null && DropDownTeams.SelectedValue != null)
                {
                    config.TeamsGroupId = DropDownTeams.SelectedValue as string;
                    config.DriveItemId = DropDownChannel.SelectedValue as string;
                    config.TeamsGroupName = ((NovagoSite)DropDownTeams.SelectedItem).Name;
                    config.DriveItemName = ((NovagoSite)DropDownChannel.SelectedItem).Name;
                    EquipeLabel.Visibility = Visibility.Hidden;
                    DropDownTeams.Visibility = Visibility.Hidden;
                    CanalLabel.Visibility = Visibility.Hidden;
                    DropDownChannel.Visibility = Visibility.Hidden;
                }

                if (MessageBox.Show("Voulez-vous enregistrer la configuration?", "Enregistrement des paramètres", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                    SaveXml.SaveData(config, Environment.CurrentDirectory + "/config.xml");

                ShowConfig();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible de sauvegarder la configuration\n" + ex.Message);
            }
        }

        private void ShowConfig()
        {
            TeamsLabel.Visibility = Visibility.Visible;
            ChannelLabel.Visibility = Visibility.Visible;
            try
            {
                PriceListConfig config = SaveXml.GetData(Environment.CurrentDirectory + "/config.xml");
                this.config = config;
                logoPath = config.LogoPath;
                LogoPath.Text = config.LogoPath.Split('/')[config.LogoPath.Split('/').Length - 1];
                FooterSet.Text = config.Footer;
                IsValidityFooter.IsChecked = config.IsValidityDateInFooter;
                SmtpServerSet.Text = config.SmtpServer;
                SmtpUsernameSet.Text = config.SmtpUsername;
                SmtpPasswordSet.Password = config.SmtpPassword;
                SmtpServerPort.Text = config.SmtpPort.ToString();
                if (config.TeamsGroupName != null && config.DriveItemName != null)
                {
                    TeamsLabel.Content = config.TeamsGroupName;
                    ChannelLabel.Content = config.DriveItemName;

                    IsSendToMsTeams.Visibility = Visibility.Visible;
                    IsSendToMsTeams.Content += "\nGROUPE : " + config.TeamsGroupName + "\nCANAL : " + config.DriveItemName;
                }

                ShowLogo(config.LogoPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible de lire le fichier de configuration. Il est peut-être vide." + ex.Message);
            }

        }

        private void SendEmail()
        {
            var dialog = new EmailPrompt();
            if ((bool)IsSendEmail.IsChecked && dialog.ShowDialog() == true)
            {
                try
                {
                    SendEmail sm = new SendEmail(SmtpServerSet.Text, Int32.Parse(SmtpServerPort.Text), SmtpUsernameSet.Text, SmtpPasswordSet.Password);
                    sm.SendPriceList(SmtpUsernameSet.Text, dialog.ResponseText.Split(';'), TitleSet.Text, pdfFileName);
                    MessageBox.Show("Le courriel a bien été envoyé");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void SendToTeams()
        {
            if ((bool)IsSendToMsTeams.IsChecked && MessageBox.Show("Voulez-vous vraiment importer le document vers Microsoft Teams", "Importation Microsoft Teams", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                try
                {
                    if (graphClient == null)
                    {
                        var graphAsync = AuthenticationHelper.GetAuthenticatedClientAsync();
                        graphAsync.Wait();
                        graphClient = graphAsync.Result;
                    }
                    string teamGroupId = config.TeamsGroupId;
                    string driveItemId = config.DriveItemId;
                    if (teamGroupId != null && driveItemId != null)
                    {
                        using (var stream = System.IO.File.Open(pdfFileName, FileMode.Open))
                        {

                            var folder = graphClient.Groups[teamGroupId].Drive.Items[driveItemId].ItemWithPath(TitleSet.Text + ".pdf").Content.Request().PutAsync<DriveItem>(stream);
                            folder.Wait();

                            MessageBox.Show("Le fichier a bien téléchargé", "Téléchargement réussi", MessageBoxButton.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Aucune équipe teams n'a été séléctionné dans les paramètres");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erreur", MessageBoxButton.OK);
                }
            }
        }

        private List<Price> GetCheckedItems()
        {
            List<Price> priceList = new List<Price>();
            foreach (var item in ListBoxPrices.Items)
            {
                var i = item as Price;
                if (i.IsChecked)
                {
                    priceList.Add(i);
                }
            }
            return priceList;
        }

        private List<NovagoSite> GetGroups(string siteId, bool isFirstExecution)
        {
            List<NovagoSite> tempSites = new List<NovagoSite>();
            var sites = graphClient.Groups.Request().GetAsync();
            sites.Wait();

            foreach (var site in sites.Result)
            {
                if (site.DeletedDateTime > DateTimeOffset.Now || site.DeletedDateTime is null)
                {
                    tempSites.Add(new NovagoSite()
                    {
                        Id = site.Id,
                        Name = site.DisplayName
                    });
                }
            }

            return tempSites;
        }

        private List<NovagoSite> GetChannels(ComboBox dropDown)
        {
            var drives = graphClient.Groups[(string)dropDown.SelectedValue].Drive.Request().GetAsync();
            drives.Wait();
            List<NovagoSite> driveList = new List<NovagoSite>();
            var items = graphClient.Drives[drives.Result.Id].Root.Children.Request().GetAsync();
            items.Wait();
            foreach (var i in items.Result)
            {
                if (i.Folder != null)
                {
                    driveList.Add(new NovagoSite()
                    {
                        Id = i.Id,
                        Name = i.Name
                    });
                }
            }


            return driveList;
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            if (showPDF())
            {
                SendEmail();
                SendToTeams();
            }
        }

        public void RemoveText(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            if (tb.Text.Equals("Séparez les adresses courriels par des points-virgules."))
                tb.Text = "";
        }

        public void AddText(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            if (string.IsNullOrWhiteSpace(tb.Text))
                tb.Text = "Séparez les adresses courriels par des points-virgules.";
        }

        private void GabaritCatalog_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(Environment.CurrentDirectory + "/Template/listedeprix_representant_animal.xls");
        }

        private void GabaritListe_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(Environment.CurrentDirectory + "/Template/listedeprix_producteur.xls");
        }

        private void LogToTeams_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var graphAsync = AuthenticationHelper.GetAuthenticatedClientAsync();
                graphAsync.Wait();
                graphClient = graphAsync.Result;
                teamsName = GetGroups(null, true);
                teamsName.Sort((x, y) => String.Compare(x.Name, y.Name));
                if (DropDownChannel.ItemsSource != null)
                    ((List<NovagoSite>)DropDownChannel.ItemsSource).Clear();
                if (DropDownTeams.ItemsSource != null)
                {
                    ((List<NovagoSite>)DropDownTeams.ItemsSource).Clear();
                    ((List<NovagoSite>)DropDownTeams.ItemsSource).AddRange(teamsName);
                }
                else
                {
                    DropDownTeams.ItemsSource = teamsName;
                }

                TeamsLabel.Visibility = Visibility.Hidden;
                ChannelLabel.Visibility = Visibility.Hidden;

                EquipeLabel.Visibility = Visibility.Visible;
                DropDownTeams.Visibility = Visibility.Visible;
                CanalLabel.Visibility = Visibility.Visible;
                DropDownChannel.Visibility = Visibility.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void TeamSelected(object sender, RoutedEventArgs e)
        {
            try
            {
                ComboBox dropDown = sender as ComboBox;
                Console.WriteLine(dropDown.SelectedValue);
                DropDownChannel.ItemsSource = GetChannels(dropDown);
            }
            catch (AggregateException ex)
            {
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButton.OK);
            }
        }

    }
}
