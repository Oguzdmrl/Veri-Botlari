using BotHizmetleri.Models;
using BotHizmetleri.Services;
using MaterialSkin;
using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace BotHizmetleri
{
    public partial class AnaSayfa : Form
    {
        private MaterialSkinManager materialSkinManager;
        ChromeOptions options = new();
        private string _lisansKey = "..";
        private bool _lisansDurum = false;
        private string _lisansBitisTarihi = "";

        public List<Sehir> Sehirler;
        public List<Ilce> Ilceler;

        public AnaSayfa()
        {
            InitializeComponent();
            loaderPictureLisans.Visible = false;


            CheckForIllegalCrossThreadCalls = false;
            listBoxTaskDurum.DrawMode = DrawMode.OwnerDrawFixed; // ListBox'ın özelleştirilebilir çizim modunu ayarla
            listBoxTaskDurum.DrawItem += new DrawItemEventHandler(listBoxTaskDurum_DrawItem); // Çizim olayını tanımla


            loaderPictureGoogleMaps.Visible = false;



            servisiDurdurBtn_Click.Enabled = false;
            MailAdresleriniTara_BtnClick.Enabled = false;
            excelKaydet_btnClick.Enabled = false;
            Sil_BtnClick.Enabled = false;
            Il_ComboBox.Visible = false;
            options.AddArgument("--headless");
            options.AddArgument("no-sandbox");
            options.AddArgument("--start-maximized");
            options.AddArgument("--window-position=-2400,-2400");
            options.AddArgument("--window-size=1920,1080"); // Pencereyi tam ekran boyutuna ayarla (1920x1080)
            options.AddArgument("--disable-dev-shm-usage");
            options.AddArgument("--disable-infobars"); // Bilgi çubuklarını gizle

            dataGridView1.ColumnCount = 5;
            dataGridView1.Columns[0].Name = "Firma";
            dataGridView1.Columns[1].Name = "Telefon";
            dataGridView1.Columns[2].Name = "Web Site";
            dataGridView1.Columns[3].Name = "E-Mail";
            dataGridView1.Columns[4].Name = "Adress";
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.RowHeadersVisible = false; // Soldaki alanı gizler
            dataGridView1.ReadOnly = true; // sadece okunabilir olması yani veri düzenleme kapalı
            dataGridView1.AllowUserToDeleteRows = false; // satırların silinmesi engelleniyor
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView1.ScrollBars = ScrollBars.Both;

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
            dataGridView1.Columns[dataGridView1.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            // Başlık hücrelerinin stilini ayarla
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Arial", 12); // Başlık fontu
                                                                                                     //dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Otomatik boyutlandırmayı devre dışı bırak
                                                                                                     //dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Otomatik boyutlandırmayı devre dışı bırak
                                                                                                     //dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Otomatik boyutlandırmayı devre dışı bırak
                                                                                                     //dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Otomatik boyutlandırmayı devre dışı bırak
                                                                                                     //dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; // Otomatik boyutlandırmayı devre dışı bırak




            dataGrid_Ilce.Enabled = false;
            dataGrid_Ilce.Visible = false;
            Il_ComboBox.Enabled = false;
            LoadData();
            Il_ComboBox.DataSource = Sehirler;
            Il_ComboBox.DisplayMember = "sehir_adi";
            Il_ComboBox.ValueMember = "sehir_id";

            dataGrid_Ilce.AllowUserToAddRows = false;


            DataGridViewTextBoxColumn ilceAdiColumn = new DataGridViewTextBoxColumn
            {
                DataPropertyName = "ilce_adi", // Bu, Ilce modelindeki property'nin adıdır
                HeaderText = "İlçe"
            };
            dataGrid_Ilce.Columns.Add(ilceAdiColumn);
            dataGrid_Ilce.Columns[dataGrid_Ilce.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;


            string selectedSehirId = Il_ComboBox.SelectedValue.ToString();
            var filteredIlceler = Ilceler.Where(i => i.sehir_id == selectedSehirId).ToList();

            dataGrid_Ilce.AutoGenerateColumns = false;
            dataGrid_Ilce.DataSource = filteredIlceler;





            #region MAİL

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // DataGridView Ayarları
            dataGridViewMail.Columns.Add("Firma", "Firma");
            dataGridViewMail.Columns.Add("Email", "Email");
            DataGridViewTextBoxColumn statusColumn = new();
            statusColumn.Name = "Durum";
            dataGridViewMail.Columns.Add(statusColumn);

            // Kolon boyutunu otomatik ayarlama
            dataGridViewMail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; // Otomatik ayarlama devre dışı
            dataGridViewMail.Columns[0].Width = 200; // Firma kolonu genişliği
            dataGridViewMail.Columns[1].Width = 200; // Email kolonu genişliği
            dataGridView1.Columns[2].Width = 100; // Durum kolonu genişliği
            // Yan kaydırmayı aktif et
            dataGridViewMail.ScrollBars = ScrollBars.Both;
            dataGridViewMail.AllowUserToAddRows = false;
            dataGridViewMail.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            #endregion



            tabPageGosterGizle();


        }
        private async void AnaSayfa_Load(object sender, EventArgs e)
        {

            var lisansManager = new LisansManager();

            _lisansKey = lisansManager.LisansDosyasiniOku().LisansKey;
            lisansTextBox.Text = _lisansKey;

            await lisansKontrolEt();





            dataGridViewWP.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize; // Başlıkları otomatik ayarla
            dataGridViewWP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // Sütun genişliklerini otomatik ayarla
            servisDurdurWp.Enabled = false;

            await webView21.EnsureCoreWebView2Async(null);
            webView21.Source = new Uri("https://web.whatsapp.com/");


            #region EMAİL BOTU

            await webViewTiny.EnsureCoreWebView2Async();

            hostingler_ComboBox.Items.Add("Hostinger");
            hostingler_ComboBox.Items.Add("Natro");
            hostingler_ComboBox.SelectedIndex = 0;
            hostingler_ComboBox.SelectedIndexChanged += hostingler_ComboBox_SelectedIndexChanged;
            PopulateTextBoxes(hostingler_ComboBox.SelectedItem.ToString());


            // HTML dosyasının projenin çalışma dizininde olduğundan emin olun
            string htmlFilePath = Path.Combine(Application.StartupPath, "tinyeditor.html");

            // TinyMCE HTML dosyasını WebView2 ile yükleyin
            webViewTiny.Source = new Uri(htmlFilePath);

            #endregion






        }

        #region GOOGLE MAPS VERİ BOTU


        #region Private Methods
        private void ExportToExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // SaveFileDialog ile dosya kaydetme konumu seç
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Dosyası (*.xlsx)|*.xlsx|Tüm Dosyalar (*.*)|*.*";
                saveFileDialog.Title = "Excel Dosyasını Kaydet";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    // Excel dosyasını oluştur
                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        // Yeni bir çalışma sayfası oluştur
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Veriler");

                        int rowCount = dataGridView1.Rows.Count;
                        int colCount = dataGridView1.Columns.Count;

                        // Sütun başlıklarını yaz
                        for (int i = 0; i < colCount; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                        }

                        // Verileri yaz
                        for (int i = 0; i < rowCount; i++)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                worksheet.Cells[i + 2, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                            }
                        }

                        // Dosyayı kaydet
                        FileInfo fileInfo = new FileInfo(filePath);
                        excelPackage.SaveAs(fileInfo);

                        MessageBox.Show("Veriler başarıyla Excel dosyasına aktarıldı.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }
        public void LoadData()
        {

            // JSON verisini deserialize ederek Sehir listesine dönüştürelim


            string sehirlerJson = File.ReadAllText("sehirler.json");
            Sehirler = JsonConvert.DeserializeObject<List<Sehir>>(sehirlerJson);

            string ilcelerJson = File.ReadAllText("ilceler.json");
            Ilceler = JsonConvert.DeserializeObject<List<Ilce>>(ilcelerJson);
        }

        #endregion

        string _SeciliIl;




        private string GetMotherboardSerialNumber()
        {
            try
            {
                // ManagementObjectSearcher ile WMI sorgusu yapıyoruz
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");

                foreach (ManagementObject obj in searcher.Get())
                {
                    // İlk anakartın seri numarasını alıyoruz
                    return obj["SerialNumber"]?.ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hata: {ex.Message}");
            }

            return null;
        }

        private void Il_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var sehir = Il_ComboBox.SelectedItem as Sehir; // YourDataType, ComboBox'a eklediğiniz nesne türü

            if (sehir != null) _SeciliIl = sehir.sehir_adi;

            string selectedSehirId = Il_ComboBox.SelectedValue.ToString();
            var filteredIlceler = Ilceler.Where(i => i.sehir_id == selectedSehirId).ToList();

            dataGrid_Ilce.AutoGenerateColumns = false;

            dataGrid_Ilce.DataSource = filteredIlceler;
        }

        private void genelTarama_CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            bool isChecked = genelTarama_CheckBox.Checked;
            dataGrid_Ilce.Visible = isChecked;
            dataGrid_Ilce.Enabled = isChecked;
            Il_ComboBox.Visible = isChecked;
            Il_ComboBox.Enabled = isChecked;
        }

        private int elapsedTime; // Geçen süre

        private CancellationTokenSource cancellationTokenSource;
        private void UpdateTaskStatus(string message, string status)// Task durumlarını göstermek için bir metot ekliyoruz:
        {
            Invoke(new Action(() =>
            {
                listBoxTaskDurum.Items.Add(new ListBoxItem { Message = message, Status = status });
                listBoxTaskDurum.Refresh(); // ListBox'ı yeniden çizmek için güncelle
            }));
        }

        private List<int> failedTasks = new List<int>(); // Hatalı task'ları toplamak için bir liste
        private List<Task> tasks = new List<Task>(); // Tüm görevleri izlemek için bir liste
        ChromeDriverService chromeDriverService;
        private async void Servisi_Baslat_BtnClick_Click(object sender, EventArgs e)
        {
            loaderPictureGoogleMaps.Visible = true;

            // Butonları devre dışı bırak
            MailAdresleriniTara_BtnClick.Enabled = false;
            excelKaydet_btnClick.Enabled = false;
            Sil_BtnClick.Enabled = false;
            servisiDurdurBtn_Click.Enabled = true;
            Servisi_Baslat_BtnClick.Enabled = false;

            // Timer başlat
            timer1.Interval = 1000;
            timer1.Tick += timer1_Tick;
            elapsedTime = 0;
            label3.Text = "00:00:00";
            timer1.Start();

            cancellationTokenSource = new CancellationTokenSource();
            CancellationToken token = cancellationTokenSource.Token;
            chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;

            if (genelTarama_CheckBox.Checked)
            {
                SemaphoreSlim semaphore = new SemaphoreSlim(2);
                int rowCount = dataGrid_Ilce.RowCount;

                for (int i = 0; i < rowCount; i++)
                {
                    string ilce = dataGrid_Ilce.Rows[i].Cells[0].Value.ToString();
                    var anahtarKelime = _SeciliIl + " " + ilce + " " + anahtarKelime_TextBox.Text;
                    anahtarKelime = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(anahtarKelime.ToLower());
                    int threadIndex = i;

                    var task = Task.Run(async () =>
                    {
                        await semaphore.WaitAsync();
                        try
                        {
                            UpdateTaskStatus($"{ilce} taraması başladı.", "running");

                            using (var driver = new ChromeDriver(chromeDriverService, options))
                            {
                                token.ThrowIfCancellationRequested();
                                await VeriCek(anahtarKelime, threadIndex, driver);
                                UpdateTaskStatus($"{ilce} taraması başarıyla tamamlandı.", "completed");
                            }
                        }
                        catch (Exception ex)
                        {
                            if (_iptalEdildiMi)
                            {
                                UpdateTaskStatus($"{ilce} Taraması iptal edildi.", "cancelled");
                                failedTasks.Add(threadIndex);
                            }
                            else
                            {
                                failedTasks.Add(threadIndex); // Hatalı task'ları kaydet
                                UpdateTaskStatus($"{ilce} taraması hata verdi.", "failed");
                            }


                            // Hata alınan satırı koyu kırmızı yap
                            Invoke(new Action(() =>
                            {
                                dataGrid_Ilce.Rows[threadIndex].DefaultCellStyle.BackColor = Color.DarkRed;
                            }));
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    });

                    tasks.Add(task);
                }

                try
                {
                    await Task.WhenAll(tasks); // Tüm görevler tamamlanana kadar bekle
                }
                catch (OperationCanceledException)
                {
                    UpdateTaskStatus("Tüm görevler iptal edildi.", "failed");
                }



                // Tüm tarama işlemi bittiğinde
                if (failedTasks.Count > 0)
                {
                    DialogResult result = MessageBox.Show("Bazı ilçelerden veri alınamadı. Tekrar denemek ister misiniz?", "Hata", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        //var task = Task.Run(async () => await TekrarDeneme());
                        var task = Task.Run(async () =>
                        {
                            await semaphore.WaitAsync();

                            token.ThrowIfCancellationRequested();
                            await TekrarDeneme();

                            semaphore.Release();

                        });

                        tasks.Add(task);
                    }
                }
                await Task.WhenAll(tasks); // Tüm görevler tamamlanana kadar bekle
                //var uniqueFirmalar = firmalarList
                //      .GroupBy(f => f.FirmaAdi) // Firma adına göre gruplama
                //      .Select(g => g.First()) // Her grup için yalnızca ilk elemanı seç
                //      .ToList(); // Sonucu listeye dönüştür

                //Invoke(new Action(() =>
                //{
                //    foreach (var firma in uniqueFirmalar)
                //    {
                //        dataGridView1.Rows.Add(firma.FirmaAdi, firma.Telefon, firma.Website, firma.EMail, firma.Adress);
                //    }
                //}));
            }
            else
            {
                var anahtarKelime = anahtarKelime_TextBox.Text;
                anahtarKelime = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(anahtarKelime.ToLower());

                List<Task> tasksTekil = new List<Task>();
                var task = Task.Run(async () =>
                {
                    UpdateTaskStatus($"{anahtarKelime} taraması başladı.", "running");

                    using (var driver = new ChromeDriver(chromeDriverService, options)) // WebDriver işlemleri burada yapılır
                    {
                        await VeriCek(anahtarKelime, 1, driver); // Zaman aşımı olmadan veri çekme işlemi
                    }
                });
                tasksTekil.Add(task);
                await Task.WhenAll(tasksTekil);
                UpdateTaskStatus($" {anahtarKelime} taraması başarıyla tamamlandı.", "completed");
                //var uniqueFirmalar = firmalarList
                //    .GroupBy(f => f.FirmaAdi) // Firma adına göre gruplama
                //    .Select(g => g.First()) // Her grup için yalnızca ilk elemanı seç
                //    .ToList(); // Sonucu listeye dönüştür

                //Invoke(new Action(() =>
                //{
                //    foreach (var firma in uniqueFirmalar)
                //    {
                //        dataGridView1.Rows.Add(firma.FirmaAdi, firma.Telefon, firma.Website, firma.EMail, firma.Adress);
                //    }
                //}));
            }
            var uniqueFirmalar = firmalarList
                   .GroupBy(f => f.FirmaAdi) // Firma adına göre gruplama
                   .Select(g => g.First()) // Her grup için yalnızca ilk elemanı seç
                   .ToList(); // Sonucu listeye dönüştür

            Invoke(new Action(() =>
            {
                foreach (var firma in uniqueFirmalar)
                {
                    dataGridView1.Rows.Add(firma.FirmaAdi, firma.Telefon, firma.Website, firma.EMail, firma.Adress);
                }
            }));
            toplamVeri_Label.Text = dataGridView1.RowCount.ToString();
            int count = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Hücre kontrolü yapıyoruz
                if (row.Cells[1].Value != null && !string.IsNullOrWhiteSpace(row.Cells[1].Value.ToString()))
                {
                    count++;
                }
            }
            totalTelefonLabel.Text = count.ToString();


            timer1.Stop();
            // Butonları tekrar aktif hale getir
            Servisi_Baslat_BtnClick.Enabled = true;
            servisiDurdurBtn_Click.Enabled = false;
            MailAdresleriniTara_BtnClick.Enabled = true;
            excelKaydet_btnClick.Enabled = true;
            Sil_BtnClick.Enabled = true;
            loaderPictureGoogleMaps.Visible = false;
        }

        private async Task TekrarDeneme()
        {

            chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;

            int maxRetries = 3; // Maksimum 3 deneme
            int timeoutIncrement = 40000; // Timeout her yeniden denemede artacak

            var failedTasksCopy = failedTasks.ToList();
            var tasksToRemove = new List<int>(); // Başarılı olanları buraya ekleyeceğiz

            foreach (int threadIndex in failedTasksCopy)
            {
                string ilce = dataGrid_Ilce.Rows[threadIndex].Cells[0].Value.ToString();
                var anahtarKelime = _SeciliIl + " " + ilce + " " + anahtarKelime_TextBox.Text;
                anahtarKelime = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(anahtarKelime.ToLower());
                int retries = 0;
                UpdateTaskStatus($"{ilce} taraması tekrar başladı.", "running");
                while (retries < maxRetries)
                {
                    try
                    {
                        using (var driver = new ChromeDriver(chromeDriverService, options))
                        {
                            // Sayfa yükleme süresini dinamik olarak artır
                            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30 + (retries * timeoutIncrement / 1000));

                            // Verileri çek
                            await VeriCek(anahtarKelime, threadIndex, driver);

                            // Görev durumunu güncelle
                            UpdateTaskStatus($"{ilce} taraması SON DENEME ile başarıyla tamamlandı.", "completed");

                            // Başarılı olursa kaldırılacaklar listesine ekle
                            tasksToRemove.Add(threadIndex);

                            break; // Başarılı olduysa döngüden çık
                        }
                    }
                    catch (Exception ex)
                    {
                        retries++;
                        UpdateTaskStatus($"{ilce} taraması SON DENEME başarısız: {ex.Message}", "failed");
                    }
                }

                if (retries == maxRetries)
                {
                    UpdateTaskStatus($"{ilce} taraması başarısız oldu.", "failed");
                }
            }

            // Tüm başarılı görevleri failedTasks listesinden kaldır
            foreach (var task in tasksToRemove)
            {
                failedTasks.Remove(task);
            }
        }

        private void GoogleMapsAcVeTextiArat(string anahtarKelime, ChromeDriver driver)
        {


            driver.Navigate().GoToUrl("https://www.google.com/maps");

            WebDriverWait wait = new(driver, TimeSpan.FromSeconds(10));
            var searchBox = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("searchboxinput")));

            searchBox.SendKeys(anahtarKelime);

            var searchButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("searchbox-searchbutton")));
            searchButton.Click();

            wait.Until(ExpectedConditions.ElementIsVisible(By.Id("searchboxinput"))); // Arama kutusunun tekrar görünmesini bekleyin

        }
        private async Task VeriCek(string anahtarKelime, int rowIndex, ChromeDriver driver)
        {
            if (genelTarama_CheckBox.Checked)
            {
                await UpdateGridRowColor(rowIndex, Color.DarkGoldenrod, Color.White);
            }
            GoogleMapsAcVeTextiArat(anahtarKelime, driver);
            Thread.Sleep(3000);
            var veriVarMi = driver.FindElements(By.CssSelector("[class^='Nv2PK']"));
            if (veriVarMi.Count > 0)
            {
                await SeleniumScroll(driver);
                await SeleniumClassReplace(driver);
                var toplamVeri = driver.FindElements(By.ClassName("Nv2PK"));


                if (adresleriTara_CheckBox.Checked) // Adres seçili ise. içerisinde mail adresi de seçiliyse kontrol eder.
                {
                    await VerileriCekVeAktar_AdressSecili(toplamVeri, driver, rowIndex);

                }
                else
                {
                    await VerileriCekVeAktar(toplamVeri, driver, rowIndex);
                }
            }
            else
            {
                if (genelTarama_CheckBox.Checked)
                {
                    await UpdateGridRowColor(rowIndex, Color.DarkGreen, Color.White);
                }
            }




        }


        private static readonly object lockObject = new object();

        private async Task VerileriCekVeAktar_AdressSecili(ReadOnlyCollection<IWebElement> toplamVeri, ChromeDriver driver, int rowIndex)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            foreach (var element in toplamVeri)
            {
                try
                {
                    // Elementi tıklanabilir hale gelene kadar bekle
                    var elementIndex = toplamVeri.ToList().IndexOf(element); // Element'in indeksi
                    string clickElementScript = $@"document.querySelectorAll('div.Nv2PK')[{elementIndex}].querySelector('.hfpxzc').click();";
                    js.ExecuteScript(clickElementScript);

                    await Task.Delay(2000); // Bekleme yerine, elemanlar hazır olana kadar WebDriverWait kullanabilirsin

                    // Verileri çek
                    string firmaAdi = TryFindElementAttribute(element, By.CssSelector("a.hfpxzc"), "aria-label").Replace("·Ziyaret edilmiş bağlantı", "").Trim();
                    string telefon = TryFindElementText(element, By.XPath(".//div[contains(@class, 'W4Efsd')]//span[contains(@class, 'UsdlK')]")) ?? "";
                    string website = TryFindElementAttribute(element, By.XPath(".//a[contains(@aria-label, 'web sitesini ziyaret et')]"), "href") ?? "";
                    string adress = (string)js.ExecuteScript(@"return document.querySelector('.Io6YTe.fontBodyMedium.kR99db.fdkmkc') ? document.querySelector('.Io6YTe.fontBodyMedium.kR99db.fdkmkc').innerText : '';") ?? "";

                    // Yeni firma bilgilerini oluştur
                    var newFirma = new FirmaBilgileri
                    {
                        FirmaAdi = firmaAdi,
                        Telefon = telefon,
                        Website = website,
                        EMail = "", // Boş e-posta alanı
                        Adress = adress
                    };

                    // Verileri thread-safe olarak listeye ekle
                    lock (lockObject)
                    {
                        firmalarList.Add(newFirma);
                    }
                }
                catch (Exception ex)
                {
                    // Hataları işle
                    Console.WriteLine($"Hata: {ex.Message}");
                }
            }

            if (genelTarama_CheckBox.Checked)
            {
                await UpdateGridRowColor(rowIndex, Color.DarkGreen, Color.White); // İlgili satırı yeşil yap
            }
        }


        #region Mail işlemelri

        private async Task<string> MailAdresleriniCek(string url)
        {
            string mail = "";
            if (!String.IsNullOrEmpty(url))
            {
                string fullUrl = url;
                string baseUrl = GetBaseUrl(fullUrl);
                string emailPageUrl = await GetEmailPageUrl(baseUrl);
                if (!string.IsNullOrEmpty(emailPageUrl))
                {
                    mail = await GetEmail(emailPageUrl);
                }
                else
                {
                    mail = await GetEmail(url);
                }

            }
            return await Task.FromResult(mail);
        }

        private async Task<string> GetEmail(string url)
        {
            var mails = await MailleriCek(url);
            return await Task.FromResult(mails);
        }

        static async Task<string> GetEmailPageUrl(string baseUrl)
        {
            string[] paths = { "/iletisim", "/contact", "/hakkimizda", "/about", "" }; // Öncelikli yollar
            foreach (string path in paths)
            {
                string url = baseUrl + path;
                if (await CheckUrlExists(url))
                {
                    return url; // URL varsa geri döndür
                }
            }
            return null; // Hiçbiri bulunmazsa null döndür
        }

        // Verilen URL'nin var olup olmadığını kontrol et
        static async Task<bool> CheckUrlExists(string url)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await client.GetAsync(url);
                    return response.IsSuccessStatusCode; // Sayfa varsa true döner
                }
                catch
                {
                    return false; // İstek başarısız olursa false döner
                }
            }
        }

        static string GetBaseUrl(string url)
        {

            Uri uri = new(url);
            string baseUrl = uri.Scheme + "://" + uri.Host; // Protokol ve ana alan adı birleştiriliyor
            return baseUrl;
        }

        static string ExtractEmails(string htmlContent)
        {
            string emailPattern = @"(?:mailto:)?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})";
            Regex regex = new Regex(emailPattern);
            MatchCollection matches = regex.Matches(htmlContent);

            // Benzersiz e-posta adreslerini saklamak için HashSet kullan
            HashSet<string> uniqueEmails = new HashSet<string>();

            foreach (Match match in matches)
            {
                // Eğer regex bir grup ile eşleşirse, o grubu al
                if (match.Groups.Count > 1)
                {
                    uniqueEmails.Add(match.Groups[1].Value); // E-posta adresini HashSet'e ekle
                }
            }

            if (uniqueEmails.Count > 0)
            {
                // E-posta adreslerini virgülle ayırarak tek bir satırda döndür
                return string.Join(",", uniqueEmails);
            }
            else
            {
                return "";
            }
        }

        private async Task<string> MailleriCek(string url)
        {
            string htmlContent = await GetHtmlContent(url);
            string emails = "";

            if (string.IsNullOrEmpty(htmlContent))
            {
                // Yönlendirilmiş URL'yi bul
                string redirectUrl = await GetRedirectUrl(url);
                if (!string.IsNullOrEmpty(redirectUrl))
                {
                    htmlContent = await GetHtmlContent(redirectUrl);
                }
            }
            if (!string.IsNullOrEmpty(htmlContent))
            {
                emails = ExtractEmails(htmlContent);
            }
            return await Task.FromResult(emails);
        }

        private async Task<string> GetRedirectUrl(string url)
        {
            try
            {
                using HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");

                HttpResponseMessage response = await client.GetAsync(url);

                // Yönlendirme olup olmadığını kontrol et
                if (response.Headers.Location != null)
                {
                    return response.Headers.Location.ToString();
                }

                // Yönlendirme yoksa null döndür
                return null;
            }
            catch (Exception)
            {
                // Herhangi bir hata durumunda null döndür
                return null;
            }
        }
        static async Task<string> GetHtmlContent(string url)
        {
            using HttpClientHandler handler = new();
            handler.AllowAutoRedirect = true; // Yönlendirmeleri otomatik takip et
            using HttpClient client = new(handler);
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");

            try
            {
                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hata: {ex.Message}");
                return null;
            }
        }


        #endregion


        List<FirmaBilgileri> firmalarList = new();

        private async Task VerileriCekVeAktar(ReadOnlyCollection<IWebElement> toplamVeri, ChromeDriver driver, int rowIndex)
        {

            foreach (var item in toplamVeri)
            {
                string firmaAdi = "";
                string telefon = "";
                string website = "";
                string mail = "";
                string adress = "";

                firmaAdi = TryFindElementAttribute(item, By.CssSelector("a.hfpxzc"), "aria-label");

                // Telefon numarasını al
                telefon = TryFindElementText(item, By.XPath(".//div[contains(@class, 'W4Efsd')]//span[contains(@class, 'UsdlK')]"));
                if (string.IsNullOrEmpty(telefon))
                {
                    telefon = "";
                }

                // Web sitesini al
                website = TryFindElementAttribute(item, By.XPath(".//a[contains(@aria-label, 'web sitesini ziyaret et')]"), "href");
                if (string.IsNullOrEmpty(website))
                {
                    website = "";
                }


                var firma = new FirmaBilgileri
                {
                    FirmaAdi = firmaAdi,
                    Telefon = telefon,
                    Website = website,
                    EMail = mail,
                    Adress = adress
                };

                firmalarList.Add(firma);


            }
            if (genelTarama_CheckBox.Checked)
            {
                await UpdateGridRowColor(rowIndex, Color.DarkGreen, Color.White);
            }

        }


        #region Private Method
        private async Task UpdateGridRowColor(int rowIndex, Color backgroundColor, Color textColor)
        {
            // UI thread'inde çalışması için Invoke kullanıyoruz
            if (dataGrid_Ilce.InvokeRequired)
            {
                dataGrid_Ilce.Invoke(new Action(async () => await UpdateGridRowColor(rowIndex, backgroundColor, textColor)));
            }
            else
            {
                // Arka plan ve metin rengini güncelle
                dataGrid_Ilce.Rows[rowIndex].DefaultCellStyle.BackColor = backgroundColor;
                dataGrid_Ilce.Rows[rowIndex].DefaultCellStyle.ForeColor = textColor;
            }
        }

        private string TryFindElementAttribute(IWebElement parent, By by, string attribute)
        {
            try
            {
                var element = parent.FindElement(by);
                return element?.GetAttribute(attribute) ?? "";
            }
            catch (NoSuchElementException)
            {
                return ""; // Eğer eleman bulunamazsa, boş string döner
            }
        }

        private string TryFindElementText(IWebElement parent, By by)
        {
            try
            {
                var element = parent.FindElement(by);
                return element?.Text ?? "";
            }
            catch (NoSuchElementException)
            {
                return ""; // Eğer eleman bulunamazsa, boş string döner
            }
        }
        string GetText(IWebElement element, string fallback = "---")
        {
            try
            {
                return element.Text;
            }
            catch (NoSuchElementException)
            {
                return fallback;
            }
        }

        private async Task SeleniumClassReplace(ChromeDriver driver)
        {
            var script = @"
            var elements = document.querySelectorAll('[class]');
            elements.forEach(function(el) {
                var classList = el.className.split(' ');
                if (classList[0] === 'Nv2PK') { // İlk sınıf ismi Nv2PK mi?
                    el.className = 'Nv2PK'; // Sadece Nv2PK olarak bırak
                }
            });
        ";
            // JavaScript ile sınıfları değiştir
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript(script);

            // Değişiklikleri doğrulamak için bekleyin
            Thread.Sleep(3000);

        }
        private async Task SeleniumScroll(ChromeDriver driver)
        {
            // Konsola "allow-pasting" yazdırma
            string allowPastingScript = @"
        (function() {
            console.log('allow-pasting');
        })();";
            ((IJavaScriptExecutor)driver).ExecuteScript(allowPastingScript);
            await Task.Delay(3000);

            // JavaScript kodunu tanımlayın
            string script = @"
        return (async function() {
            // Sayfadaki tüm div öğelerini al
            var allDivs = document.querySelectorAll('div');

            // Scroll yapabileceğiniz div öğelerini filtrele
            var scrollableDivs = Array.from(allDivs).filter(div => div.scrollHeight > div.clientHeight);

            // Scroll yapılabilir div'leri kontrol et
            if (scrollableDivs.length === 0) {
                console.log('Scroll yapabileceğiniz bir div bulunamadı.');
                return false; // Scroll yapılacak div yok
            } else {
                console.log('Scroll yapabileceğiniz divler bulundu:', scrollableDivs);

                // Scroll yapma işlemi tüm div'ler için paralel yapılacak
                await Promise.all(scrollableDivs.map(async (div, index) => {
                    console.log(`Scroll yapılıyor: Div ${index + 1}`);

                    // Scroll işlemi
                    while (true) {
                        div.scrollTop += 500; // Aşağı kaydır
                        await new Promise(resolve => setTimeout(resolve, 3000)); // Her scroll sonrası bekleme

                        const messageElement = document.querySelector('.HlvSq');
                        if (messageElement && messageElement.innerText.includes('Listenin sonuna ulaştınız.')) {
                            console.log('Listenin sonuna ulaştınız.');
                            break; // Mesaj bulunduğunda döngüden çık
                        }

                        if (div.scrollTop + div.clientHeight >= div.scrollHeight) {
                            console.log(`Div ${index + 1} en altına ulaştı.`);
                            break; // En alta ulaşıldığında döngüden çık
                        }
                    }
                    console.log(`Div ${index + 1} scroll işlemi tamamlandı.`);
                }));

                console.log('Tüm divler üzerinde scroll işlemi tamamlandı.');
                return true; // Scroll işlemi tamamlandı
            }
        })();
    ";

            // Scroll işlemi tamamlanana kadar sürekli deneme
            bool scrollCompleted = false;

            while (!scrollCompleted)
            {
                try
                {
                    // JavaScript kodunu çalıştır ve sonucu bekle
                    var result = (bool)((IJavaScriptExecutor)driver).ExecuteScript(script);

                    if (result)
                    {
                        Console.WriteLine("Scroll işlemi tamamlandı.");
                        scrollCompleted = true; // Başarıyla tamamlandığında döngüden çık
                    }
                    else
                    {
                        Console.WriteLine("Scroll yapılacak div bulunamadı.");
                        await Task.Delay(2000); // 2 saniye bekle
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("JavaScript kodu zaman aşımına uğradı. Yeniden deneme...");
                    await Task.Delay(2000); // 2 saniye bekle
                }
            }

            // Daha fazla bekleme süresi ekleyin, gerekirse
            await Task.Delay(3000);
        }



        #endregion

        private void timer1_Tick(object sender, EventArgs e)
        {
            elapsedTime++; // Geçen süreyi artır

            // Saat, dakika ve saniye hesapla
            TimeSpan time = TimeSpan.FromSeconds(elapsedTime);
            label3.Text = $"{time.Hours:D2}:{time.Minutes:D2}:{time.Seconds:D2}"; // Label'ı güncelle
        }

        private async void MailAdresleriniTara_BtnClick_Click(object sender, EventArgs e)
        {
            pictureBoxMailTara.Visible = true;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Yeni bir satır ise atla
                if (row.IsNewRow)
                    continue;

                // Hücre verilerini al
                var website = row.Cells["Web Site"].Value?.ToString(); // Website sütunu
                var eMail = row.Cells["E-Mail"].Value?.ToString(); // E-Mail sütunu

                // E-posta adresini çek
                if (!string.IsNullOrEmpty(website))
                {
                    eMail = await MailAdresleriniCek(website);
                    // Hücre değerini güncelle
                    row.Cells["E-Mail"].Value = eMail;
                }

                // Arka plan rengi ve yazı rengi ayarlama
                if (string.IsNullOrEmpty(eMail))
                {
                    // E-posta null veya boşsa koyu kırmızı arka plan
                    row.Cells["E-Mail"].Style.BackColor = Color.DarkRed;
                    row.Cells["E-Mail"].Style.ForeColor = Color.White; // Yazı rengini beyaz yap
                }
                else
                {
                    // E-posta varsa koyu yeşil arka plan
                    row.Cells["E-Mail"].Style.BackColor = Color.DarkGreen;
                    row.Cells["E-Mail"].Style.ForeColor = Color.White; // Yazı rengini beyaz yap
                }
            }
            MessageBox.Show("Mail tarama işlemi tamamlandı.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            pictureBoxMailTara.Visible = false;
            int count = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Hücre kontrolü yapıyoruz
                if (row.Cells[3].Value != null && !string.IsNullOrWhiteSpace(row.Cells[3].Value.ToString()))
                {
                    count++;
                }
            }
            totalMailLabel.Text = count.ToString();


        }

        private void Sil_BtnClick_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            totalTelefonLabel.Text = "0";
            totalMailLabel.Text = "0";
            listBoxTaskDurum.Items.Clear();
            //dataGridView1.Columns.Clear();
            servisiDurdurBtn_Click.Enabled = false;
            MailAdresleriniTara_BtnClick.Enabled = false;
            excelKaydet_btnClick.Enabled = false;
            Sil_BtnClick.Enabled = false;
            Servisi_Baslat_BtnClick.Enabled = true;
            label3.Text = "0";
            toplamVeri_Label.Text = "0";
            genelTarama_CheckBox.Checked = false;
            adresleriTara_CheckBox.Checked = false;
        }

        private void excelKaydet_btnClick_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void AnaSayfa_Resize(object sender, EventArgs e)
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
        }
        bool _iptalEdildiMi = false;

        private async void servisiDurdurBtn_Click_Click(object sender, EventArgs e)
        {
            loaderPictureGoogleMaps.Visible = true;
            // İptal işlemi başlat
            if (cancellationTokenSource != null && !cancellationTokenSource.IsCancellationRequested)
            {
                _iptalEdildiMi = true;
                cancellationTokenSource.Cancel(); // Tüm thread'lere iptal sinyali gönder
                servisiDurdurBtn_Click.Enabled = false;


                try
                {
                    // Görevlerden herhangi biri tamamlanana kadar bekleyin
                    await Task.WhenAny(tasks);

                    // Eğer iptal edildiyse driver'ları kapatın ve işlemleri durdurun
                    foreach (var task in tasks)
                    {
                        if (task.IsCanceled)
                        {
                            chromeDriverService.Dispose();
                            break; // Eğer iptal gerçekleştiyse, işlemleri bitir.
                        }
                    }

                    // UI'yı güncelle
                    Invoke(new Action(() =>
                    {
                        UpdateTaskStatus("Tüm görevler iptal edildi.", "cancelled");
                        Servisi_Baslat_BtnClick.Enabled = true; // Servisi başlat butonunu tekrar etkinleştir
                    }));
                }
                catch (Exception ex)
                {
                    // Eğer bir hata oluşursa, burada loglayabilirsiniz
                    UpdateTaskStatus($"Thread durdurulurken hata oluştu: {ex.Message}", "failed");
                }
                timer1.Stop();
                loaderPictureGoogleMaps.Visible = false;
            }
        }

        private void listBoxTaskDurum_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            ListBoxItem item = (ListBoxItem)listBoxTaskDurum.Items[e.Index];
            e.DrawBackground();

            // Status'e göre arka plan ve yazı rengi ayarlama
            if (item.Status == "running")
            {
                e.Graphics.FillRectangle(Brushes.Olive, e.Bounds); // Koyu sarı arka plan
                e.Graphics.DrawString(item.Message, e.Font, Brushes.White, e.Bounds); // Beyaz yazı
            }
            else if (item.Status == "completed")
            {
                e.Graphics.FillRectangle(Brushes.DarkGreen, e.Bounds); // Koyu yeşil arka plan
                e.Graphics.DrawString(item.Message, e.Font, Brushes.White, e.Bounds); // Beyaz yazı
            }
            else if (item.Status == "failed")
            {
                e.Graphics.FillRectangle(Brushes.DarkRed, e.Bounds); // Koyu kırmızı arka plan
                e.Graphics.DrawString(item.Message, e.Font, Brushes.White, e.Bounds); // Beyaz yazı
            }
            else if (item.Status == "cancelled")
            {
                e.Graphics.FillRectangle(Brushes.Maroon, e.Bounds); // Koyu kırmızı arka plan
                e.Graphics.DrawString(item.Message, e.Font, Brushes.White, e.Bounds); // Beyaz yazı
            }

            e.DrawFocusRectangle();
        }



        #endregion


        #region WHATSAPP BOTU



        private CancellationTokenSource cancellationTokenSourceWP; // İptal için

        private void exceldenNumaraAktar_btnClick_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string excelPath = GetExcelFilePath();

            if (!string.IsNullOrEmpty(excelPath))
            {
                // Firma, adres ve telefon bilgilerini çekip DataGridView'e ekle
                List<DataRecord> records = GetDataFromExcel(excelPath);
                dataGridViewWP.DataSource = records; // DataGridView'e bağla
            }
            else
            {
                MessageBox.Show("Bir dosya seçmediniz.");
            }
        }

        #region Private Methods

        // Excel dosya yolunu seçme
        private string GetExcelFilePath()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Dosyaları|*.xls;*.xlsx";
                openFileDialog.Title = "Bir Excel Dosyası Seçin";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return null;
        }

        // Excel'den Firma, Adres ve Telefon bilgilerini çekme
        private List<DataRecord> GetDataFromExcel(string filePath)
        {
            var records = new List<DataRecord>();
            FileInfo fileInfo = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // İlk sayfa
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // 1. satır başlık için ayrılır
                {
                    string firma = worksheet.Cells[row, 1].Text; // A kolonu (Firma)
                    string telefon = worksheet.Cells[row, 2].Text; // B kolonu (Telefon)
                    if (!string.IsNullOrWhiteSpace(telefon) && !telefon.StartsWith("02") && !telefon.StartsWith("-"))
                    {
                        records.Add(new DataRecord
                        {
                            Firma = firma,
                            Telefon = telefon,
                            Status = "Bekleniyor" // Başlangıçta bekleniyor durumu
                        });
                    }
                }
            }
            return records;
        }

        #endregion

        private void servisDurdurWp_Click(object sender, EventArgs e)
        {
            cancellationTokenSourceWP.Cancel(); // Gönderme işlemini durdur
            WpMesajGonder_BtnClick.Enabled = true;
            pictureBoxWP.Visible = false;
            servisDurdurWp.Enabled = false;
        }

        private async void WpMesajGonder_BtnClick_Click(object sender, EventArgs e)
        {
            pictureBoxWP.Visible = true;
            WpMesajGonder_BtnClick.Enabled = false;
            servisDurdurWp.Enabled = true;
            cancellationTokenSource = new CancellationTokenSource(); // İptal kaynaklarını başlat


            await SendMessages(cancellationTokenSource.Token); // Gönderme işlemini başlat


            MessageBox.Show("İşlemeler tamamlandı.", "Tamamlandı", MessageBoxButtons.OK, MessageBoxIcon.Information);


            WpMesajGonder_BtnClick.Enabled = true;
            servisDurdurWp.Enabled = false;
            pictureBoxWP.Visible = false;

        }
        private async Task LoadWebPage(string url)
        {
            // Web sayfasını yükle
            webView21.Source = new Uri(url);

            // Yüklenmesini beklemek için bir TaskCompletionSource kullan
            var tcs = new TaskCompletionSource<bool>();

            // NavigationCompleted olayını dinle
            webView21.NavigationCompleted += (s, e) =>
            {
                if (e.IsSuccess)
                {
                    tcs.SetResult(true); // Yükleme başarılıysa sonucu ayarla
                }
                else
                {
                    tcs.SetResult(false); // Yükleme başarısızsa sonucu ayarla
                }
            };

            // Sayfanın yüklenmesini bekle
            await tcs.Task;
        }

        string _urlWP = "";

        private async Task SendMessages(CancellationToken cancellationToken)
        {
            try
            {
                foreach (DataGridViewRow row in dataGridViewWP.Rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    dataGridViewWP.ClearSelection();
                    row.Selected = true;
                    dataGridViewWP.CurrentCell = row.Cells[0];

                    string phone = row.Cells["Telefon"].Value?.ToString();
                    string firma = row.Cells["Firma"].Value?.ToString();

                    if (!string.IsNullOrEmpty(phone) && !string.IsNullOrEmpty(firma))
                    {
                        labelStatus.Text = $"Gönderiliyor: {firma} - {phone}";
                        row.Cells["Status"].Value = "Bekleniyor";
                        row.DefaultCellStyle.BackColor = Color.White;

                        string message = $"{richTextBox1.Text}";
                        string url = $"https://web.whatsapp.com/send/?phone=9{phone}&text={Uri.EscapeDataString(message)}";
                        _urlWP = url;

                        // Sayfa yüklenmesini bekle
                        var navigationTask = NavigationCompletedAsync();
                        webView21.Source = new Uri(url);

                        await navigationTask; // Sayfa yüklenene kadar bekler

                        cancellationToken.ThrowIfCancellationRequested();

                        // Buradan sonrası mesaj gönderme işlemi ile devam eder...
                        if (mesajKontrol_CheckBox.Checked)
                        {
                            var deger = await MesajKontrol();
                            if (!deger)
                            {
                                await ClickSendButtonAsync();
                            }
                        }
                        else
                        {
                            await ClickSendButtonAsync();
                        }

                        await Task.Delay(2000);
                        var res = await MesajKontrol();
                        await Task.Delay(2000);
                        if (res)
                        {
                            row.Cells["Status"].Value = "Başarılı";
                            row.DefaultCellStyle.BackColor = Color.DarkGreen;
                            row.DefaultCellStyle.ForeColor = Color.White;
                        }
                        else
                        {
                            row.Cells["Status"].Value = "Hatalı";
                            row.DefaultCellStyle.BackColor = Color.DarkRed;
                            row.DefaultCellStyle.ForeColor = Color.White;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(),"HATA");
                throw ex;
            }
            
        }

        //private async Task SendMessages(CancellationToken cancellationToken)
        //{
        //    foreach (DataGridViewRow row in dataGridViewWP.Rows)
        //    {

        //        cancellationToken.ThrowIfCancellationRequested(); // İptal durumu kontrolü
        //        // Satırı seçili hale getir
        //        dataGridViewWP.ClearSelection(); // Tüm seçimleri temizle
        //        row.Selected = true; // Şu anki satırı seçili yap
        //        dataGridViewWP.CurrentCell = row.Cells[0]; // İlk hücreyi seçili yap (görsel olarak vurgulamak için)

        //        string phone = row.Cells["Telefon"].Value?.ToString();
        //        string firma = row.Cells["Firma"].Value?.ToString();

        //        if (!string.IsNullOrEmpty(phone) && !string.IsNullOrEmpty(firma))
        //        {
        //            // Label kontrolüne anlık olarak yaz
        //            labelStatus.Text = $"Gönderiliyor: {firma} - {phone}";

        //            // Başlangıçta durumunu "Bekleniyor" olarak ayarla
        //            row.Cells["Status"].Value = "Bekleniyor";
        //            row.DefaultCellStyle.BackColor = Color.White; // Arka plan rengini beyaz yap

        //            string message = $"{richTextBox1.Text}";
        //            string url = $"https://web.whatsapp.com/send/?phone=9{phone}&text={Uri.EscapeDataString(message)}";
        //            _urlWP = url;
        //            //await LoadWebPage(url);
        //            webView21.Source = new Uri(url);


        //            await Task.Delay(10000); // 10 saniye bekleyin (gerekirse süreyi ayarlayın)
        //            cancellationToken.ThrowIfCancellationRequested(); // İptal durumu kontrolü

        //            if (mesajKontrol_CheckBox.Checked)
        //            {
        //                var deger = await MesajKontrol();
        //                if (!deger) // Mesaj yok ise
        //                {
        //                    var js = @"
        //                            var interval = setInterval(function() {
        //                                var sendButton = document.querySelector('button[aria-label=""Gönder""]');
        //                                if (sendButton) {
        //                                    sendButton.click();
        //                                    clearInterval(interval);
        //                                }
        //                            }, 1000);
        //                        ";
        //                    // JavaScript kodunu çalıştırın
        //                    await webView21.ExecuteScriptAsync(js);
        //                }
        //            }
        //            else
        //            {
        //                // "Gönder" butonuna tıklama işlemi
        //                var js = @"
        //                    var interval = setInterval(function() {
        //                        var sendButton = document.querySelector('button[aria-label=""Gönder""]');
        //                        if (sendButton) {
        //                            sendButton.click();
        //                            clearInterval(interval);
        //                        }
        //                    }, 1000);
        //                ";

        //                // JavaScript kodunu çalıştırın
        //                await webView21.ExecuteScriptAsync(js);
        //            }



        //            // Mesaj gönderiminden sonra kısa bir gecikme ekleyin
        //            await Task.Delay(5000); // Mesaj gönderildikten sonra bekleme süresi
        //            var res = await MesajKontrol(); // mesaj gönderilmiş mi?
        //            // Durumu güncelleyin ve arka plan rengini ayarlayın

        //            if (res)
        //            {
        //                row.Cells["Status"].Value = "Başarılı";
        //                row.DefaultCellStyle.BackColor = Color.DarkGreen; // Koyu yeşil arka plan
        //                row.DefaultCellStyle.ForeColor = Color.White; // Beyaz yazı rengi
        //            }
        //            else if (res && mesajKontrol_CheckBox.Checked)
        //            {
        //                row.Cells["Status"].Value = "Geçildi";
        //                row.DefaultCellStyle.BackColor = Color.Yellow; // Koyu kırmızı arka plan
        //                row.DefaultCellStyle.ForeColor = Color.White; // Beyaz yazı rengi
        //            }
        //            else
        //            {
        //                row.Cells["Status"].Value = "Hatalı";
        //                row.DefaultCellStyle.BackColor = Color.DarkRed; // Koyu kırmızı arka plan
        //                row.DefaultCellStyle.ForeColor = Color.White; // Beyaz yazı rengi
        //            }
        //        }
        //    }
        //}

        // Sayfa yüklenmesini beklemek için NavigationCompleted olayını asenkron olarak dinler
        private Task NavigationCompletedAsync()
        {
            var tcs = new TaskCompletionSource<bool>();

            // Event handler'ı sadece bir kez dinlemesini sağlamak için.
            EventHandler<CoreWebView2NavigationCompletedEventArgs> handler = null;
            handler = (sender, args) =>
            {
                webView21.NavigationCompleted -= handler;  // Olayı kaldırıyoruz ki bir daha tetiklenmesin

                if (args.IsSuccess)
                    tcs.TrySetResult(true);  // Başarılı şekilde yüklendi
                else
                    tcs.TrySetResult(false); // Hata meydana geldi
            };

            webView21.NavigationCompleted += handler;  // Olayı dinliyoruz
            return tcs.Task;
        }

        // Gönder butonunu bulup tıklayan method
        private async Task ClickSendButtonAsync()
        {
            var js = @"
        var interval = setInterval(function() {
            var sendButton = document.querySelector('button[aria-label=""Gönder""]');
            if (sendButton) {
                sendButton.click();
                clearInterval(interval);
            }
        }, 1000);
    ";

            await webView21.ExecuteScriptAsync(js);
        }
        private async Task<bool> MesajKontrol()
        {
            var checkJs = @"
            var rows = document.querySelectorAll('div[role=""row""]');
            rows.length > 1;
            ";

            var mesajVarMi = await webView21.ExecuteScriptAsync(checkJs);
            bool res = mesajVarMi == "true" ? true : false;
            return res;

        }


        private void wpSilBtn_Click(object sender, EventArgs e)
        {
            dataGridViewWP.Rows.Clear();
            dataGridViewWP.Columns.Clear();
            richTextBox1.Text = "Mesaj içeriğinizi buraya giriniz...";
        }



        #endregion


        #region EMAİL BOTU


        private List<string> fileAttachments = new List<string>();

        public void AddFileAttachment(string filePath)
        {
            if (!fileAttachments.Contains(filePath))
            {
                fileAttachments.Add(filePath);
            }

        }

        private void PopulateTextBoxes(string selectedHost)
        {
            // TextBox'ları seçilen host bilgilerine göre doldur
            switch (selectedHost)
            {
                case "Hostinger":
                    smt_TextBox.Text = "smtp.hostinger.com";
                    port_TextBox.Text = "587";
                    break;
                case "Natro":
                    smt_TextBox.Text = "mail.kurumsaleposta.com";
                    port_TextBox.Text = "587";
                    break;
            }
        }
        private void hostingler_ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateTextBoxes(hostingler_ComboBox.SelectedItem.ToString());
        }

        private CancellationTokenSource cancellationTokenSourceMailGonder;
        private async void MailGonder_BtnClick_Click(object sender, EventArgs e)
        {
            cancellationTokenSourceMailGonder = new CancellationTokenSource();
            var cancellationToken = cancellationTokenSourceMailGonder.Token;
            pictureBoxMail.Visible = true;

            string htmlContent = "";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "HTML Dosyaları (*.html;*.htm)|*.html;*.htm|Tüm Dosyalar (*.*)|*.*";
                openFileDialog.Title = "Mail Şablonunu Seçin";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName; // Seçilen dosya yolunu al
                    try
                    {
                        // HTML içeriğini dosyadan oku
                        htmlContent = File.ReadAllText(selectedFilePath);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Bir hata oluştu: {ex.Message}");
                    }
                }
            }

            MessageBox.Show("Servis başlatılıyor.Aşağıdaki ekrandan takip edebilirsiniz.", "Hazır", MessageBoxButtons.OK, MessageBoxIcon.Information);
            foreach (var item in listBox1.Items)
            {
                AddFileAttachment(item.ToString());
            }

            MailGonder_BtnClick.Enabled = false;
            servisiDurdurMailBtn.Enabled = true;
            foreach (DataGridViewRow row in dataGridViewMail.Rows)
            {
                var durum = row.Cells["Durum"].Value?.ToString();

                if (string.IsNullOrEmpty(durum) || durum != "Başarılı")
                {
                    string email = row.Cells["Email"].Value?.ToString();
                    bool mailSent = await SendEmailAsync(email, baslik_TextBox.Text, htmlContent, Convert.ToInt32(port_TextBox.Text), false, username_TextBox.Text, password_TextBox.Text, smt_TextBox.Text, cancellationToken);

                    if (mailSent)
                    {
                        row.Cells["Durum"].Value = "Başarılı";
                        row.Cells["Durum"].Style.BackColor = Color.Green;
                    }
                    else
                    {
                        row.Cells["Durum"].Value = "Başarısız";
                        row.Cells["Durum"].Style.BackColor = Color.Red;
                    }
                }


            }
            MessageBox.Show("Servis tamamlandı.Aşağıdaki ekrandan kontrol edebilirsiniz.", "Tamamlandı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            MailGonder_BtnClick.Enabled = true;
            servisiDurdurMailBtn.Enabled = false;
            pictureBoxMail.Visible = false;
        }

        public async Task<bool> SendEmailAsync(string aliciMails, string subject, string body, int port, bool enableSsl, string username, string password, string smtp, CancellationToken cancellationToken)
        {
            try
            {
                // SMTP ayarları
                SmtpClient smtpClient = new(smtp) // SMTP sunucusu
                {
                    Port = port,
                    Credentials = new NetworkCredential(username, password),
                    EnableSsl = enableSsl
                };

                MailMessage mailMessage = new MailMessage
                {
                    From = new MailAddress(gonderenMail_TextBox.Text),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true // HTML içeriği göndereceksek true
                };

                //mailMessage.To.Add(aliciMail);
                string[] aliciMailListesi = aliciMails.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var aliciMail in aliciMailListesi)
                {
                    string trimmedEmail = aliciMail.Trim(); // Boşlukları temizle
                    if (!string.IsNullOrWhiteSpace(trimmedEmail))
                    {
                        mailMessage.To.Add(trimmedEmail); // Her e-posta adresini ekle
                    }
                }

                // Dosya ekleme
                foreach (var filePath in fileAttachments)
                {
                    mailMessage.Attachments.Add(new Attachment(filePath));
                }

                // E-postayı gönder
                await smtpClient.SendMailAsync(mailMessage);
                return true;
            }
            catch (OperationCanceledException)
            {
                // İşlem iptal edildiğinde bu blok çalışır
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private static readonly HttpClient httpClient = new HttpClient();
        private const string apiKey = "b2ec9071e3ae47666011cfc4555bd012"; // API anahtarınızı buraya ekleyin
        private async Task<string> UploadImage(string base64Image)
        {
            var byteArray = Convert.FromBase64String(base64Image);

            using (var content = new MultipartFormDataContent())
            {
                content.Add(new ByteArrayContent(byteArray, 0, byteArray.Length), "image", "image.png");

                var response = await httpClient.PostAsync($"https://api.imgbb.com/1/upload?expiration=600&key={apiKey}", content);
                var result = await response.Content.ReadAsStringAsync();

                // JSON yanıtını ayrıştır
                dynamic jsonResponse = Newtonsoft.Json.JsonConvert.DeserializeObject(result);

                // bool değerini almak için Value özelliğini kullanın
                if (jsonResponse.success.Value)
                {
                    return jsonResponse.data.url; // Dönüş URL'sini al
                }
                else
                {
                    throw new Exception("Resim yükleme başarısız: " + jsonResponse.status);
                }
            }
        }

        public class ListBoxItem
        {
            public string Message { get; set; }
            public string Status { get; set; }
        }

        private void dosyaSec_BtnClick_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select a File";
                openFileDialog.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"; // Dosya filtreleri

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = Path.GetFileName(openFileDialog.FileName);
                    listBox1.Items.Add(openFileDialog.FileName); // Tam yol

                    // Görsel olarak dosyayı FlowLayoutPanel içinde göstermek için:
                    FlowLayoutPanel panel = new FlowLayoutPanel
                    {
                        Width = 60, // Buton ve resim için toplam genişlik
                        Height = 80, // Yükseklik, altındaki label için yeterli alan
                        Margin = new Padding(5),
                        AutoSize = true // Otomatik boyutlandırma
                    };

                    PictureBox pictureBox = new PictureBox
                    {
                        Image = openFileDialog.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ?
                                Image.FromFile(Path.Combine(Application.StartupPath, "images", "pdf.png")) : Image.FromFile(openFileDialog.FileName),
                        SizeMode = PictureBoxSizeMode.StretchImage,
                        Width = 50,
                        Height = 50,
                        Margin = new Padding(0) // Margin kaldırılıyor
                    };

                    // Dosya adını göstermek için Label oluştur
                    Label fileNameLabel = new Label
                    {
                        Text = fileName,
                        AutoSize = true,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Margin = new Padding(0)
                    };

                    // Silme butonunu oluştur
                    Button deleteButton = new Button
                    {
                        Text = "X",
                        ForeColor = Color.Red,
                        Size = new Size(20, 20),
                        Location = new Point(40, 0), // Sağ üst köşede konumlandırma
                        FlatStyle = FlatStyle.Flat,
                        BackColor = Color.Transparent
                    };

                    deleteButton.Click += (s, args) =>
                    {
                        // Butonun bulunduğu paneli kaldır
                        flowLayoutPanel1.Controls.Remove(panel);
                        listBox1.Items.Remove(openFileDialog.FileName); // Listeyi güncelle
                    };

                    // Panelin içine resim, dosya adı ve butonu ekle
                    panel.Controls.Add(pictureBox);
                    panel.Controls.Add(fileNameLabel); // Dosya adı label'ı ekleniyor
                    panel.Controls.Add(deleteButton);
                    flowLayoutPanel1.Controls.Add(panel);
                }
            }
        }

        private void exceldenYukle_BtnClick_Click(object sender, EventArgs e)
        {
            // OpenFileDialog ile Excel dosyasını seçmek
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Dosyaları|*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                LoadExcelData(filePath);
            }
        }

        private void LoadExcelData(string filePath)
        {

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workSheet = package.Workbook.Worksheets[0]; // İlk sayfa
                int rowCount = workSheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // 2. satırdan başla (başlık satırını atla)
                {
                    string firma = workSheet.Cells[row, 1].Value?.ToString();
                    string email = workSheet.Cells[row, 4].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(firma) && !string.IsNullOrWhiteSpace(email))
                    {
                        dataGridViewMail.Rows.Add(firma, email, "Bekleniyor");

                    }
                }
            }
        }

        private async void mailSablonuKaydetBtn_Click(object sender, EventArgs e)
        {
            string fileName = dosyaAdi_TextBox.Text; // Dosya adını al
            await webViewTiny.ExecuteScriptAsync($"downloadTinyMCEContent('{fileName}');");

            // Küçük bir bekleme süresi ekleyin (indirme için zaman tanımak amacıyla)
            await Task.Delay(5000); // 5 saniye bekleyin
            string uzanti = ".htm";
            string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", fileName) + uzanti;


            try
            {
                string htmlContent = await File.ReadAllTextAsync(downloadsPath);

                // Base64 resimlerini bulmak için regex
                string pattern = @"data:image/(?<type>.+?);base64,(?<data>.+?)""";
                var matches = Regex.Matches(htmlContent, pattern);

                var tasks = new List<Task<string>>();

                foreach (Match match in matches)
                {
                    string base64Image = match.Groups["data"].Value;
                    tasks.Add(UploadImage(base64Image));
                }

                string[] urls = await Task.WhenAll(tasks);

                // URL'leri HTML içeriğinde güncelleyin
                for (int i = 0; i < matches.Count; i++)
                {
                    htmlContent = htmlContent.Replace(matches[i].Value, urls[i]);
                }

                // Güncellenmiş HTML içeriğini dosyaya kaydedin
                await File.WriteAllTextAsync(downloadsPath, htmlContent);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Bir hata oluştu: {ex.Message}");
            }




            // MailSablon klasörünün yolu
            string destinationFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MailSablon");

            // MailSablon klasörünü oluştur
            if (!Directory.Exists(destinationFolder))
            {
                Directory.CreateDirectory(destinationFolder);
            }

            // Hedef dosya yolu (uzantıyı koruyun)
            string destinationFilePath = Path.Combine(destinationFolder, Path.GetFileNameWithoutExtension(fileName)) + uzanti;
            try
            {
                // Dosyanın var olup olmadığını kontrol et
                if (File.Exists(downloadsPath))
                {
                    File.Copy(downloadsPath, destinationFilePath, true); // Varsa üzerine yaz
                }
                else
                {
                    MessageBox.Show("Dosya indirilemedi veya bulunamadı.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Bir hata oluştu: {ex.Message}");
            }
        }

        private void servisiDurdurMailBtn_Click(object sender, EventArgs e)
        {
            if (cancellationTokenSourceMailGonder != null)
            {
                cancellationTokenSourceMailGonder.Cancel(); // İptal sinyali gönder
                MailGonder_BtnClick.Enabled = true;
                servisiDurdurMailBtn.Enabled = false;
                MessageBox.Show("Mail gönderme işlemi iptal edildi.", "İptal", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion



        #region AYARLAR


        private async void aktivasyonBtn_Click_1(object sender, EventArgs e)
        {
            var lisansManager = new LisansManager();

            _lisansKey = lisansManager.LisansDosyasiniOku().LisansKey;
            lisansTextBox.Text = _lisansKey;
            await lisansKontrolEt();
        }

        bool _googleMapsVeriBotuVisible = false;
        bool _whatsappBotuVisible = false;
        bool _emailBotuVisible = false;
        async Task lisansKontrolEt()
        {
            loaderPictureLisans.Visible = true;
            lisansOkPictureBox.Visible = false;
            cancelPictureBox.Visible = false;
            var lisansManager = new LisansManager();
            var lisansService = new LisansService();

            string machineCode = GetMotherboardSerialNumber().Trim('/');

            var result = await lisansService.LisansKontrolEtAsync(_lisansKey, machineCode);
            if (result is null)
            {
                MessageBox.Show($"{result.DurumAciklama}", "Hata!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cancelPictureBox.Visible = true;
            }
            if (result.Durum)
            {

                LisansKey_Label.Text = result.LisansKey;
                Durum_Label.Text = result.Durum == true ? "Aktif" : "Pasif";
                if (!result.Durum)
                {
                    Durum_Label.ForeColor = Color.Red;
                }
                BitisTarihi_Label.Text = result.Sure?.ToString("dd/MM/yyyy");
                loaderPictureLisans.Visible = false;
                lisansOkPictureBox.Visible = true;

                _googleMapsVeriBotuVisible = result.GoogleMapsBotu;
                _whatsappBotuVisible = result.WhatsAppBotu;
                _emailBotuVisible = result.EMailBotu;

            }
            else
            {
                MessageBox.Show($"{result.DurumAciklama}", "Hata!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LisansKey_Label.Text = result.LisansKey;
                Durum_Label.Text = result.Durum == true ? "Aktif" : "Pasif";
                if (!result.Durum)
                {
                    Durum_Label.ForeColor = Color.Red;
                }
                BitisTarihi_Label.Text = result.Sure?.ToString("dd/MM/yyyy");
                cancelPictureBox.Visible = true;
            }
            loaderPictureLisans.Visible = false;

            tabPageGosterGizle();


        }

        void tabPageGosterGizle()
        {
            if (_googleMapsVeriBotuVisible)
            {
                if (!tabControl1.TabPages.Contains(tabPage7))
                {
                    tabControl1.TabPages.Add(tabPage7);
                }
            }
            else
            {
                if (tabControl1.TabPages.Contains(tabPage7))
                {
                    tabControl1.TabPages.Remove(tabPage7);
                }
            }


            if (_emailBotuVisible)
            {
                if (!tabControl1.TabPages.Contains(tabPage8))
                {
                    tabControl1.TabPages.Add(tabPage8);
                }
            }
            else
            {
                if (tabControl1.TabPages.Contains(tabPage8))
                {
                    tabControl1.TabPages.Remove(tabPage8);
                }
            }


            if (_whatsappBotuVisible)
            {
                if (!tabControl1.TabPages.Contains(tabPage9))
                {
                    tabControl1.TabPages.Add(tabPage9);
                }
            }
            else
            {
                if (tabControl1.TabPages.Contains(tabPage9))
                {
                    tabControl1.TabPages.Remove(tabPage9);
                }
            }



        }


    }




    #endregion








}

public class FirmaBilgileri
{
    public string FirmaAdi { get; set; }
    public string Telefon { get; set; }
    public string Website { get; set; }
    public string? Adress { get; set; }
    public string? EMail { get; set; }
}

public class DataRecord
{
    public string Telefon { get; set; }
    public string Firma { get; set; }
    public string Status { get; set; }
}
