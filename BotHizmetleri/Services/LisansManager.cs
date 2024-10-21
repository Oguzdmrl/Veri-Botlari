using System.Text.Json;

namespace BotHizmetleri.Services
{
    public class LisansDosyasi
    {
        public string LisansKey { get; set; }
    }

    public class LisansManager
    {
        private readonly string _dosyaYolu; // Lisans dosyasının yolu

        public LisansManager()
        {
            // Projenin çalıştığı dizinde lisans.json dosyasını oluştur
            _dosyaYolu = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "lisans.json");
        }

        public LisansDosyasi LisansDosyasiniOku()
        {
            if (File.Exists(_dosyaYolu))
            {
                var jsonData = File.ReadAllText(_dosyaYolu);
                return JsonSerializer.Deserialize<LisansDosyasi>(jsonData);
            }
            return null; // Dosya yoksa null döner
        }

        public void LisansDosyasinaYaz(LisansDosyasi lisansDosyasi)
        {
            // Dosya yoksa oluştur
            if (!File.Exists(_dosyaYolu))
            {
                // Dosyayı oluştur
                using (var dosya = File.Create(_dosyaYolu))
                {
                    dosya.Close(); // Dosya oluşturulmuş olur
                }
            }

            // Lisans anahtarını JSON dosyasına kaydet
            var jsonData = JsonSerializer.Serialize(lisansDosyasi, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_dosyaYolu, jsonData);
        }

        public void LisansDosyasiniOluştur(string lisansKey)
        {
      
                var lisansDosyasi = new LisansDosyasi { LisansKey = lisansKey };
                LisansDosyasinaYaz(lisansDosyasi);
       
        }
    }

}