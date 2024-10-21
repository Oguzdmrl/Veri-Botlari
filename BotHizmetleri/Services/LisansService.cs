using BotHizmetleri.Models;
using Newtonsoft.Json;

namespace BotHizmetleri.Services
{
    public class LisansService
    {

        public LisansService()
        {
        }
        //string baseUrl = "https://localhost:44375/api/Lisans";  // API'nin temel URL'sini buraya koy
        string baseUrl = "http://api.stoyazilim.com.tr/api";  // API'nin temel URL'sini buraya koy
        public async Task<LisansKontrolResultModel> LisansKontrolEtAsync(string lisansKey, string machineCode)
        {
            if (string.IsNullOrEmpty(lisansKey)) throw new ArgumentNullException(nameof(lisansKey));
            if (string.IsNullOrEmpty(machineCode)) throw new ArgumentNullException(nameof(machineCode));

            using (HttpClientHandler handler = new HttpClientHandler())
            {
                handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;

                using (HttpClient client = new HttpClient(handler))
                {
                    string requestUrl = $"{baseUrl}/Lisans/LisansKontrol?lisansKey={lisansKey}&machineCode={machineCode}";

                    HttpResponseMessage response = await client.GetAsync(requestUrl);

                    // Başarılı cevap alındığında işle
                    if (response.IsSuccessStatusCode)
                    {
                        string responseData = await response.Content.ReadAsStringAsync();
                        Console.WriteLine("Response Data: " + responseData);
                    }
                    else
                    {
                        Console.WriteLine($"Error: {response.StatusCode}");
                    }
                    var content = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<LisansKontrolResultModel>(content);






                    return result;
                }
            }
        }

        //public async Task<LisansKontrolResultModel> LisansKontrolEtAsync(string lisansKey, string machineCode)
        //{
        //    LisansKontrolResultModel model = new LisansKontrolResultModel();
        //    model.DurumAciklama = "";
        //    string connectionString = "Server=bdijitaldeneme.com.tr;Database=u559731046_oguz;User Id=u559731046_oguz1;Password=932dZ$gP;";

        //    using (var connection = new MySqlConnection(connectionString))
        //    {
        //        await connection.OpenAsync();

        //        // Lisans listesini çekmek için (Eğer bu liste kullanılacaksa)
        //        var lisansList = await connection.QueryAsync<Lisans>("SELECT * FROM Lisans");

        //        // Lisans key ile lisans kontrolü
        //        var lisans = await connection.QueryFirstOrDefaultAsync<Lisans>(
        //            "SELECT * FROM Lisans WHERE LisansKey = @LisansKey", new { LisansKey = lisansKey });

        //        if (lisans == null)
        //        {
        //            model.DurumAciklama = "Lisans Bulunamadı..";
        //            model.Durum = false;
        //            return model;
        //        }

        //        if (lisans.MachineCode == null) // Makine kodu eklenmemişse, makine kodunu güncelle
        //        {
        //            lisans.MachineCode = machineCode;

        //            // Lisansı güncellemek için
        //            var updateQuery = "UPDATE Lisans SET MachineCode = @MachineCode WHERE LisansKey = @LisansKey";
        //            await connection.ExecuteAsync(updateQuery, new { MachineCode = machineCode, LisansKey = lisansKey });

        //            model.Durum = true;
        //            model.DurumAciklama = "Lisans Başarıyla Aktif Edildi.";
        //            model.Sure = lisans.Sure;
        //            model.LisansKey = lisans.LisansKey;
        //            model.WhatsAppBotu = lisans.WhatsAppBotu;
        //            model.EMailBotu = lisans.EMailBotu;
        //            model.GoogleMapsBotu = lisans.GoogleMapsBotu;
        //            return model;
        //        }

        //        if (lisans.Sure < DateTime.Now) // Lisans süresi kontrolü
        //        {
        //            model.DurumAciklama = "Lisans Süresi Sona Erdi..";
        //            model.Durum = false;
        //            return model;
        //        }

        //        // Lisans key ve machine code eşleşiyor mu kontrol et
        //        var query = await connection.QueryFirstOrDefaultAsync<Lisans>(
        //            "SELECT * FROM Lisans WHERE LisansKey = @LisansKey AND MachineCode = @MachineCode",
        //            new { LisansKey = lisansKey, MachineCode = machineCode });

        //        if (query == null)
        //        {
        //            model.DurumAciklama = "Lisans Bulunamadı..";
        //            model.Durum = false;
        //            return model;
        //        }
        //        else
        //        {
        //            model.Durum = true;
        //            model.DurumAciklama = "Lisans Başarıyla Aktif Edildi.";
        //            model.Sure = query.Sure;
        //            model.LisansKey = query.LisansKey;
        //            model.WhatsAppBotu = query.WhatsAppBotu;
        //            model.EMailBotu = query.EMailBotu;
        //            model.GoogleMapsBotu = query.GoogleMapsBotu;
        //            return model;
        //        }
        //    }
        //}


    }
}
