namespace BotHizmetleri.Models;

public class Lisans
{
    public int Id { get; set; }
    public string AdSoyad { get; set; }
    public string LisansKey { get; set; }
    public string MachineCode { get; set; }
    public DateTime Sure { get; set; }
    public bool WhatsAppBotu { get; set; }
    public bool GoogleMapsBotu { get; set; }
    public bool EMailBotu { get; set; }
}
