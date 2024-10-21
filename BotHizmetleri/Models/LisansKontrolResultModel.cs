namespace BotHizmetleri.Models;

public class LisansKontrolResultModel
{
    public string? LisansKey { get; set; }
    public DateTime? Sure { get; set; }
    public bool Durum { get; set; }
    public string? DurumAciklama { get; set; }
    public string? MachineCode { get; set; }
    public bool WhatsAppBotu { get; set; }
    public bool GoogleMapsBotu { get; set; }
    public bool EMailBotu { get; set; }
}