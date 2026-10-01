using System;

namespace CharityFundApp.Models
{
    /// <summary>
    /// Сущность «Пожертвование / Взнос»
    /// </summary>
    public class Donation
    {
        public int Id { get; set; }
        public int DonorId { get; set; }
        public string? DonorName { get; set; }
        public int ProjectId { get; set; }
        public string? ProjectName { get; set; }
        public decimal Amount { get; set; }
        public DateTime DonationDate { get; set; } = DateTime.Now;
        public string PaymentMethod { get; set; } = "Банковская карта"; // "Банковская карта", "Банковский перевод", "Наличные", "Онлайн-платеж"
        public string? Notes { get; set; }

        public override string ToString() => $"№{Id}: {Amount:N2} руб. ({DonorName} -> {ProjectName})";
    }
}
