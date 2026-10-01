using System;

namespace CharityFundApp.Models
{
    /// <summary>
    /// Сущность «Донор / Благотворитель»
    /// </summary>
    public class Donor
    {
        public int Id { get; set; }
        public string FullName { get; set; } = string.Empty;
        public string DonorType { get; set; } = "Физическое лицо"; // "Физическое лицо" / "Юридическое лицо"
        public string? Phone { get; set; }
        public string? Email { get; set; }
        public string? Address { get; set; }
        public DateTime RegistrationDate { get; set; } = DateTime.Now;

        public override string ToString() => $"{FullName} ({DonorType})";
    }
}
