using System;

namespace CharityFundApp.Models
{
    /// <summary>
    /// Сущность «Получатель благотворительной помощи»
    /// </summary>
    public class Recipient
    {
        public int Id { get; set; }
        public string FullName { get; set; } = string.Empty;
        public int CategoryId { get; set; }
        public string? CategoryName { get; set; }
        public string? Phone { get; set; }
        public string? Address { get; set; }
        public string NeedDescription { get; set; } = string.Empty;
        public string Status { get; set; } = "На рассмотрении"; // "На рассмотрении", "Одобрен", "Помощь оказана", "Отклонен"
        public DateTime RegistrationDate { get; set; } = DateTime.Today;

        public override string ToString() => FullName;
    }
}
