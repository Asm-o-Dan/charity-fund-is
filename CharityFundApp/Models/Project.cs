using System;

namespace CharityFundApp.Models
{
    /// <summary>
    /// Сущность «Благотворительный проект»
    /// </summary>
    public class Project
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public int CategoryId { get; set; }
        public string? CategoryName { get; set; }
        public decimal TargetAmount { get; set; }
        public decimal CurrentAmount { get; set; }
        public DateTime StartDate { get; set; } = DateTime.Today;
        public DateTime? EndDate { get; set; }
        public string Status { get; set; } = "Активен"; // "Активен", "Завершен", "Приостановлен"
        public string? Description { get; set; }

        /// <summary>
        /// Процент собранных средств
        /// </summary>
        public decimal CompletionPercentage =>
            TargetAmount > 0 ? Math.Round((CurrentAmount / TargetAmount) * 100m, 2) : 0m;

        /// <summary>
        /// Оставшаяся сумма к сбору
        /// </summary>
        public decimal RemainingAmount =>
            TargetAmount > CurrentAmount ? TargetAmount - CurrentAmount : 0m;

        public override string ToString() => Name;
    }
}
