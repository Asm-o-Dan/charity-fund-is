using System;

namespace CharityFundApp.Models
{
    /// <summary>
    /// Сущность «Расход фонда»
    /// </summary>
    public class Expense
    {
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string? ProjectName { get; set; }
        public int? RecipientId { get; set; }
        public string? RecipientName { get; set; }
        public string ExpenseCategory { get; set; } = "Адресная помощь"; // "Адресная помощь", "Закупка медикаментов", "Оборудование", "Транспорт", "Административные расходы"
        public decimal Amount { get; set; }
        public DateTime ExpenseDate { get; set; } = DateTime.Today;
        public string DocumentNumber { get; set; } = string.Empty;
        public string? Description { get; set; }

        public override string ToString() => $"№{Id}: {Amount:N2} руб. ({ExpenseCategory} / {DocumentNumber})";
    }
}
