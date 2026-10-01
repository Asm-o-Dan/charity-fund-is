using System;

namespace CharityFundApp.Models
{
    /// <summary>
    /// Модель отчета 1: «Сумма пожертвований» (по проектам и за период)
    /// </summary>
    public class DonationSummaryReportItem
    {
        public int ProjectId { get; set; }
        public string ProjectName { get; set; } = string.Empty;
        public string CategoryName { get; set; } = string.Empty;
        public int DonationsCount { get; set; }
        public decimal TotalDonated { get; set; }
        public decimal TargetAmount { get; set; }
        public DateTime? FirstDonationDate { get; set; }
        public DateTime? LastDonationDate { get; set; }
    }

    /// <summary>
    /// Модель отчета 2: «Активные проекты» (цель, собрано, процент, остаток)
    /// </summary>
    public class ActiveProjectReportItem
    {
        public int ProjectId { get; set; }
        public string ProjectName { get; set; } = string.Empty;
        public string CategoryName { get; set; } = string.Empty;
        public decimal TargetAmount { get; set; }
        public decimal CurrentAmount { get; set; }
        public decimal CompletionPercentage { get; set; }
        public decimal RemainingAmount { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string Status { get; set; } = string.Empty;
    }

    /// <summary>
    /// Модель отчета 3: «Крупнейшие доноры» (рейтинг благотворителей)
    /// </summary>
    public class TopDonorReportItem
    {
        public int DonorId { get; set; }
        public string DonorName { get; set; } = string.Empty;
        public string DonorType { get; set; } = string.Empty;
        public string? Phone { get; set; }
        public string? Email { get; set; }
        public int DonationsCount { get; set; }
        public decimal TotalDonated { get; set; }
        public decimal MaxSingleDonation { get; set; }
        public DateTime? LastDonationDate { get; set; }
    }

    /// <summary>
    /// Модель отчета 4: «Расходы фонда» (баланс собранных и израсходованных средств)
    /// </summary>
    public class FundExpenseReportItem
    {
        public int ProjectId { get; set; }
        public string ProjectName { get; set; } = string.Empty;
        public decimal TotalCollected { get; set; }
        public decimal TotalSpent { get; set; }
        public decimal Balance { get; set; }
        public int ExpensesCount { get; set; }
    }

    /// <summary>
    /// Модель отчета 5: «Помощь по категориям» (распределение по направлениям)
    /// </summary>
    public class CategoryAssistanceReportItem
    {
        public int CategoryId { get; set; }
        public string CategoryName { get; set; } = string.Empty;
        public int ProjectsCount { get; set; }
        public int RecipientsCount { get; set; }
        public decimal TotalDonated { get; set; }
        public decimal TotalExpenses { get; set; }
    }
}
