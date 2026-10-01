using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using CharityFundApp.Models;

namespace CharityFundApp.DataAccess
{
    /// <summary>
    /// Репозиторий аналитических запросов и отчетов (ADO.NET)
    /// Реализует 5 обязательных аналитических выборок предметной области
    /// </summary>
    public class ReportRepository
    {
        /// <summary>
        /// Отчет 1: Сумма пожертвований (с группировкой по проектам и опциональным фильтром по датам)
        /// </summary>
        public List<DonationSummaryReportItem> GetDonationsSummary(DateTime? dateFrom = null, DateTime? dateTo = null)
        {
            var list = new List<DonationSummaryReportItem>();
            string sql = @"
                SELECT 
                    p.Id AS ProjectId,
                    p.Name AS ProjectName,
                    c.Name AS CategoryName,
                    p.TargetAmount,
                    COUNT(d.Id) AS DonationsCount,
                    ISNULL(SUM(d.Amount), 0.00) AS TotalDonated,
                    MIN(d.DonationDate) AS FirstDonationDate,
                    MAX(d.DonationDate) AS LastDonationDate
                FROM Projects p
                INNER JOIN Categories c ON p.CategoryId = c.Id
                LEFT JOIN Donations d ON p.Id = d.ProjectId 
                    AND (@DateFrom IS NULL OR d.DonationDate >= @DateFrom)
                    AND (@DateTo IS NULL OR d.DonationDate <= @DateTo)
                GROUP BY p.Id, p.Name, c.Name, p.TargetAmount
                ORDER BY TotalDonated DESC;";

            var parameters = new[]
            {
                new SqlParameter("@DateFrom", dateFrom.HasValue ? (object)dateFrom.Value.Date : DBNull.Value),
                new SqlParameter("@DateTo", dateTo.HasValue ? (object)dateTo.Value.Date.AddDays(1).AddTicks(-1) : DBNull.Value)
            };

            var dt = DatabaseHelper.ExecuteQuery(sql, parameters);
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new DonationSummaryReportItem
                {
                    ProjectId = Convert.ToInt32(row["ProjectId"]),
                    ProjectName = row["ProjectName"].ToString() ?? string.Empty,
                    CategoryName = row["CategoryName"].ToString() ?? string.Empty,
                    TargetAmount = Convert.ToDecimal(row["TargetAmount"]),
                    DonationsCount = Convert.ToInt32(row["DonationsCount"]),
                    TotalDonated = Convert.ToDecimal(row["TotalDonated"]),
                    FirstDonationDate = row["FirstDonationDate"] == DBNull.Value ? null : Convert.ToDateTime(row["FirstDonationDate"]),
                    LastDonationDate = row["LastDonationDate"] == DBNull.Value ? null : Convert.ToDateTime(row["LastDonationDate"])
                });
            }
            return list;
        }

        /// <summary>
        /// Отчет 2: Активные проекты (статус, цель, собрано, процент выполнения, остаток к сбору)
        /// </summary>
        public List<ActiveProjectReportItem> GetActiveProjects()
        {
            var list = new List<ActiveProjectReportItem>();
            string sql = @"
                SELECT 
                    p.Id AS ProjectId,
                    p.Name AS ProjectName,
                    c.Name AS CategoryName,
                    p.TargetAmount,
                    p.CurrentAmount,
                    CASE 
                        WHEN p.TargetAmount > 0 THEN ROUND((p.CurrentAmount / p.TargetAmount) * 100.0, 2)
                        ELSE 0.00 
                    END AS CompletionPercentage,
                    CASE 
                        WHEN p.TargetAmount > p.CurrentAmount THEN (p.TargetAmount - p.CurrentAmount)
                        ELSE 0.00 
                    END AS RemainingAmount,
                    p.StartDate,
                    p.EndDate,
                    p.Status
                FROM Projects p
                INNER JOIN Categories c ON p.CategoryId = c.Id
                WHERE p.Status = N'Активен'
                ORDER BY p.CurrentAmount DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql);
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new ActiveProjectReportItem
                {
                    ProjectId = Convert.ToInt32(row["ProjectId"]),
                    ProjectName = row["ProjectName"].ToString() ?? string.Empty,
                    CategoryName = row["CategoryName"].ToString() ?? string.Empty,
                    TargetAmount = Convert.ToDecimal(row["TargetAmount"]),
                    CurrentAmount = Convert.ToDecimal(row["CurrentAmount"]),
                    CompletionPercentage = Convert.ToDecimal(row["CompletionPercentage"]),
                    RemainingAmount = Convert.ToDecimal(row["RemainingAmount"]),
                    StartDate = Convert.ToDateTime(row["StartDate"]),
                    EndDate = row["EndDate"] == DBNull.Value ? null : Convert.ToDateTime(row["EndDate"]),
                    Status = row["Status"].ToString() ?? "Активен"
                });
            }
            return list;
        }

        /// <summary>
        /// Отчет 3: Крупнейшие доноры (топ благотворителей по общей сумме)
        /// </summary>
        public List<TopDonorReportItem> GetTopDonors(int limit = 10)
        {
            var list = new List<TopDonorReportItem>();
            string sql = @"
                SELECT TOP (@Limit)
                    d.Id AS DonorId,
                    d.FullName AS DonorName,
                    d.DonorType,
                    d.Phone,
                    d.Email,
                    COUNT(dn.Id) AS DonationsCount,
                    ISNULL(SUM(dn.Amount), 0.00) AS TotalDonated,
                    ISNULL(MAX(dn.Amount), 0.00) AS MaxSingleDonation,
                    MAX(dn.DonationDate) AS LastDonationDate
                FROM Donors d
                INNER JOIN Donations dn ON d.Id = dn.DonorId
                GROUP BY d.Id, d.FullName, d.DonorType, d.Phone, d.Email
                ORDER BY TotalDonated DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql, new SqlParameter("@Limit", limit));
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new TopDonorReportItem
                {
                    DonorId = Convert.ToInt32(row["DonorId"]),
                    DonorName = row["DonorName"].ToString() ?? string.Empty,
                    DonorType = row["DonorType"].ToString() ?? string.Empty,
                    Phone = row["Phone"] == DBNull.Value ? null : row["Phone"].ToString(),
                    Email = row["Email"] == DBNull.Value ? null : row["Email"].ToString(),
                    DonationsCount = Convert.ToInt32(row["DonationsCount"]),
                    TotalDonated = Convert.ToDecimal(row["TotalDonated"]),
                    MaxSingleDonation = Convert.ToDecimal(row["MaxSingleDonation"]),
                    LastDonationDate = row["LastDonationDate"] == DBNull.Value ? null : Convert.ToDateTime(row["LastDonationDate"])
                });
            }
            return list;
        }

        /// <summary>
        /// Отчет 4: Расходы фонда (сопоставление собранных и израсходованных средств, баланс)
        /// </summary>
        public List<FundExpenseReportItem> GetFundExpensesBalance()
        {
            var list = new List<FundExpenseReportItem>();
            string sql = @"
                SELECT 
                    p.Id AS ProjectId,
                    p.Name AS ProjectName,
                    p.CurrentAmount AS TotalCollected,
                    ISNULL(e.TotalSpent, 0.00) AS TotalSpent,
                    (p.CurrentAmount - ISNULL(e.TotalSpent, 0.00)) AS Balance,
                    ISNULL(e.ExpensesCount, 0) AS ExpensesCount
                FROM Projects p
                LEFT JOIN (
                    SELECT 
                        ProjectId, 
                        SUM(Amount) AS TotalSpent, 
                        COUNT(Id) AS ExpensesCount 
                    FROM Expenses 
                    GROUP BY ProjectId
                ) e ON p.Id = e.ProjectId
                ORDER BY TotalCollected DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql);
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new FundExpenseReportItem
                {
                    ProjectId = Convert.ToInt32(row["ProjectId"]),
                    ProjectName = row["ProjectName"].ToString() ?? string.Empty,
                    TotalCollected = Convert.ToDecimal(row["TotalCollected"]),
                    TotalSpent = Convert.ToDecimal(row["TotalSpent"]),
                    Balance = Convert.ToDecimal(row["Balance"]),
                    ExpensesCount = Convert.ToInt32(row["ExpensesCount"])
                });
            }
            return list;
        }

        /// <summary>
        /// Отчет 5: Помощь по категориям (дети, медицина, малоимущие и др.)
        /// </summary>
        public List<CategoryAssistanceReportItem> GetCategoryAssistance()
        {
            var list = new List<CategoryAssistanceReportItem>();
            string sql = @"
                SELECT 
                    c.Id AS CategoryId,
                    c.Name AS CategoryName,
                    COUNT(DISTINCT p.Id) AS ProjectsCount,
                    COUNT(DISTINCT r.Id) AS RecipientsCount,
                    ISNULL((SELECT SUM(dn.Amount) 
                            FROM Donations dn 
                            INNER JOIN Projects pr ON dn.ProjectId = pr.Id 
                            WHERE pr.CategoryId = c.Id), 0.00) AS TotalDonated,
                    ISNULL((SELECT SUM(ex.Amount) 
                            FROM Expenses ex 
                            INNER JOIN Projects pr ON ex.ProjectId = pr.Id 
                            WHERE pr.CategoryId = c.Id), 0.00) AS TotalExpenses
                FROM Categories c
                LEFT JOIN Projects p ON c.Id = p.CategoryId
                LEFT JOIN Recipients r ON c.Id = r.CategoryId
                GROUP BY c.Id, c.Name
                ORDER BY TotalDonated DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql);
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new CategoryAssistanceReportItem
                {
                    CategoryId = Convert.ToInt32(row["CategoryId"]),
                    CategoryName = row["CategoryName"].ToString() ?? string.Empty,
                    ProjectsCount = Convert.ToInt32(row["ProjectsCount"]),
                    RecipientsCount = Convert.ToInt32(row["RecipientsCount"]),
                    TotalDonated = Convert.ToDecimal(row["TotalDonated"]),
                    TotalExpenses = Convert.ToDecimal(row["TotalExpenses"])
                });
            }
            return list;
        }
    }
}
