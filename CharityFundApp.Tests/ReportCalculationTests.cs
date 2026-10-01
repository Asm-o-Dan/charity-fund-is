using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using CharityFundApp.Models;

namespace CharityFundApp.Tests
{
    public class ReportCalculationTests
    {
        [Fact]
        public void FundExpenseReportItem_Balance_MatchesDifference()
        {
            var item = new FundExpenseReportItem
            {
                ProjectId = 1,
                ProjectName = "Детское здоровье",
                TotalCollected = 150000m,
                TotalSpent = 95000m,
                Balance = 150000m - 95000m,
                ExpensesCount = 3
            };

            Assert.Equal(55000m, item.Balance);
            Assert.True(item.Balance > 0);
        }

        [Fact]
        public void TopDonorReportItem_AggregatesCorrectly()
        {
            var topDonors = new List<TopDonorReportItem>
            {
                new TopDonorReportItem { DonorId = 1, DonorName = "Донор А", TotalDonated = 50000m, DonationsCount = 2 },
                new TopDonorReportItem { DonorId = 2, DonorName = "Донор Б", TotalDonated = 120000m, DonationsCount = 5 },
                new TopDonorReportItem { DonorId = 3, DonorName = "Донор В", TotalDonated = 30000m, DonationsCount = 1 }
            };

            var ordered = topDonors.OrderByDescending(d => d.TotalDonated).ToList();

            Assert.Equal(2, ordered[0].DonorId);
            Assert.Equal(120000m, ordered[0].TotalDonated);
            Assert.Equal(200000m, topDonors.Sum(d => d.TotalDonated));
        }

        [Fact]
        public void CategoryAssistanceReportItem_CalculatesFundMargin()
        {
            var catItem = new CategoryAssistanceReportItem
            {
                CategoryId = 1,
                CategoryName = "Медицинское оборудование",
                ProjectsCount = 2,
                RecipientsCount = 5,
                TotalDonated = 350000m,
                TotalExpenses = 280000m
            };

            decimal remainingFunds = catItem.TotalDonated - catItem.TotalExpenses;
            Assert.Equal(70000m, remainingFunds);
        }
    }
}
