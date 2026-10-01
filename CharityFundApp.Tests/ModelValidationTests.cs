using System;
using Xunit;
using CharityFundApp.Models;

namespace CharityFundApp.Tests
{
    public class ModelValidationTests
    {
        [Fact]
        public void Project_CompletionPercentage_CalculatesCorrectly()
        {
            var project = new Project
            {
                TargetAmount = 100000m,
                CurrentAmount = 75000m
            };

            Assert.Equal(75.00m, project.CompletionPercentage);
        }

        [Fact]
        public void Project_CompletionPercentage_ZeroTarget_ReturnsZero()
        {
            var project = new Project
            {
                TargetAmount = 0m,
                CurrentAmount = 5000m
            };

            Assert.Equal(0m, project.CompletionPercentage);
        }

        [Fact]
        public void Project_RemainingAmount_CalculatesCorrectly()
        {
            var project = new Project
            {
                TargetAmount = 150000m,
                CurrentAmount = 115000m
            };

            Assert.Equal(35000m, project.RemainingAmount);
        }

        [Fact]
        public void Project_RemainingAmount_Overfunded_ReturnsZero()
        {
            var project = new Project
            {
                TargetAmount = 50000m,
                CurrentAmount = 60000m
            };

            Assert.Equal(0m, project.RemainingAmount);
        }

        [Fact]
        public void Donor_DefaultType_IsIndividual()
        {
            var donor = new Donor();
            Assert.Equal("Физическое лицо", donor.DonorType);
        }

        [Fact]
        public void Recipient_DefaultStatus_IsUnderReview()
        {
            var recipient = new Recipient();
            Assert.Equal("На рассмотрении", recipient.Status);
        }

        [Fact]
        public void Donation_ToString_ContainsAmountAndId()
        {
            var donation = new Donation
            {
                Id = 15,
                Amount = 12500m,
                DonorName = "Иванов И.И.",
                ProjectName = "Детское здоровье"
            };

            var str = donation.ToString();
            Assert.Contains("15", str);
            Assert.Contains("12", str);
            Assert.Contains("Иванов И.И.", str);
        }
    }
}
