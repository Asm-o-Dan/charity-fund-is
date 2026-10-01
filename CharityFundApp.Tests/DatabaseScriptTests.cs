using System;
using System.IO;
using Xunit;
using CharityFundApp.DataAccess;

namespace CharityFundApp.Tests
{
    public class DatabaseScriptTests
    {
        [Fact]
        public void DatabaseHelper_ToDbValue_HandlesNullAndNonNull()
        {
            Assert.Equal(DBNull.Value, DatabaseHelper.ToDbValue(null));
            Assert.Equal("Тест", DatabaseHelper.ToDbValue("Тест"));
            Assert.Equal(123, DatabaseHelper.ToDbValue(123));
        }

        [Fact]
        public void SqlScript_ContainsAllEntitiesAndViews()
        {
            // Ищем скрипт в проекте
            string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sql", "CharityFund_CreateDB.sql");
            if (!File.Exists(scriptPath))
            {
                scriptPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "CharityFundApp", "Sql", "CharityFund_CreateDB.sql"));
            }

            Assert.True(File.Exists(scriptPath), $"Файл скрипта не найден по пути: {scriptPath}");

            string sql = File.ReadAllText(scriptPath);

            // Проверка сущностей
            Assert.Contains("CREATE TABLE Categories", sql);
            Assert.Contains("CREATE TABLE Donors", sql);
            Assert.Contains("CREATE TABLE Projects", sql);
            Assert.Contains("CREATE TABLE Recipients", sql);
            Assert.Contains("CREATE TABLE Donations", sql);
            Assert.Contains("CREATE TABLE Expenses", sql);

            // Проверка 5 представлений
            Assert.Contains("vw_DonationsSummary", sql);
            Assert.Contains("vw_ActiveProjects", sql);
            Assert.Contains("vw_TopDonors", sql);
            Assert.Contains("vw_FundExpenses", sql);
            Assert.Contains("vw_CategoryAssistance", sql);

            // Проверка внешних ключей и индексов
            Assert.Contains("FK_Donations_Donors", sql);
            Assert.Contains("FK_Donations_Projects", sql);
            Assert.Contains("FK_Expenses_Projects", sql);
            Assert.Contains("CREATE NONCLUSTERED INDEX IX_Donations_ProjectId", sql);
        }
    }
}
