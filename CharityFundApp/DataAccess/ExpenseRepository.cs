using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using CharityFundApp.Models;

namespace CharityFundApp.DataAccess
{
    /// <summary>
    /// Репозиторий для сущности «Расходы фонда» (ADO.NET)
    /// </summary>
    public class ExpenseRepository
    {
        public List<Expense> GetAll(int? projectId = null, string? category = null, DateTime? dateFrom = null, DateTime? dateTo = null)
        {
            var list = new List<Expense>();
            string sql = @"
                SELECT 
                    e.Id, e.ProjectId, p.Name AS ProjectName,
                    e.RecipientId, r.FullName AS RecipientName,
                    e.ExpenseCategory, e.Amount, e.ExpenseDate,
                    e.DocumentNumber, e.Description
                FROM Expenses e
                INNER JOIN Projects p ON e.ProjectId = p.Id
                LEFT JOIN Recipients r ON e.RecipientId = r.Id
                WHERE 1=1";

            var parameters = new List<SqlParameter>();

            if (projectId.HasValue && projectId.Value > 0)
            {
                sql += " AND e.ProjectId = @ProjectId";
                parameters.Add(new SqlParameter("@ProjectId", projectId.Value));
            }

            if (!string.IsNullOrWhiteSpace(category) && category != "Все")
            {
                sql += " AND e.ExpenseCategory = @Category";
                parameters.Add(new SqlParameter("@Category", category));
            }

            if (dateFrom.HasValue)
            {
                sql += " AND e.ExpenseDate >= @DateFrom";
                parameters.Add(new SqlParameter("@DateFrom", dateFrom.Value.Date));
            }

            if (dateTo.HasValue)
            {
                sql += " AND e.ExpenseDate <= @DateTo";
                parameters.Add(new SqlParameter("@DateTo", dateTo.Value.Date));
            }

            sql += " ORDER BY e.ExpenseDate DESC, e.Id DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql, parameters.ToArray());
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new Expense
                {
                    Id = Convert.ToInt32(row["Id"]),
                    ProjectId = Convert.ToInt32(row["ProjectId"]),
                    ProjectName = row["ProjectName"].ToString(),
                    RecipientId = row["RecipientId"] == DBNull.Value ? null : Convert.ToInt32(row["RecipientId"]),
                    RecipientName = row["RecipientName"] == DBNull.Value ? null : row["RecipientName"].ToString(),
                    ExpenseCategory = row["ExpenseCategory"].ToString() ?? "Адресная помощь",
                    Amount = Convert.ToDecimal(row["Amount"]),
                    ExpenseDate = Convert.ToDateTime(row["ExpenseDate"]),
                    DocumentNumber = row["DocumentNumber"].ToString() ?? string.Empty,
                    Description = row["Description"] == DBNull.Value ? null : row["Description"].ToString()
                });
            }
            return list;
        }

        public Expense? GetById(int id)
        {
            string sql = @"
                SELECT 
                    e.Id, e.ProjectId, p.Name AS ProjectName,
                    e.RecipientId, r.FullName AS RecipientName,
                    e.ExpenseCategory, e.Amount, e.ExpenseDate,
                    e.DocumentNumber, e.Description
                FROM Expenses e
                INNER JOIN Projects p ON e.ProjectId = p.Id
                LEFT JOIN Recipients r ON e.RecipientId = r.Id
                WHERE e.Id = @Id;";

            var dt = DatabaseHelper.ExecuteQuery(sql, new SqlParameter("@Id", id));
            if (dt.Rows.Count == 0) return null;
            var row = dt.Rows[0];

            return new Expense
            {
                Id = Convert.ToInt32(row["Id"]),
                ProjectId = Convert.ToInt32(row["ProjectId"]),
                ProjectName = row["ProjectName"].ToString(),
                RecipientId = row["RecipientId"] == DBNull.Value ? null : Convert.ToInt32(row["RecipientId"]),
                RecipientName = row["RecipientName"] == DBNull.Value ? null : row["RecipientName"].ToString(),
                ExpenseCategory = row["ExpenseCategory"].ToString() ?? "Адресная помощь",
                Amount = Convert.ToDecimal(row["Amount"]),
                ExpenseDate = Convert.ToDateTime(row["ExpenseDate"]),
                DocumentNumber = row["DocumentNumber"].ToString() ?? string.Empty,
                Description = row["Description"] == DBNull.Value ? null : row["Description"].ToString()
            };
        }

        public int Insert(Expense exp)
        {
            string sql = @"
                INSERT INTO Expenses (ProjectId, RecipientId, ExpenseCategory, Amount, ExpenseDate, DocumentNumber, Description)
                VALUES (@ProjectId, @RecipientId, @ExpenseCategory, @Amount, @ExpenseDate, @DocumentNumber, @Description);
                SELECT SCOPE_IDENTITY();";

            var parameters = new[]
            {
                new SqlParameter("@ProjectId", exp.ProjectId),
                new SqlParameter("@RecipientId", DatabaseHelper.ToDbValue(exp.RecipientId)),
                new SqlParameter("@ExpenseCategory", exp.ExpenseCategory),
                new SqlParameter("@Amount", exp.Amount),
                new SqlParameter("@ExpenseDate", exp.ExpenseDate),
                new SqlParameter("@DocumentNumber", exp.DocumentNumber),
                new SqlParameter("@Description", DatabaseHelper.ToDbValue(exp.Description))
            };

            var newId = DatabaseHelper.ExecuteScalar(sql, parameters);
            return Convert.ToInt32(newId);
        }

        public void Update(Expense exp)
        {
            string sql = @"
                UPDATE Expenses 
                SET ProjectId = @ProjectId,
                    RecipientId = @RecipientId,
                    ExpenseCategory = @ExpenseCategory,
                    Amount = @Amount,
                    ExpenseDate = @ExpenseDate,
                    DocumentNumber = @DocumentNumber,
                    Description = @Description
                WHERE Id = @Id;";

            var parameters = new[]
            {
                new SqlParameter("@Id", exp.Id),
                new SqlParameter("@ProjectId", exp.ProjectId),
                new SqlParameter("@RecipientId", DatabaseHelper.ToDbValue(exp.RecipientId)),
                new SqlParameter("@ExpenseCategory", exp.ExpenseCategory),
                new SqlParameter("@Amount", exp.Amount),
                new SqlParameter("@ExpenseDate", exp.ExpenseDate),
                new SqlParameter("@DocumentNumber", exp.DocumentNumber),
                new SqlParameter("@Description", DatabaseHelper.ToDbValue(exp.Description))
            };

            DatabaseHelper.ExecuteNonQuery(sql, parameters);
        }

        public void Delete(int id)
        {
            string sql = "DELETE FROM Expenses WHERE Id = @Id;";
            DatabaseHelper.ExecuteNonQuery(sql, new[] { new SqlParameter("@Id", id) });
        }
    }
}
