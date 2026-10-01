using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using CharityFundApp.Models;

namespace CharityFundApp.DataAccess
{
    /// <summary>
    /// Репозиторий для сущности «Проекты» (ADO.NET)
    /// </summary>
    public class ProjectRepository
    {
        public List<Project> GetAll(string? statusFilter = null, string? search = null)
        {
            var list = new List<Project>();
            string sql = @"
                SELECT 
                    p.Id, p.Name, p.CategoryId, c.Name AS CategoryName,
                    p.TargetAmount, p.CurrentAmount, p.StartDate, p.EndDate,
                    p.Status, p.Description
                FROM Projects p
                INNER JOIN Categories c ON p.CategoryId = c.Id
                WHERE 1=1";

            var parameters = new List<SqlParameter>();

            if (!string.IsNullOrWhiteSpace(statusFilter) && statusFilter != "Все")
            {
                sql += " AND p.Status = @Status";
                parameters.Add(new SqlParameter("@Status", statusFilter));
            }

            if (!string.IsNullOrWhiteSpace(search))
            {
                sql += " AND (p.Name LIKE @Search OR p.Description LIKE @Search OR c.Name LIKE @Search)";
                parameters.Add(new SqlParameter("@Search", $"%{search.Trim()}%"));
            }

            sql += " ORDER BY p.Id DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql, parameters.ToArray());
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new Project
                {
                    Id = Convert.ToInt32(row["Id"]),
                    Name = row["Name"].ToString() ?? string.Empty,
                    CategoryId = Convert.ToInt32(row["CategoryId"]),
                    CategoryName = row["CategoryName"].ToString(),
                    TargetAmount = Convert.ToDecimal(row["TargetAmount"]),
                    CurrentAmount = Convert.ToDecimal(row["CurrentAmount"]),
                    StartDate = Convert.ToDateTime(row["StartDate"]),
                    EndDate = row["EndDate"] == DBNull.Value ? null : Convert.ToDateTime(row["EndDate"]),
                    Status = row["Status"].ToString() ?? "Активен",
                    Description = row["Description"] == DBNull.Value ? null : row["Description"].ToString()
                });
            }
            return list;
        }

        public Project? GetById(int id)
        {
            string sql = @"
                SELECT 
                    p.Id, p.Name, p.CategoryId, c.Name AS CategoryName,
                    p.TargetAmount, p.CurrentAmount, p.StartDate, p.EndDate,
                    p.Status, p.Description
                FROM Projects p
                INNER JOIN Categories c ON p.CategoryId = c.Id
                WHERE p.Id = @Id;";

            var dt = DatabaseHelper.ExecuteQuery(sql, new SqlParameter("@Id", id));
            if (dt.Rows.Count == 0) return null;
            var row = dt.Rows[0];

            return new Project
            {
                Id = Convert.ToInt32(row["Id"]),
                Name = row["Name"].ToString() ?? string.Empty,
                CategoryId = Convert.ToInt32(row["CategoryId"]),
                CategoryName = row["CategoryName"].ToString(),
                TargetAmount = Convert.ToDecimal(row["TargetAmount"]),
                CurrentAmount = Convert.ToDecimal(row["CurrentAmount"]),
                StartDate = Convert.ToDateTime(row["StartDate"]),
                EndDate = row["EndDate"] == DBNull.Value ? null : Convert.ToDateTime(row["EndDate"]),
                Status = row["Status"].ToString() ?? "Активен",
                Description = row["Description"] == DBNull.Value ? null : row["Description"].ToString()
            };
        }

        public int Insert(Project p)
        {
            string sql = @"
                INSERT INTO Projects (Name, CategoryId, TargetAmount, CurrentAmount, StartDate, EndDate, Status, Description)
                VALUES (@Name, @CategoryId, @TargetAmount, @CurrentAmount, @StartDate, @EndDate, @Status, @Description);
                SELECT SCOPE_IDENTITY();";

            var parameters = new[]
            {
                new SqlParameter("@Name", p.Name),
                new SqlParameter("@CategoryId", p.CategoryId),
                new SqlParameter("@TargetAmount", p.TargetAmount),
                new SqlParameter("@CurrentAmount", p.CurrentAmount),
                new SqlParameter("@StartDate", p.StartDate),
                new SqlParameter("@EndDate", DatabaseHelper.ToDbValue(p.EndDate)),
                new SqlParameter("@Status", p.Status),
                new SqlParameter("@Description", DatabaseHelper.ToDbValue(p.Description))
            };

            var newId = DatabaseHelper.ExecuteScalar(sql, parameters);
            return Convert.ToInt32(newId);
        }

        public void Update(Project p)
        {
            string sql = @"
                UPDATE Projects 
                SET Name = @Name,
                    CategoryId = @CategoryId,
                    TargetAmount = @TargetAmount,
                    CurrentAmount = @CurrentAmount,
                    StartDate = @StartDate,
                    EndDate = @EndDate,
                    Status = @Status,
                    Description = @Description
                WHERE Id = @Id;";

            var parameters = new[]
            {
                new SqlParameter("@Id", p.Id),
                new SqlParameter("@Name", p.Name),
                new SqlParameter("@CategoryId", p.CategoryId),
                new SqlParameter("@TargetAmount", p.TargetAmount),
                new SqlParameter("@CurrentAmount", p.CurrentAmount),
                new SqlParameter("@StartDate", p.StartDate),
                new SqlParameter("@EndDate", DatabaseHelper.ToDbValue(p.EndDate)),
                new SqlParameter("@Status", p.Status),
                new SqlParameter("@Description", DatabaseHelper.ToDbValue(p.Description))
            };

            DatabaseHelper.ExecuteNonQuery(sql, parameters);
        }

        public void Delete(int id)
        {
            string sql = "DELETE FROM Projects WHERE Id = @Id;";
            DatabaseHelper.ExecuteNonQuery(sql, new[] { new SqlParameter("@Id", id) });
        }
    }
}
