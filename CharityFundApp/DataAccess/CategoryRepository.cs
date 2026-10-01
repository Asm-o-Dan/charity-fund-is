using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using CharityFundApp.Models;

namespace CharityFundApp.DataAccess
{
    /// <summary>
    /// Репозиторий для сущности «Категории» (ADO.NET)
    /// </summary>
    public class CategoryRepository
    {
        public List<Category> GetAll()
        {
            var list = new List<Category>();
            string sql = "SELECT Id, Name, Description FROM Categories ORDER BY Name;";
            var dt = DatabaseHelper.ExecuteQuery(sql);

            foreach (DataRow row in dt.Rows)
            {
                list.Add(new Category
                {
                    Id = Convert.ToInt32(row["Id"]),
                    Name = row["Name"].ToString() ?? string.Empty,
                    Description = row["Description"] == DBNull.Value ? null : row["Description"].ToString()
                });
            }
            return list;
        }

        public Category? GetById(int id)
        {
            string sql = "SELECT Id, Name, Description FROM Categories WHERE Id = @Id;";
            var dt = DatabaseHelper.ExecuteQuery(sql, new SqlParameter("@Id", id));

            if (dt.Rows.Count == 0) return null;
            var row = dt.Rows[0];
            return new Category
            {
                Id = Convert.ToInt32(row["Id"]),
                Name = row["Name"].ToString() ?? string.Empty,
                Description = row["Description"] == DBNull.Value ? null : row["Description"].ToString()
            };
        }

        public int Insert(Category cat)
        {
            string sql = @"
                INSERT INTO Categories (Name, Description) 
                VALUES (@Name, @Description);
                SELECT SCOPE_IDENTITY();";

            var parameters = new[]
            {
                new SqlParameter("@Name", cat.Name),
                new SqlParameter("@Description", DatabaseHelper.ToDbValue(cat.Description))
            };

            var newId = DatabaseHelper.ExecuteScalar(sql, parameters);
            return Convert.ToInt32(newId);
        }

        public void Update(Category cat)
        {
            string sql = @"
                UPDATE Categories 
                SET Name = @Name, Description = @Description 
                WHERE Id = @Id;";

            var parameters = new[]
            {
                new SqlParameter("@Id", cat.Id),
                new SqlParameter("@Name", cat.Name),
                new SqlParameter("@Description", DatabaseHelper.ToDbValue(cat.Description))
            };

            DatabaseHelper.ExecuteNonQuery(sql, parameters);
        }

        public void Delete(int id)
        {
            string sql = "DELETE FROM Categories WHERE Id = @Id;";
            DatabaseHelper.ExecuteNonQuery(sql, new[] { new SqlParameter("@Id", id) });
        }

        public bool IsUsed(int id)
        {
            string sql = @"
                SELECT 
                    (SELECT COUNT(*) FROM Projects WHERE CategoryId = @Id) +
                    (SELECT COUNT(*) FROM Recipients WHERE CategoryId = @Id);";

            var count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(sql, new[] { new SqlParameter("@Id", id) }) ?? 0);
            return count > 0;
        }
    }
}
