using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using CharityFundApp.Models;

namespace CharityFundApp.DataAccess
{
    /// <summary>
    /// Репозиторий для сущности «Получатели помощи» (ADO.NET)
    /// </summary>
    public class RecipientRepository
    {
        public List<Recipient> GetAll(string? statusFilter = null, string? search = null)
        {
            var list = new List<Recipient>();
            string sql = @"
                SELECT 
                    r.Id, r.FullName, r.CategoryId, c.Name AS CategoryName,
                    r.Phone, r.Address, r.NeedDescription, r.Status, r.RegistrationDate
                FROM Recipients r
                INNER JOIN Categories c ON r.CategoryId = c.Id
                WHERE 1=1";

            var parameters = new List<SqlParameter>();

            if (!string.IsNullOrWhiteSpace(statusFilter) && statusFilter != "Все")
            {
                sql += " AND r.Status = @Status";
                parameters.Add(new SqlParameter("@Status", statusFilter));
            }

            if (!string.IsNullOrWhiteSpace(search))
            {
                sql += " AND (r.FullName LIKE @Search OR r.NeedDescription LIKE @Search OR r.Address LIKE @Search)";
                parameters.Add(new SqlParameter("@Search", $"%{search.Trim()}%"));
            }

            sql += " ORDER BY r.Id DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql, parameters.ToArray());
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new Recipient
                {
                    Id = Convert.ToInt32(row["Id"]),
                    FullName = row["FullName"].ToString() ?? string.Empty,
                    CategoryId = Convert.ToInt32(row["CategoryId"]),
                    CategoryName = row["CategoryName"].ToString(),
                    Phone = row["Phone"] == DBNull.Value ? null : row["Phone"].ToString(),
                    Address = row["Address"] == DBNull.Value ? null : row["Address"].ToString(),
                    NeedDescription = row["NeedDescription"].ToString() ?? string.Empty,
                    Status = row["Status"].ToString() ?? "На рассмотрении",
                    RegistrationDate = Convert.ToDateTime(row["RegistrationDate"])
                });
            }
            return list;
        }

        public Recipient? GetById(int id)
        {
            string sql = @"
                SELECT 
                    r.Id, r.FullName, r.CategoryId, c.Name AS CategoryName,
                    r.Phone, r.Address, r.NeedDescription, r.Status, r.RegistrationDate
                FROM Recipients r
                INNER JOIN Categories c ON r.CategoryId = c.Id
                WHERE r.Id = @Id;";

            var dt = DatabaseHelper.ExecuteQuery(sql, new SqlParameter("@Id", id));
            if (dt.Rows.Count == 0) return null;
            var row = dt.Rows[0];

            return new Recipient
            {
                Id = Convert.ToInt32(row["Id"]),
                FullName = row["FullName"].ToString() ?? string.Empty,
                CategoryId = Convert.ToInt32(row["CategoryId"]),
                CategoryName = row["CategoryName"].ToString(),
                Phone = row["Phone"] == DBNull.Value ? null : row["Phone"].ToString(),
                Address = row["Address"] == DBNull.Value ? null : row["Address"].ToString(),
                NeedDescription = row["NeedDescription"].ToString() ?? string.Empty,
                Status = row["Status"].ToString() ?? "На рассмотрении",
                RegistrationDate = Convert.ToDateTime(row["RegistrationDate"])
            };
        }

        public int Insert(Recipient r)
        {
            string sql = @"
                INSERT INTO Recipients (FullName, CategoryId, Phone, Address, NeedDescription, Status, RegistrationDate)
                VALUES (@FullName, @CategoryId, @Phone, @Address, @NeedDescription, @Status, @RegistrationDate);
                SELECT SCOPE_IDENTITY();";

            var parameters = new[]
            {
                new SqlParameter("@FullName", r.FullName),
                new SqlParameter("@CategoryId", r.CategoryId),
                new SqlParameter("@Phone", DatabaseHelper.ToDbValue(r.Phone)),
                new SqlParameter("@Address", DatabaseHelper.ToDbValue(r.Address)),
                new SqlParameter("@NeedDescription", r.NeedDescription),
                new SqlParameter("@Status", r.Status),
                new SqlParameter("@RegistrationDate", r.RegistrationDate)
            };

            var newId = DatabaseHelper.ExecuteScalar(sql, parameters);
            return Convert.ToInt32(newId);
        }

        public void Update(Recipient r)
        {
            string sql = @"
                UPDATE Recipients 
                SET FullName = @FullName,
                    CategoryId = @CategoryId,
                    Phone = @Phone,
                    Address = @Address,
                    NeedDescription = @NeedDescription,
                    Status = @Status,
                    RegistrationDate = @RegistrationDate
                WHERE Id = @Id;";

            var parameters = new[]
            {
                new SqlParameter("@Id", r.Id),
                new SqlParameter("@FullName", r.FullName),
                new SqlParameter("@CategoryId", r.CategoryId),
                new SqlParameter("@Phone", DatabaseHelper.ToDbValue(r.Phone)),
                new SqlParameter("@Address", DatabaseHelper.ToDbValue(r.Address)),
                new SqlParameter("@NeedDescription", r.NeedDescription),
                new SqlParameter("@Status", r.Status),
                new SqlParameter("@RegistrationDate", r.RegistrationDate)
            };

            DatabaseHelper.ExecuteNonQuery(sql, parameters);
        }

        public void Delete(int id)
        {
            string sql = "DELETE FROM Recipients WHERE Id = @Id;";
            DatabaseHelper.ExecuteNonQuery(sql, new[] { new SqlParameter("@Id", id) });
        }
    }
}
