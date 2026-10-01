using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using CharityFundApp.Models;

namespace CharityFundApp.DataAccess
{
    /// <summary>
    /// Репозиторий для сущности «Доноры» (ADO.NET)
    /// </summary>
    public class DonorRepository
    {
        public List<Donor> GetAll(string? search = null)
        {
            var list = new List<Donor>();
            string sql = "SELECT Id, FullName, DonorType, Phone, Email, Address, RegistrationDate FROM Donors";
            var parameters = new List<SqlParameter>();

            if (!string.IsNullOrWhiteSpace(search))
            {
                sql += " WHERE FullName LIKE @Search OR Phone LIKE @Search OR Email LIKE @Search OR Address LIKE @Search";
                parameters.Add(new SqlParameter("@Search", $"%{search.Trim()}%"));
            }

            sql += " ORDER BY FullName;";

            var dt = DatabaseHelper.ExecuteQuery(sql, parameters.ToArray());
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new Donor
                {
                    Id = Convert.ToInt32(row["Id"]),
                    FullName = row["FullName"].ToString() ?? string.Empty,
                    DonorType = row["DonorType"].ToString() ?? "Физическое лицо",
                    Phone = row["Phone"] == DBNull.Value ? null : row["Phone"].ToString(),
                    Email = row["Email"] == DBNull.Value ? null : row["Email"].ToString(),
                    Address = row["Address"] == DBNull.Value ? null : row["Address"].ToString(),
                    RegistrationDate = Convert.ToDateTime(row["RegistrationDate"])
                });
            }
            return list;
        }

        public Donor? GetById(int id)
        {
            string sql = "SELECT Id, FullName, DonorType, Phone, Email, Address, RegistrationDate FROM Donors WHERE Id = @Id;";
            var dt = DatabaseHelper.ExecuteQuery(sql, new SqlParameter("@Id", id));

            if (dt.Rows.Count == 0) return null;
            var row = dt.Rows[0];
            return new Donor
            {
                Id = Convert.ToInt32(row["Id"]),
                FullName = row["FullName"].ToString() ?? string.Empty,
                DonorType = row["DonorType"].ToString() ?? "Физическое лицо",
                Phone = row["Phone"] == DBNull.Value ? null : row["Phone"].ToString(),
                Email = row["Email"] == DBNull.Value ? null : row["Email"].ToString(),
                Address = row["Address"] == DBNull.Value ? null : row["Address"].ToString(),
                RegistrationDate = Convert.ToDateTime(row["RegistrationDate"])
            };
        }

        public int Insert(Donor donor)
        {
            string sql = @"
                INSERT INTO Donors (FullName, DonorType, Phone, Email, Address, RegistrationDate)
                VALUES (@FullName, @DonorType, @Phone, @Email, @Address, @RegistrationDate);
                SELECT SCOPE_IDENTITY();";

            var parameters = new[]
            {
                new SqlParameter("@FullName", donor.FullName),
                new SqlParameter("@DonorType", donor.DonorType),
                new SqlParameter("@Phone", DatabaseHelper.ToDbValue(donor.Phone)),
                new SqlParameter("@Email", DatabaseHelper.ToDbValue(donor.Email)),
                new SqlParameter("@Address", DatabaseHelper.ToDbValue(donor.Address)),
                new SqlParameter("@RegistrationDate", donor.RegistrationDate)
            };

            var newId = DatabaseHelper.ExecuteScalar(sql, parameters);
            return Convert.ToInt32(newId);
        }

        public void Update(Donor donor)
        {
            string sql = @"
                UPDATE Donors 
                SET FullName = @FullName, 
                    DonorType = @DonorType, 
                    Phone = @Phone, 
                    Email = @Email, 
                    Address = @Address
                WHERE Id = @Id;";

            var parameters = new[]
            {
                new SqlParameter("@Id", donor.Id),
                new SqlParameter("@FullName", donor.FullName),
                new SqlParameter("@DonorType", donor.DonorType),
                new SqlParameter("@Phone", DatabaseHelper.ToDbValue(donor.Phone)),
                new SqlParameter("@Email", DatabaseHelper.ToDbValue(donor.Email)),
                new SqlParameter("@Address", DatabaseHelper.ToDbValue(donor.Address))
            };

            DatabaseHelper.ExecuteNonQuery(sql, parameters);
        }

        public void Delete(int id)
        {
            string sql = "DELETE FROM Donors WHERE Id = @Id;";
            DatabaseHelper.ExecuteNonQuery(sql, new[] { new SqlParameter("@Id", id) });
        }
    }
}
