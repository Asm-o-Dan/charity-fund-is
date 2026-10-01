using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using CharityFundApp.Models;

namespace CharityFundApp.DataAccess
{
    /// <summary>
    /// Репозиторий для сущности «Пожертвования» (ADO.NET с транзакционным обновлением прогресса проекта)
    /// </summary>
    public class DonationRepository
    {
        public List<Donation> GetAll(int? projectId = null, int? donorId = null, DateTime? dateFrom = null, DateTime? dateTo = null)
        {
            var list = new List<Donation>();
            string sql = @"
                SELECT 
                    d.Id, d.DonorId, dn.FullName AS DonorName,
                    d.ProjectId, p.Name AS ProjectName,
                    d.Amount, d.DonationDate, d.PaymentMethod, d.Notes
                FROM Donations d
                INNER JOIN Donors dn ON d.DonorId = dn.Id
                INNER JOIN Projects p ON d.ProjectId = p.Id
                WHERE 1=1";

            var parameters = new List<SqlParameter>();

            if (projectId.HasValue && projectId.Value > 0)
            {
                sql += " AND d.ProjectId = @ProjectId";
                parameters.Add(new SqlParameter("@ProjectId", projectId.Value));
            }

            if (donorId.HasValue && donorId.Value > 0)
            {
                sql += " AND d.DonorId = @DonorId";
                parameters.Add(new SqlParameter("@DonorId", donorId.Value));
            }

            if (dateFrom.HasValue)
            {
                sql += " AND d.DonationDate >= @DateFrom";
                parameters.Add(new SqlParameter("@DateFrom", dateFrom.Value.Date));
            }

            if (dateTo.HasValue)
            {
                sql += " AND d.DonationDate <= @DateTo";
                parameters.Add(new SqlParameter("@DateTo", dateTo.Value.Date.AddDays(1).AddTicks(-1)));
            }

            sql += " ORDER BY d.DonationDate DESC, d.Id DESC;";

            var dt = DatabaseHelper.ExecuteQuery(sql, parameters.ToArray());
            foreach (DataRow row in dt.Rows)
            {
                list.Add(new Donation
                {
                    Id = Convert.ToInt32(row["Id"]),
                    DonorId = Convert.ToInt32(row["DonorId"]),
                    DonorName = row["DonorName"].ToString(),
                    ProjectId = Convert.ToInt32(row["ProjectId"]),
                    ProjectName = row["ProjectName"].ToString(),
                    Amount = Convert.ToDecimal(row["Amount"]),
                    DonationDate = Convert.ToDateTime(row["DonationDate"]),
                    PaymentMethod = row["PaymentMethod"].ToString() ?? "Банковская карта",
                    Notes = row["Notes"] == DBNull.Value ? null : row["Notes"].ToString()
                });
            }
            return list;
        }

        public Donation? GetById(int id)
        {
            string sql = @"
                SELECT 
                    d.Id, d.DonorId, dn.FullName AS DonorName,
                    d.ProjectId, p.Name AS ProjectName,
                    d.Amount, d.DonationDate, d.PaymentMethod, d.Notes
                FROM Donations d
                INNER JOIN Donors dn ON d.DonorId = dn.Id
                INNER JOIN Projects p ON d.ProjectId = p.Id
                WHERE d.Id = @Id;";

            var dt = DatabaseHelper.ExecuteQuery(sql, new SqlParameter("@Id", id));
            if (dt.Rows.Count == 0) return null;
            var row = dt.Rows[0];

            return new Donation
            {
                Id = Convert.ToInt32(row["Id"]),
                DonorId = Convert.ToInt32(row["DonorId"]),
                DonorName = row["DonorName"].ToString(),
                ProjectId = Convert.ToInt32(row["ProjectId"]),
                ProjectName = row["ProjectName"].ToString(),
                Amount = Convert.ToDecimal(row["Amount"]),
                DonationDate = Convert.ToDateTime(row["DonationDate"]),
                PaymentMethod = row["PaymentMethod"].ToString() ?? "Банковская карта",
                Notes = row["Notes"] == DBNull.Value ? null : row["Notes"].ToString()
            };
        }

        /// <summary>
        /// Добавляет пожертвование и автоматически увеличивает текущий баланс проекта в единой транзакции
        /// </summary>
        public int InsertWithTransaction(Donation donation)
        {
            using var conn = DatabaseHelper.CreateConnection();
            conn.Open();
            using var transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                // 1. Вставка пожертвования
                string insertSql = @"
                    INSERT INTO Donations (DonorId, ProjectId, Amount, DonationDate, PaymentMethod, Notes)
                    VALUES (@DonorId, @ProjectId, @Amount, @DonationDate, @PaymentMethod, @Notes);
                    SELECT " + (DatabaseHelper.CurrentProvider == DatabaseProvider.Sqlite ? "last_insert_rowid();" : "SCOPE_IDENTITY();");

                int newId;
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = insertSql;

                    void AddParam(string name, object? val)
                    {
                        var p = cmd.CreateParameter();
                        p.ParameterName = name;
                        p.Value = val ?? DBNull.Value;
                        cmd.Parameters.Add(p);
                    }

                    AddParam("@DonorId", donation.DonorId);
                    AddParam("@ProjectId", donation.ProjectId);
                    AddParam("@Amount", donation.Amount);
                    AddParam("@DonationDate", donation.DonationDate);
                    AddParam("@PaymentMethod", donation.PaymentMethod);
                    AddParam("@Notes", donation.Notes);

                    newId = Convert.ToInt32(cmd.ExecuteScalar());
                }

                // 2. Обновление собранных средств в проекте
                string updateProjectSql = @"
                    UPDATE Projects
                    SET CurrentAmount = CurrentAmount + @Amount
                    WHERE Id = @ProjectId;";

                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = updateProjectSql;

                    var pAmount = cmd.CreateParameter();
                    pAmount.ParameterName = "@Amount";
                    pAmount.Value = donation.Amount;
                    cmd.Parameters.Add(pAmount);

                    var pProj = cmd.CreateParameter();
                    pProj.ParameterName = "@ProjectId";
                    pProj.Value = donation.ProjectId;
                    cmd.Parameters.Add(pProj);

                    cmd.ExecuteNonQuery();
                }

                transaction.Commit();
                return newId;
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }

        /// <summary>
        /// Удаляет пожертвование и корректирует баланс проекта в транзакции
        /// </summary>
        public void DeleteWithTransaction(int donationId)
        {
            using var conn = DatabaseHelper.CreateConnection();
            conn.Open();
            using var transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);

            try
            {
                var donation = GetById(donationId);
                if (donation == null) return;

                // 1. Уменьшение собранной суммы проекта
                string updateProjectSql = @"
                    UPDATE Projects
                    SET CurrentAmount = CASE 
                        WHEN CurrentAmount >= @Amount THEN CurrentAmount - @Amount 
                        ELSE 0 
                    END
                    WHERE Id = @ProjectId;";

                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = updateProjectSql;

                    var pAmount = cmd.CreateParameter();
                    pAmount.ParameterName = "@Amount";
                    pAmount.Value = donation.Amount;
                    cmd.Parameters.Add(pAmount);

                    var pProj = cmd.CreateParameter();
                    pProj.ParameterName = "@ProjectId";
                    pProj.Value = donation.ProjectId;
                    cmd.Parameters.Add(pProj);

                    cmd.ExecuteNonQuery();
                }

                // 2. Удаление самого пожертвования
                string deleteSql = "DELETE FROM Donations WHERE Id = @Id;";
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = deleteSql;

                    var pId = cmd.CreateParameter();
                    pId.ParameterName = "@Id";
                    pId.Value = donationId;
                    cmd.Parameters.Add(pId);

                    cmd.ExecuteNonQuery();
                }

                transaction.Commit();
            }
            catch
            {
                transaction.Rollback();
                throw;
            }
        }
    }
}
