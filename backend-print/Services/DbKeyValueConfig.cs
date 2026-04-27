using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace backend_print.Services
{
    /// <summary>
    /// DB の m_key（k/v）から設定値を取得する。
    /// Web.config へのフォールバックは行わない（m_key に存在しない場合は例外）。
    /// </summary>
    public static class DbKeyValueConfig
    {
        private const string DefaultTableName = "dbo.m_key";

        public static string GetRequiredString(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new ArgumentException("key が空です。");

            var fromDb = GetFromDbOrThrow(key.Trim());
            return fromDb.Trim();
        }

        private static string GetFromDbOrThrow(string key)
        {
            try
            {
                var cs = ConfigurationManager.ConnectionStrings["MyDbConnection"]?.ConnectionString;
                if (string.IsNullOrWhiteSpace(cs))
                    throw new InvalidOperationException("connectionStrings['MyDbConnection'] が未設定です。");

                using (IDbConnection db = new SqlConnection(cs))
                using (var cmd = db.CreateCommand())
                {
                    db.Open();
                    cmd.CommandText = $"SELECT v FROM {DefaultTableName} WHERE k = @k;";
                    var p = cmd.CreateParameter();
                    p.ParameterName = "@k";
                    p.Value = key;
                    cmd.Parameters.Add(p);
                    var obj = cmd.ExecuteScalar();
                    var v = obj == null || obj == DBNull.Value ? null : Convert.ToString(obj);
                    if (string.IsNullOrWhiteSpace(v))
                        throw new InvalidOperationException($"m_key にキー '{key}' が存在しないか、値が空です。");
                    return v;
                }
            }
            catch
            {
                throw;
            }
        }
    }
}

