using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace PACKINGSEAL.Utils
{
    public class AppCommonModule
    {
        private static readonly string connectStr = ConfigurationManager.ConnectionStrings["SQLConnectionString"].ConnectionString;

        public string GetCustomer(string PONo)
        {   
            string mySQL = string.Empty;
            string customner = string.Empty;
            mySQL = $@"SELECT [X_CONSIGNEE_NAME]
                     FROM [ENVIETNAMPO].[dbo].[_TBL_PO_H]
                     WHERE X_PO_NO = '{PONo}'";
            using (SqlConnection conn = new SqlConnection(connectStr))
            {
                using (SqlCommand cmd = new SqlCommand(mySQL, conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            customner = dt.Rows[0]["X_CONSIGNEE_NAME"].ToString();
                        }
                    }
                }
            }
            return customner;
        }

        public string GetFirstPart(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            string[] parts = input.Split('_');
            return parts.Length > 0 ? parts[0] : string.Empty;
        }

        public string GetContractNo(string PONo)
        {
            string mySQL = string.Empty;
            string contracNo = string.Empty;
            mySQL = $@"SELECT [X_CONTRACT_NO]
                     FROM [ENVIETNAMPO].[dbo].[_TBL_PO_H]
                     WHERE X_PO_NO = '{PONo}'";
            using (SqlConnection conn = new SqlConnection(connectStr))
            {
                using (SqlCommand cmd = new SqlCommand(mySQL, conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            contracNo = dt.Rows[0]["X_CONTRACT_NO"].ToString();
                        }
                    }
                }
            }
            return contracNo;
        }

        public string GetXNote1(string PONo)
        {
            string mySQL = string.Empty;
            string xNote1 = string.Empty;
            mySQL = $@"SELECT [X_NOTE1]
                     FROM [ENVIETNAMPO].[dbo].[_TBL_PO_H]
                     WHERE X_PO_NO = '{PONo}'";
            using (SqlConnection conn = new SqlConnection(connectStr))
            {
                using (SqlCommand cmd = new SqlCommand(mySQL, conn))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                            xNote1 = dt.Rows[0]["X_NOTE1"].ToString();
                        }
                    }
                }
            }
            return xNote1;
        }

        public string GetUnitByItemCode(string itemCode)
        {
            string unit = "PCS";
            string query = "SELECT TOP 1 _UNIT FROM [ENVNDIVDB].[dbo].[TBL_MST_GOODS] WHERE _GOODS_CD = @ItemCode";
            using (SqlConnection conn = new SqlConnection(connectStr))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@ItemCode", itemCode);
                conn.Open();
                var result = cmd.ExecuteScalar();
                if (result != null)
                {
                    unit = result.ToString();
                }
            }
            return unit;
        }
    }
}
