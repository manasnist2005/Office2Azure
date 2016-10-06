using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office2Azure
{
    public class DBHelper
    {
        public static void SaveData(List<Person> lstPerson)
        {
            Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: *******Data insertion started. *******");
            SqlConnection sqlConn = new SqlConnection(SettingsHelper.AzureDBConnStr);
            try
            { 
                string stmt = "INSERT INTO dbo.Person(Name,Gender,Age,City,Company) VALUES(@Name,@Gender,@Age,@City,@Company)";
                SqlCommand cmd = new SqlCommand(stmt, sqlConn);
            
                cmd.Parameters.Add("@Name", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@Gender", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@Age", SqlDbType.Int);
                cmd.Parameters.Add("@City", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@Company", SqlDbType.VarChar, 50);

                sqlConn.Open();
                foreach (var prsn in lstPerson)
                {                
                    cmd.Parameters["@Name"].Value = prsn.Name;
                    cmd.Parameters["@Gender"].Value = prsn.Gender;
                    cmd.Parameters["@Age"].Value = prsn.Age;
                    cmd.Parameters["@City"].Value = prsn.City;
                    cmd.Parameters["@Company"].Value = prsn.Company;

                    cmd.ExecuteNonQuery();
                    Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: Data inserted for "+ prsn.Name);
                }
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + ":: *******Data insertion completed. *******");

            }
            catch(Exception ex)
            {
                Console.WriteLine(DateTime.UtcNow.ToString("yyyy-MM-dd-HH-mm:ss") + "::Error while data insertion. :" + ex.Message);
            }
            sqlConn.Close();
        }
    }
}
