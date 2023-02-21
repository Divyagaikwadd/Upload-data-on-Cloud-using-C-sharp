using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;  //OLE DB (Object Linking and Embedding, Database) technology. OLE DB is a Microsoft data access technology
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestAccess1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {   

            //  insert data inside the microsoft access

            OleDbConnection conn = null;
            conn = new OleDbConnection();
            conn.ConnectionString = ConfigurationManager.ConnectionStrings["Connection"].ToString();
            conn.Open();
            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandText = "insert into[Employee](Name1)values(@nm)";
            cmd.Parameters.AddWithValue("@nm", textBox1.Text);
            cmd.Connection = conn;
            int a = cmd.ExecuteNonQuery();
            if (a > 0)
            {
                MessageBox.Show("inserted");
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //code to fech the data from access database to in the list

                                      //pass the path of microsoft access database
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\ADMIN\\Documents\\Database3.accdb";
            OleDbConnection connection = new OleDbConnection(connectionString);

            // Open the connection
            connection.Open();

            // Create a command to retrieve the column names
            OleDbCommand command = new OleDbCommand("SELECT * FROM  Employee", connection);

            // Execute the command and retrieve the column names
            OleDbDataReader reader = command.ExecuteReader();
            List<string> columnNames = new List<string>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                columnNames.Add(reader.GetName(i));
            }

            // Create a list to store the data
            List<List<object>> data = new List<List<object>>();

            // Retrieve the data for each row
            while (reader.Read())
            {
                List<object> row = new List<object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row.Add(reader.GetValue(i));
                }
                data.Add(row);
            }
          



            //  code to connect with mongodb atlas database  uploading the data through list
                                          
            var client = new MongoClient("pass the link of mongodb atlas application with your username & password");

            var database = client.GetDatabase("AccessDB"); //database name on mongodb atlas

            // Create a new collection                           //column name
            var collection = database.GetCollection<BsonDocument>("Employee");

            // Create a list of BsonDocuments to insert
            List<BsonDocument> documents = new List<BsonDocument>();
            foreach (var row in data)
            {
                var document = new BsonDocument();
                for (int i = 0; i < columnNames.Count; i++)
                {
                    document.Add(columnNames[i], BsonValue.Create(row[i]));
                }
                documents.Add(document);
            }

            // Insert the documents into the collection
            collection.InsertMany(documents);
            MessageBox.Show("Successful on cloud");
           





        }
    }
}
