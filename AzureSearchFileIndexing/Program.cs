using Azure;

using Azure.Search.Documents;

using Azure.Search.Documents.Indexes;

using Azure.Search.Documents.Indexes.Models;

using Azure.Search.Documents.Models;
using AzureSearchFileIndexing;


using Microsoft.Extensions.Azure;
using Microsoft.Office.Interop.Excel;

using System.Data;
using System.Data.OleDb;

using System.Net;

using System.Net.Sockets;

using DataTable = System.Data.DataTable;



DataSet ds = new DataSet();

string search_service_name = "<Your AI Search Service Name>";

string search_service_apikey = "<Your AI Search Service API key>";

string search_index_name = "<Your AI Search Service indexName Name>";

// Create a Uri SearchIndexClient to to send create/delete index commands

Uri serviceEndpoint = new($"https://{search_service_name}.search.windows.net/");

AzureKeyCredential credential = new AzureKeyCredential(search_service_apikey);




SearchIndexClient admin_client = new SearchIndexClient(serviceEndpoint, credential);

// Create a a SearchClient to load and query documents


SearchClient ingesterClient = admin_client.GetSearchClient(search_index_name);

// Load documents

Console.WriteLine("(0)", "Uploading documents...\n");

UploadDocuments(ingesterClient);

//Load documents

Console.WriteLine("{0}", "Uploading documents...\n"); UploadDocuments(ingesterClient);

// Wait 2 seconds for Indexing to complete before starting queries (for demo and console-app purposes only)

Console.WriteLine("Waiting for indexing...\n ");

System.Threading.Thread.Sleep(2000);

#region Create Index method

//static void CreateIndex(string search_index_name, SearchIndexClient admin_client)
//{

//    FieldBuilder fieldBuilder = new FieldBuilder();

//    var searchFields = fieldBuilder.Build(typeof(SupportDataModel));

//    var definition = new SearchIndex(search_index_name, searchFields);

//    var suggester = new SearchSuggester("sg", new[] { "TicketType", "TicketSubject", "TicketDescription", "TicketStatus", "TicketPriority", "TicketChannel" });
//    definition.Suggesters.Add(suggester);

//    admin_client.CreateOrUpdateIndex(definition);
//}
#endregion

void UploadDocuments(SearchClient searchClient)
{
    //?Use file directory methods to resolve below path
    string filePath = "..data\\customer_support_tickets2.xls";

    string file_excel_connection_string = @"Provider-Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + "; Extended Properties Excel 12.0; Persist Security Info-False";

    try
    {


        using (OleDbConnection conn = new OleDbConnection(file_excel_connection_string))
        {
            conn.Open(); //Added this line
            DataTable data_table_activity = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (data_table_activity != null)

            {



                string worksheetName;

                for (int cnt = 0; cnt < data_table_activity.Rows.Count; cnt++)

                {

                    worksheetName = data_table_activity.Rows[cnt]["TABLE_NAME"].ToString();

                    if (worksheetName.Contains('\''))

                    {

                        worksheetName = worksheetName.Replace('\'', ' ').Trim();
                    }


                    if (worksheetName.EndsWith("$"))
                    {

                        OleDbCommand onlineConnection = new OleDbCommand("SELECT FROM [" + worksheetName + "]", conn);

                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(onlineConnection);

                        myDataAdapter.Fill(ds, "MyTableName");

                        var data = ds.Tables["MyTableName"];

                        for (int i = 0; i < ds.Tables["MyTableName"].Rows.Count; i++)

                        {

                            IndexDocumentsBatch<SupportDataModel> batch = IndexDocumentsBatch.Create(
                                IndexDocumentsAction.Upload(
                                    new SupportDataModel()

                                    {

                                        TicketID = ds.Tables["MyTableName"].Rows[i]["Ticket ID"].ToString(),
                                        TicketType = ds.Tables["MyTableName"].Rows[i]["Ticket Type"].ToString(),
                                        TicketSubject = ds.Tables["MyTableName"].Rows[i]["Ticket Subject"].ToString(),
                                        TicketDescription = ds.Tables["MyTableName"].Rows[i]["Ticket Description"].ToString(),
                                        TicketStatus = ds.Tables["MyTableName"].Rows[1]["Ticket Status"].ToString(),
                                        TicketPriority = ds.Tables["MyTableName"].Rows[i]["Ticket Priority"].ToString(),
                                        TicketChannel = ds.Tables["MyTableName"].Rows[i]["Ticket Channel"].ToString()
                                    }
                                    )

                                );
                            IndexDocumentsResult result = searchClient.IndexDocuments(batch);
                        }
                    }
                }
            }
        }
    }





    catch (Exception ex)
    {
        Console.WriteLine(ex);
    }
}

