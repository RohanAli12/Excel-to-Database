using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel_to_Database.Controllers.Model;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace practise.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UploadExcel : ControllerBase
    {
        public readonly IConfiguration _configuration;

        public UploadExcel(IConfiguration configuration)
        {
            _configuration = configuration;
        }


        [HttpPost]
        [Route("GetUploadExcel")]

        public async Task<IActionResult> GetUploadExcel(IFormFile excelFile)
        {

            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest("Please select an Excel file.");
            }

            // Read data from the uploaded Excel file
            using (var stream = new MemoryStream())
            {
                await excelFile.CopyToAsync(stream);

                using (var workbook = new XLWorkbook(stream))
                {
                    var worksheet = workbook.Worksheet(1); // Assuming data is on the first worksheet

                    // Configure your ADO.NET connection
                    string connectionString = _configuration.GetConnectionString("EmployeeAppCon").ToString();
                    using (var con = new SqlConnection(connectionString))
                    {
                        con.Open();



                        // Mapping Table info
                        
                        SqlDataAdapter da = new SqlDataAdapter("select * from mapping_info", con);
                        // DataTable mappingTable = new DataTable();
                        // da.Fill(mappingTable);

                        SqlCommand cm = new SqlCommand("select excel_cellid,db_columnname from mapping_info",con);
                        SqlDataReader reader = cm.ExecuteReader();


                        // Dictionary<int, string> maps = new Dictionary<int, string>();


                        int branch_manager_index = 0;
                        int field_1_index = 0;
                        int field_2_index = 0;
                        int field_3_index = 0;
                        int field_4_index = 0;
                        int field_5_index = 0;
                        int field_6_index = 0;
                        int field_7_index = 0;
                        int field_8_index = 0;
                        int field_9_index = 0;
                        int field_10_index = 0;
                        

                        while (reader.Read())
                        {
                            int id = (int)reader["excel_cellid"];
                            string db = (string)reader["db_columnname"];
                            // string combinedValue = (int)reader["excel_cellid"] + " " + (string)reader["db_columnname"];
                            // maps.Add(id, combinedValue);
                            
                            if (db == "Branch_Manager")
                            {
                                branch_manager_index = id;    
                            }else if (db == "field1")
                            {
                                field_1_index = id;
                            }else if (db == "field2")
                            {
                                field_2_index = id;
                            }else if (db == "field3")
                            {
                                field_3_index = id;
                            }else if (db == "field4")
                            {
                                field_4_index = id;
                            }else if (db == "field5")
                            {
                                field_5_index = id;
                            }else if (db == "field6")
                            {
                                field_6_index = id;
                            }else if (db == "field7")
                            {
                                field_7_index = id;
                            }else if (db == "field8")
                            {
                                field_8_index = id;
                            }else if (db == "field9")
                            {
                                field_9_index = id;
                            }else if (db == "field10")
                            {
                                field_10_index = id;
                            }
                        }

                        reader.Close();

                        string map1;
                        string map2;
                        string map3;
                        string map4;
                        string map5;
                        string map6;
                        string map7;
                        string map8;
                        string map9;
                        string map10;
                        string map11;


                  /* foreach (var item in maps)
                        {
                            int id = item.Key;
                            string combined = item.Value;
                            if(id == 1)
                            {
                                map1 = combined;
                            }else if (id == 2)
                            {
                                map2 = combined;
                            }else if (id == 3)
                            {
                                map3 = combined;
                            }else if (id == 4)
                            {
                                map4 = combined;
                            }else if (id == 5)
                            {
                                map5 = combined;
                            }else if (id == 6)
                            {
                                map6 = combined;
                            }else if (id == 7)
                            {
                                map7 = combined;
                            }else if (id == 8)
                            {
                                map8 = combined;
                            }else if (id == 9)
                            {
                                map9 = combined;
                            }else if (id == 10)
                            {
                                map10 = combined;
                            }else if (id == 11)
                            {
                                map11 = combined;
                            }


                        }*/

                        //  Console.WriteLine(mappingTable.Columns[0].ColumnName);
                        // Dictionary<int, string> mappingTable1 = new Dictionary<int, string>();
                        //using (SqlDataReader reader = new(mappingTable)
                        //{ rea

                        //  while (reader.Read())
                        //  {
                        //        mappingTable1.Add(reader["Excel_Column"].ToString(),
                        //                      reader["DB_ColumnName"].ToString());
                        //}

                        //}

                        /* map = mappingTable.Columns[1].DefaultValue == mappingTable.Columns[2].DefaultValue;

                           foreach (DataRow row in mappingTable.Columns[1].Colo)
                           {
                               for (int colIndex = 25; colIndex <= 50; colIndex++)
                               {
                                   Console.WriteLine(row[colIndex]);
                                   Console.WriteLine(myDataTable.Columns[colIndex].ColumnName);
                               }
                           }*/


                        // Check Cell1 Mapping with Which Column like (Branch_Manager)

                        foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip header row
                        {   
                            map1 = row.Cell(branch_manager_index).Value.ToString();
                            map2 = row.Cell(field_1_index).Value.ToString();
                            map3 =row.Cell(field_2_index).Value.ToString();
                            map4 = row.Cell(field_3_index).Value.ToString();
                            map5 = row.Cell(field_4_index).Value.ToString();
                            map6 = row.Cell(field_5_index).Value.ToString();
                            map7 = row.Cell(field_6_index).Value.ToString();
                            map8 = row.Cell(field_7_index).Value.ToString();
                            map9 = row.Cell(field_8_index).Value.ToString();
                            map10 = row.Cell(field_9_index).Value.ToString();
                            map11 = row.Cell(field_10_index).Value.ToString();
                        // Extract other cell values as needed

                        // Define your SQL query to insert data into the database
                        string insertQuery = "INSERT INTO performances VALUES (@Column1, @Column2,@Column3,@Column4,@Column5,@Column6,@Column7,@Column8,@Column9,@Column10,@column11)";

                            using (var cmd = new SqlCommand(insertQuery, con))
                            {
                                cmd.Parameters.AddWithValue("@Column1", map1);
                                cmd.Parameters.AddWithValue("@Column2", map2);
                                cmd.Parameters.AddWithValue("@Column3", map3);
                                cmd.Parameters.AddWithValue("@Column4", map4);
                                cmd.Parameters.AddWithValue("@Column5", map5);
                                cmd.Parameters.AddWithValue("@Column6", map6);
                                cmd.Parameters.AddWithValue("@Column7", map7);
                                cmd.Parameters.AddWithValue("@Column8", map8);
                                cmd.Parameters.AddWithValue("@Column9", map9);
                                cmd.Parameters.AddWithValue("@Column10", map10);
                                cmd.Parameters.AddWithValue("@Column11", map11);
                                cmd.ExecuteNonQuery();
                            }
                        }

                        con.Close();
                    }
                }
            }

            return Ok("Data successfully imported from Excel.");
        }
    }
}

