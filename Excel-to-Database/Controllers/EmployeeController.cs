using Excel_to_Database.Controllers.Model;
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
    public class EmployeeController : ControllerBase
    {
        public readonly IConfiguration _configuration;

        public EmployeeController(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        [HttpGet]
        [Route("GetAllEmployees")]

        public string GetEmployees()
        {
            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("EmployeeAppCon").ToString());
            SqlDataAdapter da = new SqlDataAdapter("select * from performances", con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            List<Employee> employeeList = new List<Employee>();
            Response response = new Response();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Employee employee = new Employee();
                    employee.id = Convert.ToInt32(dt.Rows[i]["id"]);
                    employee.branch_manager = Convert.ToString(dt.Rows[i]["branch_manager"]);
                    employee.field1 = Convert.ToString(dt.Rows[i]["field1"]);
                    employee.field2 = Convert.ToString(dt.Rows[i]["field2"]);
                    employee.field3 = Convert.ToString(dt.Rows[i]["field3"]);
                    employee.field4 = Convert.ToString(dt.Rows[i]["field4"]);
                    employee.field5 = Convert.ToString(dt.Rows[i]["field5"]);
                    employee.field6 = Convert.ToString(dt.Rows[i]["field6"]);
                    employee.field7 = Convert.ToString(dt.Rows[i]["field7"]);
                    employee.field8 = Convert.ToString(dt.Rows[i]["field8"]);
                    employee.field9 = Convert.ToString(dt.Rows[i]["field9"]);
                    employee.field10 = Convert.ToString(dt.Rows[i]["field10"]);
                    employeeList.Add(employee);
                }
            }
            if (employeeList.Count > 0)
            {
               return JsonConvert.SerializeObject(employeeList); 


            }
            else
            {
                response.StatusCode = 100;
                response.ErrorMessage = "No Data Found";
                return JsonConvert.SerializeObject(response);
            }
        }

    }
}
