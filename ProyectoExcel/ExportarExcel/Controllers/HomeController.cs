using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.Data.SqlClient;
using System.Data;
using ExportarExcel.Models;
using System.Text;
using ClosedXML.Excel;
using System.IO;

namespace ExportarExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public FileResult exportar(string empleado,string fechainicio, string fechafin) {

            DataTable dt = new DataTable();
            using (SqlConnection cn = new SqlConnection("Data Source=.;Initial Catalog=Northwind;Integrated Security=True"))
            {
                StringBuilder consulta = new StringBuilder();
                consulta.AppendLine("SET DATEFORMAT dmy;");
                consulta.AppendLine("select * from [dbo].[Orders] where employeeID = iif(@employee =0,employeeID,@employee) and convert(date,OrderDate) between @fechainicio and @fechafin");
                

                SqlCommand cmd = new SqlCommand(consulta.ToString(), cn);
                cmd.Parameters.AddWithValue("@employee", empleado);
                cmd.Parameters.AddWithValue("@fechainicio", fechainicio);
                cmd.Parameters.AddWithValue("@fechafin", fechafin);
                cmd.CommandType = CommandType.Text;

                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                {
                    da.Fill(dt);
                }
            }

            dt.TableName = "Datos";

            using (XLWorkbook libro = new XLWorkbook())
            {
                var hoja = libro.Worksheets.Add(dt);

                hoja.ColumnsUsed().AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    libro.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Reporte " + DateTime.Now.ToString() + ".xlsx");
                }
            }

            
        }


        public JsonResult obtenerEmpleado() {
            List<Empleado> listaEmpleado = new List<Empleado>();

            using (SqlConnection cn = new SqlConnection("Data Source=.;Initial Catalog=Northwind;Integrated Security=True")) {
                SqlCommand cmd = new SqlCommand("select EmployeeID,concat(FirstName,' ',LastName)[Nombres] from [dbo].[Employees]",cn);
                cmd.CommandType = CommandType.Text;
                cn.Open();
                using (SqlDataReader dr = cmd.ExecuteReader()) {
                    while (dr.Read()) {
                        listaEmpleado.Add(new Empleado()
                        {
                            _IdEmpleado = Convert.ToInt32(dr["EmployeeID"]),
                            _Nombres = dr["Nombres"].ToString()
                        });
                    }
                }
            }

            return Json(listaEmpleado ,JsonRequestBehavior.AllowGet);
        }


    }
}