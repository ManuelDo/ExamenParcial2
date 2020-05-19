using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace API_REST_Northwind_2.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    [RoutePrefix("v2/Analysis/NorthwindCube")]
    public class Northwind_Controller : ApiController
    {

        [HttpGet]
        [Route("Histograma/{dim}/{order}")]
        public HttpResponseMessage Histograma(string dim, string order = "DESC")
        {
            string dimension;

            List<string> clients = new List<string>();
            List<string> employees = new List<string>();
            List<string> products = new List<string>();
            List<int> years = new List<int>();


            switch (dim)
            {
                case "Cliente":
                    dimension = "[Dim Cliente].[Dim Cliente Nombre].CHILDREN";
                    break;
                case "Producto":
                    dimension = "[Dim Producto].[Dim Producto Nombre].CHILDREN";
                    break;
                case "Empleado":
                    dimension = "[Dim Empleado].[Dim Empleado Nombre].CHILDREN";
                    break;
                case "Año":
                    dimension = "[Dim Tiempo].[Dim Tiempo Año].CHILDREN";
                    break;
                default:
                    dimension = "[Dim Cliente].[Dim Cliente Nombre].CHILDREN";
                    break;
            }

            string WITH = @"
                WITH 
                SET [TopVentas] AS 
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Ventas], " + order +
                    @")
                )
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Ventas]
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                {
                    HEAD([TopVentas], 10)
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DHW Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result2;

            switch (dim)
            {
                case "Cliente":
                    result2 = new{
                        datosDimension = clients,
                        datosVenta = ventas,
                        datosTabla = lstTabla
                    };
                    break;
                case "Producto":
                    result2 = new
                    {
                        datosDimension = products,
                        datosVenta = ventas,
                        datosTabla = lstTabla
                    };
                    break;
                case "Empleado":
                    result2 = new
                    {
                        datosDimension = employees,
                        datosVenta = ventas,
                        datosTabla = lstTabla
                    };
                    break;
                case "Año":
                    result2 = new
                    {
                        datosDimension = years,
                        datosVenta = ventas,
                        datosTabla = lstTabla
                    };
                    break;
                default:
                    result2 = new
                    {
                        datosDimension = clients,
                        datosVenta = ventas,
                        datosTabla = lstTabla
                    };
                    break;
            }

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", dimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            switch (dim)
                            {
                                case "Cliente":
                                    clients.Add(dr.GetString(0));
                                    break;
                                case "Producto":
                                    products.Add(dr.GetString(0));
                                    break;
                                case "Empleado":
                                    employees.Add(dr.GetString(0));
                                    break;
                                case "Año":
                                    years.Add(dr.GetInt32(0));
                                    break;
                                default:
                                    clients.Add(dr.GetString(0));
                                    break;
                            }
                            
                            ventas.Add(Math.Round(dr.GetDecimal(1)));

                            dynamic objTabla = new
                            {
                                descripcion = dr.GetString(0),
                                valor = Math.Round(dr.GetDecimal(1))
                            };

                            lstTabla.Add(objTabla);
                        }
                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result2);
        }

        [HttpGet]
        [Route("GetItemsByDimension/{dim}/{order}")]
        public HttpResponseMessage GetItemsByDimension(string dim, string order)
        {

            string WITH = @"
                WITH 
                SET [OrderDimension] AS 
                NONEMPTY(
                    ORDER(
                        {0}.CHILDREN,
                        {0}.CURRENTMEMBER.MEMBER_NAME, " + order +
                    @")
                )
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Ventas]
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                    [OrderDimension]
                ON ROWS
            ";

            string CUBO_NAME = "[DHW Northwind]";
            WITH = string.Format(WITH, dim);
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();

            dynamic result = new
            {
                datosDimension = dimension
            };

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", dim);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                        }
                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpPost]
        [Route("GetDataPieByDimension")]
        public HttpResponseMessage GetDataPieByDimension(string[] values)
        {
            string WITH = @"
                WITH 
                SET [OrderDimension] AS 
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Ventas], DESC
                    )
                )
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Ventas]
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                {
                    [OrderDimension]
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DHW Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = dimension,
                datosVenta = ventas,
                datosTabla = lstTabla
            };

            string valoresDimension = string.Empty;
            foreach (var item in values)
            {
                if(item == "1996" || item == "1997" || item == "1998")
                {
                    valoresDimension += "[Dim Tiempo].[Dim Tiempo Año].[" + item + "],";
                }
                else if (item == "Enero" || item == "Febrero" || item == "Marzo" || item == "Abril" || item == "Mayo" || item == "Junio" || item == "Julio" 
                    || item == "Agosto" || item == "Septiembre" || item == "Octubre" || item == "Noviembre" || item == "Diciembre")
                {
                    valoresDimension += "[Dim Tiempo].[Dim Tiempo Mes].[" + item + "],";
                }
                else
                {
                    valoresDimension += "[Dim Cliente].[Dim Cliente Nombre].[" + item + "],";
                }
            }

            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension);
            valoresDimension = @"{" + valoresDimension + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                            ventas.Add(Math.Round(dr.GetDecimal(1)));

                            dynamic objTabla = new {
                                descripcion = dr.GetString(0),
                                valor = Math.Round(dr.GetDecimal(1))
                            };

                            lstTabla.Add(objTabla);
                        }

                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }

        [HttpPost]
        [Route("GetDataBar")]
        public HttpResponseMessage GetDataBar(string[] values)
        {
            string WITH = @"
                WITH 
                SET [OrderDimension] AS 
                NONEMPTY(
                    ORDER(
                        STRTOSET(@Dimension),
                        [Measures].[Ventas], DESC
                    )
                )
            ";

            string COLUMNS = @"
                NON EMPTY
                {
                    [Measures].[Ventas]
                }
                ON COLUMNS,    
            ";

            string ROWS = @"
                NON EMPTY
                {
                    [OrderDimension]
                }
                ON ROWS
            ";

            string CUBO_NAME = "[DHW Northwind]";
            string MDX_QUERY = WITH + @"SELECT " + COLUMNS + ROWS + " FROM " + CUBO_NAME;

            Debug.Write(MDX_QUERY);

            List<string> dimension = new List<string>();
            List<decimal> ventas = new List<decimal>();
            List<dynamic> lstTabla = new List<dynamic>();

            dynamic result = new
            {
                datosDimension = dimension,
                datosVenta = ventas,
                datosTabla = lstTabla
            };

            string valoresDimension = string.Empty;
            foreach (var item in values)
            {
                if (item == "1996" || item == "1997" || item == "1998")
                {
                    valoresDimension += "[Dim Tiempo].[Dim Tiempo Año].[" + item + "],";
                }
                else if (item == "Enero" || item == "Febrero" || item == "Marzo" || item == "Abril" || item == "Mayo" || item == "Junio" || item == "Julio"
                    || item == "Agosto" || item == "Septiembre" || item == "Octubre" || item == "Noviembre" || item == "Diciembre")
                {
                    valoresDimension += "[Dim Tiempo].[Dim Tiempo Mes].[" + item + "],";
                }
                else
                {
                    valoresDimension += "[Dim Cliente].[Dim Cliente Nombre].[" + item + "],";
                }
            }

            valoresDimension = valoresDimension.TrimEnd(',');
            valoresDimension = string.Format(valoresDimension);
            valoresDimension = @"{" + valoresDimension + "}";

            using (AdomdConnection cnn = new AdomdConnection(ConfigurationManager.ConnectionStrings["CuboNorthwind"].ConnectionString))
            {
                cnn.Open();
                using (AdomdCommand cmd = new AdomdCommand(MDX_QUERY, cnn))
                {
                    cmd.Parameters.Add("Dimension", valoresDimension);
                    using (AdomdDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        while (dr.Read())
                        {
                            dimension.Add(dr.GetString(0));
                            ventas.Add(Math.Round(dr.GetDecimal(1)));

                            dynamic objTabla = new
                            {
                                descripcion = dr.GetString(0),
                                valor = Math.Round(dr.GetDecimal(1))
                            };

                            lstTabla.Add(objTabla);
                        }

                        dr.Close();
                    }
                }
            }

            return Request.CreateResponse(HttpStatusCode.OK, (object)result);
        }
    }
}
