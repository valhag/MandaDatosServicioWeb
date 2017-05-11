using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using Microsoft.Win32;
using System.Configuration;
using System.IO;
using System.Data;
//using MyExcel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace MandaDatosServicioWeb
{

    public class reg
    {
        public string site;
        public string batch;
        public string name;
        public string phone;
        public string email;
        public string RFC;
        public string Activation;
        public string Tipo;
    }

    public class RegProducto
    {
        public int IdProducto;
        public string CodigoProducto;
        public string NombreProducto;
        public double Existencia;
        public double EntradasPeriodo;
        public double SalidasPeriodo;
        public int metodocosteo;
        public string Clasif1;
        public string Clasif2;
        public string Clasif3;
        public string Clasif4;
        public string Clasif5;
        public string Clasif6;
        
    }

    public class RegCapas
    {
        public int IdProducto;
        public string Fecha;
        public decimal ExistenciaInicial;
        public decimal ExistenciaEntradasPeriodo;
        public decimal ExistenciaSalidasPeriodo;
        public decimal ExistenciaFinal;
        public long idcapae;
        public long idcapas;
        public long idcapa;
        public decimal costo;
        public string  almacen;
    }

    public class RegConcepto
    {
        private string _Codigo;

        public string Codigo
        {
            get { return _Codigo; }
            set { _Codigo = value; }
        }
        private string _Nombre;

        public string Nombre
        {
            get { return _Nombre; }
            set { _Nombre = value; }
        }
        private string _sTipocfd;

        public string Tipocfd
        {
            get { return _sTipocfd; }
            set { _sTipocfd = value; }
        }
        private long _id;

        public long id
        {
            get { return _id; }
            set { _id = value; }
        }

    }

    
    class Class1
    {
        public string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ";
        public OleDbConnection _conexion;
        private DataTable DatosFacturaAbono = null;

        public List<RegConcepto> _RegClasificaciones = new List<RegConcepto>();

        List<RegProducto> listaprods = new List<RegProducto>();
        List<RegCapas> listacapas = new List<RegCapas>();
        List<RegCapas> sortedlist = new List<RegCapas>();

        public DataSet Datos = null;
        
        public class RegEmpresa
        {
            private string _Nombre;

            public string Nombre
            {
                get { return _Nombre; }
                set { _Nombre = value; }
            }
            private string _Ruta;

            public string Ruta
            {
                get { return _Ruta; }
                set { _Ruta = value; }
            }
        }
        public class RegConcepto
        {
            private string _Codigo;

            public string Codigo
            {
                get { return _Codigo; }
                set { _Codigo = value; }
            }
            private string _Nombre;

            public string Nombre
            {
                get { return _Nombre; }
                set { _Nombre = value; }
            }
            private string _sTipocfd;

            public string Tipocfd
            {
                get { return _sTipocfd; }
                set { _sTipocfd = value; }
            }
            private long _id;

            public long id
            {
                get { return _id; }
                set { _id = value; }
            }

        }

        public List<RegEmpresa> mCargarEmpresas(out string amensaje)
        {

            OleDbConnection lconexion = new OleDbConnection();

            lconexion = mAbrirRutaGlobal(out amensaje);

            List<RegEmpresa> _RegEmpresas = new List<RegEmpresa>();
            //amensaje = lconexion.ConnectionString;

            if (amensaje == "")
            {
                //lconexion = miconexion.mAbrirConexionDestino();
                try
                {

                    OleDbCommand lsql = new OleDbCommand("select cnombree01,crutadatos from mgw00001 where cidempresa > 1 ", lconexion);
                    OleDbDataReader lreader;
                    //long lIdDocumento = 0;
                    lreader = lsql.ExecuteReader();
                    _RegEmpresas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegEmpresa lRegEmpresa = new RegEmpresa();
                            lRegEmpresa.Nombre = lreader[0].ToString();
                            lRegEmpresa.Ruta = lreader[1].ToString();
                            _RegEmpresas.Add(lRegEmpresa);
                        }
                    }
                    lreader.Close();

                }
                catch (Exception eeeee)
                {
                    amensaje = eeeee.Message;
                }

            }



            return _RegEmpresas;




        }
        public OleDbConnection mAbrirRutaGlobal(out string amensaje)
        {
            amensaje = "";
            RegistryKey hklp = Registry.LocalMachine;
            hklp = hklp.OpenSubKey(llaveregistry);
            Object obc = hklp.GetValue("DIRECTORIODATOS");
            if (obc == null)
            {
                amensaje = "No existe instalacion de Adminpaq en este computadora";
                return null;
            }
            _conexion = new OleDbConnection();
            _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + obc.ToString();
            try
            {
                _conexion.Open();
            }
            catch (Exception eeee)
            {
                amensaje = eeee.Message;
            }
            return _conexion;

        }
        //public MyExcel.Workbook mIniciarExcel()
        //{
        //    MyExcel.Application excelApp = new MyExcel.Application();
        //    excelApp.Visible = true;
        //    MyExcel.Workbook newWorkbook = excelApp.Workbooks.Add();
        //    newWorkbook.Worksheets.Add();
        //    return newWorkbook;

        //}


        public void mTraerDataset(List<string> lquery, string mEmpresa)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen(mEmpresa);
            DataSet ds = new DataSet();
            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            string nombretabla = "Tabla";
            int indice =1;
            foreach (string lista in lquery)
            {
                OleDbCommand mySqlCommand = new OleDbCommand(lista, lconexion);
                mySqlDataAdapter.SelectCommand = mySqlCommand;
                mySqlDataAdapter.Fill(ds,nombretabla + indice.ToString() );
                indice++;
            }


            //connection.Open();
            //oledbAdapter = new OleDbDataAdapter(firstSql, connection);
            //oledbAdapter.Fill(ds, "First Table");
            //oledbAdapter.SelectCommand.CommandText = secondSql;
            //oledbAdapter.Fill(ds, "Second Table");
            //oledbAdapter.Dispose();
            //connection.Close();

            Datos = ds;

        }


        public void mTraerInformacionPrimerReporte(string lquery, string mEmpresa)
        {
            OleDbConnection lconexion = new OleDbConnection();
            lconexion = mAbrirConexionOrigen(mEmpresa);

            OleDbCommand mySqlCommand = new OleDbCommand(lquery, lconexion);


            DataSet ds = new DataSet();

            OleDbDataAdapter mySqlDataAdapter = new OleDbDataAdapter();
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);

            DatosFacturaAbono = ds.Tables[0];

        }

        
        public OleDbConnection mAbrirConexionOrigen(string mEmpresa)
        {
            _conexion = null;
            string rutaorigen = mEmpresa;
            if (rutaorigen != "c:\\" && rutaorigen != "VentasPorConcepto.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion = new OleDbConnection();
                _conexion.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen;
                _conexion.Open();
            }
            return _conexion;

        }

        
        

        




        

        

        

        private void mConfigurarObjetosImpresion()
        {
            DataTable Productos = Datos.Tables[0];
            DataTable InventarioInicialentradas = Datos.Tables[1];
            DataTable InventarioInicialsalidas = Datos.Tables[2];
            DataTable capasinicialentradas = Datos.Tables[3];
            DataTable capasinicialsalidas = Datos.Tables[4];
            DataTable Movimientosentradas = Datos.Tables[5];
            DataTable Movimientossalidas = Datos.Tables[6];
            DataTable capasenperiodoentradas = Datos.Tables[7];
            DataTable capasenperiodosalidas = Datos.Tables[8];

            listaprods.Clear();
            listacapas.Clear();

            DataRow xxx;
            DataTable table1 = new DataTable();
            table1.Columns.Add("uno", typeof(double));
            table1.Columns.Add("dos", typeof(double));
            table1.Columns.Add("tres", typeof(double));
            table1.Columns.Add("cuatro", typeof(double));
            table1.Columns.Add("cinco", typeof(double));
            table1.Columns.Add("seis", typeof(double));
            table1.Columns.Add("siete", typeof(double));
            xxx = table1.Rows.Add(0, 0, 0,0,0,0,0);

            var inventarioinicial = from p in Productos.AsEnumerable()
                                    //from e in InventarioInicialentradas.AsEnumerable()
                                    join e in InventarioInicialentradas.AsEnumerable() on (string)p["idprodue"].ToString() equals (string)e["idprodue"].ToString() into tempp
                                    from e1 in tempp.DefaultIfEmpty(xxx)
                                    join s in InventarioInicialsalidas.AsEnumerable() on (string)e1[0].ToString() equals (string)s["idprodus"].ToString() into temp
                                    from s1 in temp.DefaultIfEmpty(xxx)
                                    join me in Movimientosentradas.AsEnumerable() on (string)s1[0].ToString() equals (string)me["idprodue"].ToString() into temp1
                                    from move in temp1.DefaultIfEmpty(xxx)
                                    join ms in Movimientossalidas.AsEnumerable() on (string)s1[0].ToString() equals (string)ms["idprodus"].ToString() into temp2
                                    from movs in temp2.DefaultIfEmpty(xxx)
                                    select new
                                    {
                                        //Id = e.Field<decimal>(0).ToString(),
                                        Id = p.Field<decimal>(0).ToString(),
                                        //Nombre = e.Field<string>(2).ToString(),
                                        Nombre = p.Field<string>(2).ToString(),
                                        Salidas = s1.Field<double>(4).ToString() ?? string.Empty,
                                        //Entradas = e.Field<double>(4),
                                        Entradas = e1.Field<double>(1),
                                        Codigo = p.Field<string>(1).ToString(),
                                        Metodo = p.Field<decimal>(3).ToString(),
                                        MovEntradas = move.Field<double>(1),
                                        MovSalidas = movs.Field<double>(1),
                                        clasif1 = p.Field<string>(4).ToString(),
                                        clasif2 = p.Field<string>(5).ToString(),
                                        clasif3 = p.Field<string>(6).ToString(),
                                        clasif4 = p.Field<string>(7).ToString(),
                                        clasif5 = p.Field<string>(8).ToString(),
                                        clasif6 = p.Field<string>(9).ToString()
                                    };
            //Codigo = e.Field<string>(1).ToString()
            //Metodo = e.Field<int>(3).ToString()
            string nombre = "";
            double saldo = 0;
            int cuantosprods = 0;
            foreach (var saldos in inventarioinicial)
            {
                RegProducto lprod = new RegProducto();
                nombre = saldos.Nombre;
                lprod.IdProducto = int.Parse(saldos.Id);
                lprod.NombreProducto = saldos.Nombre;
                lprod.CodigoProducto = saldos.Codigo;
                lprod.metodocosteo = int.Parse(saldos.Metodo);
                saldo = saldos.Entradas - double.Parse(saldos.Salidas);
                lprod.Existencia = saldo;
                lprod.EntradasPeriodo = saldos.MovEntradas;
                lprod.SalidasPeriodo = saldos.MovSalidas;
                lprod.Clasif1 = saldos.clasif1;
                lprod.Clasif2 = saldos.clasif2;
                lprod.Clasif3 = saldos.clasif3;
                lprod.Clasif4 = saldos.clasif4;
                lprod.Clasif5 = saldos.clasif5;
                lprod.Clasif6 = saldos.clasif6;
                listaprods.Add(lprod);

            }

            //and (string)e["cidcapa"].ToString() equals (string)s["cidcapa"].ToString() into UP

            DataRow zz;
            DataTable table = new DataTable();
            table.Columns.Add("uno", typeof(double));
            table.Columns.Add("dos", typeof(double));
            table.Columns.Add("tres", typeof(double));
            table.Columns.Add("cuatro", typeof(double));
            table.Columns.Add("cinco", typeof(double));
            table.Columns.Add("seis", typeof(double));
            table.Columns.Add("siete", typeof(double));
            table.Columns.Add("ocho", typeof(double));
            table.Columns.Add("nueve", typeof(double));
            table.Columns.Add("diez", typeof(double));
            table.Columns.Add("once", typeof(double));
            zz = table.Rows.Add(0, 0, 0, 0, 0, 0,0, 0, 0,0, 0);

            var capasinicial = from capaseinicial in capasinicialentradas.AsEnumerable()
                               join capassinicial in capasinicialsalidas.AsEnumerable() on
                               new 
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassinicial["cidprodu01"].ToString(),
                                   cidcapa = capassinicial["cidcapa"].ToString()
                               } into temp
                               from s1 in temp.DefaultIfEmpty(zz)
                               join capassalidasperiodo in capasenperiodosalidas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassalidasperiodo["cidprodu01"].ToString(),
                                   cidcapa = capassalidasperiodo["cidcapa"].ToString()
                               } into temp2
                               from s2 in temp2.DefaultIfEmpty(zz)
                               /*join capasentradasperiodo in capasenperiodoentradas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   //cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                                * 
                               equals
                               new
                               {
                                   cidprodu01 = capasentradasperiodo["cidprodu01"].ToString(),
                                   //cidcapa = capasentradasperiodo["cidcapa"].ToString()
                               } into temp3
                               from s3 in temp3.DefaultIfEmpty(zz)*/
                               select new
                               {
                                   cidprodu = capaseinicial.Field<decimal>(0),
                                   cidcapa = capaseinicial.Field<decimal>(1),
                                   cfecha = capaseinicial.Field<string>(2),
                                   unidadesentrada = capaseinicial.Field<double>(3),
                                   unidadessalida = s1.Field<double>(2).ToString() ?? string.Empty,
                                   unidadesperiodosalida = s2.Field<double>(2).ToString() ?? string.Empty,
                                   costo = capaseinicial.Field<double>(4),
                                   almacen = capaseinicial.Field<string>(5)
                               };
            decimal existenciacapae = 0;
            decimal existenciacapas = 0;

            foreach (var capa in capasinicial)
            {
                RegCapas capalocal = new RegCapas();
                capalocal.Fecha = capa.cfecha;
                capalocal.IdProducto = int.Parse(capa.cidprodu.ToString());
                existenciacapae = decimal.Parse(capa.unidadesentrada.ToString()); //- capas.unidadessalida;

                existenciacapas = decimal.Parse(capa.unidadessalida.ToString()); //- capas.unidadessalida;
                capalocal.ExistenciaInicial = existenciacapae - existenciacapas;
                //capalocal.ExistenciaEntradasPeriodo = decimal.Parse(capa.unidadesperiodoentrada);
                //if (capalocal.ExistenciaEntradasPeriodo > 0)
                   // capalocal.ExistenciaInicial = 0;
                capalocal.ExistenciaSalidasPeriodo = decimal.Parse(capa.unidadesperiodosalida);
                capalocal.ExistenciaFinal = capalocal.ExistenciaEntradasPeriodo + capalocal.ExistenciaSalidasPeriodo;
                //capalocal.idcapae = long.Parse(capa.idcapae);
                //capalocal.idcapas = long.Parse(capa.idcapas);
                capalocal.idcapa = long.Parse(capa.cidcapa.ToString());
                capalocal.costo = decimal.Parse(capa.costo.ToString());
                capalocal.almacen = capa.almacen;
                if (existenciacapae - existenciacapas != 0)
                        listacapas.Add(capalocal);
            }
            var capassoloentrdas = from capaseinicial in capasenperiodoentradas.AsEnumerable()
                                   join capassalidasperiodo in capasenperiodosalidas.AsEnumerable() on
                               new
                               {
                                   cidprodu01 = capaseinicial["cidprodu01"].ToString(),
                                   cidcapa = capaseinicial["cidcapa"].ToString()
                               }
                               equals
                               new
                               {
                                   cidprodu01 = capassalidasperiodo["cidprodu01"].ToString(),
                                   cidcapa = capassalidasperiodo["cidcapa"].ToString()
                               } into temp2
                                   from s2 in temp2.DefaultIfEmpty(zz)
                                   select new
                                   {
                                       cidprodu = capaseinicial.Field<decimal>(0),
                                       cidcapa = capaseinicial.Field<decimal>(1),
                                       //cfecha = capaseinicial.Field<string>(2),
                                       unidadesentrada = 0,
                                       unidadesperiodoentrada = capaseinicial.Field<double>(2),
                                       costo = capaseinicial.Field<double>(3),
                                       unidadesperiodosalida = s2.Field<double>(2).ToString() ?? string.Empty,
                                       almacen = capaseinicial.Field<string>(4)
                                   };
            foreach (var capa in capassoloentrdas)
            {
                RegCapas capalocal = new RegCapas();
                //capalocal.Fecha = capa.cfecha;
                capalocal.IdProducto = int.Parse(capa.cidprodu.ToString());
                existenciacapae = decimal.Parse(capa.unidadesentrada.ToString()); //- capas.unidadessalida;
                existenciacapas = 0; //- capas.unidadessalida;
                capalocal.ExistenciaInicial = existenciacapae - existenciacapas;
                capalocal.ExistenciaEntradasPeriodo = decimal.Parse(capa.unidadesperiodoentrada.ToString());
                capalocal.ExistenciaSalidasPeriodo = decimal.Parse(capa.unidadesperiodosalida);
                capalocal.ExistenciaFinal = capalocal.ExistenciaEntradasPeriodo + capalocal.ExistenciaSalidasPeriodo;
                capalocal.costo = decimal.Parse(capa.costo.ToString());
                capalocal.almacen = capa.almacen;
                    listacapas.Add(capalocal);
            }
            listacapas = listacapas.OrderBy(o => o.IdProducto).ToList();
            //sortedlist.Clear();
            //sortedlist = listacapas.OrderBy(o => o.IdProducto).ToList();
            //listacapas.Clear();
            //listacapas = sortedlist;

        }

        



        
        public List<RegConcepto> mCargarConceptos(string mEmpresa)
        {
             List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            if (mEmpresa.IndexOf("\\") != -1)
            {
                OleDbConnection lconexion = new OleDbConnection();

                lconexion = mAbrirConexionOrigen(mEmpresa);
                
                if (lconexion != null)
                {
                    OleDbCommand lsql = new OleDbCommand("select cidconce01,ccodigoc01,cnombrec01 from mgw10006 where ciddocum01 = 4", lconexion);
                    OleDbDataReader lreader;
                    lreader = lsql.ExecuteReader();
                    _RegFacturas.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegConcepto lRegConcepto = new RegConcepto();
                            lRegConcepto.Codigo = lreader[1].ToString();
                            lRegConcepto.Nombre = lreader[2].ToString();
                            lRegConcepto.id = long.Parse(lreader[0].ToString());
                            _RegFacturas.Add(lRegConcepto);
                        }
                    }
                    lreader.Close();
                }
            }
            return _RegFacturas;
        }

        public void mCargarClasificaciones(string mEmpresa, int clasificacion)
        {

            int clasif = clasificacion + 24;
            //List<RegConcepto> _RegFacturas = new List<RegConcepto>();
            _RegClasificaciones.Clear();
            if (mEmpresa.IndexOf("\\") != -1)
            {
                OleDbConnection lconexion = new OleDbConnection();

                lconexion = mAbrirConexionOrigen(mEmpresa);

                if (lconexion != null)
                {
                    OleDbCommand lsql = new OleDbCommand("select cidvalor01,ccodigov01,cvalorcl01 from mgw10020 where cidclasi01 = " + clasif, lconexion);
                    OleDbDataReader lreader = null ;
                    try
                    {
                        lreader = lsql.ExecuteReader();
                    }
                    catch (Exception eeee)
                    { 

                    }
                    _RegClasificaciones.Clear();
                    if (lreader.HasRows)
                    {
                        while (lreader.Read())
                        {
                            RegConcepto lRegConcepto = new RegConcepto();
                            lRegConcepto.Codigo = lreader[1].ToString();
                            lRegConcepto.Nombre = lreader[2].ToString();
                            lRegConcepto.id = long.Parse(lreader[0].ToString());
                            _RegClasificaciones.Add(lRegConcepto);
                        }
                    }
                    lreader.Close();
                }
            }
            //return _RegClasificaciones;
        }

        public void mBorraElememento(RegConcepto clasif)
        {
            _RegClasificaciones.Remove(clasif);
        }

        

        public string mRegresarCatalogoValido(int tipo,  string codigo, string mEmpresa)
        {
            OleDbConnection lconexion = new OleDbConnection();
            string regresa = "";
            lconexion = mAbrirConexionOrigen(mEmpresa);
            OleDbCommand lsql = new OleDbCommand();
            if (lconexion != null)
            {
                if (tipo == 2) // proveedores
                    lsql.CommandText = "select cidclien01,crazonso01 from mgw10002 where ccodigoc01 = '" + codigo + "' and ctipocli01 >= 1";
                lsql.Connection = lconexion;
                OleDbDataReader lreader;
                lreader = lsql.ExecuteReader();
                
                if (lreader.HasRows)
                {
                    lreader.Read();
                    {
                        if (lreader[0].ToString() != "")
                        {
                            regresa = lreader[1].ToString();
                        }
                    }
                }
                lreader.Close();
                
            }
            return regresa;

        }

    }

}
