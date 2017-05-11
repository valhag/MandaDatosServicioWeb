using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Linq;
using System.IO;
using System.Net;
using System.Web;
using System.Net.Mail;
using System.Web.Script.Serialization;
using System.Threading;
using System.Globalization;
using System.Configuration;
using Microsoft.Win32;
using System.Data.OleDb;
using System.Text.RegularExpressions;



namespace MandaDatosServicioWeb
{


    public partial class Form1 : Form
    {

        Class1 x = new Class1();
        string Cadenaconexion;
        string Empresa="";
        public Form1()
        {
            InitializeComponent();
        }


        private void mProcesarDBF(ref DataTable SaldoTotal,ref DataTable Vencido,ref DataTable Detalle, string empresa)
        {

            //string empresa = comboBox1.SelectedValue.ToString().Trim();
            OleDbConnection lconexion2 = mAbrirConexionOrigenvfp(empresa);
            DataSet ds = new DataSet();

            /*
             * string sql1 = "select m2.CIDCLIENTEPROVEEDOR, m2.CRAZONSOCIAL, m11.cemail, sum(cpendiente),  m0.CNOMBREEMPRESA, m0.CRFCEMPRESA " +
                ",m11a.CTELEFONO1,m11a.CNOMBRECALLE + ' ' + m11a.CNUMEROEXTERIOR + ' '  + m11a.CCOLONIA + ' ' + m11a.CESTADO + ' ' + m11a.CPAIS as direccion, " +
                " m11.CNOMBRECALLE, m11.CNUMEROEXTERIOR, m11a.CCOLONIA,m11.CESTADO,m11.ctelefono1, m2.crfc " +
             */




            string sql11 = "select m2.cidclien01 as cidclienteproveedor, m2.CRAZONSO01 as crazonsocial, m11.cemail, sum(cpendiente), m0.CNOMBREE01 as cnombreempresa, m0.CRFCEMPR01 as crfcempresa, " +
                " m11a.CTELEFONO1  ,m11a.CNOMBREC01 + ' ' + m11a.CNUMEROE01 + ' '  + m11a.CCOLONIA + ' ' + m11a.CESTADO as direccion, " +
             " m11.CNOMBREC01 as cnombrecalle, m11.CNUMEROE01 as cnombreexterior, m11a.CCOLONIA, m11.CESTADO,m11.ctelefono1,m2.crfc,m2.cemail1, m2.cemail2,m2.cemail3 " +
             " from MGW10008 m8 " +
             " join mgw10002 m2 on m2.CIDCLIEN01 = m8.CIDCLIEN01 and m2.CTIPOCLI01 <=2 " +
             " join mgw10011 m11 on m11.CIDCATAL01 = m2.CIDCLIEN01 and m11.CTIPOCAT01 =1 and m2.CEMAIL1 <> ' '  " +
             " left join mgw10000 m0 on m0.cidempresa > 0 " +
             " left join mgw10011 m11a on m0.cidempresa = m11a.cidcatal01 and m11a.CTIPOCAT01 = 4" +
             " where m8.CNATURAL01 = 0 and CAFECTADO = 1 " +
             " group by m2.CIDCLIEN01, m2.CRAZONSO01,m11.CEMAIL,  m0.CNOMBREE01, m0.CRFCEMPR01,m11.CNOMBREC01, m11.CNUMEROE01, m11.CESTADO,m11.ctelefono1,m2.crfc,m11a.CCOLONIA, " +
               " m11a.CTELEFONO1  ,m11a.CNOMBREC01, m11a.CCOLONIA, m11a.CNUMEROE01, m11a.CCOLONIA, m11a.CESTADO, m2.cemail1, m2.cemail2, m2.cemail3;";

            OleDbDataAdapter mySqlDataAdapter1 = new OleDbDataAdapter();
            OleDbCommand mySqlCommand = new OleDbCommand(sql11, lconexion2);
            //sql.CommandText = sql1;
            mySqlDataAdapter1.SelectCommand = mySqlCommand;
            mySqlDataAdapter1.Fill(ds);

           SaldoTotal = ds.Tables[0];

           ds = null;
           ds = new DataSet();

            string sql33 = "";
            // saldo vencido total
            sql33 = "select m2.cidclien01 as CIDCLIENTEPROVEEDOR, m2.crazonso01 as CRAZONSOCIAL, m11.cemail, sum(cpendiente)  " +
            " from mgw10008 m8 join mgw10002 m2 on m2.CIDCLIEN01 = m8.CIDCLIEN01 and m2.CTIPOCLI01 <=2 " +
            " join mgw10011 m11 on m11.CIDCATAL01 = m2.CIDCLIEN01 and m11.CTIPOCAT01 =1 and m11.CEMAIL <> ' ' " +
            " where m8.CNATURAL01 = 0 and CAFECTADO = 1 and m8.cpendiente > 0  and date() > m8.CFECHAVE01" +
                //" and m2.ccodigocliente = 'ECV940426G35'" +
            " group by m2.CIDCLIEN01, m2.CRAZONSO01,m11.CEMAIL; ";
            //sql11 += sql33;

            mySqlCommand.CommandText = sql33;
            mySqlDataAdapter1.SelectCommand = mySqlCommand;
            mySqlDataAdapter1.Fill(ds);

            Vencido = ds.Tables[0];

            string sql22;
            // detalle de documentos
            sql22 = "select m8.cidclien01 as CIDCLIENTEPROVEEDOR,  m8.cfolio, m2.crazonso01 as CRAZONSOCIAL,  dtos(m8.cfecha), dtos(m8.cfechave01) as cfechavencimiento, m6.cnombrec01 as CNOMBRECONCEPTO " +
            " ,m1.cnombrea01 as CNOMBREAGENTE,m8.ctotal,m8.cpendiente, m5.cnombrep01 as CNOMBREPRODUCTO, m8.cseriedo01 as cseriedocumento,m7.cdescrip01 as cdescripcion, m10.cobserva01 as cobservaciones, str(m8.cidmoneda) as cidmoneda" +
            " from mgw10008 m8 join mgw10002 m2 on m2.CIDCLIEN01 = m8.CIDCLIEN01 and m2.CTIPOCLI01 <=2 and m2.cemail1 <> ' ' " +
            " join mgw10011 m11 on m11.CIDCATAL01 = m2.CIDCLIEN01 and m11.CTIPOCAT01 =1 " +
            " join mgw10006 m6 on m8.CIDCONCE01 = m6.CIDCONCE01 " +
            " join mgw10007 m7 on m7.CIDDOCUM01 = m6.CIDDOCUM01 " +
            " join mgw10001 m1 on m8.cidagente = m1.cidagente " +
            " join mgw10010 m10 on m10.CIDDOCUM01 = m8.CIDDOCUM01 and m10.CNUMEROM01=100 " +
            " join mgw10005 m5 on m5.CIDPRODU01 = m10.CIDPRODU01 " +
            " where m8.CNATURAL01 = 0 and cafectado = 1 " +
            " and m8.cpendiente > 0 " +
            " and m2.ctipocli01 = 1 " +
                //" and m2.ccodigocliente = 'ECV940426G35'" +
            " order by m8.CIDCLIEN01 ";


            //OleDbDataAdapter mySqlDataAdapter1 = new OleDbDataAdapter();
            //OleDbCommand mySqlCommand = new OleDbCommand(sql11, lconexion2);
            ds = new DataSet();
            mySqlCommand.CommandText = sql22;
            mySqlDataAdapter1.SelectCommand = mySqlCommand;
            mySqlDataAdapter1.Fill(ds);

            Detalle = ds.Tables[0];



            lconexion2.Close();

 
        }

        private void mProcesarSQL(ref DataSet ds, string empresa)
        {

            //ListView.SelectedListViewItemCollection empresas = listView1.SelectedItems;

            
                
            

            SqlConnection lconexion = new SqlConnection();
            //DataSet ds = new DataSet();

          

            SqlCommand sql = new SqlCommand();

            // saldo total
            string sql1 = "select cast(m2.CIDCLIENTEPROVEEDOR as decimal(18,0)) as CIDCLIENTEPROVEEDOR, m2.CRAZONSOCIAL, m11.cemail, sum(cpendiente),  m0.CNOMBREEMPRESA, m0.CRFCEMPRESA " +
                ",m11a.CTELEFONO1,m11a.CNOMBRECALLE + ' ' + m11a.CNUMEROEXTERIOR + ' '  + m11a.CCOLONIA + ' ' + m11a.CESTADO + ' ' + m11a.CPAIS as direccion, " +
                " m11.CNOMBRECALLE, m11.CNUMEROEXTERIOR, m11a.CCOLONIA,m11.CESTADO,m11.ctelefono1, m2.crfc, m2.cemail1,m2.cemail2, m2.cemail3 " +
"from admDocumentos m8 join admClientes m2 on m2.CIDCLIENTEPROVEEDOR = m8.CIDCLIENTEPROVEEDOR and m2.CTIPOCLIENTE <=2 " +
"join admDomicilios m11 on m11.CIDCATALOGO = m2.CIDCLIENTEPROVEEDOR and m11.CTIPOCATALOGO =1 and m11.CEMAIL <> ' ' " +
" left join admDomicilios m11a on m11a.CtipoCATALOGO = 4 " +
"cross join admParametros m0 " +
"where m8.CNATURALEZA = 0 and CAFECTADO = 1 " +
                //"and m2.ccodigocliente = 'ECV940426G35'" +
"group by m2.CIDCLIENTEPROVEEDOR, m2.CRAZONSOCIAL,m11.CEMAIL ,  m0.CNOMBREEMPRESA, m0.CRFCEMPRESA , m11a.CTELEFONO1 " +
" ,m11a.CNOMBRECALLE, m11a.CCOLONIA, m11a.CNUMEROEXTERIOR, m11a.CCOLONIA, m11a.CESTADO,m11a.CPAIS, " +
"m11.CNOMBRECALLE, m11.CNUMEROEXTERIOR, m11a.CCOLONIA,m11.CESTADO,m11.ctelefono1,m2.crfc, m2.cemail1,m2.cemail2, m2.cemail3 ";

            string sql3 = "";
            // saldo vencido total
            sql3 = "select cast(m2.CIDCLIENTEPROVEEDOR as decimal(18,0)) as CIDCLIENTEPROVEEDOR, m2.CRAZONSOCIAL, m11.cemail, sum(cpendiente)  " +
            " from admDocumentos m8 join admClientes m2 on m2.CIDCLIENTEPROVEEDOR = m8.CIDCLIENTEPROVEEDOR and m2.CTIPOCLIENTE <=2 " +
            " join admDomicilios m11 on m11.CIDCATALOGO = m2.CIDCLIENTEPROVEEDOR and m11.CTIPOCATALOGO =1 and m11.CEMAIL <> ' ' " +
            " where m8.CNATURALEZA = 0 and CAFECTADO = 1 and m8.cpendiente > 0  and getdate() > m8.CFECHAVENCIMIENTO" +
                //" and m2.ccodigocliente = 'ECV940426G35'" +                                                                                                            
            " group by m2.CIDCLIENTEPROVEEDOR, m2.CRAZONSOCIAL,m11.CEMAIL ";
            sql1 += sql3;

            string sql2;
            // detalle de documentos
            sql2 = "select cast(m8.CIDCLIENTEPROVEEDOR as decimal(18,0)) as CIDCLIENTEPROVEEDOR,  m8.cfolio, m2.CRAZONSOCIAL,  Convert(varchar(10),CONVERT(date,m8.cfecha,106),103) as cfecha, Convert(varchar(10),CONVERT(date,m8.cfechavencimiento,106),103) as cfechavencimiento, m6.CNOMBRECONCEPTO " +
            " ,m1.CNOMBREAGENTE,m8.ctotal,m8.cpendiente, m5.CNOMBREPRODUCTO, m8.cseriedocumento,m7.cdescripcion, m10.cobservamov as cobservaciones, str(m8.cidmoneda) as cidmoneda " +
            " from admdocumentos m8 join admClientes m2 on m2.CIDCLIENTEPROVEEDOR = m8.CIDCLIENTEPROVEEDOR and m2.CTIPOCLIENTE <=2 " +
            " join admDomicilios m11 on m11.CIDCATALOGO = m2.CIDCLIENTEPROVEEDOR and m11.CTIPOCATALOGO =1 and m11.cemail <> ' ' " +
            " join admConceptos m6 on m8.CIDCONCEPTODOCUMENTO = m6.CIDCONCEPTODOCUMENTO " +
            " join admDocumentosModelo m7 on m7.CIDDOCUMENTODE = m6.CIDDOCUMENTODE " +
            " join admAgentes m1 on m8.cidagente = m1.cidagente " +
            " join admMovimientos m10 on m10.CIDDOCUMENTO = m8.CIDDOCUMENTO and m10.CNUMEROMOVIMIENTO=1 " +
            " join admProductos m5 on m5.CIDPRODUCTO = m10.CIDPRODUCTO " +
            " where m8.CNATURALEZA = 0 and cafectado = 1 " +
            " and m8.cpendiente > 0 " +
            " and m2.ctipocliente = 1 " +
                //" and m2.ccodigocliente = 'ECV940426G35'" +
            " order by m8.CIDCLIENTEPROVEEDOR ";

            sql1 += sql2;

            lconexion = mAbrirConexionOrigen(empresa);



            SqlDataAdapter mySqlDataAdapter = new SqlDataAdapter();
            SqlCommand mySqlCommand = new SqlCommand(sql1, lconexion);
            //sql.CommandText = sql1;
            mySqlDataAdapter.SelectCommand = mySqlCommand;
            mySqlDataAdapter.Fill(ds);

            
            
        }

        private string mArmaXML(string empresa)
        {
            DataTable SaldoTotal;
            DataTable Vencido;
            DataTable Detalle;
            DataSet ds = new DataSet();
            if (radioButton1.Checked == true)
            {

                mProcesarSQL(ref ds, empresa);
                SaldoTotal = ds.Tables[0];
                Vencido = ds.Tables[1];
                Detalle = ds.Tables[2];
            }
            else
            {
                SaldoTotal = null;
                Vencido = null;
                Detalle = null;
                mProcesarDBF(ref SaldoTotal, ref Vencido, ref Detalle, empresa);
                //mProcesarDBF( DataTable SaldoTotal,  DataTable Vencido,  DataTable Detalle );
                

            }
            

            DataRow zz;
            DataTable table = new DataTable();
            table.Columns.Add("uno", typeof(decimal));
            table.Columns.Add("dos", typeof(string));
            table.Columns.Add("tres", typeof(string));
            table.Columns.Add("cuatro", typeof(double));
            table.Columns.Add("cinco", typeof(double));
            table.Columns.Add("seis", typeof(string));
            table.Columns.Add("siete", typeof(string));
            table.Columns.Add("ocho", typeof(string));
            table.Columns.Add("nueve", typeof(string));
            table.Columns.Add("diez", typeof(string));
            table.Columns.Add("once", typeof(string));
            table.Columns.Add("doce", typeof(string));
            table.Columns.Add("trece", typeof(string));
            table.Columns.Add("catorce", typeof(string));
            table.Columns.Add("quince", typeof(string));
            zz = table.Rows.Add(0, "", "", 0, 0, "", "", "", "", "", "", "", "", "", "");

            //m2.CIDCLIENTEPROVEEDOR, m2.CRAZONSOCIAL, m11.cemail, sum(cpendiente)


            var saldos = from total in SaldoTotal.AsEnumerable()
                         join vencido in Vencido.AsEnumerable() on (string)total["CIDCLIENTEPROVEEDOR"].ToString() equals (string)vencido["CIDCLIENTEPROVEEDOR"].ToString() into temp
                         from s1 in temp.DefaultIfEmpty(zz)
                         select new
                         {

                             idcliente = total.Field<decimal>(0),
                             razonsocial = total.Field<string>(1),
                             email = total.Field<string>(2),
                             total = total.Field<double>(3),
                             vencido = s1.Field<double>(3).ToString() ?? string.Empty,
                             nombreempresa = total.Field<string>(4).ToString() ?? string.Empty,
                             rfc = total.Field<string>(5).ToString() ?? string.Empty,
                             telefonoempresa = total.Field<string>(6).ToString() ?? string.Empty,
                             direccionempresa = total.Field<string>(7).ToString() ?? string.Empty,
                             callecliente = total.Field<string>(8).ToString() ?? string.Empty,
                             numerocliente = total.Field<string>(9).ToString() ?? string.Empty,
                             coloniacliente = total.Field<string>(10).ToString() ?? string.Empty,
                             estadocliente = total.Field<string>(11).ToString() ?? string.Empty,
                             telefonocliente = total.Field<string>(12).ToString() ?? string.Empty,
                             rfccliente = total.Field<string>(13).ToString() ?? string.Empty,
                             email1 = total.Field<string>(14).ToString() ?? string.Empty,
                             email2 = total.Field<string>(15).ToString() ?? string.Empty,
                             email3 = total.Field<string>(16).ToString() ?? string.Empty
                         };
            Empresa = SaldoTotal.Rows[0][4].ToString();

            DataRow zzz;
            DataTable table1 = new DataTable();
            table1.Columns.Add("cero", typeof(decimal));
            table1.Columns.Add("uno", typeof(double));
            table1.Columns.Add("dos", typeof(string));
            table1.Columns.Add("tres", typeof(string));
            table1.Columns.Add("cuatro", typeof(string));
            table1.Columns.Add("cinco", typeof(string));
            table1.Columns.Add("seis", typeof(string));
            table1.Columns.Add("siete", typeof(double));
            table1.Columns.Add("ocho", typeof(double));
            table1.Columns.Add("nuevo", typeof(string));
            table1.Columns.Add("diez", typeof(string));
            table1.Columns.Add("once", typeof(string));
            table1.Columns.Add("doce", typeof(string));
            table1.Columns.Add("trece", typeof(string));
            table1.Columns.Add("catorce", typeof(string));
            zzz = table1.Rows.Add(0, 0, "", "", "", "", "", 0, 0, "", "", "","","","");


            System.Xml.Linq.XElement clientes = new XElement("Clientes");
            double porvencer = 0;
            foreach (var saldo in saldos)
            {
                porvencer = double.Parse(saldo.total.ToString()) - double.Parse(saldo.vencido.ToString());
                System.Xml.Linq.XElement cliente = new XElement("Cliente",
                new XElement("Id", saldo.idcliente),
                new XElement("RazonSocial", saldo.razonsocial),
                new XElement("NombreEmpresa", saldo.nombreempresa),
                new XElement("TelefonoEmpresa", saldo.telefonoempresa),
                new XElement("DireccionEmpresa", saldo.direccionempresa),
                new XElement("RFC", saldo.rfc),
                new XElement("Email", saldo.email),
                new XElement("Total", saldo.total),
                new XElement("Vencido", saldo.vencido),
                new XElement("PorVencer", porvencer),
                new XElement("RFCCliente", saldo.rfccliente),
                new XElement("TelefonoCliente", saldo.telefonocliente),
                new XElement("CalleCliente", saldo.callecliente),
                new XElement("NumeroCliente", saldo.numerocliente),
                new XElement("ColoniaCliente", saldo.coloniacliente),
                new XElement("TelefonoCliente", saldo.telefonocliente),
                new XElement("EstadoCliente", saldo.estadocliente),
                new XElement("Banco", Properties.Settings.Default.Banco),
                new XElement("Cuenta", Properties.Settings.Default.Cuenta),
                new XElement("CLABE", Properties.Settings.Default.CLABE),
                new XElement("RFCBanco", Properties.Settings.Default.RFCBanco),
                new XElement("RazonSocialBanco", Properties.Settings.Default.RazonSocialBanco),
                new XElement("correoconfirmacion", Properties.Settings.Default.correoconfirmacion),
                new XElement("email1", saldo.email1),
                new XElement("email2", saldo.email2),
                new XElement("email3", saldo.email3)
                );

                // aqui seria mejor buscar documentos al cabo es en memoria

                var doctos = from total in SaldoTotal.AsEnumerable()
                             join Doctos in Detalle.AsEnumerable() on (string)total["CIDCLIENTEPROVEEDOR"].ToString() equals (string)Doctos["CIDCLIENTEPROVEEDOR"].ToString() into temp
                             from s1 in temp.DefaultIfEmpty(zzz)
                             where total.Field<decimal>(0) == saldo.idcliente
                             //where total["CIDCLIENTEPROVEEDOR"] == s1["CIDCLIENTEPROVEEDOR"].ToString()
                             select new
                             {
                                 /* select cast(m8.CIDCLIENTEPROVEEDOR as decimal(18,0)) as CIDCLIENTEPROVEEDOR,  m8.cfolio, m2.CRAZONSOCIAL,  Convert(varchar(10),CONVERT(date,m8.cfecha,106),103) as cfecha, Convert(varchar(10),CONVERT(date,m8.cfechavencimiento,106),103) as cfechavencimiento, m6.CNOMBRECONCEPTO " +
            " ,m1.CNOMBREAGENTE,m8.ctotal,m8.cpendiente, m5.CNOMBREPRODUCTO, m8.cseriedocumento,m7.cdescripcion " +" +*/
                                 /*
                                 sql22 = "select m8.cidclien01 as CIDCLIENTEPROVEEDOR,  m8.cfolio, m2.crazonso01 as CRAZONSOCIAL,  dtos(m8.cfecha), dtos(m8.cfechavencimiento) as cfechavencimiento, m6.cnombrec01 as CNOMBRECONCEPTO " +
            " ,m1.cnombrea01 as CNOMBREAGENTE,m8.ctotal,m8.cpendiente, m5.cnombrep01 as CNOMBREPRODUCTO, 
                                  * m8.cseriedo01 as cseriedocumento,m7.cdescrip01 as cdescripcion " +*/
                                 idcliente = total.Field<decimal>(0),
                                 Folio = s1.Field<double>(1).ToString() ?? string.Empty,
                                 RazonSocial = s1.Field<string>(2).ToString() ?? string.Empty,
                                 Fecha = s1.Field<string>(3).ToString() ?? string.Empty,
                                 Vencimiento = s1.Field<string>(4).ToString() ?? string.Empty,
                                 Concepto = s1.Field<string>(5) ?? string.Empty,
                                 Agente = s1.Field<string>(6) ?? string.Empty,
                                 Total = s1.Field<double>(7).ToString() ?? string.Empty,
                                 Pendiente = s1.Field<double>(8).ToString() ?? string.Empty,
                                 Producto = s1.Field<string>(9) ?? string.Empty,
                                 Serie = s1.Field<string>(10) ?? string.Empty,
                                 documentomodelo = s1.Field<string>(11) ?? string.Empty,
                                 Observaciones = s1.Field<string>(12) ?? string.Empty,
                                 Moneda = s1.Field<string>(13).ToString() ?? string.Empty
                             };

                XElement documentos = new XElement("Documentos");
                int lentro = 0;
                foreach (var docto in doctos)
                {
                    lentro = 1;
                    string lseriemasfolio = docto.Folio;
                    if (docto.Serie != "")
                        lseriemasfolio = docto.Serie + "/" + docto.Folio.ToString();

                    string lmoneda = "Dolares";
                    if (docto.Moneda.Trim() == "1")
                        lmoneda = "Pesos";

                    System.Xml.Linq.XElement doctito = new XElement("Documento",
                    new XElement("Fecha", docto.Fecha),
                    new XElement("Serie", docto.Serie),
                    new XElement("Folio", lseriemasfolio),
                    new XElement("Vencimiento", docto.Vencimiento),
                    new XElement("Concepto", docto.Concepto),
                    new XElement("Agente", docto.Agente),
                    new XElement("Total", docto.Total),
                    new XElement("Pendiente", docto.Pendiente),
                    new XElement("Producto", docto.Producto),
                    new XElement("IdCliente", docto.idcliente),
                    new XElement("DocumentoModelo", docto.documentomodelo),
                    new XElement("Observaciones", docto.Observaciones),
                    new XElement("Moneda", lmoneda)
                    );
                    if (docto.Folio == "0")
                        lentro = 0;
                    else
                        documentos.Add(doctito);
                }
                if (lentro == 1)
                {
                    cliente.Add(documentos);
                    clientes.Add(cliente);


                }
            }

            var xdoc = new XDocument(new XElement(clientes));
            XmlDocument newxml = new XmlDocument();
            string rawXml = "";
            if (clientes.HasElements == true)
            {
                newxml.LoadXml(xdoc.ToString());
                rawXml = newxml.OuterXml;
            }

            return rawXml;



            

            
        }

        public void mTest(XElement clientes)
        {

            var xdoc = new XDocument(new XElement(clientes));
            XmlDocument newxml = new XmlDocument();
            newxml.LoadXml(xdoc.ToString());
            string rawXml = newxml.OuterXml;

            string strURL = "http://localhost:1067/Service1.asmx/GetTimeString?XmlDoc='uno'";
            //?myint=12345
            //XmlDocument xmlDoc = new XmlDocument();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL);

            request.Method = "POST";
            request.ContentType = "text/xml";
            request.Timeout = 30 * 1000;
            //open the pipe?
            Stream request_stream = request.GetRequestStream();
            //write the XML to the open pipe (e.g. stream)
            newxml.Save(request_stream);
            //CLOSE THE PIPE !!! Very important or next step will time out!!!!
            request_stream.Close();

            //get the response from the webservice
            HttpWebResponse response = (HttpWebResponse) request.GetResponse();
            Stream r_stream = response.GetResponseStream();
            //convert it
            StreamReader response_stream = new
            StreamReader(r_stream,System.Text.Encoding.GetEncoding("utf-8"));
            string sOutput =response_stream.ReadToEnd();

            //display it
            //this.txtAbstract.Text = sOutput;
            MessageBox.Show(sOutput);

            //clean up!
            response_stream.Close();
        }

        public void mTest11(XElement clientes)
        {
            string destinationUrl = "http://localhost:1067/Service1.asmx/GetTimeString";
            var xdoc = new XDocument(new XElement(clientes));
            XmlDocument newxml = new XmlDocument();
            newxml.LoadXml(xdoc.ToString());
            string requestXml = newxml.OuterXml;
            

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(destinationUrl);
            byte[] bytes;
            bytes = System.Text.Encoding.ASCII.GetBytes(requestXml);
            request.ContentType = "text/xml; encoding='utf-8'";
            request.ContentLength = bytes.Length;
            request.Method = "POST";
            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);
            requestStream.Close();
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                Stream responseStream = response.GetResponseStream();
                string responseStr = new StreamReader(responseStream).ReadToEnd();
                //return responseStr;
            }
            //return null;

        }

        public SqlConnection mAbrirConexionOrigen(string mEmpresa)
        {
            SqlConnection _conexion;
            _conexion = null;
            string rutaorigen = mEmpresa;
            if (rutaorigen != "c:\\" && rutaorigen != "VentasPorConcepto.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion = new SqlConnection();
                /*_conexion.ConnectionString = "Server=" + Properties.Settings.Default.server + ";Database=" + Properties.Settings.Default.database + ";User Id=" + Properties.Settings.Default.user + ";Password=" + Properties.Settings.Default.password;

                ListView.SelectedListViewItemCollection empresas = listView1.SelectedItems;
                string empresaseleccionada = "";
	            foreach ( ListViewItem item in empresas )
	            {
                    empresaseleccionada = item.SubItems[1].Text;
                    break;
	            }*/

                mEmpresa = mEmpresa.Substring(mEmpresa.IndexOf("ad"));
                string x = "Server=" + Properties.Settings.Default.server + ";Database=" + mEmpresa + ";User Id=" + Properties.Settings.Default.user + ";Password=" + Properties.Settings.Default.password;
                //MessageBox.Show(x);

                _conexion.ConnectionString = "Server=" + Properties.Settings.Default.server + ";Database=" + mEmpresa + ";User Id=" + Properties.Settings.Default.user + ";Password=" + Properties.Settings.Default.password;
                _conexion.Open();
            }
            return _conexion;

        }
        public OleDbConnection mAbrirConexionOrigenvfp(string mEmpresa)
        {
            OleDbConnection _conexion1;
            _conexion1 = null;
            string rutaorigen = mEmpresa;
            if (rutaorigen != "c:\\" && rutaorigen != "VentasPorConcepto.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion1 = new OleDbConnection();
                _conexion1.ConnectionString = "Provider=vfpoledb.1;Data Source=" + rutaorigen;
                _conexion1.Open();
            }
            return _conexion1;

        }

        public SqlConnection mAbrirConexionOrigen2(string mEmpresa)
        {
            SqlConnection _conexion;
            _conexion = null;
            string rutaorigen = mEmpresa;
            if (rutaorigen != "c:\\" && rutaorigen != "VentasPorConcepto.RegEmpresa" && rutaorigen != "Ruta")
            {
                _conexion = new SqlConnection();
                _conexion.ConnectionString = "Server=" + Properties.Settings.Default.server + ";Database=" + Properties.Settings.Default.database + ";User Id=" + Properties.Settings.Default.user + ";Password=" + Properties.Settings.Default.password;
                _conexion.Open();
            }
            return _conexion;

        }

        private string mProcesarCorreos()
        {
            Properties.Settings.Default.correoprueba = textPrueba.Text;
            Properties.Settings.Default.correoreply = textReply.Text;
            Properties.Settings.Default.server = txtServer.Text;
            //Properties.Settings.Default.database = listView1.;
            Properties.Settings.Default.user = txtUser.Text;
            Properties.Settings.Default.password = txtPass.Text;
            Properties.Settings.Default.correoSalida = textCuentaMail.Text;

            Properties.Settings.Default.Save();

            ListView.ListViewItemCollection empresas = listView1.Items;

            foreach (ListViewItem empresa in empresas)
            {
                textBox1.Text = "Consultando informacion";
                this.Refresh();
                string miXml = mArmaXML(empresa.SubItems[1].Text);
                if (miXml == "")
                    return "No existen registros con saldos pendientes";
                Sitiodesarrollosoftwarecontable.WebService1 obj = new Sitiodesarrollosoftwarecontable.WebService1();
                //aloneservice.WebService1 obj = new aloneservice.WebService1();
                List<Correo> listacorreos = new List<Correo>();
                string x = "";
                try
                {
                    textBox1.Text = "Validando informacion";
                    this.Refresh();
                    x = obj.ConXML(miXml, textReply.Text, textPrueba.Text, txtCodigoSitio.Text);
                }
                catch (Exception ee)
                {
                    return ee.Message;
                }


                JavaScriptSerializer serializador = new JavaScriptSerializer();

                serializador.MaxJsonLength = int.MaxValue;

                listacorreos = serializador.Deserialize<List<Correo>>(x);

                if (listacorreos.Count == 0)
                {
                    return "Configuracion no activa o no existe, Validar con su implementador";
                }

                if (listacorreos.Count == 1)
                {
                    if (listacorreos[0].asunto == "No registrado")
                        return "Codigo de sitio no Registrado";
                }


                if (textCuentaMail.Text != "" && textPWDCuenta.Text != "")
                {
                    
                    //client.Host = textHost.Text;
                    //client.Port = int.Parse(textPort.Text);

                    //client.Credentials = new System.Net.NetworkCredential(textCuentaMail.Text, textPWDCuenta.Text);
                    //client.Timeout = 20000;
                    
                }

                List<Correo> lista2 = new List<Correo>();
                int ii = 1;
                foreach (Correo mailx in listacorreos)
                {
                    textBox1.Text = "Procesando correo " + ii.ToString() + " de " + listacorreos.Count.ToString();
                    this.Refresh();
                    /*try
                    {*/
                    mailx.asunto = "Documentos con Saldo Pendiente";
                    string respuesta = "Correo mal escrito";
                    if (mailx.para != "")
                        respuesta = mEnviaCorreo(mailx.para, mailx.asunto, mailx.cuerpo, mailx.reply);
                        //respuesta = "sdkjfsdlkfjslkdjflksdasdfdsfsdfdsjflksdjflksadjflksadjflksdajflsdjflksadjflkasdjlkfjsdalkfjasdlkfjasdlk;jfl;kasdjf;lkasdjflkasdjflkdsajlfkjsdalkfjasd;fjasdjfklsadjflksdajflkasdjflkasdjflksdajflkdsj";
                        //respuesta = "OK";
                     
                    if (respuesta =="OK")
                        mailx.status = "Enviado";
                    else
                    {
                        if (respuesta.Length > 100)
                            mailx.status = respuesta.Substring(1,100);
                        else
                            mailx.status = respuesta;
                    }
                    mailx.de = textCuentaMail.Text;
                    lista2.Add(mailx);
                    ii++;
                }
                string correos = serializador.Serialize(lista2);
                textBox1.Text = "Grabando historial ";
                this.Refresh();
                string xxx = obj.GrabarEmails(correos,txtCodigoSitio.Text,Empresa);

                textBox1.Text = "Correos Enviados ";
                this.Refresh();
            }
            return "Correos Enviados";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(mProcesarCorreos());
            

        }

        private void mEnviargmail()
        { 
            string result = "Message Sent Successfully..!!";
     string senderID = "hectorvalagui@gmail.com";// use sender’s email id here..
     const string senderPassword = "1"; // sender password here…
     try
     {
       SmtpClient smtp = new SmtpClient
       {
         Host = "smtp.gmail.com", // smtp server address here…
         Port = 587,
         EnableSsl = true,
         DeliveryMethod = SmtpDeliveryMethod.Network,
         Credentials = new System.Net.NetworkCredential(senderID, senderPassword),
         Timeout = 30000,
       };
       MailMessage message = new MailMessage(senderID, "hectorvalagui@hotmail.com", "uno", "dos");
       smtp.Send(message);
     }
     catch (Exception ex)
     {
       result = "Error sending email.!!!";
     }
 }




        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                textHost.Text = Properties.Settings.Default.hostsmtp;
                textPort.Text = Properties.Settings.Default.puerto;
                textCuentaMail.Text = Properties.Settings.Default.cuentaemail;
                txtEmailDestino.Text = Properties.Settings.Default.emaildestino;
                textPWDCuenta.Text = Properties.Settings.Default.pwdMail;




                label25.Text = "Notifika " + this.ProductVersion;


                //smtp3.hp.com
                this.SetStyle(
        ControlStyles.AllPaintingInWmPaint |
        ControlStyles.UserPaint |
        ControlStyles.DoubleBuffer,
        true);

                listView1.Columns.Add("uno");
                listView1.Columns.Add("dos");
                listView1.View = View.Details;
                listView1.Columns[0].Width = 500;
                listView1.Columns[1].Width = 0;
                listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
                //listView1.Columns[0].AutoResize(ColumnHeaderAutoResizeStyle.None);
                //listView1.Columns[0].Width = 7000;
                //listView1.Columns[0].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);



                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker1.CustomFormat = "HH:mm";


                //For 12 H format
                //dateTimePicker1.CustomFormat = "hh:mm:ss tt";
                dateTimePicker1.ShowUpDown = true;

                dateTimePicker1.Value = DateTime.Today;
                string xx = Properties.Settings.Default.Hora;

                if (xx == "")
                    dateTimePicker1.Value = DateTime.Now.AddHours(-1);
                else
                    dateTimePicker1.Value = DateTime.Parse(Properties.Settings.Default.Hora);





                txtServer.Text = Properties.Settings.Default.server;
                txtBD.Text = "CompacWAdmin";
                txtUser.Text = Properties.Settings.Default.user;
                txtPass.Text = Properties.Settings.Default.password;
                txtPass.PasswordChar = '*';

                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                    ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                    "; password = " + Properties.Settings.Default.password + ";";

                empresasComercial1.Populate(Cadenaconexion);

                textPrueba.Text = Properties.Settings.Default.correoprueba;
                textReply.Text = Properties.Settings.Default.correoreply;

                //textCuentaMail.Text = Properties.Settings.Default.correoSalida;

                string lperiodicidad = Properties.Settings.Default.Periodicidad;
                if (lperiodicidad == "")
                    lperiodicidad = "Mensual";
                string ldia = Properties.Settings.Default.Dia;
                if (ldia == "")
                    ldia = "1";
                string lhora = Properties.Settings.Default.Hora;
                if (lhora == "")
                    lhora = "08:00";
                cbPeriodicidad.Items.Add("Mensual");
                cbPeriodicidad.Items.Add("Semanal");
                cbPeriodicidad.Items.Add("Diaria");





                switch (lperiodicidad)
                {
                    case "Mensual":
                        cbPeriodicidad.SelectedIndex = 0;
                        for (int i = 1; i <= 31; i++)
                        {
                            cbDia.Items.Add(i);
                        }
                        cbDia.SelectedIndex = int.Parse(ldia) - 1;
                        break;
                    case "Semanal":
                        cbPeriodicidad.SelectedIndex = 1;
                        cbDia.Items.Add("Lunes");
                        cbDia.Items.Add("Martes");
                        cbDia.Items.Add("Miercoles");
                        cbDia.Items.Add("Jueves");
                        cbDia.Items.Add("Viernes");
                        cbDia.Items.Add("Sabado");
                        cbDia.Items.Add("Domingo");
                        switch (Properties.Settings.Default.Dia.ToString())
                        {
                            case "Lunes": cbDia.SelectedIndex = 0; break;
                            case "Martes": cbDia.SelectedIndex = 1; break;
                            case "Miercoles": cbDia.SelectedIndex = 2; break;
                            case "Jueves": cbDia.SelectedIndex = 3; break;
                            case "Viernes": cbDia.SelectedIndex = 4; break;
                            case "Sabado": cbDia.SelectedIndex = 5; break;
                            case "Domingo": cbDia.SelectedIndex = 6; break;
                            default: cbDia.SelectedIndex = 0; break;
                        }
                        break;
                    case "Diaria":
                        cbPeriodicidad.SelectedIndex = 2; break;
                    default:
                        int ii = 0;

                        cbPeriodicidad.SelectedIndex = 0;
                        //break;
                        for (int i = 1; i <= 31; i++)
                        {
                            cbDia.Items.Add(i);
                        }
                        cbDia.SelectedIndex = int.Parse(ldia);
                        break;

                }





                timer1.Interval = 10000;

                string lcodigo = mBuscaRegistry("SOFTWARE\\Computación en Acción, SA CV\\AppKey\\CONTPAQ_I_COMERCIAL", "SiteCode");
                /*string llaveregistry = "SOFTWARE\\Computación en Acción, SA CV\\AppKey\\CONTPAQ_I_COMERCIAL";
                RegistryKey hklp = Registry.LocalMachine;
                hklp = hklp.OpenSubKey(llaveregistry);
                Object obc = hklp.GetValue("SiteCode");
                string lruta1 = obc.ToString();*/

                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
                if (lcodigo != "")
                {
                    txtCodigoSitio.Text = lcodigo;
                    radioButton1.Enabled = true;
                }
                /*
                HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Computación en Acción, SA CV\AdminPAQ
HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Computación en Acción, SA CV\CONTPAQ I Facturacion
HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Computación en Acción, SA CV\AppKey\Contpaq_i
                */


                lcodigo = mBuscaRegistry("SOFTWARE\\Computación en Acción, SA CV\\AppKey\\AdminPAQ", "SiteCode");
                if (lcodigo != "")
                {
                    txtCodigoSitio.Text = lcodigo;
                    radioButton2.Enabled = true;
                    // buscar eldirectorio 
                    lcodigo = mBuscaRegistry("SOFTWARE\\Computación en Acción, SA CV\\AdminPAQ", "DIRECTORIODATOS");
                    txtDirectorioAdmin.Text = @lcodigo;
                }
                else
                    txtDirectorioAdmin.Text = "";


                //HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Computación en Acción, SA CV\AppKey\FacturacionI
                lcodigo = mBuscaRegistry("SOFTWARE\\Computación en Acción, SA CV\\AppKey\\FacturacionI", "SiteCode");
                if (lcodigo != "")
                {
                    txtCodigoSitio.Text = lcodigo;
                    radioButton3.Enabled = true;
                    // buscar eldirectorio 
                    lcodigo = mBuscaRegistry("SOFTWARE\\Computación en Acción, SA CV\\CONTPAQ I Facturacion", "DIRECTORIODATOS");
                    txtDirectorioFE.Text = @lcodigo;
                }


                try
                {
                    textHost.Text = Properties.Settings.Default.hostsmtp;
                    txtBanco.Text = Properties.Settings.Default.Banco;
                    txtRFCBanco.Text = Properties.Settings.Default.RFCBanco;
                    txtCuenta.Text = Properties.Settings.Default.Cuenta;
                    txtCLABE.Text = Properties.Settings.Default.CLABE;
                    txtCorreoConfirmacion.Text = Properties.Settings.Default.correoconfirmacion;
                    txtRazonSocialBanco.Text = Properties.Settings.Default.RazonSocialBanco;
                }
                catch (Exception ee)
                {
                }
            }
            catch (Exception eee)
            {
                MessageBox.Show(eee.Message);
            }


        }

      

        private string mBuscaRegistry(string llave, string subllave)
        {
            string lregreso = "";
            RegistryKey hklp = Registry.LocalMachine;
            try
            {
                hklp = hklp.OpenSubKey(llave);
                Object obc = hklp.GetValue(subllave);
                lregreso = obc.ToString();
            }
            catch (Exception eee)
            { }
            return lregreso;
        }

        private bool mValida()
        {
            string Cadenaconexion = "data source =" + txtServer.Text + ";initial catalog =" + txtBD.Text + ";user id = " + txtUser.Text + "; password = " + txtPass.Text + ";";
            SqlConnection _con = new SqlConnection();

            _con.ConnectionString = Cadenaconexion;
            try
            {
                _con.Open();
                // si se conecto grabar los datos en el cnf
                _con.Close();
                return true;
            }
            catch (Exception ee)
            {
                return false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (mValida())
            {
                Properties.Settings.Default.server = txtServer.Text;
                Properties.Settings.Default.database = txtBD.Text;
                Properties.Settings.Default.user = txtUser.Text;
                Properties.Settings.Default.password = txtPass.Text;

                Properties.Settings.Default.Save();
                
                Cadenaconexion = "data source =" + Properties.Settings.Default.server +
                ";initial catalog =" + Properties.Settings.Default.database + " ;user id = " + Properties.Settings.Default.user +
                "; password = " + Properties.Settings.Default.password + ";";
                /*y.mllenarcomboempresas();
                y.Visible = true;*/
                MessageBox.Show("Conexion Exitosa");
                empresasComercial1.Populate(Cadenaconexion);
            }
            else
                MessageBox.Show("Valores de conexion incorrectos");
        }

        string lhora2="0";
        private void timer1_Tick(object sender, EventArgs e)
        {
            int lchecarhora = 0;
            string periodo = Properties.Settings.Default.Periodicidad;
            string hora = Properties.Settings.Default.Hora;
            
            
            DateTime lhora = DateTime.Parse(hora);

/*            CultureInfo mexicanSpanishCi = CultureInfo.GetCultureInfo("es-MX");
            Thread.CurrentThread.CurrentCulture = mexicanSpanishCi;
            Thread.CurrentThread.CurrentUICulture = mexicanSpanishCi;*/

            System.Globalization.CultureInfo cultureInfo = new System.Globalization.CultureInfo("es-MX", true);



            switch (Properties.Settings.Default.Periodicidad)
            { 
                case "Mensual":
                    int ldia = System.DateTime.Today.Day;
                    if (ldia == int.Parse(Properties.Settings.Default.Dia.ToString()))
                        lchecarhora = 1;
                    break;
                case "Semanal":
                    int x = (int)System.Globalization.CultureInfo
        .InvariantCulture.Calendar.GetDayOfWeek(DateTime.Now);

                    string ldiax = Properties.Settings.Default.Dia.ToString();
                    string ldiay = "";
                    switch (x)
                    {
                        case 1: ldiay = "Lunes"; break;
                        case 2: ldiay = "Martes"; break;
                        case 3: ldiay = "Miercoles"; break;
                        case 4: ldiay = "Jueves"; break;
                        case 5: ldiay = "Viernes"; break;
                        case 6: ldiay = "Sabado"; break;
                        case 7: ldiay = "Domingo"; break;
                    }
                    
                    if (ldiax == ldiay)
                        lchecarhora = 1;
                    break;
                case "Diaria":
                    lchecarhora = 1;
                    break;
            }
            if (lchecarhora == 1)
            {
                if (lhora.Hour == System.DateTime.Now.Hour && lhora.Minute == System.DateTime.Now.Minute)
                {
                    mProcesarCorreos();
                    //MessageBox.Show("chido");
                    while (lhora.Minute == System.DateTime.Now.Minute) ;
                }
                else
                    lchecarhora = 0;

            }
            

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cbPeriodicidad.SelectedItem.ToString())
            //switch (lperiodicidad)
            {
               case "Mensual":
                    cbDia.Items.Clear();
                    for (int i = 1; i <= 31; i++)
                    {
                        cbDia.Items.Add(i);
                    }
                    cbDia.SelectedIndex = 0;
                    break;
                case "Semanal":
                    cbDia.Items.Clear();
                        cbDia.Items.Add("Lunes");
                        cbDia.Items.Add("Martes");
                        cbDia.Items.Add("Miercoles");
                        cbDia.Items.Add("Jueves");
                        cbDia.Items.Add("Viernes");
                        cbDia.Items.Add("Sabado");
                        cbDia.Items.Add("Domingo");
                        switch (Properties.Settings.Default.Dia.ToString())
                        {
                            case "Lunes": cbDia.SelectedIndex = 0; break;
                            case "Martes": cbDia.SelectedIndex = 1; break;
                            case "Miercoles": cbDia.SelectedIndex = 2; break;
                            case "Jueves": cbDia.SelectedIndex = 3; break;
                            case "Viernes": cbDia.SelectedIndex = 4; break;
                            case "Sabado": cbDia.SelectedIndex = 5; break;
                            case "Domingo": cbDia.SelectedIndex = 6; break;
                            default: cbDia.SelectedIndex = 0; break;
                        }
                    break;
                case "Diaria": 
                    cbPeriodicidad.SelectedIndex = 2; break;
                default:
                    for (int i = 1; i <= 31; i++)
                    {
                        cbDia.Items.Add(i);
                    }
                    cbDia.SelectedIndex = 0;
                    break;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            Properties.Settings.Default.Periodicidad = cbPeriodicidad.SelectedItem.ToString();
            if (cbDia.Items.Count==0)
                Properties.Settings.Default.Dia = "0";
            else
                Properties.Settings.Default.Dia = cbDia.SelectedItem.ToString();
            Properties.Settings.Default.Hora = dateTimePicker1.Value.Hour.ToString() + ":" + dateTimePicker1.Value.Minute.ToString().PadLeft(2,'0');

            /*Properties.Settings.Default.server = txtServer.Text;
            Properties.Settings.Default.database = txtBD.Text;
            Properties.Settings.Default.user = txtUser.Text ;
            Properties.Settings.Default.password = txtPass.Text ;
            Properties.Settings.Default.correoprueba = textPrueba.Text ;
            Properties.Settings.Default.correoreply = textReply.Text ;*/

            Properties.Settings.Default.Save();
            timer1.Start();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Banco = txtBanco.Text;
            Properties.Settings.Default.Cuenta = txtCuenta.Text;
            Properties.Settings.Default.CLABE = txtCLABE.Text ;
            Properties.Settings.Default.correoconfirmacion = txtCorreoConfirmacion.Text ;
            Properties.Settings.Default.RazonSocialBanco = txtRazonSocialBanco.Text;
            Properties.Settings.Default.RFCBanco = txtRFCBanco.Text;
            Properties.Settings.Default.Save();
            
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            /*
            if (Cadenaconexion != "")
            {
                ciCompanyList11.Populate(Cadenaconexion);
            }
            else
            {
                this.Visible = false;
                Form4 x = new Form4();
                x.asignaform1(this);
                x.Show();
            }
            if (Archivo != "")
                botonExcel1.mSetNombre(Archivo);
            */
            this.Text = " Notifik Documentos con Saldo Pendiente " + this.ProductVersion;
            comboBox2.SelectedIndex = 0;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            
        }

        private void tabPage1_Enter(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                empresasComercial1.Visible = true;
                label30.Visible = false;
                comboBox1.Visible = false;
            }
            else
            {
                empresasComercial1.Visible = false;
                label30.Visible = true;
                comboBox1.Visible = true;
                string mensaje="";
                this.comboBox1.DataSource = null;
                this.comboBox1.Items.Clear();
                
                this.comboBox1.DataSource = x.mCargarEmpresas(out mensaje);
                comboBox1.DisplayMember = "Nombre";
                comboBox1.ValueMember = "Ruta";
                try
                {
                    this.comboBox1.SelectedIndex = 1;
                    this.comboBox1.SelectedIndex = 0;
                }
                catch (Exception ee)
                {
                    this.comboBox1.SelectedIndex = 0;
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count == 5)
            {
                MessageBox.Show("Solo es permitido configurar 5 empresas");
                return;
            }
            

            if (radioButton1.Checked == true)
            {
                ListViewItem itemF = listView1.FindItemWithText(empresasComercial1.nombreempresa);
                if (itemF != null)
                {
                    MessageBox.Show("Empresa ya seleccionada");
                    return;
                }
                //ListViewItem item = new ListViewItem(new[] { empresasComercial1.aliasbdd });
                //listView1.Items.Add(empresasComercial1.nombreempresa, empresasComercial1.aliasbdd);

                ListViewItem item = new ListViewItem(new[] { empresasComercial1.nombreempresa, empresasComercial1.aliasbdd });
                listView1.Items.Add(item);

            }
            else
            {
                ListViewItem itemF = listView1.FindItemWithText(comboBox1.Text);
                if (itemF != null)
                {
                    MessageBox.Show("Empresa ya seleccionada");
                    return;
                }
                ListViewItem item = new ListViewItem(new[] { comboBox1.Text, comboBox1.SelectedValue.ToString()});
                listView1.Items.Add(item);
            }
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem listViewItem in listView1.SelectedItems)
            {
                listViewItem.Remove();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string result = "Message Sent Successfully..!!";
            string senderID = textCuentaMail.Text ;// use sender’s email id here..
            const string senderPassword = "1"; // sender password here…

            bool ssl = false; //gmail
            if (radioButton4.Checked == true)
                         ssl = true; //gmail

            try
            {
                SmtpClient smtp = new SmtpClient
                {
                    Host = textHost.Text, // smtp server address here…
                    Port = int.Parse(textPort.Text),
                    
                    EnableSsl = ssl,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new System.Net.NetworkCredential(senderID, senderPassword),
                    Timeout = 30000,
                };
                MailMessage message = new MailMessage(senderID, txtEmailDestino.Text,"Notifika Test", "Notifika Test");
                smtp.Send(message);
                MessageBox.Show("Correo enviado correctamente");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Correo no enviado, Revise configuracion"); 
            }

        }


        private string mEnviaCorreo(string para, string asunto,string cuerpo, string reply)
        {
            System.Net.Mail.SmtpClient cliente = new System.Net.Mail.SmtpClient();

            //Hay que crear las credenciales del correo emisor
            cliente.Credentials =
                new System.Net.NetworkCredential(textCuentaMail.Text, textPWDCuenta.Text);

            
            //Lo siguiente es obligatorio si enviamos el mensaje desde Gmail
            /*
            cliente.Port = 587;
            cliente.EnableSsl = true;
            */
            if (textHost.Text.IndexOf("gmail")>=0)
            {
                cliente.Port = 587;
                cliente.EnableSsl = true;
            }
            cliente.Port = int.Parse(textPort.Text);
            cliente.EnableSsl = false; //gmail
            if (radioButton4.Checked == true)
                cliente.EnableSsl = true; //gmail

            cliente.Host = textHost.Text; //Para Gmail "smtp.gmail.com";


            /*-------------------------ENVIO DE CORREO----------------------*/

            try
            {
                //Enviamos el mensaje      
                //if (email_bien_escrito(para) == true)
                //{
                if (para.IndexOf('ñ') == -1)
                {
                    MailMessage message = new MailMessage(textCuentaMail.Text, para, asunto, cuerpo);
                    message.IsBodyHtml = true;
                    if (reply != "")
                    {
                        message.ReplyTo = new MailAddress(reply);
                    }
                    //smtp.Send(message);
                    cliente.Send(message);
                }
                else
                 return "Email mal escrito";
                return "OK";
            }
            catch (System.Net.Mail.SmtpException ex)
            {
                return ex.Message; //Aquí gestionamos los errores al intentar enviar el correo
            }
        }

        private Boolean email_bien_escrito(String email)
        {
            String expresion;
            expresion = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
            if (Regex.IsMatch(email, expresion))
            {
                if (Regex.Replace(email, expresion, String.Empty).Length == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        private void mProbarCorreo1()
        {
            Properties.Settings.Default.hostsmtp = textHost.Text;
            Properties.Settings.Default.puerto = textPort.Text;
            Properties.Settings.Default.cuentaemail = textCuentaMail.Text;
            Properties.Settings.Default.emaildestino = txtEmailDestino.Text;
            Properties.Settings.Default.pwdMail = textPWDCuenta.Text;

            Properties.Settings.Default.Save();


            if (textHost.Text.IndexOf("gmail") >= 0)
            {
                bool ssl = false; //gmail
                if (radioButton4.Checked == true)
                    ssl = true; //gmail
            }


            string result = "Message Sent Successfully..!!";
            string senderID = textCuentaMail.Text;// use sender’s email id here..
            const string senderPassword = "1"; // sender password here…



            try
            {
                SmtpClient smtp = new SmtpClient
                {
                    Host = textHost.Text, // smtp server address here…
                    Port = int.Parse(textPort.Text),

                    //EnableSsl = ssl,
                    //DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new System.Net.NetworkCredential(senderID, senderPassword),
                    //Timeout = 30000,
                };
                MailMessage message = new MailMessage(senderID, txtEmailDestino.Text, "Notifika Test", "Notifika Test");
                smtp.Send(message);
                MessageBox.Show("Correo enviado correctamente");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Correo no enviado, Revise configuracion");
            }


        }

        private void button8_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.hostsmtp = textHost.Text;
            Properties.Settings.Default.puerto = textPort.Text;
            Properties.Settings.Default.cuentaemail = textCuentaMail.Text;
            Properties.Settings.Default.emaildestino = txtEmailDestino.Text;
            Properties.Settings.Default.pwdMail = textPWDCuenta.Text;

            Properties.Settings.Default.Save();
            string respuesta = mEnviaCorreo(txtEmailDestino.Text, "Notifika Test", "Notifika Test","");
            if (respuesta != "OK")
                MessageBox.Show(respuesta);
            else
                MessageBox.Show("Correo enviado correctamente");
            
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            Sitiodesarrollosoftwarecontable.WebService1 obj = new Sitiodesarrollosoftwarecontable.WebService1();            
            //aloneservice.WebService1 obj = new aloneservice.WebService1();
            reg nuevo = new reg();

            nuevo.Activation = txtCodigoSitio.Text;
            nuevo.batch = txtLote.Text;
            if (txtLote.Text == "") { MessageBox.Show("Capture Lote"); return; }
            
            nuevo.email = txtEmail.Text;
            if (txtEmail.Text == "") { MessageBox.Show("Capture Email"); return; }
            nuevo.name = txtRazonSocial.Text;
            if (txtRazonSocial.Text == "") { MessageBox.Show("Capture Razon Social"); return; }
            nuevo.phone = txtTelefono.Text;
            if (txtTelefono.Text == "") { MessageBox.Show("Capture Telefono"); return; }
            nuevo.RFC = txtRFC.Text;
            if (txtRFC.Text == "") { MessageBox.Show("Capture RFC"); return; }
            nuevo.Tipo = comboBox2.SelectedItem.ToString() ;


            JavaScriptSerializer serializador = new JavaScriptSerializer();

                serializador.MaxJsonLength = int.MaxValue;

            //reg regs = serializador.Deserialize<reg>(cadena);

            string xxxxx;
            xxxxx = serializador.Serialize(nuevo);

            string zzzzz = obj.Grabar(xxxxx);
            MessageBox.Show(zzzzz);
        }

        private void txtTelefono_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                //MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }
    }
}
