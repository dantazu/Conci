using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Concilia.ui;
using ExcelDataReader;

namespace Concilia.ui
{
    public class ExcelUtil
    {
        public static List<BancoUpeu> ListaUpeu = null;

        public static List<BancoBCP> ListaBcp = null;

        public static List<BancoBCP> ListaBcpOriginal = null;

        public static List<BancoBCP> ListaBcpTeclaso = null;

        public static List<BancoUpeu> ListaPenUpeu = null;

        public static List<BancoBCP> ListaPenBanco = null;

        public static List<BancoUpeu> ListaPenUpeuAnon = null;

        public static List<BancoBCP> ListaPenBancoAnon = null;

        public static List<ComparabancosComision> Listacom = null;

        public static List<ComparabancosComision> Listacom2 = null;

        public static List<ComparabancosComision> Listadet = null;

        private static string CadenaConexion;

        private static OleDbConnection cnn;

        private static OleDbDataAdapter Adaptador;

        private static OleDbCommandBuilder Constructor;

        private static DataTable Tabla;

        public static List<string> ListaSheets { get; set; }

        public static DataTableCollection GetDatatale(string FileName)
        {
            //using FileStream fileStream = File.Open(FileName, FileMode.Open, FileAccess.Read);
            //using IExcelDataReader self = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
             FileStream fileStream = File.Open(FileName, FileMode.Open, FileAccess.Read);
             IExcelDataReader self = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
            DataSet dataSet = self.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = (IExcelDataReader _) => new ExcelDataTableConfiguration()
            });
            return dataSet.Tables;
        }

        public static DataTableCollection GetNameSheets(string FileName)
        {
            DataTableCollection datatale = GetDatatale(FileName);
            List<string> list = new List<string>();
            foreach (object item in datatale)
            {
                list.Add(item.ToString());
            }
            ListaSheets = list; 
            return datatale;
        }
       
        public static List<BancoBCP> PopulateinCollectionBCP(DataTable table)
        {
            ListaBcp = new List<BancoBCP>();
            ListaBcpOriginal = new List<BancoBCP>();
            ListaBcpTeclaso = new List<BancoBCP>();
            Listacom = new List<ComparabancosComision>();
            Listacom2 = new List<ComparabancosComision>();
            Listadet = new List<ComparabancosComision>();
            IFormatProvider formatProvider = new CultureInfo("es-ES", useUserOverride: false);
            for (int i = 5; i < table.Rows.Count + 1; i++)
            {
                try
                {
                    BancoBCP bancoBCP = new BancoBCP();
                    DateTime dateTime = DateTime.Parse(table.Rows[i - 1][0].ToString());
                    bancoBCP.FechaOperacion = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day);
                    string empty = string.Empty;
                    string text = table.Rows[i - 1][2].ToString();
                    string empty2 = string.Empty;
                    if (table.Rows[i - 1][1].ToString().Equals("8928129"))
                    {
                        string empty3 = string.Empty;
                    }
                    if (!table.Rows[i - 1][1].ToString().Trim().Equals("DETRAC"))
                    {
                        if (text.Contains("PAGO DETRAC"))
                        {
                            empty = text.Substring(text.Length - 10).TrimStart('0');
                            bancoBCP.NroOpe = empty;
                        }
                        else if (text.Contains("PROV TLC"))
                        {
                            empty = text.Trim().Substring(text.Trim().Length - 5).TrimStart('0');
                            bancoBCP.NroOpe = empty;
                        }
                        else if (text.Contains("HABER TLC"))
                        {
                            empty = text.Trim().Substring(text.Trim().Length - 5).TrimStart('0');
                            bancoBCP.NroOpe = empty;
                        }
                        else if (text.Contains("DEVOL. PAGO"))
                        {
                            empty = text.Trim().Substring(text.Trim().Length - 5).TrimStart('0');
                            bancoBCP.NroOpe = empty;
                        }
                        else if (text.Contains("CTS TLC"))
                        {
                            empty = text.Trim().Substring(text.Trim().Length - 4).TrimStart('0');
                            bancoBCP.NroOpe = empty;
                        }
                        else if (text.Contains("CHEQUE"))
                        {
                            string text2 = text.Substring(0, 6);
                            if (text2.Equals("CHEQUE"))
                            {
                                empty = text.Trim().Substring(text.Trim().Length - 4).TrimStart('0');
                                bancoBCP.NroOpe = empty;
                            }
                            else
                            {
                                bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                            }
                        }
                        else if (text.Contains("CHEQ"))
                        {
                            string text2 = text.Substring(0, 6);
                            if (text2.Equals("CHEQ.P"))
                            {
                                empty = text.Trim().Substring(text.Trim().Length - 4).TrimStart('0');
                                bancoBCP.NroOpe = empty;
                            }
                            else
                            {
                                bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                            }
                        }
                        else
                        {
                            bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                        }
                    }
                    else
                    {
                        bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                    }
                    bancoBCP.Descripcion = table.Rows[i - 1][2].ToString();
                    decimal num = default(decimal);
                    if (string.IsNullOrEmpty(table.Rows[i - 1][3].ToString()))
                    {
                        num = default(decimal);
                    }
                    else
                    {
                        num = decimal.Parse(table.Rows[i - 1][3].ToString());
                        if (num > 0m)
                        {
                            bancoBCP.Dh = 1;
                        }
                        else
                        {
                            bancoBCP.Dh = 2;
                            num = -1m * num;
                        }
                    }
                    bancoBCP.Importe = num;
                    bancoBCP.CodigoPos = "0";
                    if (bancoBCP.NroOpe.ToUpper().Equals("COMIS"))
                    {
                        ComparabancosComision item = new ComparabancosComision
                        {
                            FechaOpe = bancoBCP.FechaOperacion,
                            NroOpe = bancoBCP.NroOpe,
                            Descripcion = bancoBCP.Descripcion,
                            Importe = bancoBCP.Importe,
                            Dh = bancoBCP.Dh
                        };
                        Listacom.Add(item);
                    }
                    else if (bancoBCP.NroOpe.ToUpper().Equals("DETRAC"))
                    {
                        ComparabancosComision item2 = new ComparabancosComision
                        {
                            FechaOpe = bancoBCP.FechaOperacion,
                            NroOpe = bancoBCP.NroOpe,
                            Descripcion = bancoBCP.Descripcion,
                            Importe = bancoBCP.Importe,
                            Dh = bancoBCP.Dh
                        };
                        Listadet.Add(item2);
                    }
                    else
                    {
                        ListaBcp.Add(bancoBCP);
                    }
                    if (bancoBCP.NroOpe.Equals("VISANET"))
                    {
                        bancoBCP.Whoyo = "Banco-Visanet";
                    }
                    else
                    {
                        bancoBCP.Whoyo = "Banco";
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error en Hoja BANCO: " + Environment.NewLine + "En la fila " + i + Environment.NewLine + ex.Message);
                }
            }
            if (Listacom.Count > 0)
            {
                BancoBCP item3 = new BancoBCP
                {
                    FechaOperacion = Listacom.First().FechaOpe,
                    NroOpe = "COMIS",
                    Descripcion = "COMIS.RECAUDACION",
                    Importe = Listacom.Sum((ComparabancosComision W) => W.Importe),
                    Dh = Listacom.First().Dh,
                    CodigoPos = "0"
                };
                ListaBcp.Add(item3);
            }
            if (Listacom2.Count > 0)
            {
                BancoBCP item4 = new BancoBCP
                {
                    FechaOperacion = Listacom2.First().FechaOpe,
                    NroOpe = "xxxx",
                    Descripcion = "COMIS.RECAUDACION",
                    Importe = Listacom2.Sum((ComparabancosComision W) => W.Importe),
                    Dh = Listacom2.First().Dh,
                    CodigoPos = "0"
                };
                ListaBcp.Add(item4);
            }
            if (Listadet.Count > 0)
            {
                BancoBCP item5 = new BancoBCP
                {
                    FechaOperacion = Listadet.First().FechaOpe,
                    NroOpe = "DETRAC",
                    Descripcion = "COMIS.DETRACCION",
                    Importe = Listadet.Sum((ComparabancosComision W) => W.Importe),
                    Dh = Listadet.First().Dh,
                    CodigoPos = "0"
                };
                ListaBcp.Add(item5);
            }
            return ListaBcp;
        }

        public static List<BancoBCP> PopulateinCollectionPEN_Upeu(DataTable table)
        {
            ListaPenBanco = new List<BancoBCP>();
            for (int i = 5; i < table.Rows.Count + 1; i++)
            {
                try
                {
                    BancoBCP bancoBCP = new BancoBCP();
                    bancoBCP.Whoyo = "P-Upe";
                    bancoBCP.FechaOperacion = DateTime.Parse(table.Rows[i - 1][0].ToString());
                    if (table.Rows[i - 1][1].ToString().Contains("COM_"))
                    {
                        string empty = string.Empty;
                    }
                    bancoBCP.Descripcion = table.Rows[i - 1][2].ToString();
                    if (string.IsNullOrEmpty(table.Rows[i - 1][7].ToString()))
                    {
                        bancoBCP.CodigoPos = "0";
                    }
                    else
                    {
                        bancoBCP.CodigoPos = table.Rows[i - 1][7].ToString();
                    }
                    if (!bancoBCP.CodigoPos.Equals("0"))
                    {
                        if (!string.IsNullOrEmpty(table.Rows[i - 1][1].ToString().Trim()) && table.Rows[i - 1][1].ToString().Trim().Length >= 4)
                        {
                            if (!table.Rows[i - 1][1].ToString().Contains("COM_"))
                            {
                                string empty2 = string.Empty;
                            }
                            if (!table.Rows[i - 1][1].ToString().Equals("VISANET") && !table.Rows[i - 1][1].ToString().Contains("COM_"))
                            {
                                bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim().Substring(table.Rows[i - 1][1].ToString().Trim().Length - 4)
                                    .TrimStart('0')
                                    .Trim();
                            }
                            else
                            {
                                bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                            }
                        }
                        else
                        {
                            bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                        }
                    }
                    else
                    {
                        bancoBCP.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                    }
                    decimal num = default(decimal);
                    if (string.IsNullOrEmpty(table.Rows[i - 1][3].ToString()))
                    {
                        num = default(decimal);
                        if (string.IsNullOrEmpty(table.Rows[i - 1][4].ToString()))
                        {
                            num = default(decimal);
                        }
                        else
                        {
                            num = Math.Round(decimal.Parse(table.Rows[i - 1][4].ToString()), 2);
                            bancoBCP.Dh = 2;
                        }
                    }
                    else
                    {
                        num = Math.Round(decimal.Parse(table.Rows[i - 1][3].ToString()), 2);
                        bancoBCP.Dh = 1;
                    }
                    bancoBCP.Importe = num;
                    if (!string.IsNullOrEmpty(table.Rows[i - 1][3].ToString()) || !string.IsNullOrEmpty(table.Rows[i - 1][4].ToString()))
                    {
                        ListaPenBanco.Add(bancoBCP);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error en Hoja PEN_: " + Environment.NewLine + "En la fila " + i + Environment.NewLine + ex.Message);
                }
            }
            return ListaPenBanco;
        }

        public static List<BancoUpeu> PopulateinCollectionPEN_Banco(DataTable table)
        {
            ListaPenUpeu = new List<BancoUpeu>();
            for (int i = 5; i < table.Rows.Count + 1; i++)
            {
                BancoUpeu bancoUpeu = new BancoUpeu();
                bancoUpeu.Whoyo = "P-Ban";
                bancoUpeu.Pendiente = true;
                bancoUpeu.FechaRegistro = DateTime.Parse(table.Rows[i - 1][0].ToString());
                bancoUpeu.FechaOperacion = table.Rows[i - 1][0].ToString();
                bancoUpeu.Descripcion = table.Rows[i - 1][2].ToString();
                if (string.IsNullOrEmpty(table.Rows[i - 1][7].ToString()))
                {
                    bancoUpeu.CodigoPos = "0";
                }
                else
                {
                    bancoUpeu.CodigoPos = table.Rows[i - 1][7].ToString();
                }
                if (!bancoUpeu.CodigoPos.Equals("0"))
                {
                    if (!string.IsNullOrEmpty(table.Rows[i - 1][1].ToString().Trim()) && table.Rows[i - 1][1].ToString().Trim().Length >= 4)
                    {
                        bancoUpeu.NroOpe = table.Rows[i - 1][1].ToString().Trim().Substring(table.Rows[i - 1][1].ToString().Trim().Length - 4)
                            .TrimStart('0')
                            .Trim();
                    }
                    else
                    {
                        bancoUpeu.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                    }
                }
                else
                {
                    bancoUpeu.NroOpe = table.Rows[i - 1][1].ToString().Trim();
                }
                decimal num = default(decimal);
                if (string.IsNullOrEmpty(table.Rows[i - 1][5].ToString()))
                {
                    num = default(decimal);
                    if (string.IsNullOrEmpty(table.Rows[i - 1][6].ToString()))
                    {
                        num = default(decimal);
                    }
                    else
                    {
                        num = decimal.Parse(table.Rows[i - 1][6].ToString());
                        bancoUpeu.Dh = 2;
                    }
                }
                else
                {
                    num = decimal.Parse(table.Rows[i - 1][5].ToString());
                    bancoUpeu.Dh = 1;
                }
                bancoUpeu.Importe = num;
                if (!string.IsNullOrEmpty(table.Rows[i - 1][5].ToString()) || !string.IsNullOrEmpty(table.Rows[i - 1][6].ToString()))
                {
                    ListaPenUpeu.Add(bancoUpeu);
                }
            }
            return ListaPenUpeu;
        }

        public static List<BancoUpeu> PopulateinCollectionUpeu(DataTable table)
        {
            ListaUpeu = new List<BancoUpeu>();
            Listacom = new List<ComparabancosComision>();
            for (int i = 5; i < table.Rows.Count + 1; i++)
            {
                try
                {
                    BancoUpeu bancoUpeu = new BancoUpeu();
                    bancoUpeu.Whoyo = "M-Upe";
                    bancoUpeu.FechaRegistro = DateTime.Parse(table.Rows[i - 1][0].ToString());
                    bancoUpeu.ReferenciaLibros = table.Rows[i - 1][1].ToString().TrimEnd();
                    bancoUpeu.Descripcion = table.Rows[i - 1][2].ToString();
                    if (bancoUpeu.ReferenciaLibros.Equals("MB 26-2.36"))
                    {
                        string empty = string.Empty;
                    }
                    if (table.Rows[i - 1][4].ToString() != string.Empty)
                    {
                        bancoUpeu.FechaOperacion = DateTime.Parse(table.Rows[i - 1][4].ToString()).ToShortDateString();
                    }
                    else
                    {
                        bancoUpeu.FechaOperacion = table.Rows[i - 1][4].ToString();
                    }
                    decimal num = default(decimal);
                    decimal num2 = default(decimal);
                    num2 = ((!string.IsNullOrEmpty(table.Rows[i - 1][5].ToString())) ? decimal.Parse(table.Rows[i - 1][5].ToString()) : default(decimal));
                    num = ((!string.IsNullOrEmpty(table.Rows[i - 1][6].ToString())) ? decimal.Parse(table.Rows[i - 1][6].ToString()) : default(decimal));
                    if (num2 == 0m)
                    {
                        bancoUpeu.Dh = 2;
                        bancoUpeu.Importe = num;
                    }
                    else
                    {
                        bancoUpeu.Dh = 1;
                        bancoUpeu.Importe = num2;
                    }
                    if (string.IsNullOrEmpty(table.Rows[i - 1][7].ToString()))
                    {
                        bancoUpeu.CodigoPos = "0";
                    }
                    else
                    {
                        bancoUpeu.CodigoPos = table.Rows[i - 1][7].ToString();
                    }
                    if (!bancoUpeu.CodigoPos.Equals("0"))
                    {
                        if (!string.IsNullOrEmpty(table.Rows[i - 1][3].ToString().Trim()) && table.Rows[i - 1][3].ToString().Trim().Length >= 4)
                        {
                            bancoUpeu.NroOpe = table.Rows[i - 1][3].ToString().Trim().Substring(table.Rows[i - 1][3].ToString().Trim().Length - 4)
                                .TrimStart('0')
                                .Trim();
                        }
                        else
                        {
                            bancoUpeu.NroOpe = table.Rows[i - 1][3].ToString().Trim();
                        }
                    }
                    else
                    {
                        bancoUpeu.NroOpe = table.Rows[i - 1][3].ToString().Trim();
                    }
                    ListaUpeu.Add(bancoUpeu);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error en Hoja UPEU: " + Environment.NewLine + "En la fila " + i + Environment.NewLine + ex.Message);
                }
            }
            return ListaUpeu;
        }
              

        public static List<BancoBCP> GetListVisanetNew(DataTable table)
        {
            List<BancoBCP> list = new List<BancoBCP>();
            for (int i = 5; i < table.Rows.Count + 1; i++)
            {
                try
                {
                    BancoBCP bancoBCP = new BancoBCP();
                    bancoBCP.Whoyo = "V-NET";
                    bancoBCP.CodigoPos = table.Rows[i - 1][0].ToString().Trim();
                    bancoBCP.Descripcion = table.Rows[i - 1][1].ToString();
                    bancoBCP.FechaOperacion = DateTime.Parse(table.Rows[i - 1][2].ToString());
                    string text = table.Rows[i - 1][3].ToString();
                    string text2 = table.Rows[i - 1][4].ToString();
                    if (string.IsNullOrEmpty(text))
                    {
                        text = "0";
                    }
                    if (string.IsNullOrEmpty(text2))
                    {
                        text2 = "0";
                    }
                    bancoBCP.Importe = decimal.Parse(text);
                    bancoBCP.NetoAbonar = decimal.Parse(text2);
                    if (bancoBCP.Importe < 0m)
                    {
                        bancoBCP.Dh = 2;
                    }
                    else
                    {
                        bancoBCP.Dh = 1;
                    }
                    bancoBCP.FechaAbono = table.Rows[i - 1][5].ToString();
                    bancoBCP.ReferenciaVoucher = table.Rows[i - 1][6].ToString().Trim();
                    if (!string.IsNullOrEmpty(bancoBCP.ReferenciaVoucher) && bancoBCP.ReferenciaVoucher.Length >= 4)
                    {
                        bancoBCP.NroOpe = bancoBCP.ReferenciaVoucher.Substring(bancoBCP.ReferenciaVoucher.Length - 4).TrimStart('0').Trim();
                    }
                    else if (!string.IsNullOrEmpty(bancoBCP.ReferenciaVoucher) && bancoBCP.ReferenciaVoucher.Length < 4)
                    {
                        bancoBCP.NroOpe = bancoBCP.ReferenciaVoucher;
                    }
                    else
                    {
                        bancoBCP.NroOpe = string.Empty;
                    }
                    list.Add(bancoBCP);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error en Hoja VISA " + Environment.NewLine + "En la fila " + i + Environment.NewLine + ex.Message);
                }
            }
            return list;
        }

        public static List<BancoBCP> GetListMastercardNew(DataTable table)
        {
            List<BancoBCP> list = new List<BancoBCP>();
            for (int i = 5; i < table.Rows.Count + 1; i++)
            {
                try
                {
                    BancoBCP bancoBCP = new BancoBCP();
                    bancoBCP.Whoyo = "M-CARD";
                    bancoBCP.Dh = 1;
                    bancoBCP.CodigoPos = table.Rows[i - 1][0].ToString().Trim();
                    bancoBCP.Descripcion = table.Rows[i - 1][1].ToString().Trim();
                    bancoBCP.FechaOperacion = DateTime.Parse(table.Rows[i - 1][2].ToString());
                    string text = table.Rows[i - 1][3].ToString();
                    string text2 = table.Rows[i - 1][4].ToString();
                    if (string.IsNullOrEmpty(text))
                    {
                        text = "0";
                    }
                    if (string.IsNullOrEmpty(text2))
                    {
                        text2 = "0";
                    }
                    bancoBCP.Importe = decimal.Parse(text);
                    bancoBCP.NetoAbonar = decimal.Parse(text2);
                    if (!string.IsNullOrEmpty(table.Rows[i - 1][5].ToString()))
                    {
                        bancoBCP.FechaAbono = ToNewFormatyyymmdd(table.Rows[i - 1][5].ToString());
                    }
                    bancoBCP.ReferenciaVoucher = table.Rows[i - 1][6].ToString().Trim();
                    if (!string.IsNullOrEmpty(bancoBCP.ReferenciaVoucher) && bancoBCP.ReferenciaVoucher.Length >= 4)
                    {
                        bancoBCP.NroOpe = bancoBCP.ReferenciaVoucher.Substring(bancoBCP.ReferenciaVoucher.Length - 4).TrimStart('0').Trim();
                    }
                    else if (!string.IsNullOrEmpty(bancoBCP.ReferenciaVoucher) && bancoBCP.ReferenciaVoucher.Length < 4)
                    {
                        bancoBCP.NroOpe = bancoBCP.ReferenciaVoucher;
                    }
                    else
                    {
                        bancoBCP.NroOpe = string.Empty;
                    }
                    if (bancoBCP.NroOpe.Equals("6596"))
                    {
                        string empty = string.Empty;
                    }
                    if (!bancoBCP.Descripcion.Equals("99"))
                    {
                        list.Add(bancoBCP);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Error en Hoja MC: " + Environment.NewLine + "En la fila " + i + Environment.NewLine + ex.Message);
                }
            }
            return list;
        }

        public static List<BancoBCP> GetListAmericanEExpNew(DataTable table)
        {
            List<BancoBCP> list = new List<BancoBCP>();
            for (int i = 5; i < table.Rows.Count + 1; i++)
            {
                try
                {
                    BancoBCP bancoBCP = new BancoBCP();
                    bancoBCP.Whoyo = "A-EXP";
                    bancoBCP.Dh = 1;
                    bancoBCP.CodigoPos = table.Rows[i - 1][0].ToString().Trim();
                    bancoBCP.Descripcion = table.Rows[i - 1][1].ToString().Trim();
                    bancoBCP.FechaOperacion = DateTime.Parse(table.Rows[i - 1][2].ToString());
                    string text = table.Rows[i - 1][3].ToString();
                    string text2 = table.Rows[i - 1][4].ToString();
                    if (string.IsNullOrEmpty(text))
                    {
                        text = "0";
                    }
                    if (string.IsNullOrEmpty(text2))
                    {
                        text2 = "0";
                    }
                    bancoBCP.Importe = decimal.Parse(text);
                    bancoBCP.NetoAbonar = decimal.Parse(text2);
                    if (!string.IsNullOrEmpty(table.Rows[i - 1][5].ToString()))
                    {
                        bancoBCP.FechaAbono = ToNewFormatyyymmdd(table.Rows[i - 1][5].ToString());
                    }
                    bancoBCP.ReferenciaVoucher = table.Rows[i - 1][6].ToString().Trim();
                    if (!string.IsNullOrEmpty(bancoBCP.ReferenciaVoucher) && bancoBCP.ReferenciaVoucher.Length >= 4)
                    {
                        bancoBCP.NroOpe = bancoBCP.ReferenciaVoucher.Substring(bancoBCP.ReferenciaVoucher.Length - 4).TrimStart('0').Trim();
                    }
                    else if (!string.IsNullOrEmpty(bancoBCP.ReferenciaVoucher) && bancoBCP.ReferenciaVoucher.Length < 4)
                    {
                        bancoBCP.NroOpe = bancoBCP.ReferenciaVoucher;
                    }
                    else
                    {
                        bancoBCP.NroOpe = string.Empty;
                    }
                    list.Add(bancoBCP);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error en Hoja AE: " + Environment.NewLine + "En la fila " + i + Environment.NewLine + ex.Message);
                }
            }
            return list;
        }

        public static string conectar()
        {
            try
            {
                CadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\\Otros/dtConcilia.accdb;Persist Security Info=False;";
                cnn = new OleDbConnection(CadenaConexion);
                cnn.Open();
                return " Conectado";
            }
            catch
            {
                return " Sin conexión";
            }
        }

        public static string ToNewFormatyyymmdd(string date)
        {
            string empty = string.Empty;
            string empty2 = string.Empty;
            string empty3 = string.Empty;
            string empty4 = string.Empty;
            if (string.IsNullOrEmpty(date.Trim()))
            {
                return string.Empty;
            }
            if (date.Equals("0"))
            {
                return string.Empty;
            }
            try
            {
                empty4 = DateTime.Parse(date).ToString("dd/MM/yyyy");
            }
            catch
            {
                empty = date.Substring(0, 4);
                empty2 = date.Substring(4, 2);
                empty3 = date.Substring(6, 2);
                empty4 = DateTime.ParseExact(empty3 + "/" + empty2 + "/" + empty, "dd/MM/yyyy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy");
            }
            return empty4;
        }

      
        public static List<TablaUpeu> GetDescripcionTerFile()
        {
            List<TablaUpeu> list = new List<TablaUpeu>();
            string path = Application.StartupPath + "\\Terminales.txt";
            List<string> list2 = (from l in File.ReadAllLines(path)
                                  select (l)).ToList();
            foreach (string item2 in list2)
            {
                string terminal = item2.Split('$')[0];
                string descripcion = item2.Split('$')[1];
                TablaUpeu item = new TablaUpeu
                {
                    Terminal = terminal,
                    Descripcion = descripcion
                };
                list.Add(item);
            }
            return list;
        }
    }
}
