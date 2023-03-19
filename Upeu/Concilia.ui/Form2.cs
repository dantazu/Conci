using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using MoreLinq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Globalization;
using System.Reflection;
using System.Collections;
using System.Threading;

namespace Concilia.ui
{
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();
        }

        private bool check;

        private OpenFileDialog diag;

        private List<ComparabancosBcp> ListComparabcp = null;

        private List<ComparabancosUpeu> ListComparaupeu = null;

        private List<ComparabancosComision> listComparaComi = null;

        private List<ComparabancosUpeu> ListPendienteExtra = null;

        private List<ComparabancosUpeu> ListComiAbo = null;

        private List<ComparabancosUpeu> ListComiNoAbo = null;

        private List<ComparabancosUpeu> ListAgrupaNoAbo = null;

        private List<ComparabancosUpeu> ListComiTotAbo = null;

        private List<ComparabancosUpeu> ListComiCOM = null;

        private List<UnidosUpeuyBancos> ListaUnidosUyB = null;

        private List<UnidosUpeuyBancos> ListaUnidosOnly = null;

        private List<BancoUpeu> listaPendBanco = null;

        private List<BancoBCP> listaPendUpeu = null;

        private List<BancoBCP> listaPendUpeuVisanet = null;

        private List<BancoUpeu> listaPendBancoVisa = null;

        private List<BancoBCP> listaBancoGeneral = null;

        private List<BancoUpeu> listaMayorUpeu = null;

        private List<BancoUpeu> listaMayorUpeuT = null;

        private List<BancoUpeu> listaMayorUpeuST = null;

        private List<BancoBCP> listaBancos = null;

        private List<BancoBCP> listaVisaMCAE = null;

        private List<BancoBCP> listaCafetinAux = null;

        private List<BancoBCP> listaVisaMCAEGroup = null;

        private List<BancoBCP> listaBancoVisa = null;

        private List<BancoUpeu> listaMayorVisa = null;

        private List<BancoBCP> listaMasterC = null;

        private List<BancoUpeu> listasPendienteBancoYmayorUpeu = null;

        private List<BancoBCP> listasPendienteUpeuYBancos = null;

        private List<BancoBCP> listaNoexisten = null;

        private List<BancoUpeu> listaNoexisten1 = null;

        private List<BancoBCP> listaFechaMonto = null;

        private List<BancoBCP> listaExisten = null;

        private List<BancoBCP> listaMasterCGroup = null;

        private List<BancoBCP> listaMasterAEx = null;

        private List<BancoBCP> listaMasterAExGroup = null;

        private List<BancoBCP> listaPenUpe = null;

        private List<BancoUpeu> listaPenBanco = null;

        private DataTable dtVisanet;

        private DataTable dtMasterC;

        private DataTable dtMasterAmerEx;

        private DataTableCollection dtc = null;

        private int MesOperacion = 0;

        public int termino = 0;

        private List<BancoBCP> listAgrupavn = null;

        private ComparabancosUpeu objCompupe;

        private ComparabancosBcp objCompbcp;

        private List<ComparabancosBcp> listAuxbcp = null;

        private List<ComparabancosUpeu> listAuxupe = null;
        string msgPen = string.Empty;
        public decimal SumadebeUpeu { get; set; }

        public decimal SumahaberUpeu { get; set; }

        public decimal SumadebeBanco { get; set; }

        public decimal SumahaberBanco { get; set; }

        public string CuentaContableUpeu { get; set; }

        public decimal SaldoIniUpeu { get; set; }

        public decimal SaldoFiNUpeu { get; set; }

        public string CuentaContableBanco { get; set; }

        public decimal SaldoIniBanco { get; set; }

        public decimal SaldoFiNBanco { get; set; }

        public string NombreBanco { get; set; }

        public decimal SaldoIniPendiente { get; set; }

        public DateTime MesTrabajo { get; set; }






        private void Form2_Load(object sender, EventArgs e)
        {
          var objTerminales=  ExcelUtil.GetDescripcionTerFile();
            var term = objTerminales.Where(w => w.Terminal.Equals( "9999999999")).ToList();
            if (term.Count>0)
            {
                string empresa = term.First().Descripcion;
                Image empr=null;
                if (empresa.ToUpper().Equals("MAC"))
                {
                     empr = Concilia.ui.Properties.Resources.banmac;
                }
                else if (empresa.ToUpper().Equals("UPEU"))
                {
                     empr = Concilia.ui.Properties.Resources.banupeu;
                }
             
               // Image empr = Concilia.ui.Properties.Resources.upeu1;
                this.pictureBox2.Image = empr;
            }
            BtnProcesar.Visible = false;
            progressBar1.Visible = false;
            Label7.Visible = false;
            BtnLeerDatos.Visible = false;
            // ExcelUtil.conectar();
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = version.ToString();
        }

        public DateTime GetLastDayOf(DateTime date)
        {
            return new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
        }

        private void ClearLists()
        {
            ListPendienteExtra = new List<ComparabancosUpeu>();
            ListaUnidosUyB = new List<UnidosUpeuyBancos>();
            ListaUnidosOnly = new List<UnidosUpeuyBancos>();
            ListComparaupeu = new List<ComparabancosUpeu>();
            ListComparabcp = new List<ComparabancosBcp>();
        }

        private void ExceltoListNewv(DataTableCollection resultDataColl)
        {
             msgPen = string.Empty;
            new Logconci("Ingresando leer excel ", true);
            listaPendUpeu = new List<BancoBCP>();
            listaPendBanco = new List<BancoUpeu>();
            listaNoexisten = new List<BancoBCP>();
            listaExisten = new List<BancoBCP>();
            listaNoexisten1 = new List<BancoUpeu>();
            listComparaComi = new List<ComparabancosComision>();
            listaVisaMCAE = new List<BancoBCP>();
            listaCafetinAux = new List<BancoBCP>();
            listaVisaMCAEGroup = new List<BancoBCP>();
            listaMasterAExGroup = new List<BancoBCP>();
            listaMasterC = new List<BancoBCP>();
            listaMasterCGroup = new List<BancoBCP>();
            listaFechaMonto = new List<BancoBCP>();
            listaMasterAEx = new List<BancoBCP>();
            listaMayorVisa = new List<BancoUpeu>();
            listaPendBancoVisa = new List<BancoUpeu>();
            listaMayorUpeu = new List<BancoUpeu>();
            listaMayorUpeuT = new List<BancoUpeu>();
            listaMayorUpeuST = new List<BancoUpeu>();
            listaBancoVisa = new List<BancoBCP>();
            listaBancos = new List<BancoBCP>();
            listaBancoGeneral = new List<BancoBCP>();
            listAgrupavn = new List<BancoBCP>();
            listasPendienteBancoYmayorUpeu = new List<BancoUpeu>();
            listasPendienteUpeuYBancos = new List<BancoBCP>();
            ListAgrupaNoAbo = new List<ComparabancosUpeu>();
            listaPendUpeuVisanet = new List<BancoBCP>();
            listaPenUpe = new List<BancoBCP>();
            listaPenBanco = new List<BancoUpeu>();
            List<string> listaSheets = ExcelUtil.ListaSheets;
            ClearLists();
            dtVisanet = new DataTable();
            dtVisanet = resultDataColl[listaSheets[0]];
            try
            {
                MesOperacion = DateTime.Parse(dtVisanet.Rows[2][0].ToString()).Month;
            }
            catch (Exception ex)
            {
                throw new Exception("Error en Hoja VISA: " + Environment.NewLine + ex.Message);
            }
            int count = dtVisanet.Columns.Count;
            if (count >= 9)
            {
                MessageBox.Show("La hoja " + listaSheets[0] + " tiene " + count + " filas Solo debe tener 8  de A a la H ");
                termino = 1;
                return;
            }
            listaCafetinAux = ExcelUtil.GetListVisanetNew(dtVisanet);
            var enumerable = from p in listaCafetinAux
                             group p by new { p.CodigoPos, p.FechaAbono };
            foreach (var item3 in enumerable)
            {
                string codigoPos = item3.Key.CodigoPos;
                string fechaAbono = item3.Key.FechaAbono;
                decimal num = default(decimal);
                decimal num2 = default(decimal);
                foreach (BancoBCP item4 in item3)
                {
                    if (item4.Importe < 0m)
                    {
                        string empty = string.Empty;
                    }
                    if (!item4.NroOpe.Equals(string.Empty))
                    {
                        num2 += item4.Importe;
                    }
                    num += item4.NetoAbonar;
                }
                BancoBCP item2 = new BancoBCP
                {
                    CodigoPos = item3.Key.CodigoPos,
                    FechaAbono = item3.Key.FechaAbono,
                    Importe = num2,
                    NetoAbonar = num,
                    Diferencia = num2 - num,
                    Whoyo = item3.First().Whoyo
                };
                listAgrupavn.Add(item2);
            }
            DataTable dataTable = resultDataColl[listaSheets[1]];
            int count2 = dataTable.Columns.Count;
            if (count2 >= 9)
            {
                MessageBox.Show("La hoja " + listaSheets[1] + " tiene " + count2 + " filas solo debe tener 8  de A a la H ");
                termino = 1;
                return;
            }

            listaPenUpe = ExcelUtil.PopulateinCollectionPEN_Upeu(dataTable);

            decimal SaldoIniPenUpeu = decimal.Parse(dataTable.Rows[3][3].ToString());
            decimal debeUpeu = listaPenUpe.Where(w=>w.Dh==1).Sum(w=>w.Importe);
            decimal haberUpeu = listaPenUpe.Where(w => w.Dh == 2).Sum(w => w.Importe);
            decimal difupeudebhab = (debeUpeu+ SaldoIniPenUpeu) - haberUpeu;


            listaPendUpeu = listaPenUpe.Where((BancoBCP w) => w.NroOpe != "VISANET").ToList();

            listaPendUpeuVisanet = listaPenUpe.Where((BancoBCP w) => w.NroOpe == "VISANET").ToList();

            listaPenBanco = ExcelUtil.PopulateinCollectionPEN_Banco(dataTable);

            decimal SaldoIniPenBanco = decimal.Parse(dataTable.Rows[3][5].ToString());
            decimal debeBanco = listaPenBanco.Where(w => w.Dh == 1).Sum(w => w.Importe);
            decimal haberBanco = listaPenBanco.Where(w => w.Dh == 2).Sum(w => w.Importe);

            decimal difbancodebhab = (debeBanco+ SaldoIniPenBanco) - haberBanco;
            if (difupeudebhab==difbancodebhab)
            {
               
                msgPen = "*";
            }
            else
            {
                msgPen = "Pendientes: plantilla no esta cuadrado ";
            }

            listaPendBanco = listaPenBanco.Where((BancoUpeu w) => w.NroOpe != "VISANET").ToList();

            DataTable dataTable2 = resultDataColl[listaSheets[2]];
            int count3 = dataTable2.Columns.Count;
            if (count3 >= 9)
            {
                MessageBox.Show("La hoja " + listaSheets[2] + " tiene " + count3 + " filas solo debe tener 8  de A a la I ");
                termino = 1;
                return;
            }
            listaMayorUpeu = ExcelUtil.PopulateinCollectionUpeu(dataTable2);
            decimal mayorupeudebe = listaMayorUpeu.Where(w => w.Dh == 1).Sum(W => W.Importe);
            decimal mayorupeuhaber = listaMayorUpeu.Where(w => w.Dh == 2).Sum(W => W.Importe);
            decimal totalupeu = 0;

            List<BancoUpeu> list = listaMayorUpeu.Where((BancoUpeu w) => w.NroOpe.Equals("9076")).ToList();
            listaMayorUpeuST = listaMayorUpeu.Where((BancoUpeu W) => W.CodigoPos == "0").ToList();
            listaMayorUpeuT = listaMayorUpeu.Where((BancoUpeu W) => W.CodigoPos != "0").ToList();
            try
            {
                CuentaContableUpeu = dataTable2.Rows[1][1].ToString();
                SaldoIniUpeu = decimal.Parse(dataTable2.Rows[2][1].ToString());
                SaldoFiNUpeu = decimal.Parse(dataTable2.Rows[3][1].ToString());
                totalupeu =( mayorupeudebe - mayorupeuhaber)+ SaldoIniUpeu;
            }
            catch (Exception ex2)
            {
                MessageBox.Show("La hoja " + listaSheets[2] + " Mala configuracion del encabezado B2: Nro cuenta B3: Saldo Inicial B4: saldo final o" + Environment.NewLine + ex2.Message);
                return;
            }
            DataTable dataTable3 = resultDataColl[listaSheets[3]];
            int count4 = dataTable3.Columns.Count;
            if (count4 >= 6)
            {
                MessageBox.Show("La hoja " + listaSheets[3] + " tiene " + count4 + " filas solo debe tener 5  de A a la E ");
                termino = 1;
                return;
            }

            listaBancoGeneral = ExcelUtil.PopulateinCollectionBCP(dataTable3);
            decimal totalbanco = 0;
            listaBancoVisa = listaBancoGeneral.Where((BancoBCP w) => w.NroOpe.Equals("VISANET")).ToList();
            listaBancos = listaBancoGeneral.Where((BancoBCP w) => w.NroOpe != "VISANET").ToList();
            List<BancoBCP> list2 = listaBancos.Where((BancoBCP w) => w.NroOpe.Equals("11292")).ToList();
            try
            {
                CuentaContableBanco = dataTable3.Rows[1][1].ToString();
                NombreBanco = dataTable3.Rows[1][2].ToString();
                SaldoIniBanco = decimal.Parse(dataTable3.Rows[2][1].ToString());
                SaldoFiNBanco = decimal.Parse(dataTable3.Rows[3][1].ToString());
                MesTrabajo = DateTime.Parse(dataTable3.Rows[2][2].ToString());
                decimal bancodebe = listaBancoGeneral.Where(w => w.Dh == 1).Sum(w => w.Importe);
                decimal bancohabe = listaBancoGeneral.Where(w => w.Dh == 2).Sum(w => w.Importe);
                totalbanco = (bancodebe-bancohabe)+ SaldoIniBanco;
            }
            catch (Exception ex3)
            {
                MessageBox.Show("Error en la  hoja " + listaSheets[3]+" " + ex3.Message+" ojo en cuentacontable,nombre banco saldo inicial,final y fecha");
                termino = 1;
                return;
            }
            if (NombreBanco.Equals(string.Empty))
            {
                MessageBox.Show("Ingrese Nombre del banco ");
                termino = 1;
                return;
            }
            dtMasterC = new DataTable();
            dtMasterC = resultDataColl[listaSheets[4]];
            try
            {
                MesOperacion = DateTime.Parse(dtMasterC.Rows[2][0].ToString()).Month;
            }
            catch (Exception ex4)
            {
                throw new Exception("Error en Hoja MC: " + Environment.NewLine + ex4.Message);
            }
            int count5 = dtMasterC.Columns.Count;
            if (count5 >= 8)
            {
                MessageBox.Show("La hoja " + listaSheets[4] + " tiene " + count5 + " filas Solo debe tener 7  de A a la G ");
                termino = 1;
                return;
            }
            if (SaldoIniBanco != SaldoIniPenBanco)
            {
                msgPen += "-Saldo incial Pendiente banco debe ser igual al saldo inicial de Banco -"+Environment.NewLine;
            }
            if (SaldoIniUpeu != SaldoIniPenUpeu)
            {
                msgPen += "-Saldo incial Pendiente upeu debe ser igual al saldo final de Upeu -" + Environment.NewLine;
            }
            if (totalbanco!=SaldoFiNBanco)
            {//compara saldo final banco con los importes totales de banco si no son iguales esta mal el saldo o falta un registro o mas
                msgPen += "-Importe total banco es diferente al saldo final banco-";
            }
            if (totalupeu != SaldoFiNUpeu)
            {//compara saldo final banco con los importes totales de banco si no son iguales esta mal el saldo o falta un registro o mas
                msgPen += "-Importe total upeu es diferente al saldo final upeu-";
            }
            listaMasterC = ExcelUtil.GetListMastercardNew(dtMasterC);

            List<BancoBCP> list3 = listaMasterC.Where((BancoBCP w) => w.NroOpe == "6596").ToList();
            listaMasterCGroup = (from r in listaMasterC
                                 group r by new { r.CodigoPos, r.FechaAbono, r.Descripcion } into g
                                 orderby g.Key.FechaAbono
                                 select new BancoBCP
                                 {
                                     CodigoPos = g.Key.CodigoPos,
                                     FechaAbono = g.Key.FechaAbono,
                                     FechaOperacion = g.Last().FechaOperacion,
                                     NetoAbonar = g.Sum((BancoBCP w) => w.NetoAbonar),
                                     Importe = g.Sum((BancoBCP w) => w.Importe),
                                     Diferencia = g.Sum((BancoBCP w) => w.Importe) - g.Sum((BancoBCP w) => w.NetoAbonar),
                                     Whoyo = g.First().Whoyo,
                                     Descripcion = g.Key.Descripcion
                                 }).ToList();
            List<BancoBCP> list4 = listaMasterCGroup.Where((BancoBCP w) => w.Importe == 330m).ToList();
            dtMasterAmerEx = new DataTable();
            dtMasterAmerEx = resultDataColl[listaSheets[5]];
            try
            {
                MesOperacion = DateTime.Parse(dtMasterAmerEx.Rows[2][0].ToString()).Month;
            }
            catch (Exception ex5)
            {
                throw new Exception("Error en Hoja AE: " + Environment.NewLine + ex5.Message);
            }
            int count6 = dtMasterC.Columns.Count;
            if (count6 >= 8)
            {
                MessageBox.Show("La hoja " + listaSheets[5] + " tiene " + count6 + " filas Solo debe tener 7  de A a la G ");
                termino = 1;
                return;
            }
            listaMasterAEx = ExcelUtil.GetListAmericanEExpNew(dtMasterAmerEx);
            listaMasterAExGroup = (from r in listaMasterAEx
                                   group r by new { r.CodigoPos, r.FechaAbono, r.Descripcion } into g
                                   orderby g.Key.FechaAbono
                                   select new BancoBCP
                                   {
                                       CodigoPos = g.Key.CodigoPos,
                                       FechaOperacion = g.First().FechaOperacion,
                                       Descripcion = g.First().Descripcion,
                                       FechaAbono = g.Key.FechaAbono,
                                       NetoAbonar = g.Sum((BancoBCP w) => w.NetoAbonar),
                                       Importe = g.Sum((BancoBCP w) => w.Importe),
                                       Diferencia = g.Sum((BancoBCP w) => w.Importe) - g.Sum((BancoBCP w) => w.NetoAbonar),
                                       Whoyo = g.First().Whoyo
                                   }).ToList();
            listaVisaMCAE = listaCafetinAux.Where((BancoBCP w) => w.NroOpe != string.Empty).Concat(listaMasterC.Where((BancoBCP w) => w.NroOpe != string.Empty)).Concat(listaMasterAEx.Where((BancoBCP w) => w.NroOpe != string.Empty))
                .ToList();
            listaVisaMCAEGroup = listAgrupavn.Concat(listaMasterCGroup).Concat(listaMasterAExGroup).ToList();
            List<BancoBCP> list5 = listaVisaMCAEGroup.DistinctBy((BancoBCP x) => x.Terminal).ToList();
            foreach (BancoBCP item in list5)
            {
                List<TablaUpeu> descripcionTerFile = ExcelUtil.GetDescripcionTerFile();
                List<TablaUpeu> list6 = descripcionTerFile.Where((TablaUpeu w) => w.Terminal == item.CodigoPos).ToList();
                if (list6.Count == 0)
                {
                    MessageBox.Show("El terminal " + item.CodigoPos + " No esta registrado en el archivo txt Terminales se encuentra en  Apllication Folder;el archivo contine los terminales y sus nombres serparados por el sigono $");
                    termino = 1;
                    break;
                }
            }
        }
       
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                for (int i = 0; i < 101; i+=10)
                {

                    if (backgroundWorker1.CancellationPending)
                    {
                        e.Cancel = true;
                        break;
                    }

                    backgroundWorker1.ReportProgress(i);

                    //Thread.Sleep(20);
                    ProcesaPrincipal();

                }

                e.Result = 0;

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            new Logconci("Ingresando dibujar excel ", true);
            DivujaExcel();
            clearTxtLbl();
            BtnImportar.Focus();
            MessageBox.Show("Proceso concluido ", "Informacion");
        }
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lblNivel.Text = e.ProgressPercentage + "% complete";
        }

        private void ProcesaPrincipal()
        {
            ClearLists();
            new Logconci("Ingresando  ProcesaVNMCAE ", true);
            ProcesaVNMCAE();
            new Logconci("ok  ProcesaVNMCAE ", true);
            new Logconci("Ingresando  ProcesarAgrupadosT ", true);
            ProcesarAgrupadosT();
            new Logconci("ok  ProcesarAgrupadosT ", true);
            new Logconci("Ingresando  ProcesaMayorBancos ", true);
            ProcesaMayorBancos();
            new Logconci("ok  ProcesaMayorBancos ", true);
            new Logconci("Ingresando  ProcesaMayorUpeu ", true);
            ProcesaMayorUpeu();
            new Logconci("ok  ProcesaMayorUpeu ", true);
            new Logconci("Ingresando  RecorreMayoruUpeuT ", true);
            RecorreMayoruUpeuT();
            new Logconci("ok  RecorreMayoruUpeuT ", true);
            foreach (BancoBCP item in listaPendUpeuVisanet)
            {
                if (item.Descripcion.Equals("9092108-0000942738 DINERS CLUB"))
                {
                    var leo = string.Empty;
                }

                if (!CompruebaSiExisteVisanet(item.NroOpe, item.CodigoPos, item.Importe, item.Descripcion))
                {
                    objCompupe = new ComparabancosUpeu
                    {
                        FechaOpe = item.FechaOperacion,
                        Descripcion = item.Descripcion,
                        NroOpe = item.NroOpe,
                        Importe = item.Importe,
                        Dh = item.Dh,
                        Terminal = item.CodigoPos,
                        Observacion = "PEN-VISANET" + item.Whoyo,
                        Mb = item.ReferenciaVoucher
                    };
                    ListComparaupeu.Add(objCompupe);
                }
            }
            new Logconci("Fin recorrido listaPendUpeuVisanet", true);
        }

        private void Corregir()
        {
            listAuxbcp = new List<ComparabancosBcp>();
            listAuxupe = new List<ComparabancosUpeu>();
            List<string> list = new List<string>();
            foreach (BancoUpeu item in listaMayorUpeu)
            {
                string auxitem = item.ReferenciaLibros.Trim();
                List<ComparabancosUpeu> list2 = ListComparaupeu.Where((ComparabancosUpeu w) => w.Mb == auxitem).ToList();
                List<UnidosUpeuyBancos> list3 = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.ReferenciaLibrosU == auxitem).ToList();
                List<UnidosUpeuyBancos> list4 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.ReferenciaLibrosU == auxitem).ToList();
                if (list2.Count == 0 && list3.Count == 0 && list3.Count == 0)
                {
                    list.Add(item.NroOpe + "-" + item.ReferenciaLibros);
                }
            }
            if (list.Count <= 0)
            {
                return;
            }
            foreach (string item2 in list)
            {
                MessageBox.Show(item2.ToString());
            }
        }

        private void ProcesoUnion()
        {
            List<BancoBCP> list = listaBancos.Where((BancoBCP w) => w.NroOpe.Equals("11292")).ToList();
            listasPendienteUpeuYBancos = listaPendUpeu.Where((BancoBCP w) => w.NroOpe != string.Empty).Concat(listaBancos.Where((BancoBCP w) => w.NroOpe != string.Empty)).ToList();
            List<BancoBCP> list2 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe.Equals("11292")).ToList();
            listasPendienteBancoYmayorUpeu = listaPendBanco.Where((BancoUpeu w) => w.NroOpe != string.Empty).Concat(listaMayorUpeu.Where((BancoUpeu w) => w.NroOpe != string.Empty)).ToList();
        }

        public void ProcesaVNMCAE()
        {
            foreach (BancoBCP item in listaVisaMCAE)
            {
                if (item.NroOpe.Equals("305"))
                {
                    string empty = string.Empty;
                }
                if (item.NroOpe.Equals("7673"))
                {
                    string empty2 = string.Empty;
                }
                List<BancoBCP> list = listaVisaMCAE.Where((BancoBCP W) => W.NroOpe == item.NroOpe && W.CodigoPos == item.CodigoPos && W.Importe == item.Importe).ToList();
                if (list.Count == 1)
                {
                    string codigoPos2 = string.Empty;
                    List<BancoUpeu> list2 = new List<BancoUpeu>();
                    if (item.CodigoPos == "4702775" || item.CodigoPos == "4702775A" || item.CodigoPos == "4702775D")
                    {
                        codigoPos2 = "138122203";
                    }
                    else
                    {
                        codigoPos2 = item.CodigoPos;
                    }
                    list2 = listaMayorUpeuT.Where((BancoUpeu w) => w.NroOpe == item.NroOpe && w.CodigoPos == codigoPos2 && w.Importe == item.Importe).ToList();
                    if (list2.Count == 0)
                    {
                        list2 = listaPendBanco.Where((BancoUpeu w) => w.NroOpe == item.NroOpe && w.CodigoPos == codigoPos2 && w.Importe == item.Importe).ToList();
                    }
                    if (list2.Count == 0)
                    {
                        List<BancoUpeu> listMayor = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe == item.NroOpe && w.CodigoPos == item.CodigoPos).ToList();
                        if (listMayor.Count <= 1)
                        {
                            objCompupe = new ComparabancosUpeu
                            {
                                FechaOpe = item.FechaOperacion,
                                Descripcion = item.Descripcion,
                                NroOpe = item.NroOpe,
                                Importe = ((item.Importe < 0m) ? (-1m * item.Importe) : item.Importe),
                                Dh = item.Dh,
                                Terminal = ((codigoPos2 == item.CodigoPos) ? codigoPos2 : (codigoPos2 + "-" + item.CodigoPos)),
                                Observacion = (item.Whoyo ?? ""),
                                Mb = item.ReferenciaVoucher
                            };
                            ListComparaupeu.Add(objCompupe);
                        }
                        else
                        {
                            // modifiado 06-03-2023
                            decimal sumtotal = listMayor.Sum(w => w.Importe);
                            if (item.Importe == sumtotal)
                            {
                                foreach (var list21 in listMayor)
                                {
                                    UnidosUpeuyBancos unidosUpeuyBancos4 = new UnidosUpeuyBancos();
                                    unidosUpeuyBancos4.NroOpeU = list21.NroOpe;
                                    unidosUpeuyBancos4.FechaRegistroU = list21.FechaRegistro.ToShortDateString();
                                    unidosUpeuyBancos4.ReferenciaLibrosU = list21.ReferenciaLibros;
                                    unidosUpeuyBancos4.DescripcionU = list21.Descripcion;
                                    unidosUpeuyBancos4.FechaOperacionU = list21.FechaOperacion;
                                    unidosUpeuyBancos4.ImporteU = list21.Importe;
                                    unidosUpeuyBancos4.DhU = list21.Dh;
                                    unidosUpeuyBancos4.NroOpeB = string.Empty;
                                    unidosUpeuyBancos4.FechaOperacionB = string.Empty;
                                    unidosUpeuyBancos4.ImporteB = 0;
                                    unidosUpeuyBancos4.DescripcionB = string.Empty;
                                    unidosUpeuyBancos4.DhB = 0;
                                    unidosUpeuyBancos4.Terminal = list21.CodigoPos;
                                    unidosUpeuyBancos4.Wherepath = "Mayoryyyyy";
                                    ListaUnidosUyB.Add(unidosUpeuyBancos4);
                                }
                                UnidosUpeuyBancos unidosUpeuyBancos6 = new UnidosUpeuyBancos();
                                unidosUpeuyBancos6.NroOpeU = item.NroOpe;
                                unidosUpeuyBancos6.FechaRegistroU = string.Empty;
                                unidosUpeuyBancos6.ReferenciaLibrosU = string.Empty;
                                unidosUpeuyBancos6.DescripcionU = string.Empty;
                                unidosUpeuyBancos6.FechaOperacionU = string.Empty;
                                unidosUpeuyBancos6.ImporteU = 0;
                                unidosUpeuyBancos6.DhU = 0;
                                unidosUpeuyBancos6.NroOpeB = item.NroOpe;
                                unidosUpeuyBancos6.FechaOperacionB = item.FechaOperacion.ToShortDateString();
                                unidosUpeuyBancos6.ImporteB = item.Importe;
                                unidosUpeuyBancos6.DescripcionB = item.Descripcion;
                                unidosUpeuyBancos6.Terminal = item.CodigoPos;
                                unidosUpeuyBancos6.DhB = 1;
                                unidosUpeuyBancos6.Wherepath = "Visanetyyyyy";
                                ListaUnidosUyB.Add(unidosUpeuyBancos6);
                            }

                        }
                       
                    }
                    else if (list2.Count == 1)
                    {
                        if (list2.First().NroOpe.Equals("2046994"))
                        {
                            string empty3 = string.Empty;
                        }
                        UnidosUpeuyBancos unidosUpeuyBancos = new UnidosUpeuyBancos();
                        unidosUpeuyBancos.NroOpeU = list2.First().NroOpe;
                        unidosUpeuyBancos.FechaRegistroU = list2.First().FechaOperacion;
                        unidosUpeuyBancos.ReferenciaLibrosU = list2.First().ReferenciaLibros;
                        unidosUpeuyBancos.DescripcionU = list2.First().Descripcion;
                        unidosUpeuyBancos.FechaOperacionU = list2.First().FechaOperacion;
                        unidosUpeuyBancos.ImporteU = list2.First().Importe;
                        unidosUpeuyBancos.DhU = list2.First().Dh;
                        unidosUpeuyBancos.NroOpeB = item.NroOpe;
                        unidosUpeuyBancos.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                        unidosUpeuyBancos.ImporteB = item.Importe;
                        unidosUpeuyBancos.DescripcionB = item.Descripcion;
                        unidosUpeuyBancos.DhB = item.Dh;
                        unidosUpeuyBancos.Terminal = ((codigoPos2 == item.CodigoPos) ? codigoPos2 : codigoPos2);
                        unidosUpeuyBancos.Wherepath = item.Whoyo;
                        ListaUnidosOnly.Add(unidosUpeuyBancos);
                    }
                    else
                    {
                        if (list2.Count <= 1)
                        {
                            continue;
                        }
                        string empty4 = string.Empty;
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = item.FechaOperacion,
                            Descripcion = item.Descripcion,
                            NroOpe = item.NroOpe,
                            Importe = item.Importe,
                            Dh = item.Dh,
                            Terminal = ((item.CodigoPos == codigoPos2) ? codigoPos2 : (codigoPos2 + "-" + item.CodigoPos)),
                            Observacion = ">2" + item.Whoyo,
                            Mb = item.ReferenciaVoucher,
                            Pintar = 100
                        };
                        ListComparaupeu.Add(objCompupe);
                        foreach (BancoUpeu item2 in list2)
                        {
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = item2.ReferenciaLibros,
                                FechaOpe = item2.FechaRegistro,
                                Descripcion = item2.ReferenciaLibros + "-" + item2.Descripcion,
                                NroOpe = item2.NroOpe,
                                Importe = item2.Importe,
                                Dh = ((item2.Dh != 1) ? 1 : 2),
                                Terminal = item2.CodigoPos,
                                Observacion = "dupli o mas en mayour " + item2.Whoyo,
                                Whoyo = item2.Whoyo
                            };
                            ListComparabcp.Add(objCompbcp);
                        }
                    }
                    continue;
                }
                string empty5 = string.Empty;
                string codigoPos = string.Empty;
                if (item.CodigoPos == "4702775" || item.CodigoPos == "4702775A")
                {
                    codigoPos = "138122203";
                }
                else
                {
                    codigoPos = item.CodigoPos;
                }
                if (CompruebaSiExiste(item.NroOpe, codigoPos, item.Importe))
                {
                    continue;
                }
                foreach (BancoBCP itemv in list)
                {
                    decimal num = default(decimal);
                    List<BancoUpeu> list3 = listaMayorUpeuT.Where((BancoUpeu w) => w.NroOpe == itemv.NroOpe && w.Importe == item.Importe && w.CodigoPos == codigoPos).ToList();
                    if (list3.Count == 0)
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = itemv.FechaOperacion,
                            Descripcion = itemv.Descripcion,
                            NroOpe = itemv.NroOpe,
                            Importe = itemv.Importe,
                            Dh = itemv.Dh,
                            Terminal = ((itemv.CodigoPos == codigoPos) ? codigoPos : (codigoPos + "-" + itemv.CodigoPos)),
                            Observacion = itemv.Whoyo,
                            Mb = itemv.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                        continue;
                    }
                    if (list3.Count == 2)
                    {
                        if (list.Count == 2)
                        {
                            objCompupe = new ComparabancosUpeu
                            {
                                FechaOpe = list[0].FechaOperacion,
                                Descripcion = list[0].Descripcion,
                                NroOpe = list[0].NroOpe,
                                Importe = list[0].Importe,
                                Dh = list[0].Dh,
                                Terminal = ((list[0].CodigoPos == codigoPos) ? codigoPos : (codigoPos + "-" + list[0].CodigoPos)),
                                Observacion = list[0].Whoyo,
                                Mb = list[0].ReferenciaVoucher
                            };
                            ListComparaupeu.Add(objCompupe);
                            objCompupe = new ComparabancosUpeu
                            {
                                FechaOpe = list[1].FechaOperacion,
                                Descripcion = list[1].Descripcion,
                                NroOpe = list[1].NroOpe,
                                Importe = list[1].Importe,
                                Dh = list[1].Dh,
                                Terminal = ((list[1].CodigoPos == codigoPos) ? codigoPos : (codigoPos + "-" + list[1].CodigoPos)),
                                Observacion = list[1].Whoyo,
                                Mb = list[1].ReferenciaVoucher
                            };
                            ListComparaupeu.Add(objCompupe);
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = list3[0].ReferenciaLibros,
                                FechaOpe = list3[0].FechaRegistro,
                                Descripcion = list3[0].ReferenciaLibros + "-" + list3[0].Descripcion,
                                NroOpe = list3[0].NroOpe,
                                Importe = list3[0].Importe,
                                Dh = ((list3[0].Dh != 1) ? 1 : 2),
                                Terminal = list3[0].CodigoPos,
                                Observacion = "duplicados" + list3[0].Whoyo
                            };
                            ListComparabcp.Add(objCompbcp);
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = list3[1].ReferenciaLibros,
                                FechaOpe = list3[1].FechaRegistro,
                                Descripcion = list3[1].ReferenciaLibros + "-" + list3[1].Descripcion,
                                NroOpe = list3[1].NroOpe,
                                Importe = list3[1].Importe,
                                Dh = ((list3[1].Dh != 1) ? 1 : 2),
                                Terminal = list3[1].CodigoPos,
                                Observacion = "duplicados" + list3[1].Whoyo
                            };
                            ListComparabcp.Add(objCompbcp);
                        }
                        if (list.Count > 2)
                        {
                            string empty6 = string.Empty;
                        }
                        break;
                    }
                    if (list3.Count > 2)
                    {
                        string empty7 = string.Empty;
                    }
                }
            }
        }

        public void ProcesarAgrupadosT()
        {
            DateTime fechaaux = default(DateTime);
            ListComiAbo = new List<ComparabancosUpeu>();
            ListComiNoAbo = new List<ComparabancosUpeu>();
            List<Cafetin> list = new List<Cafetin>();
            int x = 0;
            foreach (BancoBCP itemv in listaVisaMCAEGroup)
            {
                x++;
                if (NombreBanco.StartsWith("Scoti"))
                {
                    fechaaux = DateTime.Parse(itemv.FechaAbono);
                }
                else if (NombreBanco.StartsWith("BCP"))
                {
                    if (itemv.Whoyo.Equals("V-NET"))
                    {
                        fechaaux = DateTime.Parse(itemv.FechaAbono).AddDays(-1.0);
                    }
                    else if (itemv.Whoyo.Equals("A-EXP"))
                    {
                        fechaaux = DateTime.Parse(itemv.FechaAbono);
                    }
                    else
                    {
                        fechaaux = DateTime.Parse(itemv.FechaAbono);
                    }
                }
                else
                {
                    fechaaux = DateTime.Parse(itemv.FechaAbono).AddDays(-1.0);
                }
                List<BancoBCP> listsearchBanco = new List<BancoBCP>();
                if (itemv.Whoyo == "A-EXP")
                {
                    if (itemv.Importe > 0m)
                    {
                        if (itemv.NetoAbonar == 213.51m)
                        {
                            string empty = string.Empty;
                        }
                        listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar && w.FechaOperacion == fechaaux).ToList();
                        if (listsearchBanco.Count == 0)
                        {
                            listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar && w.Descripcion.Contains("CIA DE SERV")).ToList();
                            if (listsearchBanco.Count == 1)
                            {
                                List<BancoBCP> list2 = listaVisaMCAEGroup.Where((BancoBCP w) => w.NetoAbonar == listsearchBanco.First().Importe && w.FechaAbono == listsearchBanco.First().FechaOperacion.ToShortDateString()).ToList();
                                if (list2.Count == 1)
                                {
                                    listsearchBanco = new List<BancoBCP>();
                                }
                            }
                            else if (listsearchBanco.Count == 2)
                            {
                                foreach (BancoBCP itemse2 in listsearchBanco)
                                {
                                    List<BancoBCP> list3 = listaVisaMCAEGroup.Where((BancoBCP w) => w.NetoAbonar == itemse2.Importe && w.FechaAbono == itemse2.FechaOperacion.ToShortDateString()).ToList();
                                    if (list3.Count == 0)
                                    {
                                        listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemse2.Importe && w.FechaOperacion == itemse2.FechaOperacion).ToList();
                                    }
                                }
                            }
                        }
                        if (listsearchBanco.Count == 0)
                        {
                            List<BancoBCP> list4 = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar).ToList();
                        }
                    }
                }
                else if (itemv.Whoyo == "M-CARD")
                {
                    if (itemv.Importe > 0m)
                    {
                        if (itemv.Importe == 330m)
                        {
                            string empty2 = string.Empty;
                        }
                        listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar && w.FechaOperacion == DateTime.Parse(itemv.FechaAbono)).ToList();
                        if (listsearchBanco.Count == 0)
                        {
                            listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar && w.Descripcion.Contains("PROCESOS DE")).ToList();
                            if (listsearchBanco.Count == 1)
                            {
                                List<BancoBCP> list5 = listaVisaMCAEGroup.Where((BancoBCP w) => w.NetoAbonar == listsearchBanco.First().Importe && w.FechaAbono == listsearchBanco.First().FechaOperacion.ToShortDateString()).ToList();
                                if (list5.Count == 1)
                                {
                                    listsearchBanco = new List<BancoBCP>();
                                }
                            }
                            else if (listsearchBanco.Count == 2)
                            {
                                foreach (BancoBCP itemse in listsearchBanco)
                                {
                                    List<BancoBCP> list6 = listaVisaMCAEGroup.Where((BancoBCP w) => w.NetoAbonar == itemse.Importe && w.FechaAbono == itemse.FechaOperacion.ToShortDateString()).ToList();
                                    if (list6.Count == 0)
                                    {
                                        listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemse.Importe && w.FechaOperacion == itemse.FechaOperacion).ToList();
                                    }
                                }
                            }
                        }
                        if (listsearchBanco.Count == 0)
                        {
                            List<BancoBCP> list7 = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar).ToList();
                            if (list7.Count == 1)
                            {
                                string empty3 = string.Empty;
                            }
                        }
                    }
                }
                else if (itemv.Whoyo == "V-NET")
                {
                    if (itemv.Importe > 0m)
                    {
                        if (itemv.NetoAbonar == 213.51m)
                        {
                            string empty4 = string.Empty;
                        }
                        listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.FechaOperacion == fechaaux && w.Importe == itemv.NetoAbonar).ToList();
                        if (listsearchBanco.Count == 0)
                        {
                            listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar && w.Descripcion.Contains("ENT. CREDIBANK VISANET")).ToList();
                        }
                        if (listsearchBanco.Count == 0)
                        {
                            List<BancoBCP> list8 = listaBancoVisa.Where((BancoBCP w) => w.Importe == itemv.NetoAbonar).ToList();
                            if (list8.Count == 1)
                            {
                                listsearchBanco = list8;
                            }
                        }
                    }
                    else
                    {
                        listsearchBanco = listaBancoVisa.Where((BancoBCP w) => w.FechaOperacion == fechaaux && w.Importe == itemv.NetoAbonar).ToList();
                    }
                }
                else
                {
                    string empty5 = string.Empty;
                }
                if (listsearchBanco.Count == 1)
                {
                    bool flag = CompruebaSiExisteVisanet(listsearchBanco.First().NroOpe, listsearchBanco.First().Terminal, listsearchBanco.First().Importe, listsearchBanco.First().Descripcion);
                    UnidosUpeuyBancos unidosUpeuyBancos = new UnidosUpeuyBancos();
                    unidosUpeuyBancos.NroOpeU = string.Empty;
                    unidosUpeuyBancos.FechaRegistroU = string.Empty;
                    unidosUpeuyBancos.ReferenciaLibrosU = string.Empty;
                    unidosUpeuyBancos.DescripcionU = string.Empty;
                    unidosUpeuyBancos.FechaOperacionU = string.Empty;
                    unidosUpeuyBancos.ImporteU = 0m;
                    unidosUpeuyBancos.DhU = 1;
                    unidosUpeuyBancos.NroOpeB = listsearchBanco.First().NroOpe;
                    unidosUpeuyBancos.FechaOperacionB = listsearchBanco.First().FechaOperacion.ToString("dd/MM/yyyy");
                    unidosUpeuyBancos.ImporteB = listsearchBanco.First().Importe;
                    unidosUpeuyBancos.DescripcionB = listsearchBanco.First().Descripcion;
                    unidosUpeuyBancos.DhB = listsearchBanco.First().Dh;
                    unidosUpeuyBancos.Terminal = listsearchBanco.First().Terminal;
                    unidosUpeuyBancos.Wherepath = "VISA-AGRUPA-B";
                    ListaUnidosOnly.Add(unidosUpeuyBancos);
                    UnidosUpeuyBancos unidosUpeuyBancos2 = new UnidosUpeuyBancos();
                    unidosUpeuyBancos2.NroOpeU = string.Empty;
                    unidosUpeuyBancos2.FechaRegistroU = string.Empty;
                    unidosUpeuyBancos2.ReferenciaLibrosU = string.Empty;
                    unidosUpeuyBancos2.DescripcionU = string.Empty;
                    unidosUpeuyBancos2.FechaOperacionU = string.Empty;
                    unidosUpeuyBancos2.ImporteU = 0m;
                    unidosUpeuyBancos2.DhU = 1;
                    unidosUpeuyBancos2.NroOpeB = "VISANET";
                    unidosUpeuyBancos2.FechaOperacionB = itemv.FechaOperacion.ToString("dd/MM/yyyy");
                    unidosUpeuyBancos2.ImporteB = itemv.NetoAbonar;
                    unidosUpeuyBancos2.DescripcionB = itemv.Whoyo;
                    unidosUpeuyBancos2.DhB = itemv.Dh;
                    unidosUpeuyBancos2.Terminal = itemv.CodigoPos;
                    unidosUpeuyBancos2.Wherepath = "VISA-AGRUPA-G";
                    ListaUnidosOnly.Add(unidosUpeuyBancos2);
                    objCompupe = new ComparabancosUpeu
                    {
                        Descripcion = "VISA AGRUPADOS-" + itemv.Whoyo,
                        NroOpe = string.Empty,
                        Importe = itemv.Diferencia,
                        ImporteAbono = itemv.NetoAbonar,
                        ImporteTransac = itemv.Importe,
                        Dh = 1,
                        Terminal = itemv.CodigoPos,
                        Observacion = "Diferencia",
                        Whoyo = itemv.Whoyo,
                        FechaAbono = itemv.FechaAbono,
                        Mb = itemv.ReferenciaVoucher
                    };
                    ListComiAbo.Add(objCompupe);
                    continue;
                }
                if (listsearchBanco.Count == 2)
                {
                    string text = "Sospechoso";
                    if (!CompruebaSiExiste(listsearchBanco[0].NroOpe, listsearchBanco[0].CodigoPos, listsearchBanco[0].Importe))
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = listsearchBanco[0].FechaOperacion,
                            Descripcion = listsearchBanco[0].Descripcion,
                            NroOpe = listsearchBanco[0].NroOpe,
                            Importe = listsearchBanco[0].Importe,
                            Dh = 1,
                            Terminal = listsearchBanco[0].CodigoPos,
                            Observacion = "duplicado en banco monto y fecha:" + itemv.Whoyo,
                            Mb = listsearchBanco[0].ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = listsearchBanco[1].FechaOperacion,
                            Descripcion = listsearchBanco[1].Descripcion,
                            NroOpe = listsearchBanco[1].NroOpe,
                            Importe = listsearchBanco[1].Importe,
                            Dh = 1,
                            Terminal = listsearchBanco[1].CodigoPos,
                            Observacion = "duplicado en banco monto y fecha:" + itemv.Whoyo,
                            Mb = listsearchBanco[1].ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                    if (!CompruebaSiExisteDup(itemv.Importe, itemv.CodigoPos, itemv.FechaAbono))
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = itemv.FechaOperacion,
                            Descripcion = itemv.Descripcion,
                            NroOpe = itemv.NroOpe,
                            Importe = itemv.NetoAbonar,
                            Dh = 2,
                            Terminal = itemv.CodigoPos,
                            Observacion = "duplicado en banco monto y fecha:" + itemv.Whoyo,
                            Mb = itemv.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                        if (itemv.NetoAbonar != 0m)
                        {
                            objCompupe = new ComparabancosUpeu
                            {
                                Descripcion = "Comision no Abono-" + itemv.Whoyo + ":" + itemv.Descripcion,
                                FechaOpe = DateTime.Parse(itemv.FechaAbono),
                                NroOpe = "COM_" + MesTrabajo.Month.ToString("00") + "_" + itemv.Whoyo,
                                Importe = itemv.Diferencia,
                                ImporteAbono = itemv.NetoAbonar,
                                ImporteTransac = itemv.Importe,
                                Dh = 2,
                                Terminal = itemv.CodigoPos,
                                Observacion = "Comision-No Abonado",
                                Whoyo = itemv.Whoyo,
                                FechaAbono = itemv.FechaAbono,
                                Mb = itemv.ReferenciaVoucher
                            };
                            ListComiNoAbo.Add(objCompupe);
                        }
                    }
                    continue;
                }
                if (listsearchBanco.Count == 3)
                {
                    string text2 = "Sospechoso";
                    continue;
                }
                List<TablaUpeu> descripcionTer = ExcelUtil.GetDescripcionTerFile();
                if (itemv.NetoAbonar != 0m)
                {
                    objCompupe = new ComparabancosUpeu
                    {
                        Descripcion = "Comision no Abono-" + itemv.Whoyo + ":" + itemv.Descripcion,
                        FechaOpe = DateTime.Parse(itemv.FechaAbono),
                        NroOpe = "COM_" + MesTrabajo.Month.ToString("00") + "_" + itemv.Whoyo,
                        Importe = itemv.Diferencia,
                        ImporteAbono = itemv.NetoAbonar,
                        ImporteTransac = itemv.Importe,
                        Dh = 2,
                        Terminal = itemv.CodigoPos,
                        Observacion = "Comision-No Abonado",
                        Whoyo = itemv.Whoyo,
                        FechaAbono = itemv.FechaAbono,
                        Mb = itemv.ReferenciaVoucher
                    };
                    ListComiNoAbo.Add(objCompupe);
                }
                if (itemv.Importe > 0m)
                {
                    string descripcion = descripcionTer.Where((TablaUpeu w) => w.Terminal == itemv.CodigoPos).FirstOrDefault().Descripcion;
                    if (itemv.NetoAbonar == 0m)
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            Descripcion = "Pendiente de Abono del: " + itemv.FechaAbono + " _ " + itemv.Whoyo + " _ " + descripcion + ":" + itemv.Descripcion,
                            FechaOpe = DateTime.Parse(itemv.FechaAbono),
                            NroOpe = "VISANET",
                            Importe = itemv.Importe,
                            ImporteAbono = itemv.NetoAbonar,
                            ImporteTransac = itemv.Importe,
                            Dh = 2,
                            Terminal = itemv.CodigoPos,
                            Observacion = "VISA-AGRUPA-G",
                            Whoyo = itemv.Whoyo,
                            FechaAbono = itemv.FechaAbono,
                            Mb = itemv.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                        ListAgrupaNoAbo.Add(objCompupe);
                    }
                    else
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            Descripcion = "No Abono del: " + itemv.FechaAbono + " _ " + itemv.Whoyo + " _ " + descripcion + ":" + itemv.Descripcion,
                            FechaOpe = DateTime.Parse(itemv.FechaAbono),
                            NroOpe = "VISANET",
                            Importe = itemv.NetoAbonar,
                            ImporteAbono = itemv.NetoAbonar,
                            ImporteTransac = itemv.Importe,
                            Dh = 2,
                            Terminal = itemv.CodigoPos,
                            Observacion = "VISA-AGRUPA-G",
                            Whoyo = itemv.Whoyo,
                            FechaAbono = itemv.FechaAbono,
                            Mb = itemv.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                        ListAgrupaNoAbo.Add(objCompupe);
                    }
                }
            }
            var list9 = (from p in ListComiAbo
                         group p by new { p.Whoyo }).ToList();
            foreach (var item2 in list9)
            {
                decimal num = item2.Sum((ComparabancosUpeu w) => w.Importe);
                string auxnro = string.Empty;
                if (item2.Key.Whoyo.Equals("V-NET"))
                {
                    auxnro = "COM-VN";
                }
                else if (item2.Key.Whoyo.Equals("M-CARD"))
                {
                    auxnro = "COM-MC";
                }
                else if (item2.Key.Whoyo.Equals("A-EXP"))
                {
                    auxnro = "COM-AE";
                }
                List<BancoUpeu> list10 = listaMayorUpeuST.Where((BancoUpeu w) => w.NroOpe == auxnro).ToList();
                if (list10.Count == 1)
                {
                    decimal num2 = list10.Sum((BancoUpeu w) => w.Importe);
                    if (num == num2)
                    {
                        UnidosUpeuyBancos unidosUpeuyBancos3 = new UnidosUpeuyBancos();
                        unidosUpeuyBancos3.NroOpeU = list10.First().NroOpe;
                        unidosUpeuyBancos3.FechaRegistroU = list10.First().FechaRegistro.ToShortDateString();
                        unidosUpeuyBancos3.ReferenciaLibrosU = list10.First().ReferenciaLibros;
                        unidosUpeuyBancos3.DescripcionU = list10.First().Descripcion;
                        unidosUpeuyBancos3.FechaOperacionU = list10.First().FechaOperacion;
                        unidosUpeuyBancos3.ImporteU = list10.First().Importe;
                        unidosUpeuyBancos3.DhU = list10.First().Dh;
                        unidosUpeuyBancos3.NroOpeB = list10.First().NroOpe;
                        unidosUpeuyBancos3.FechaOperacionB = item2.First().FechaOpe.ToShortDateString();
                        unidosUpeuyBancos3.ImporteB = num;
                        unidosUpeuyBancos3.DescripcionB = item2.First().Descripcion;
                        unidosUpeuyBancos3.DhB = item2.First().Dh;
                        unidosUpeuyBancos3.Terminal = item2.First().Terminal;
                        unidosUpeuyBancos3.Wherepath = "Comision";
                        ListaUnidosUyB.Add(unidosUpeuyBancos3);
                        continue;
                    }
                    if (!CompruebaSiExiste(auxnro, "0", num))
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            Descripcion = item2.Key.Whoyo,
                            FechaOpe = DateTime.Parse(item2.First().FechaAbono),
                            NroOpe = auxnro,
                            Importe = num,
                            ImporteAbono = num,
                            ImporteTransac = 0m,
                            Dh = 2,
                            Terminal = "0",
                            Observacion = "Agrupado no hay en MAyor",
                            Whoyo = item2.Key.Whoyo,
                            Mb = string.Empty
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                    foreach (BancoUpeu itemx2 in list10)
                    {
                        List<BancoBCP> list11 = listasPendienteUpeuYBancos.Where((BancoBCP W) => W.NroOpe == itemx2.NroOpe && W.CodigoPos == itemx2.CodigoPos && W.Importe == itemx2.Importe).ToList();
                        if (list11.Count == 0 && !CompruebaSiExiste(itemx2.NroOpe, itemx2.CodigoPos, itemx2.Importe))
                        {
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = itemx2.ReferenciaLibros,
                                Descripcion = itemx2.ReferenciaLibros + "-" + itemx2.Descripcion,
                                FechaOpe = itemx2.FechaRegistro,
                                NroOpe = itemx2.NroOpe,
                                Importe = itemx2.Importe,
                                Dh = ((itemx2.Dh != 1) ? 1 : 2),
                                Terminal = "0",
                                Observacion = "no concila",
                                Whoyo = itemx2.Whoyo
                            };
                            ListComparabcp.Add(objCompbcp);
                        }
                    }
                }
                else if (list10.Count > 1)
                {
                    decimal num3 = list10.Sum((BancoUpeu w) => w.Importe);
                    if (num == num3)
                    {
                        UnidosUpeuyBancos unidosUpeuyBancos4 = new UnidosUpeuyBancos();
                        unidosUpeuyBancos4.NroOpeU = auxnro;
                        unidosUpeuyBancos4.FechaRegistroU = string.Empty;
                        unidosUpeuyBancos4.ReferenciaLibrosU = string.Empty;
                        unidosUpeuyBancos4.DescripcionU = string.Empty;
                        unidosUpeuyBancos4.FechaOperacionU = string.Empty;
                        unidosUpeuyBancos4.ImporteU = 0m;
                        unidosUpeuyBancos4.DhU = 2;
                        unidosUpeuyBancos4.NroOpeB = auxnro;
                        unidosUpeuyBancos4.FechaOperacionB = item2.First().FechaOpe.ToShortDateString();
                        unidosUpeuyBancos4.ImporteB = num;
                        unidosUpeuyBancos4.DescripcionB = item2.First().Descripcion;
                        unidosUpeuyBancos4.DhB = item2.First().Dh;
                        unidosUpeuyBancos4.Terminal = item2.First().Terminal;
                        unidosUpeuyBancos4.Wherepath = "Comision";
                        ListaUnidosUyB.Add(unidosUpeuyBancos4);
                        foreach (BancoUpeu item3 in list10)
                        {
                            UnidosUpeuyBancos unidosUpeuyBancos5 = new UnidosUpeuyBancos();
                            unidosUpeuyBancos5.NroOpeU = item3.NroOpe;
                            unidosUpeuyBancos5.FechaRegistroU = item3.FechaRegistro.ToShortDateString();
                            unidosUpeuyBancos5.ReferenciaLibrosU = item3.ReferenciaLibros;
                            unidosUpeuyBancos5.DescripcionU = item3.Descripcion;
                            unidosUpeuyBancos5.FechaOperacionU = item3.FechaOperacion;
                            unidosUpeuyBancos5.ImporteU = item3.Importe;
                            unidosUpeuyBancos5.DhU = item3.Dh;
                            unidosUpeuyBancos5.NroOpeB = item3.NroOpe;
                            unidosUpeuyBancos5.FechaOperacionB = string.Empty;
                            unidosUpeuyBancos5.ImporteB = 0m;
                            unidosUpeuyBancos5.DescripcionB = string.Empty;
                            unidosUpeuyBancos5.DhB = 1;
                            unidosUpeuyBancos5.Terminal = item3.CodigoPos;
                            unidosUpeuyBancos5.Wherepath = "Comision";
                            ListaUnidosUyB.Add(unidosUpeuyBancos5);
                        }
                        continue;
                    }
                    if (!CompruebaSiExiste(auxnro, "0", num))
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            Descripcion = item2.Key.Whoyo,
                            FechaOpe = DateTime.Parse(item2.First().FechaAbono),
                            NroOpe = auxnro,
                            Importe = num,
                            ImporteAbono = num,
                            ImporteTransac = 0m,
                            Dh = 2,
                            Terminal = "0",
                            Observacion = "Agrupado no hay en MAyor",
                            Whoyo = item2.Key.Whoyo,
                            Mb = string.Empty
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                    foreach (BancoUpeu itemx in list10)
                    {
                        List<BancoBCP> list12 = listasPendienteUpeuYBancos.Where((BancoBCP W) => W.NroOpe == itemx.NroOpe && W.CodigoPos == itemx.CodigoPos && W.Importe == itemx.Importe).ToList();
                        if (list12.Count == 0 && !CompruebaSiExiste(itemx.NroOpe, itemx.CodigoPos, itemx.Importe))
                        {
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = itemx.ReferenciaLibros,
                                Descripcion = itemx.ReferenciaLibros + "-" + itemx.Descripcion,
                                FechaOpe = itemx.FechaRegistro,
                                NroOpe = itemx.NroOpe,
                                Importe = itemx.Importe,
                                Dh = ((itemx.Dh != 1) ? 1 : 2),
                                Terminal = "0",
                                Observacion = "no concila",
                                Whoyo = itemx.Whoyo
                            };
                            ListComparabcp.Add(objCompbcp);
                        }
                    }
                }
                else
                {
                    objCompupe = new ComparabancosUpeu
                    {
                        Descripcion = item2.Key.Whoyo,
                        FechaOpe = DateTime.Parse(item2.First().FechaAbono),
                        NroOpe = auxnro,
                        Importe = num,
                        ImporteAbono = num,
                        ImporteTransac = 0m,
                        Dh = 2,
                        Terminal = "0",
                        Observacion = "Agrupado no hay en MAyor",
                        Whoyo = item2.Key.Whoyo,
                        Mb = string.Empty
                    };
                    ListComparaupeu.Add(objCompupe);
                }
            }
            foreach (BancoBCP item in listaBancoVisa)
            {
                if (item.Descripcion.Equals("0001593187 DINERS CLUB"))
                {
                    string empty6 = string.Empty;
                }
                if (item.Importe == 213.51m)
                {
                    string empty7 = string.Empty;
                }
                if (CompruebaSiExisteVisanet(item.NroOpe, item.CodigoPos, item.Importe, item.Descripcion))
                {
                    continue;
                }
                List<BancoBCP> list13 = listaBancoVisa.Where((BancoBCP w) => w.Importe == item.Importe).ToList();
                if (list13.Count == 1)
                {
                    List<BancoBCP> list14 = listaPendUpeuVisanet.Where((BancoBCP w) => w.Importe == item.Importe).ToList();
                    if (list14.Count == 1)
                    {
                        if (item.Dh != list14.First().Dh)
                        {
                            UnidosUpeuyBancos unidosUpeuyBancos6 = new UnidosUpeuyBancos();
                            unidosUpeuyBancos6.NroOpeU = string.Empty;
                            unidosUpeuyBancos6.FechaRegistroU = string.Empty;
                            unidosUpeuyBancos6.ReferenciaLibrosU = string.Empty;
                            unidosUpeuyBancos6.DescripcionU = string.Empty;
                            unidosUpeuyBancos6.FechaOperacionU = string.Empty;
                            unidosUpeuyBancos6.ImporteU = 0m;
                            unidosUpeuyBancos6.DhU = 1;
                            unidosUpeuyBancos6.NroOpeB = list14.First().NroOpe;
                            unidosUpeuyBancos6.FechaOperacionB = list14.First().FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos6.ImporteB = list14.First().Importe;
                            unidosUpeuyBancos6.DescripcionB = list14.First().Descripcion;
                            unidosUpeuyBancos6.DhB = list14.First().Dh;
                            unidosUpeuyBancos6.Terminal = list14.First().CodigoPos;
                            unidosUpeuyBancos6.Wherepath = list14.First().Whoyo;
                            ListaUnidosOnly.Add(unidosUpeuyBancos6);
                            UnidosUpeuyBancos unidosUpeuyBancos7 = new UnidosUpeuyBancos();
                            unidosUpeuyBancos7.NroOpeU = string.Empty;
                            unidosUpeuyBancos7.FechaRegistroU = string.Empty;
                            unidosUpeuyBancos7.ReferenciaLibrosU = string.Empty;
                            unidosUpeuyBancos7.DescripcionU = string.Empty;
                            unidosUpeuyBancos7.FechaOperacionU = string.Empty;
                            unidosUpeuyBancos7.ImporteU = 0m;
                            unidosUpeuyBancos7.DhU = 1;
                            unidosUpeuyBancos7.NroOpeB = item.NroOpe;
                            unidosUpeuyBancos7.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos7.ImporteB = item.Importe;
                            unidosUpeuyBancos7.DescripcionB = item.Descripcion;
                            unidosUpeuyBancos7.DhB = item.Dh;
                            unidosUpeuyBancos7.Terminal = item.CodigoPos;
                            unidosUpeuyBancos7.Wherepath = item.Whoyo;
                            ListaUnidosOnly.Add(unidosUpeuyBancos7);
                        }
                        else
                        {
                            objCompupe = new ComparabancosUpeu
                            {
                                Descripcion = item.Descripcion,
                                FechaOpe = item.FechaOperacion,
                                NroOpe = item.NroOpe,
                                Importe = item.Importe,
                                Dh = item.Dh,
                                Terminal = item.CodigoPos,
                                Observacion = "dh son iguales",
                                Mb = item.ReferenciaVoucher
                            };
                            ListComparaupeu.Add(objCompupe);
                            objCompupe = new ComparabancosUpeu
                            {
                                Descripcion = list14.First().Descripcion,
                                FechaOpe = list14.First().FechaOperacion,
                                NroOpe = list14.First().NroOpe,
                                Importe = list14.First().Importe,
                                Dh = list14.First().Dh,
                                Terminal = list14.First().CodigoPos,
                                Observacion = "dh son iguales",
                                Mb = list14.First().ReferenciaVoucher
                            };
                            ListComparaupeu.Add(objCompupe);
                        }
                    }
                    else if (list14.Count >= 2)
                    {
                        string empty8 = string.Empty;
                    }
                    else
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            Descripcion = item.Descripcion,
                            FechaOpe = item.FechaOperacion,
                            NroOpe = item.NroOpe,
                            Importe = item.Importe,
                            Dh = item.Dh,
                            Terminal = item.CodigoPos,
                            Observacion = "NO existe en Pend-" + item.Whoyo,
                            Mb = item.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                }
                else
                {
                    objCompupe = new ComparabancosUpeu
                    {
                        Descripcion = item.Descripcion,
                        FechaOpe = item.FechaOperacion,
                        NroOpe = item.NroOpe,
                        Importe = item.Importe,
                        Dh = item.Dh,
                        Terminal = item.CodigoPos,
                        Observacion = "recorrido visa solo: " + item.Whoyo,
                        Mb = item.ReferenciaVoucher
                    };
                    ListComparaupeu.Add(objCompupe);
                }
            }
            foreach (var item4 in (from p in ListComiNoAbo
                                   group p by new { p.Whoyo }).ToList())
            {
                objCompupe = new ComparabancosUpeu
                {
                    Descripcion = "Comision no Abono-" + item4.Key.Whoyo,
                    FechaOpe = DateTime.Parse(item4.First().FechaAbono),
                    NroOpe = item4.First().NroOpe,
                    Importe = item4.Sum((ComparabancosUpeu w) => w.Importe),
                    ImporteAbono = 0m,
                    ImporteTransac = 0m,
                    Dh = 2,
                    Terminal = item4.First().Terminal,
                    Observacion = "Comision-No Abonado",
                    Whoyo = item4.Key.Whoyo,
                    FechaAbono = item4.First().FechaAbono,
                    Mb = string.Empty
                };
                ListComparaupeu.Add(objCompupe);
            }
        }

       
        public void ProcesaMayorBancos()
        {
            foreach (BancoBCP item in listasPendienteUpeuYBancos)
            {
                if (item.NroOpe.Equals("DETRAC0"))
                {
                    string empty = string.Empty;
                }
                if (item.NroOpe.Equals("305"))
                {
                    
                    var xx = string.Empty;
                }
                if (item.NroOpe.Equals("556371"))
                {
                    string empty2 = string.Empty;
                }
                if (item.NroOpe.Equals("2067306"))
                {
                    string empty3 = string.Empty;
                }
                if (item.NroOpe.Equals("139184559"))
                {
                    string empty4 = string.Empty;   
                }
                if (item.NroOpe.Equals("28533"))
                {
                    var xx = string.Empty;
                }
                if (item.NroOpe.Equals("DETRAC"))
                {
                    List<BancoUpeu> list = (from w in listasPendienteBancoYmayorUpeu
                                            where w.NroOpe.Equals(item.NroOpe) && w.CodigoPos == item.CodigoPos
                                            orderby w.Importe descending
                                            select w).ToList();
                    if (list.Count > 0)
                    {
                        decimal num = list.Sum((BancoUpeu w) => w.Importe);
                        if (item.Importe == num)
                        {
                            if (ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeU == item.NroOpe).ToList().Count != 0)
                            {
                                continue;
                            }
                            UnidosUpeuyBancos unidosUpeuyBancos = new UnidosUpeuyBancos();
                            unidosUpeuyBancos.NroOpeU = list.First().NroOpe;
                            unidosUpeuyBancos.FechaRegistroU = string.Empty;
                            unidosUpeuyBancos.ReferenciaLibrosU = string.Empty;
                            unidosUpeuyBancos.DescripcionU = string.Empty;
                            unidosUpeuyBancos.FechaOperacionU = string.Empty;
                            unidosUpeuyBancos.ImporteU = 0m;
                            unidosUpeuyBancos.DhU = 1;
                            unidosUpeuyBancos.NroOpeB = item.NroOpe;
                            unidosUpeuyBancos.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos.ImporteB = item.Importe;
                            unidosUpeuyBancos.DescripcionB = item.Descripcion;
                            unidosUpeuyBancos.DhB = item.Dh;
                            unidosUpeuyBancos.Terminal = item.CodigoPos;
                            unidosUpeuyBancos.Wherepath = "Concilia Detraccion : igual montos";
                            ListaUnidosUyB.Add(unidosUpeuyBancos);
                            foreach (BancoUpeu item2 in list)
                            {
                                if (!CompruebaSiExiste(item2.NroOpe, item2.CodigoPos, item2.Importe))
                                {
                                    UnidosUpeuyBancos unidosUpeuyBancos2 = new UnidosUpeuyBancos();
                                    unidosUpeuyBancos2.NroOpeU = item2.NroOpe;
                                    unidosUpeuyBancos2.FechaRegistroU = item2.FechaRegistro.ToString("dd/MM/yyyy");
                                    unidosUpeuyBancos2.ReferenciaLibrosU = item2.ReferenciaLibros;
                                    unidosUpeuyBancos2.DescripcionU = item2.Descripcion;
                                    unidosUpeuyBancos2.FechaOperacionU = item2.FechaOperacion;
                                    unidosUpeuyBancos2.ImporteU = item2.Importe;
                                    unidosUpeuyBancos2.DhU = item2.Dh;
                                    unidosUpeuyBancos2.NroOpeB = item2.NroOpe;
                                    unidosUpeuyBancos2.FechaOperacionB = string.Empty;
                                    unidosUpeuyBancos2.ImporteB = 0m;
                                    unidosUpeuyBancos2.DescripcionB = string.Empty;
                                    unidosUpeuyBancos2.DhB = 1;
                                    unidosUpeuyBancos2.Terminal = item2.CodigoPos;
                                    unidosUpeuyBancos2.Wherepath = "Concilia Detraccion : igual montos";
                                    ListaUnidosUyB.Add(unidosUpeuyBancos2);
                                }
                            }
                        }
                        else
                        {
                            if (ListComparabcp.Where((ComparabancosBcp w) => w.NroOpe == item.NroOpe).ToList().Count != 0)
                            {
                                continue;
                            }
                            foreach (BancoUpeu itemM in list)
                            {
                                List<BancoBCP> list2 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == itemM.NroOpe && w.Importe == itemM.Importe).ToList();
                                if (list2.Count == 0)
                                {
                                    objCompbcp = new ComparabancosBcp
                                    {
                                        Referencialibros = itemM.ReferenciaLibros,
                                        FechaOpe = itemM.FechaRegistro,
                                        Descripcion = itemM.ReferenciaLibros + "-" + itemM.Descripcion,
                                        NroOpe = itemM.NroOpe,
                                        Importe = itemM.Importe,
                                        Dh = ((itemM.Dh != 1) ? 1 : 2),
                                        Terminal = itemM.CodigoPos,
                                        Observacion = "DETRAC UPEU A PENDI"
                                    };
                                    ListComparabcp.Add(objCompbcp);
                                }
                            }
                            objCompupe = new ComparabancosUpeu
                            {
                                FechaOpe = item.FechaOperacion,
                                Descripcion = item.Descripcion,
                                NroOpe = item.NroOpe,
                                Importe = item.Importe,
                                Dh = item.Dh,
                                Terminal = item.CodigoPos,
                                Observacion = "DETRAC BANCO A PEND" + item.Whoyo,
                                Mb = item.ReferenciaVoucher
                            };
                            ListComparaupeu.Add(objCompupe);
                        }
                    }
                    else
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = item.FechaOperacion,
                            Descripcion = item.Descripcion,
                            NroOpe = item.NroOpe,
                            Importe = item.Importe,
                            Dh = item.Dh,
                            Terminal = item.CodigoPos,
                            Observacion = "DETRAC BANCO A PEND" + item.Whoyo,
                            Mb = item.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                }
                else if (item.NroOpe.Equals("COMIS"))
                {
                    List<BancoUpeu> listM = (from w in listasPendienteBancoYmayorUpeu
                                             where w.NroOpe.Equals(item.NroOpe) && w.CodigoPos == item.CodigoPos
                                             orderby w.Importe descending
                                             select w).ToList();
                    if (listM.Count == 1)
                    {
                        decimal num2 = listM.Sum((BancoUpeu w) => w.Importe);
                        if (item.Importe == num2)
                        {
                            UnidosUpeuyBancos unidosUpeuyBancos3 = new UnidosUpeuyBancos();
                            unidosUpeuyBancos3.NroOpeU = item.NroOpe;
                            unidosUpeuyBancos3.FechaRegistroU = listM.First().FechaRegistro.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos3.ReferenciaLibrosU = listM.First().ReferenciaLibros;
                            unidosUpeuyBancos3.DescripcionU = listM.First().Descripcion;
                            unidosUpeuyBancos3.FechaOperacionU = listM.First().FechaOperacion;
                            unidosUpeuyBancos3.ImporteU = num2;
                            unidosUpeuyBancos3.DhU = listM.First().Dh;
                            unidosUpeuyBancos3.NroOpeB = item.NroOpe;
                            unidosUpeuyBancos3.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos3.ImporteB = item.Importe;
                            unidosUpeuyBancos3.DescripcionB = item.Descripcion;
                            unidosUpeuyBancos3.DhB = item.Dh;
                            unidosUpeuyBancos3.Terminal = item.CodigoPos;
                            unidosUpeuyBancos3.Wherepath = "Concilia Detraccion : igual montos";
                            ListaUnidosUyB.Add(unidosUpeuyBancos3);
                            continue;
                        }
                        List<BancoBCP> list3 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == listM.First().NroOpe && w.CodigoPos == listM.First().CodigoPos && w.Importe == listM.First().Importe).ToList();
                        if (list3.Count == 0)
                        {
                            if (!CompruebaSiExiste(listM.First().NroOpe, listM.First().CodigoPos, listM.First().Importe))
                            {
                                objCompbcp = new ComparabancosBcp
                                {
                                    Referencialibros = listM.First().ReferenciaLibros,
                                    FechaOpe = listM.First().FechaRegistro,
                                    Descripcion = listM.First().ReferenciaLibros + "-" + listM.First().Descripcion,
                                    NroOpe = listM.First().NroOpe,
                                    Importe = listM.First().Importe,
                                    Dh = ((listM.First().Dh != 1) ? 1 : 2),
                                    Terminal = listM.First().CodigoPos,
                                    Observacion = "Comision mayor a pend-" + listM.First().Whoyo
                                };
                                ListComparabcp.Add(objCompbcp);
                            }
                            if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                            {
                                objCompupe = new ComparabancosUpeu
                                {
                                    FechaOpe = item.FechaOperacion,
                                    Descripcion = item.Descripcion,
                                    NroOpe = item.NroOpe,
                                    Importe = item.Importe,
                                    Dh = item.Dh,
                                    Terminal = item.CodigoPos,
                                    Observacion = "COMISION BANCO A PEND" + item.Whoyo,
                                    Mb = item.ReferenciaVoucher
                                };
                                ListComparaupeu.Add(objCompupe);
                            }
                        }
                        else if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                        {
                            objCompupe = new ComparabancosUpeu
                            {
                                FechaOpe = item.FechaOperacion,
                                Descripcion = item.Descripcion,
                                NroOpe = item.NroOpe,
                                Importe = item.Importe,
                                Dh = item.Dh,
                                Terminal = item.CodigoPos,
                                Observacion = "COMISION BANCO A PEND" + item.Whoyo,
                                Mb = item.ReferenciaVoucher
                            };
                            ListComparaupeu.Add(objCompupe);
                        }
                    }
                    else if (listM.Count > 1)
                    {
                        decimal num3 = listM.Sum((BancoUpeu w) => w.Importe);
                        if (item.Importe == num3)
                        {
                            UnidosUpeuyBancos unidosUpeuyBancos4 = new UnidosUpeuyBancos();
                            unidosUpeuyBancos4.NroOpeU = item.NroOpe;
                            unidosUpeuyBancos4.FechaRegistroU = listM.First().FechaRegistro.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos4.ReferenciaLibrosU = listM.First().ReferenciaLibros;
                            unidosUpeuyBancos4.DescripcionU = listM.First().Descripcion;
                            unidosUpeuyBancos4.FechaOperacionU = listM.First().FechaOperacion;
                            unidosUpeuyBancos4.ImporteU = 0m;
                            unidosUpeuyBancos4.DhU = listM.First().Dh;
                            unidosUpeuyBancos4.NroOpeB = item.NroOpe;
                            unidosUpeuyBancos4.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos4.ImporteB = item.Importe;
                            unidosUpeuyBancos4.DescripcionB = item.Descripcion;
                            unidosUpeuyBancos4.DhB = item.Dh;
                            unidosUpeuyBancos4.Terminal = item.CodigoPos;
                            unidosUpeuyBancos4.Wherepath = "Concilia Detraccion : igual montos";
                            ListaUnidosUyB.Add(unidosUpeuyBancos4);
                            foreach (BancoUpeu item3 in listM)
                            {
                                UnidosUpeuyBancos unidosUpeuyBancos5 = new UnidosUpeuyBancos();
                                unidosUpeuyBancos5.NroOpeU = item3.NroOpe;
                                unidosUpeuyBancos5.FechaRegistroU = item3.FechaRegistro.ToString("dd/MM/yyyy");
                                unidosUpeuyBancos5.ReferenciaLibrosU = item3.ReferenciaLibros;
                                unidosUpeuyBancos5.DescripcionU = item3.Descripcion;
                                unidosUpeuyBancos5.FechaOperacionU = item3.FechaOperacion;
                                unidosUpeuyBancos5.ImporteU = item3.Importe;
                                unidosUpeuyBancos5.DhU = item3.Dh;
                                unidosUpeuyBancos5.NroOpeB = item3.NroOpe;
                                unidosUpeuyBancos5.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                unidosUpeuyBancos5.ImporteB = 0m;
                                unidosUpeuyBancos5.DescripcionB = string.Empty;
                                unidosUpeuyBancos5.DhB = 1;
                                unidosUpeuyBancos5.Terminal = item3.CodigoPos;
                                unidosUpeuyBancos5.Wherepath = "COMIS";
                                ListaUnidosUyB.Add(unidosUpeuyBancos5);
                            }
                            continue;
                        }
                        foreach (BancoUpeu item4 in listM)
                        {
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = item4.ReferenciaLibros,
                                FechaOpe = item4.FechaRegistro,
                                Descripcion = item4.ReferenciaLibros + "-" + item4.Descripcion,
                                NroOpe = item4.NroOpe,
                                Importe = item4.Importe,
                                Dh = ((item4.Dh != 1) ? 1 : 2),
                                Terminal = item4.CodigoPos,
                                Observacion = "COMISION UPEU A PENDI"
                            };
                            ListComparabcp.Add(objCompbcp);
                        }
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = item.FechaOperacion,
                            Descripcion = item.Descripcion,
                            NroOpe = item.NroOpe,
                            Importe = item.Importe,
                            Dh = item.Dh,
                            Terminal = item.CodigoPos,
                            Observacion = "COMISION BANCO A PEND" + item.Whoyo,
                            Mb = item.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                    else if (listM.Count == 0)
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            FechaOpe = item.FechaOperacion,
                            Descripcion = item.Descripcion,
                            NroOpe = item.NroOpe,
                            Importe = item.Importe,
                            Dh = item.Dh,
                            Terminal = item.CodigoPos,
                            Observacion = "COMISION BANCO A PEND" + item.Whoyo,
                            Mb = item.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                }
                else if (item.NroOpe.Contains("VISANET"))
                {
                    List<BancoBCP> list4 = (from w in listasPendienteUpeuYBancos
                                            where w.NroOpe.Equals(item.NroOpe)
                                            orderby w.Importe descending
                                            select w).ToList();
                    if (list4.Count == 1)
                    {
                        if (item.Dh == list4.First().Dh)
                        {
                            if (ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeU == item.NroOpe).ToList().Count == 0)
                            {
                                UnidosUpeuyBancos unidosUpeuyBancos6 = new UnidosUpeuyBancos();
                                unidosUpeuyBancos6.NroOpeU = list4.First().NroOpe;
                                unidosUpeuyBancos6.FechaRegistroU = list4.First().FechaOperacion.ToString("dd/MM/yyyy");
                                unidosUpeuyBancos6.ReferenciaLibrosU = string.Empty;
                                unidosUpeuyBancos6.DescripcionU = list4.First().Descripcion;
                                unidosUpeuyBancos6.FechaOperacionU = list4.First().FechaOperacion.ToShortDateString();
                                unidosUpeuyBancos6.ImporteU = list4.First().Importe;
                                unidosUpeuyBancos6.DhU = list4.First().Dh;
                                unidosUpeuyBancos6.NroOpeB = item.NroOpe;
                                unidosUpeuyBancos6.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                unidosUpeuyBancos6.ImporteB = item.Importe;
                                unidosUpeuyBancos6.DescripcionB = item.Descripcion;
                                unidosUpeuyBancos6.DhB = item.Dh;
                                unidosUpeuyBancos6.Terminal = item.CodigoPos;
                                unidosUpeuyBancos6.Wherepath = "Con Fecha y terminal y importe";
                                ListaUnidosOnly.Add(unidosUpeuyBancos6);
                            }
                        }
                        else
                        {
                            UnidosUpeuyBancos unidosUpeuyBancos7 = new UnidosUpeuyBancos();
                            unidosUpeuyBancos7.NroOpeU = item.NroOpe;
                            unidosUpeuyBancos7.FechaRegistroU = item.FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos7.ReferenciaLibrosU = string.Empty;
                            unidosUpeuyBancos7.DescripcionU = string.Empty;
                            unidosUpeuyBancos7.FechaOperacionU = string.Empty;
                            unidosUpeuyBancos7.ImporteU = 0m;
                            unidosUpeuyBancos7.DhU = 1;
                            unidosUpeuyBancos7.NroOpeB = item.NroOpe;
                            unidosUpeuyBancos7.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos7.ImporteB = item.Importe;
                            unidosUpeuyBancos7.DescripcionB = item.Descripcion;
                            unidosUpeuyBancos7.DhB = item.Dh;
                            unidosUpeuyBancos7.Terminal = item.CodigoPos;
                            unidosUpeuyBancos7.Wherepath = "visa debe haber";
                            ListaUnidosUyB.Add(unidosUpeuyBancos7);
                            UnidosUpeuyBancos unidosUpeuyBancos8 = new UnidosUpeuyBancos();
                            unidosUpeuyBancos8.FechaRegistroU = item.FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos8.NroOpeU = item.NroOpe;
                            unidosUpeuyBancos8.ReferenciaLibrosU = string.Empty;
                            unidosUpeuyBancos8.DescripcionU = string.Empty;
                            unidosUpeuyBancos8.FechaOperacionU = string.Empty;
                            unidosUpeuyBancos8.ImporteU = 0m;
                            unidosUpeuyBancos8.DhU = 1;
                            unidosUpeuyBancos8.NroOpeB = list4.First().NroOpe;
                            unidosUpeuyBancos8.FechaOperacionB = list4.First().FechaOperacion.ToString("dd/MM/yyyy");
                            unidosUpeuyBancos8.ImporteB = list4.First().Importe;
                            unidosUpeuyBancos8.DescripcionB = list4.First().Descripcion;
                            unidosUpeuyBancos8.DhB = list4.First().Dh;
                            unidosUpeuyBancos8.Terminal = list4.First().CodigoPos;
                            unidosUpeuyBancos8.Wherepath = "visa debe haber";
                            ListaUnidosUyB.Add(unidosUpeuyBancos8);
                        }
                    }
                    else
                    {
                        objCompupe = new ComparabancosUpeu
                        {
                            Descripcion = item.Descripcion,
                            FechaOpe = item.FechaOperacion,
                            NroOpe = item.NroOpe,
                            Importe = item.Importe,
                            Dh = item.Dh,
                            Terminal = item.CodigoPos,
                            Observacion = "CARACOx: " + item.Whoyo,
                            Mb = item.ReferenciaVoucher
                        };
                        ListComparaupeu.Add(objCompupe);
                    }
                }
                else
                {
                    if (CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                    {
                        continue;
                    }
                    List<BancoBCP> listaduplicado = (from w in listasPendienteUpeuYBancos
                                                     where w.NroOpe.ToUpper().Equals(item.NroOpe.ToUpper()) && w.CodigoPos == item.CodigoPos
                                                     orderby w.Importe descending
                                                     select w).ToList();
                    if (listaduplicado.Count() == 1)
                    {
                        List<BancoUpeu> list5 = (from w in listasPendienteBancoYmayorUpeu
                                                 where w.NroOpe.Equals(item.NroOpe) && w.CodigoPos == item.CodigoPos
                                                 orderby w.Importe descending
                                                 select w).ToList();
                        if (list5.Count == 1)
                        {
                            if (item.Dh == list5.First().Dh)
                            {
                                if (item.Importe == list5.First().Importe)
                                {
                                    if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                    {
                                        UnidosUpeuyBancos unidosUpeuyBancos9 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos9.NroOpeU = list5.First().NroOpe;
                                        unidosUpeuyBancos9.FechaRegistroU = list5.First().FechaRegistro.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos9.ReferenciaLibrosU = list5.First().ReferenciaLibros;
                                        unidosUpeuyBancos9.DescripcionU = list5.First().Descripcion;
                                        unidosUpeuyBancos9.FechaOperacionU = list5.First().FechaOperacion;
                                        unidosUpeuyBancos9.ImporteU = list5.First().Importe;
                                        unidosUpeuyBancos9.DhU = list5.First().Dh;
                                        unidosUpeuyBancos9.NroOpeB = item.NroOpe;
                                        unidosUpeuyBancos9.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos9.ImporteB = item.Importe;
                                        unidosUpeuyBancos9.DescripcionB = item.Descripcion;
                                        unidosUpeuyBancos9.DhB = item.Dh;
                                        unidosUpeuyBancos9.Terminal = item.CodigoPos;
                                        unidosUpeuyBancos9.Wherepath = item.Whoyo;
                                        ListaUnidosOnly.Add(unidosUpeuyBancos9);
                                    }
                                    continue;
                                }
                                if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                {
                                    objCompupe = new ComparabancosUpeu
                                    {
                                        FechaOpe = item.FechaOperacion,
                                        Descripcion = item.Descripcion,
                                        NroOpe = item.NroOpe,
                                        Importe = item.Importe,
                                        Dh = item.Dh,
                                        Terminal = item.CodigoPos,
                                        Observacion = "Cuidado dos a pendientes" + item.Whoyo,
                                        Mb = item.ReferenciaVoucher
                                    };
                                    ListComparaupeu.Add(objCompupe);
                                }
                                if (!CompruebaSiExiste(list5.First().NroOpe, list5.First().CodigoPos, list5.First().Importe))
                                {
                                    objCompbcp = new ComparabancosBcp
                                    {
                                        Referencialibros = list5.First().ReferenciaLibros,
                                        FechaOpe = list5.First().FechaRegistro,
                                        Descripcion = list5.First().ReferenciaLibros + "-" + list5.First().Descripcion,
                                        NroOpe = list5.First().NroOpe,
                                        Importe = list5.First().Importe,
                                        Dh = ((list5.First().Dh != 1) ? 1 : 2),
                                        Terminal = list5.First().CodigoPos,
                                        Observacion = "Cuidado dos a pendientes" + list5.First().Whoyo
                                    };
                                    ListComparabcp.Add(objCompbcp);
                                }
                            }
                            else
                            {
                                string empty5 = string.Empty;
                                objCompupe = new ComparabancosUpeu
                                {
                                    FechaOpe = item.FechaOperacion,
                                    Descripcion = item.Descripcion,
                                    NroOpe = item.NroOpe,
                                    Importe = item.Importe,
                                    Dh = item.Dh,
                                    Terminal = item.CodigoPos,
                                    Observacion = "Banco no esta en mayor :" + item.Whoyo,
                                    Mb = item.ReferenciaVoucher
                                };
                                ListComparaupeu.Add(objCompupe);
                                objCompbcp = new ComparabancosBcp
                                {
                                    Referencialibros = list5.First().ReferenciaLibros,
                                    FechaOpe = list5.First().FechaRegistro,
                                    Descripcion = list5.First().ReferenciaLibros + "-" + list5.First().Descripcion,
                                    NroOpe = list5.First().NroOpe,
                                    Importe = list5.First().Importe,
                                    Dh = ((list5.First().Dh != 1) ? 1 : 2),
                                    Terminal = list5.First().CodigoPos,
                                    Observacion = "Mayor no en banco " + list5.First().Whoyo
                                };
                                ListComparabcp.Add(objCompbcp);
                            }
                        }
                        else if (list5.Count == 2)
                        {
                            if (list5[0].Dh != list5[1].Dh)
                            {
                                decimal num4 = list5[0].Importe - list5[1].Importe;
                                if (item.Importe == num4)
                                {
                                    int num5 = 0;
                                    if (ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeU == item.NroOpe && w.ImporteU == item.Importe && w.Terminal == "0").ToList().Count() != 0)
                                    {
                                        continue;
                                    }
                                    if (item.Dh == 1)
                                    {
                                        BancoUpeu bancoUpeu = list5.Where((BancoUpeu w) => w.Dh == 2).ToList().First();
                                        UnidosUpeuyBancos unidosUpeuyBancos10 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos10.NroOpeU = bancoUpeu.NroOpe;
                                        unidosUpeuyBancos10.FechaRegistroU = bancoUpeu.FechaRegistro.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos10.ReferenciaLibrosU = bancoUpeu.ReferenciaLibros;
                                        unidosUpeuyBancos10.DescripcionU = bancoUpeu.Descripcion;
                                        unidosUpeuyBancos10.FechaOperacionU = bancoUpeu.FechaOperacion;
                                        unidosUpeuyBancos10.ImporteU = bancoUpeu.Importe;
                                        unidosUpeuyBancos10.DhU = bancoUpeu.Dh;
                                        unidosUpeuyBancos10.NroOpeB = item.NroOpe;
                                        unidosUpeuyBancos10.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos10.ImporteB = item.Importe;
                                        unidosUpeuyBancos10.DescripcionB = item.Descripcion;
                                        unidosUpeuyBancos10.DhB = item.Dh;
                                        unidosUpeuyBancos10.Terminal = item.CodigoPos;
                                        unidosUpeuyBancos10.Wherepath = "M-UPEU";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos10);
                                        BancoUpeu bancoUpeu2 = list5.Where((BancoUpeu w) => w.Dh == 1).ToList().First();
                                        UnidosUpeuyBancos unidosUpeuyBancos11 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos11.NroOpeU = bancoUpeu2.NroOpe;
                                        unidosUpeuyBancos11.FechaRegistroU = bancoUpeu2.FechaRegistro.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos11.ReferenciaLibrosU = bancoUpeu2.ReferenciaLibros;
                                        unidosUpeuyBancos11.DescripcionU = bancoUpeu2.Descripcion;
                                        unidosUpeuyBancos11.FechaOperacionU = bancoUpeu2.FechaOperacion;
                                        unidosUpeuyBancos11.ImporteU = bancoUpeu2.Importe;
                                        unidosUpeuyBancos11.DhU = bancoUpeu2.Dh;
                                        unidosUpeuyBancos11.NroOpeB = item.NroOpe;
                                        unidosUpeuyBancos11.FechaOperacionB = string.Empty;
                                        unidosUpeuyBancos11.ImporteB = 0m;
                                        unidosUpeuyBancos11.DescripcionB = string.Empty;
                                        unidosUpeuyBancos11.DhB = 2;
                                        unidosUpeuyBancos11.Terminal = item.CodigoPos;
                                        unidosUpeuyBancos11.Wherepath = bancoUpeu2.Whoyo;
                                        ListaUnidosUyB.Add(unidosUpeuyBancos11);
                                    }
                                    else
                                    {
                                        string empty6 = string.Empty;
                                    }
                                    continue;
                                }
                                if (num4 == 0m)
                                {
                                    int num6 = 0;
                                    foreach (BancoUpeu item5 in list5)
                                    {
                                        if (item5.Dh == item.Dh && item5.Importe == item.Importe)
                                        {
                                            UnidosUpeuyBancos unidosUpeuyBancos12 = new UnidosUpeuyBancos();
                                            unidosUpeuyBancos12.NroOpeU = item5.NroOpe;
                                            unidosUpeuyBancos12.FechaRegistroU = item5.FechaRegistro.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos12.ReferenciaLibrosU = item5.ReferenciaLibros;
                                            unidosUpeuyBancos12.DescripcionU = item5.Descripcion;
                                            unidosUpeuyBancos12.FechaOperacionU = item5.FechaOperacion;
                                            unidosUpeuyBancos12.ImporteU = item5.Importe;
                                            unidosUpeuyBancos12.DhU = item5.Dh;
                                            unidosUpeuyBancos12.NroOpeB = item.NroOpe;
                                            unidosUpeuyBancos12.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos12.ImporteB = item.Importe;
                                            unidosUpeuyBancos12.DescripcionB = item.Descripcion;
                                            unidosUpeuyBancos12.DhB = item.Dh;
                                            unidosUpeuyBancos12.Terminal = item.CodigoPos;
                                            unidosUpeuyBancos12.Wherepath = "M-UPEU";
                                            ListaUnidosUyB.Add(unidosUpeuyBancos12);
                                            num6++;
                                        }
                                        else
                                        {
                                            objCompbcp = new ComparabancosBcp
                                            {
                                                Referencialibros = item5.ReferenciaLibros,
                                                FechaOpe = item5.FechaRegistro,
                                                Descripcion = item5.ReferenciaLibros + "-" + item5.Descripcion,
                                                NroOpe = item5.NroOpe,
                                                Importe = item5.Importe,
                                                Dh = ((item5.Dh != 1) ? 1 : 2),
                                                Terminal = item5.CodigoPos,
                                                Observacion = "Error revisar" + item5.Whoyo
                                            };
                                            ListComparabcp.Add(objCompbcp);
                                        }
                                    }
                                    if (num6 == 0)
                                    {
                                        objCompupe = new ComparabancosUpeu
                                        {
                                            FechaOpe = item.FechaOperacion,
                                            Descripcion = item.Descripcion,
                                            NroOpe = item.NroOpe,
                                            Importe = item.Importe,
                                            Dh = item.Dh,
                                            Terminal = item.CodigoPos,
                                            Observacion = "Error-revisar :" + item.Whoyo,
                                            Mb = item.ReferenciaVoucher
                                        };
                                        ListComparaupeu.Add(objCompupe);
                                    }
                                    continue;
                                }
                                int num7 = 0;
                                foreach (BancoUpeu itemr in list5)
                                {
                                    if (itemr.Dh == item.Dh && itemr.Importe == item.Importe)
                                    {
                                        UnidosUpeuyBancos unidosUpeuyBancos13 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos13.NroOpeU = itemr.NroOpe;
                                        unidosUpeuyBancos13.FechaRegistroU = itemr.FechaRegistro.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos13.ReferenciaLibrosU = itemr.ReferenciaLibros;
                                        unidosUpeuyBancos13.DescripcionU = itemr.Descripcion;
                                        unidosUpeuyBancos13.FechaOperacionU = itemr.FechaOperacion;
                                        unidosUpeuyBancos13.ImporteU = itemr.Importe;
                                        unidosUpeuyBancos13.DhU = itemr.Dh;
                                        unidosUpeuyBancos13.NroOpeB = item.NroOpe;
                                        unidosUpeuyBancos13.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos13.ImporteB = item.Importe;
                                        unidosUpeuyBancos13.DescripcionB = item.Descripcion;
                                        unidosUpeuyBancos13.DhB = item.Dh;
                                        unidosUpeuyBancos13.Terminal = item.CodigoPos;
                                        unidosUpeuyBancos13.Wherepath = "M-UPEU";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos13);
                                        num7++;
                                    }
                                    else
                                    {
                                        List<BancoBCP> list6 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.FechaOperacion == DateTime.Parse(itemr.FechaOperacion) && w.CodigoPos == itemr.CodigoPos && w.Importe == itemr.Importe).ToList();
                                        if (list6.Count == 0 && !CompruebaSiExiste(itemr.NroOpe, itemr.CodigoPos, itemr.Importe))
                                        {
                                            objCompbcp = new ComparabancosBcp
                                            {
                                                FechaOpe = itemr.FechaRegistro,
                                                Descripcion = itemr.ReferenciaLibros + "-" + itemr.Descripcion,
                                                NroOpe = itemr.NroOpe,
                                                Importe = itemr.Importe,
                                                Dh = ((itemr.Dh != 1) ? 1 : 2),
                                                Terminal = itemr.CodigoPos,
                                                Observacion = "Error revisar" + itemr.Whoyo
                                            };
                                            ListComparabcp.Add(objCompbcp);
                                        }
                                    }
                                }
                                if (num7 == 0 && !CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                {
                                    objCompupe = new ComparabancosUpeu
                                    {
                                        FechaOpe = item.FechaOperacion,
                                        Descripcion = item.Descripcion,
                                        NroOpe = item.NroOpe,
                                        Importe = item.Importe,
                                        Dh = item.Dh,
                                        Terminal = item.CodigoPos,
                                        Observacion = "Error-revisar :" + item.Whoyo,
                                        Mb = item.ReferenciaVoucher
                                    };
                                    ListComparaupeu.Add(objCompupe);
                                }
                                continue;
                            }
                            decimal num8 = list5[0].Importe + list5[1].Importe;
                            if (item.Importe == num8)
                            {
                                UnidosUpeuyBancos unidosUpeuyBancos14 = new UnidosUpeuyBancos();
                                unidosUpeuyBancos14.NroOpeU = item.NroOpe;
                                unidosUpeuyBancos14.FechaRegistroU = string.Empty;
                                unidosUpeuyBancos14.ReferenciaLibrosU = string.Empty;
                                unidosUpeuyBancos14.DescripcionU = string.Empty;
                                unidosUpeuyBancos14.FechaOperacionU = string.Empty;
                                unidosUpeuyBancos14.ImporteU = 0m;
                                unidosUpeuyBancos14.DhU = 0;
                                unidosUpeuyBancos14.NroOpeB = item.NroOpe;
                                unidosUpeuyBancos14.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                unidosUpeuyBancos14.ImporteB = item.Importe;
                                unidosUpeuyBancos14.DescripcionB = item.Descripcion;
                                unidosUpeuyBancos14.DhB = item.Dh;
                                unidosUpeuyBancos14.Terminal = item.CodigoPos;
                                unidosUpeuyBancos14.Wherepath = "M-UPEU";
                                ListaUnidosUyB.Add(unidosUpeuyBancos14);
                                UnidosUpeuyBancos unidosUpeuyBancos15 = new UnidosUpeuyBancos();
                                unidosUpeuyBancos15.NroOpeU = list5[0].NroOpe;
                                unidosUpeuyBancos15.FechaRegistroU = list5[0].FechaRegistro.ToString("dd/MM/yyyy");
                                unidosUpeuyBancos15.ReferenciaLibrosU = list5[0].ReferenciaLibros;
                                unidosUpeuyBancos15.DescripcionU = list5[0].Descripcion;
                                unidosUpeuyBancos15.FechaOperacionU = list5[0].FechaOperacion;
                                unidosUpeuyBancos15.ImporteU = list5[0].Importe;
                                unidosUpeuyBancos15.DhU = list5[0].Dh;
                                unidosUpeuyBancos15.NroOpeB = list5[0].NroOpe;
                                unidosUpeuyBancos15.FechaOperacionB = string.Empty;
                                unidosUpeuyBancos15.ImporteB = 0m;
                                unidosUpeuyBancos15.DescripcionB = string.Empty;
                                unidosUpeuyBancos15.DhB = 1;
                                unidosUpeuyBancos15.Terminal = list5[0].CodigoPos;
                                unidosUpeuyBancos15.Wherepath = "M-UPEU";
                                ListaUnidosUyB.Add(unidosUpeuyBancos15);
                                UnidosUpeuyBancos unidosUpeuyBancos16 = new UnidosUpeuyBancos();
                                unidosUpeuyBancos16.NroOpeU = list5[1].NroOpe;
                                unidosUpeuyBancos16.FechaRegistroU = list5[1].FechaRegistro.ToString("dd/MM/yyyy");
                                unidosUpeuyBancos16.ReferenciaLibrosU = list5[1].ReferenciaLibros;
                                unidosUpeuyBancos16.DescripcionU = list5[1].Descripcion;
                                unidosUpeuyBancos16.FechaOperacionU = list5[1].FechaOperacion;
                                unidosUpeuyBancos16.ImporteU = list5[1].Importe;
                                unidosUpeuyBancos16.DhU = list5[1].Dh;
                                unidosUpeuyBancos16.NroOpeB = list5[1].NroOpe;
                                unidosUpeuyBancos16.FechaOperacionB = string.Empty;
                                unidosUpeuyBancos16.ImporteB = 0m;
                                unidosUpeuyBancos16.DescripcionB = string.Empty;
                                unidosUpeuyBancos16.DhB = 1;
                                unidosUpeuyBancos16.Terminal = list5[1].CodigoPos;
                                unidosUpeuyBancos16.Wherepath = "M-UPEU";
                                ListaUnidosUyB.Add(unidosUpeuyBancos16);
                                int num9 = 0;
                                continue;
                            }
                            int num10 = 0;
                            foreach (BancoUpeu item6 in list5)
                            {
                                if (item6.Dh == item.Dh && item6.Importe == item.Importe)
                                {
                                    if (num10 == 0)
                                    {
                                        if (!CompruebaSiExiste(item6.NroOpe, item6.CodigoPos, item6.Importe))
                                        {
                                            UnidosUpeuyBancos unidosUpeuyBancos17 = new UnidosUpeuyBancos();
                                            unidosUpeuyBancos17.NroOpeU = item6.NroOpe;
                                            unidosUpeuyBancos17.FechaRegistroU = item6.FechaRegistro.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos17.ReferenciaLibrosU = item6.ReferenciaLibros;
                                            unidosUpeuyBancos17.DescripcionU = item6.Descripcion;
                                            unidosUpeuyBancos17.FechaOperacionU = item6.FechaOperacion;
                                            unidosUpeuyBancos17.ImporteU = item6.Importe;
                                            unidosUpeuyBancos17.DhU = item6.Dh;
                                            unidosUpeuyBancos17.NroOpeB = item.NroOpe;
                                            unidosUpeuyBancos17.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos17.ImporteB = item.Importe;
                                            unidosUpeuyBancos17.DescripcionB = item.Descripcion;
                                            unidosUpeuyBancos17.DhB = item.Dh;
                                            unidosUpeuyBancos17.Wherepath = "M-UPEU";
                                            unidosUpeuyBancos17.Terminal = item.CodigoPos;
                                            if (item6.NroOpe == item.NroOpe)
                                            {
                                                ListaUnidosOnly.Add(unidosUpeuyBancos17);
                                            }
                                            else
                                            {
                                                ListaUnidosUyB.Add(unidosUpeuyBancos17);
                                            }
                                            num10++;
                                        }
                                    }
                                    else if (list5[0].Importe == list5[1].Importe)
                                    {
                                        if (item.Importe == list5[0].Importe)
                                        {
                                            objCompbcp = new ComparabancosBcp
                                            {
                                                Referencialibros = item6.ReferenciaLibros,
                                                FechaOpe = item6.FechaRegistro,
                                                Descripcion = item6.ReferenciaLibros + "-" + item6.Descripcion,
                                                NroOpe = item6.NroOpe,
                                                Importe = item6.Importe,
                                                Dh = ((item6.Dh != 1) ? 1 : 2),
                                                Terminal = item6.CodigoPos,
                                                Observacion = "1 conciliado y uno pendiente",
                                                Whoyo = item6.Whoyo
                                            };
                                            ListComparabcp.Add(objCompbcp);
                                        }
                                    }
                                    else if (!CompruebaSiExiste(item6.NroOpe, item6.CodigoPos, item6.Importe))
                                    {
                                        objCompbcp = new ComparabancosBcp
                                        {
                                            Referencialibros = item6.ReferenciaLibros,
                                            FechaOpe = item6.FechaRegistro,
                                            Descripcion = item6.ReferenciaLibros + "-" + item6.Descripcion,
                                            NroOpe = item6.NroOpe,
                                            Importe = item6.Importe,
                                            Dh = ((item6.Dh != 1) ? 1 : 2),
                                            Terminal = item6.CodigoPos,
                                            Observacion = "1 conciliado y uno pendiente",
                                            Whoyo = item6.Whoyo
                                        };
                                        ListComparabcp.Add(objCompbcp);
                                    }
                                }
                                else if (!CompruebaSiExiste(item6.NroOpe, item6.CodigoPos, item6.Importe))
                                {
                                    objCompbcp = new ComparabancosBcp
                                    {
                                        Referencialibros = item6.ReferenciaLibros,
                                        FechaOpe = item6.FechaRegistro,
                                        Descripcion = item6.ReferenciaLibros + "-" + item6.Descripcion,
                                        NroOpe = item6.NroOpe,
                                        Importe = item6.Importe,
                                        Dh = ((item6.Dh != 1) ? 1 : 2),
                                        Terminal = item6.CodigoPos,
                                        Observacion = "los dos pendientes y",
                                        Whoyo = item6.Whoyo
                                    };
                                    ListComparabcp.Add(objCompbcp);
                                }
                            }
                            if (num10 == 0 && !CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                            {
                                objCompupe = new ComparabancosUpeu
                                {
                                    FechaOpe = item.FechaOperacion,
                                    Descripcion = item.ReferenciaVoucher + "-" + item.Descripcion,
                                    NroOpe = item.NroOpe,
                                    Importe = item.Importe,
                                    Dh = item.Dh,
                                    Terminal = item.CodigoPos,
                                    Observacion = "duplicado no conciliable:" + item.Whoyo,
                                    Mb = item.ReferenciaVoucher
                                };
                                ListComparaupeu.Add(objCompupe);
                            }
                        }
                        else if (list5.Count == 0)
                        {
                            _ = item.FechaOperacion;
                            if (true)
                            {
                                string Fechaformat = item.FechaOperacion.ToString("dd/MM/yyyy");
                                List<BancoUpeu> listMayorUp = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.FechaOperacion == Fechaformat && w.CodigoPos == item.CodigoPos && w.Importe == item.Importe).ToList();
                                if (listMayorUp.Count == 1)
                                {
                                    List<BancoBCP> list7 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == listMayorUp.First().NroOpe && w.CodigoPos == listMayorUp.First().CodigoPos && w.Importe == listMayorUp.First().Importe).ToList();
                                    if (list7.Count == 0)
                                    {
                                        List<BancoUpeu> list8 = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe == listMayorUp.First().NroOpe && w.CodigoPos == listMayorUp.First().CodigoPos).ToList();
                                        if (list8.Count == 1)
                                        {
                                            if (!CompruebaSiExiste(listMayorUp.First().NroOpe, listMayorUp.First().CodigoPos, listMayorUp.First().Importe))
                                            {
                                                if (item.Dh == listMayorUp.First().Dh)
                                                {
                                                    UnidosUpeuyBancos unidosUpeuyBancos18 = new UnidosUpeuyBancos();
                                                    unidosUpeuyBancos18.NroOpeU = listMayorUp.First().NroOpe;
                                                    unidosUpeuyBancos18.FechaRegistroU = listMayorUp.First().FechaRegistro.ToString("dd/MM/yyyy");
                                                    unidosUpeuyBancos18.ReferenciaLibrosU = listMayorUp.First().ReferenciaLibros;
                                                    unidosUpeuyBancos18.DescripcionU = listMayorUp.First().Descripcion;
                                                    unidosUpeuyBancos18.FechaOperacionU = listMayorUp.First().FechaOperacion;
                                                    unidosUpeuyBancos18.ImporteU = listMayorUp.First().Importe;
                                                    unidosUpeuyBancos18.DhU = listMayorUp.First().Dh;
                                                    unidosUpeuyBancos18.NroOpeB = item.NroOpe;
                                                    unidosUpeuyBancos18.FechaOperacionB = item.FechaOperacion.ToString("dd/MM/yyyy");
                                                    unidosUpeuyBancos18.ImporteB = item.Importe;
                                                    unidosUpeuyBancos18.DescripcionB = item.Descripcion;
                                                    unidosUpeuyBancos18.DhB = item.Dh;
                                                    unidosUpeuyBancos18.Terminal = item.CodigoPos;
                                                    unidosUpeuyBancos18.Wherepath = "M-COM";
                                                    if (item.NroOpe == listMayorUp.First().NroOpe)
                                                    {
                                                        ListaUnidosOnly.Add(unidosUpeuyBancos18);
                                                    }
                                                    else
                                                    {
                                                        ListaUnidosUyB.Add(unidosUpeuyBancos18);
                                                    }
                                                }
                                                else if (listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == listMayorUp.First().NroOpe && w.Importe == listMayorUp.First().Importe && w.CodigoPos == listMayorUp.First().CodigoPos).ToList().Count == 0)
                                                {
                                                    objCompupe = new ComparabancosUpeu
                                                    {
                                                        FechaOpe = item.FechaOperacion,
                                                        Descripcion = item.Descripcion,
                                                        NroOpe = item.NroOpe,
                                                        Importe = item.Importe,
                                                        Dh = item.Dh,
                                                        Terminal = item.CodigoPos,
                                                        Observacion = "No se encuentra en mayor:" + item.Whoyo,
                                                        Mb = item.ReferenciaVoucher
                                                    };
                                                    ListComparaupeu.Add(objCompupe);
                                                    objCompbcp = new ComparabancosBcp
                                                    {
                                                        Referencialibros = listMayorUp.First().ReferenciaLibros,
                                                        FechaOpe = listMayorUp.First().FechaRegistro,
                                                        Descripcion = listMayorUp.First().ReferenciaLibros + "-" + listMayorUp.First().Descripcion,
                                                        NroOpe = listMayorUp.First().NroOpe,
                                                        Importe = listMayorUp.First().Importe,
                                                        Dh = ((listMayorUp.First().Dh != 1) ? 1 : 2),
                                                        Terminal = listMayorUp.First().CodigoPos,
                                                        Observacion = "1 concil y 1 pendiente-1",
                                                        Whoyo = listMayorUp.First().Whoyo
                                                    };
                                                    ListComparabcp.Add(objCompbcp);
                                                }
                                            }
                                            else if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                            {
                                                objCompupe = new ComparabancosUpeu
                                                {
                                                    FechaOpe = item.FechaOperacion,
                                                    Descripcion = item.Descripcion,
                                                    NroOpe = item.NroOpe,
                                                    Importe = item.Importe,
                                                    Dh = item.Dh,
                                                    Terminal = item.CodigoPos,
                                                    Observacion = "Conciliado x Fecha y Monto :" + item.Whoyo,
                                                    Mb = item.ReferenciaVoucher
                                                };
                                                ListComparaupeu.Add(objCompupe);
                                            }
                                        }
                                        else if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                        {
                                            objCompupe = new ComparabancosUpeu
                                            {
                                                FechaOpe = item.FechaOperacion,
                                                Descripcion = item.Descripcion,
                                                NroOpe = item.NroOpe,
                                                Importe = item.Importe,
                                                Dh = item.Dh,
                                                Terminal = item.CodigoPos,
                                                Observacion = "Conciliado x Fecha y Monto :" + item.Whoyo,
                                                Mb = item.ReferenciaVoucher
                                            };
                                            ListComparaupeu.Add(objCompupe);
                                        }
                                    }
                                    else if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                    {
                                        objCompupe = new ComparabancosUpeu
                                        {
                                            FechaOpe = item.FechaOperacion,
                                            Descripcion = item.Descripcion,
                                            NroOpe = item.NroOpe,
                                            Importe = item.Importe,
                                            Dh = item.Dh,
                                            Terminal = item.CodigoPos,
                                            Observacion = "777:" + item.Whoyo,
                                            Mb = item.ReferenciaVoucher
                                        };
                                        ListComparaupeu.Add(objCompupe);
                                    }
                                }
                                else if (listMayorUp.Count == 0)
                                {
                                    if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe) && item.Importe > 0m)
                                    {
                                        objCompupe = new ComparabancosUpeu
                                        {
                                            FechaOpe = item.FechaOperacion,
                                            Descripcion = item.Descripcion,
                                            NroOpe = item.NroOpe,
                                            Importe = item.Importe,
                                            Dh = item.Dh,
                                            Terminal = item.CodigoPos,
                                            Observacion = "No se encuentra en mayor:" + item.Whoyo,
                                            Mb = item.ReferenciaVoucher
                                        };
                                        ListComparaupeu.Add(objCompupe);
                                    }
                                }
                                else
                                {
                                    if (listMayorUp.Count < 2)
                                    {
                                        continue;
                                    }
                                    string empty7 = string.Empty;
                                    foreach (BancoUpeu itemb in listMayorUp)
                                    {
                                        List<BancoBCP> list9 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == itemb.NroOpe && w.CodigoPos == itemb.CodigoPos && w.Importe == itemb.Importe).ToList();
                                        if (list9.Count == 0)
                                        {
                                            List<BancoUpeu> list10 = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe == itemb.NroOpe && w.CodigoPos == itemb.CodigoPos && w.Importe == itemb.Importe).ToList();
                                            if (list10.Count == 0 && !CompruebaSiExiste(itemb.NroOpe, itemb.CodigoPos, itemb.Importe))
                                            {
                                                objCompbcp = new ComparabancosBcp
                                                {
                                                    Referencialibros = itemb.ReferenciaLibros,
                                                    FechaOpe = itemb.FechaRegistro,
                                                    Descripcion = itemb.ReferenciaLibros + "-" + itemb.Descripcion,
                                                    NroOpe = itemb.NroOpe,
                                                    Importe = itemb.Importe,
                                                    Dh = ((itemb.Dh != 1) ? 1 : 2),
                                                    Terminal = itemb.CodigoPos,
                                                    Observacion = "grrrr:" + itemb.Whoyo
                                                };
                                                ListComparabcp.Add(objCompbcp);
                                            }
                                        }
                                    }
                                    if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                    {
                                        objCompupe = new ComparabancosUpeu
                                        {
                                            FechaOpe = item.FechaOperacion,
                                            Descripcion = item.Descripcion,
                                            NroOpe = item.NroOpe,
                                            Importe = item.Importe,
                                            Dh = item.Dh,
                                            Terminal = item.CodigoPos,
                                            Observacion = "DETRAC BANCO A PEND" + item.Whoyo,
                                            Mb = item.ReferenciaVoucher
                                        };
                                        ListComparaupeu.Add(objCompupe);
                                    }
                                }
                            }
                            else
                            {
                                objCompupe = new ComparabancosUpeu
                                {
                                    FechaOpe = item.FechaOperacion,
                                    Descripcion = item.Descripcion,
                                    NroOpe = item.NroOpe,
                                    Importe = item.Importe,
                                    Dh = item.Dh,
                                    Terminal = item.CodigoPos,
                                    Observacion = "No se encuenra en mayor: Fecha mala configuracion" + item.Whoyo,
                                    Mb = item.ReferenciaVoucher
                                };
                                ListComparaupeu.Add(objCompupe);
                            }
                        }
                        else if (list5.Count > 2)
                        {
                            decimal sumtotal = list5.Sum(w => w.Importe);
                            if (item.Importe == sumtotal)
                            {
                                // conciliar otros
                                //modificado 07-03-2023
                                foreach (var list2 in list5)
                                {
                                    UnidosUpeuyBancos unidosUpeuyBancos4 = new UnidosUpeuyBancos();
                                    unidosUpeuyBancos4.NroOpeU = list2.NroOpe;
                                    unidosUpeuyBancos4.FechaRegistroU = list2.FechaRegistro.ToShortDateString();
                                    unidosUpeuyBancos4.ReferenciaLibrosU = list2.ReferenciaLibros;
                                    unidosUpeuyBancos4.DescripcionU = list2.Descripcion;
                                    unidosUpeuyBancos4.FechaOperacionU = list2.FechaOperacion;
                                    unidosUpeuyBancos4.ImporteU = list2.Importe;
                                    unidosUpeuyBancos4.DhU = list2.Dh;
                                    unidosUpeuyBancos4.NroOpeB = string.Empty;
                                    unidosUpeuyBancos4.FechaOperacionB = string.Empty;
                                    unidosUpeuyBancos4.ImporteB = 0;
                                    unidosUpeuyBancos4.DescripcionB = string.Empty;
                                    unidosUpeuyBancos4.DhB = 0;
                                    unidosUpeuyBancos4.Terminal = list2.CodigoPos;
                                    unidosUpeuyBancos4.Wherepath = "Mayorextra";
                                    ListaUnidosUyB.Add(unidosUpeuyBancos4);
                                }
                                UnidosUpeuyBancos unidosUpeuyBancos6 = new UnidosUpeuyBancos();
                                unidosUpeuyBancos6.NroOpeU = item.NroOpe;
                                unidosUpeuyBancos6.FechaRegistroU = string.Empty;
                                unidosUpeuyBancos6.ReferenciaLibrosU = string.Empty;
                                unidosUpeuyBancos6.DescripcionU = string.Empty;
                                unidosUpeuyBancos6.FechaOperacionU = string.Empty;
                                unidosUpeuyBancos6.ImporteU = 0;
                                unidosUpeuyBancos6.DhU = 0;
                                unidosUpeuyBancos6.NroOpeB = item.NroOpe;
                                unidosUpeuyBancos6.FechaOperacionB = item.FechaOperacion.ToShortDateString();
                                unidosUpeuyBancos6.ImporteB = item.Importe;
                                unidosUpeuyBancos6.DescripcionB = item.Descripcion;
                                unidosUpeuyBancos6.Terminal = item.CodigoPos;
                                unidosUpeuyBancos6.DhB = 1;
                                unidosUpeuyBancos6.Wherepath = "Visanetextra";
                                ListaUnidosUyB.Add(unidosUpeuyBancos6);
                            }
                            else
                            { //pendientes
                                objCompupe = new ComparabancosUpeu
                                {
                                    FechaOpe = item.FechaOperacion,
                                    Descripcion = item.Descripcion,
                                    NroOpe = item.NroOpe,
                                    Importe = item.Importe,
                                    Dh = item.Dh,
                                    Terminal = item.CodigoPos,
                                    Observacion = "1 -> n :" + item.Whoyo,
                                    Mb = item.ReferenciaVoucher,
                                    Pintar = 100
                                };
                                ListComparaupeu.Add(objCompupe);
                            }
                            
                        }
                    }
                    else if (listaduplicado.Count == 2)
                    {
                        int num11 = 1;
                        decimal num12 = default(decimal);
                        List<BancoUpeu> list11 = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe.Equals(item.NroOpe) && w.CodigoPos == item.CodigoPos).ToList();
                        if (list11.Count == 1)
                        {
                            num12 = listaduplicado.Sum((BancoBCP w) => w.Importe);
                            if (num12 == list11.First().Importe)
                            {
                                int num13 = 0;
                                foreach (BancoBCP item7 in listaduplicado.OrderByDescending((BancoBCP w) => w.Importe))
                                {
                                    if (item7.NroOpe.Equals("139184559"))
                                    {
                                        string empty8 = string.Empty;
                                    }
                                    if (item7.Dh == list11.First().Dh)
                                    {
                                        if (num13 == 0)
                                        {
                                            bool flag = CompruebaSiExiste(list11.First().NroOpe, list11.First().CodigoPos, list11.First().Importe);
                                            bool flag2 = CompruebaSiExiste(list11.First().NroOpe, list11.First().CodigoPos, list11.First().Importe);
                                            if (!flag && !flag2)
                                            {
                                                UnidosUpeuyBancos unidosUpeuyBancos19 = new UnidosUpeuyBancos();
                                                unidosUpeuyBancos19.NroOpeU = list11.First().NroOpe;
                                                unidosUpeuyBancos19.FechaRegistroU = list11.First().FechaRegistro.ToString("dd/MM/yyyy");
                                                unidosUpeuyBancos19.ReferenciaLibrosU = list11.First().ReferenciaLibros;
                                                unidosUpeuyBancos19.DescripcionU = list11.First().Descripcion;
                                                unidosUpeuyBancos19.FechaOperacionU = list11.First().FechaOperacion;
                                                unidosUpeuyBancos19.ImporteU = list11.First().Importe;
                                                unidosUpeuyBancos19.DhU = list11.First().Dh;
                                                unidosUpeuyBancos19.NroOpeB = item7.NroOpe;
                                                unidosUpeuyBancos19.FechaOperacionB = item7.FechaOperacion.ToString("dd/MM/yyyy");
                                                unidosUpeuyBancos19.ImporteB = item7.Importe;
                                                unidosUpeuyBancos19.DescripcionB = item7.Descripcion;
                                                unidosUpeuyBancos19.DhB = item7.Dh;
                                                unidosUpeuyBancos19.Terminal = item7.CodigoPos;
                                                unidosUpeuyBancos19.Wherepath = "M-COM";
                                                ListaUnidosUyB.Add(unidosUpeuyBancos19);
                                                num13++;
                                            }
                                        }
                                        else
                                        {
                                            UnidosUpeuyBancos unidosUpeuyBancos20 = new UnidosUpeuyBancos();
                                            unidosUpeuyBancos20.NroOpeU = list11.First().NroOpe;
                                            unidosUpeuyBancos20.FechaRegistroU = string.Empty;
                                            unidosUpeuyBancos20.ReferenciaLibrosU = string.Empty;
                                            unidosUpeuyBancos20.DescripcionU = string.Empty;
                                            unidosUpeuyBancos20.FechaOperacionU = string.Empty;
                                            unidosUpeuyBancos20.ImporteU = 0m;
                                            unidosUpeuyBancos20.DhU = 1;
                                            unidosUpeuyBancos20.NroOpeB = item7.NroOpe;
                                            unidosUpeuyBancos20.FechaOperacionB = item7.FechaOperacion.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos20.ImporteB = item7.Importe;
                                            unidosUpeuyBancos20.DescripcionB = item7.Descripcion;
                                            unidosUpeuyBancos20.DhB = item7.Dh;
                                            unidosUpeuyBancos20.Wherepath = "M-COM";
                                            unidosUpeuyBancos20.Terminal = item7.CodigoPos;
                                            ListaUnidosUyB.Add(unidosUpeuyBancos20);
                                        }
                                    }
                                    else
                                    {
                                        string empty9 = string.Empty;
                                    }
                                }
                                continue;
                            }
                            num12 = listaduplicado[0].Importe - listaduplicado[1].Importe;
                            if (num12 == list11.First().Importe)
                            {
                                if (ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeB == item.NroOpe && w.Terminal == item.CodigoPos).ToList().Count() == 0)
                                {
                                    foreach (BancoBCP item8 in listaduplicado)
                                    {
                                        if (item8.Dh == list11.First().Dh)
                                        {
                                            UnidosUpeuyBancos unidosUpeuyBancos21 = new UnidosUpeuyBancos();
                                            unidosUpeuyBancos21.NroOpeU = list11.First().NroOpe;
                                            unidosUpeuyBancos21.FechaRegistroU = list11.First().FechaRegistro.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos21.ReferenciaLibrosU = list11.First().ReferenciaLibros;
                                            unidosUpeuyBancos21.DescripcionU = list11.First().Descripcion;
                                            unidosUpeuyBancos21.FechaOperacionU = list11.First().FechaOperacion;
                                            unidosUpeuyBancos21.ImporteU = list11.First().Importe;
                                            unidosUpeuyBancos21.DhU = list11.First().Dh;
                                            unidosUpeuyBancos21.NroOpeB = item8.NroOpe;
                                            unidosUpeuyBancos21.FechaOperacionB = item8.FechaOperacion.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos21.ImporteB = item8.Importe;
                                            unidosUpeuyBancos21.DescripcionB = item8.Descripcion;
                                            unidosUpeuyBancos21.DhB = item8.Dh;
                                            unidosUpeuyBancos21.Wherepath = "M-COM";
                                            unidosUpeuyBancos21.Terminal = item8.CodigoPos;
                                            ListaUnidosUyB.Add(unidosUpeuyBancos21);
                                        }
                                        else
                                        {
                                            UnidosUpeuyBancos unidosUpeuyBancos22 = new UnidosUpeuyBancos();
                                            unidosUpeuyBancos22.NroOpeU = item8.NroOpe;
                                            unidosUpeuyBancos22.FechaRegistroU = string.Empty;
                                            unidosUpeuyBancos22.ReferenciaLibrosU = string.Empty;
                                            unidosUpeuyBancos22.DescripcionU = string.Empty;
                                            unidosUpeuyBancos22.FechaOperacionU = string.Empty;
                                            unidosUpeuyBancos22.ImporteU = 0m;
                                            unidosUpeuyBancos22.DhU = 1;
                                            unidosUpeuyBancos22.NroOpeB = item8.NroOpe;
                                            unidosUpeuyBancos22.FechaOperacionB = item8.FechaOperacion.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos22.ImporteB = item8.Importe;
                                            unidosUpeuyBancos22.DescripcionB = item8.Descripcion;
                                            unidosUpeuyBancos22.DhB = item8.Dh;
                                            unidosUpeuyBancos22.Terminal = item8.CodigoPos;
                                            unidosUpeuyBancos22.Wherepath = "M-COM";
                                            ListaUnidosUyB.Add(unidosUpeuyBancos22);
                                        }
                                    }
                                }
                                else
                                {
                                    string empty10 = string.Empty;
                                }
                            }
                            else
                            {
                                if (ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeU == item.NroOpe && w.Terminal == item.CodigoPos).ToList().Count() != 0)
                                {
                                    continue;
                                }
                                foreach (BancoBCP item9 in listaduplicado)
                                {
                                    if (item9.Dh == list11.First().Dh && item9.Importe == list11.First().Importe)
                                    {
                                        UnidosUpeuyBancos unidosUpeuyBancos23 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos23.NroOpeU = list11.First().NroOpe;
                                        unidosUpeuyBancos23.FechaRegistroU = list11.First().FechaRegistro.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos23.ReferenciaLibrosU = list11.First().ReferenciaLibros;
                                        unidosUpeuyBancos23.DescripcionU = list11.First().Descripcion;
                                        unidosUpeuyBancos23.FechaOperacionU = list11.First().FechaOperacion;
                                        unidosUpeuyBancos23.ImporteU = list11.First().Importe;
                                        unidosUpeuyBancos23.DhU = list11.First().Dh;
                                        unidosUpeuyBancos23.NroOpeB = item9.NroOpe;
                                        unidosUpeuyBancos23.FechaOperacionB = item9.FechaOperacion.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos23.ImporteB = item9.Importe;
                                        unidosUpeuyBancos23.DescripcionB = item9.Descripcion;
                                        unidosUpeuyBancos23.DhB = item9.Dh;
                                        unidosUpeuyBancos23.Terminal = item9.CodigoPos;
                                        unidosUpeuyBancos23.Wherepath = "M-COM";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos23);
                                    }
                                    else
                                    {
                                        objCompupe = new ComparabancosUpeu
                                        {
                                            FechaOpe = item9.FechaOperacion,
                                            Descripcion = item9.Descripcion,
                                            NroOpe = item9.NroOpe,
                                            Importe = item9.Importe,
                                            Dh = item9.Dh,
                                            Terminal = item9.CodigoPos,
                                            Observacion = "Duplica 1 de 2: " + item.Whoyo,
                                            Whoyo = "M-UPEU",
                                            Mb = item9.ReferenciaVoucher
                                        };
                                        ListComparaupeu.Add(objCompupe);
                                    }
                                }
                            }
                        }
                        else if (list11.Count == 2)
                        {
                            foreach (BancoBCP item10 in listaduplicado)
                            {
                                foreach (BancoUpeu item11 in list11)
                                {
                                    if (item11.Dh == item10.Dh && item11.Importe == item10.Importe)
                                    {
                                        bool flag3 = CompruebaSiExiste(item11.NroOpe, item11.CodigoPos, item11.Importe);
                                        bool flag4 = CompruebaSiExiste(item10.NroOpe, item10.CodigoPos, item10.Importe);
                                        if (!flag3 && !flag4)
                                        {
                                            UnidosUpeuyBancos unidosUpeuyBancos24 = new UnidosUpeuyBancos();
                                            unidosUpeuyBancos24.NroOpeU = item11.NroOpe;
                                            unidosUpeuyBancos24.FechaRegistroU = item11.FechaRegistro.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos24.ReferenciaLibrosU = item11.ReferenciaLibros;
                                            unidosUpeuyBancos24.DescripcionU = item11.Descripcion;
                                            unidosUpeuyBancos24.FechaOperacionU = item11.FechaOperacion;
                                            unidosUpeuyBancos24.ImporteU = item11.Importe;
                                            unidosUpeuyBancos24.DhU = item11.Dh;
                                            unidosUpeuyBancos24.NroOpeB = item10.NroOpe;
                                            unidosUpeuyBancos24.FechaOperacionB = item10.FechaOperacion.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos24.ImporteB = item10.Importe;
                                            unidosUpeuyBancos24.DescripcionB = item10.Descripcion;
                                            unidosUpeuyBancos24.DhB = item10.Dh;
                                            unidosUpeuyBancos24.Terminal = item10.CodigoPos;
                                            unidosUpeuyBancos24.Wherepath = "M-UPEU";
                                            if (item11.NroOpe == item10.NroOpe)
                                            {
                                                ListaUnidosOnly.Add(unidosUpeuyBancos24);
                                            }
                                            else
                                            {
                                                ListaUnidosUyB.Add(unidosUpeuyBancos24);
                                            }
                                        }
                                        else
                                        {
                                            string empty11 = string.Empty;
                                        }
                                    }
                                    else
                                    {
                                        string empty12 = string.Empty;
                                    }
                                }
                            }
                        }
                        else if (list11.Count == 3)
                        {
                            string empty13 = string.Empty;
                        }
                        else if (list11.Count == 0)
                        {
                            if (listaduplicado[0].Dh != listaduplicado[1].Dh)
                            {
                                if (listaduplicado[0].Importe == listaduplicado[1].Importe)
                                {
                                    if (ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeU == listaduplicado[0].NroOpe).ToList().Count == 0)
                                    {
                                        UnidosUpeuyBancos unidosUpeuyBancos25 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos25.NroOpeU = listaduplicado[0].NroOpe;
                                        unidosUpeuyBancos25.FechaRegistroU = string.Empty;
                                        unidosUpeuyBancos25.ReferenciaLibrosU = string.Empty;
                                        unidosUpeuyBancos25.DescripcionU = string.Empty;
                                        unidosUpeuyBancos25.FechaOperacionU = string.Empty;
                                        unidosUpeuyBancos25.ImporteU = 0m;
                                        unidosUpeuyBancos25.DhU = listaduplicado[0].Dh;
                                        unidosUpeuyBancos25.NroOpeB = listaduplicado[0].NroOpe;
                                        unidosUpeuyBancos25.FechaOperacionB = listaduplicado[0].FechaOperacion.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos25.ImporteB = listaduplicado[0].Importe;
                                        unidosUpeuyBancos25.DescripcionB = listaduplicado[0].Descripcion;
                                        unidosUpeuyBancos25.DhB = listaduplicado[0].Dh;
                                        unidosUpeuyBancos25.Terminal = listaduplicado[0].CodigoPos;
                                        unidosUpeuyBancos25.Wherepath = "M-UPEU";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos25);
                                        UnidosUpeuyBancos unidosUpeuyBancos26 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos26.NroOpeU = listaduplicado[1].NroOpe;
                                        unidosUpeuyBancos26.FechaRegistroU = string.Empty;
                                        unidosUpeuyBancos26.ReferenciaLibrosU = string.Empty;
                                        unidosUpeuyBancos26.DescripcionU = string.Empty;
                                        unidosUpeuyBancos26.FechaOperacionU = string.Empty;
                                        unidosUpeuyBancos26.ImporteU = 0m;
                                        unidosUpeuyBancos26.DhU = listaduplicado[1].Dh;
                                        unidosUpeuyBancos26.NroOpeB = listaduplicado[1].NroOpe;
                                        unidosUpeuyBancos26.FechaOperacionB = listaduplicado[1].FechaOperacion.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos26.ImporteB = listaduplicado[1].Importe;
                                        unidosUpeuyBancos26.DescripcionB = listaduplicado[1].Descripcion;
                                        unidosUpeuyBancos26.DhB = listaduplicado[1].Dh;
                                        unidosUpeuyBancos26.Terminal = listaduplicado[1].CodigoPos;
                                        unidosUpeuyBancos26.Wherepath = "M-UPEU";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos26);
                                    }
                                }
                                else if (ListComparabcp.Where((ComparabancosBcp w) => w.NroOpe == item.NroOpe).ToList().Count == 0)
                                {
                                    objCompupe = new ComparabancosUpeu
                                    {
                                        FechaOpe = listaduplicado[0].FechaOperacion,
                                        Descripcion = listaduplicado[0].Descripcion,
                                        NroOpe = listaduplicado[0].NroOpe,
                                        Importe = listaduplicado[0].Importe,
                                        Dh = listaduplicado[0].Dh,
                                        Terminal = listaduplicado[0].CodigoPos,
                                        Observacion = "No se enc. en banco y :" + listaduplicado[0].Whoyo,
                                        Whoyo = "M-UPEU",
                                        Mb = listaduplicado[0].ReferenciaVoucher
                                    };
                                    ListComparaupeu.Add(objCompupe);
                                    objCompupe = new ComparabancosUpeu
                                    {
                                        FechaOpe = listaduplicado[1].FechaOperacion,
                                        Descripcion = listaduplicado[1].Descripcion,
                                        NroOpe = listaduplicado[1].NroOpe,
                                        Importe = listaduplicado[1].Importe,
                                        Dh = listaduplicado[1].Dh,
                                        Terminal = listaduplicado[1].CodigoPos,
                                        Observacion = "No se enc. en banco y :" + listaduplicado[1].Whoyo,
                                        Whoyo = "M-UPEU",
                                        Mb = listaduplicado[1].ReferenciaVoucher
                                    };
                                    ListComparaupeu.Add(objCompupe);
                                }
                            }
                            else if (listaduplicado[0].Importe > 0m && listaduplicado[1].Importe > 0m && ListComparaupeu.Where((ComparabancosUpeu w) => w.NroOpe == item.NroOpe).ToList().Count == 0)
                            {
                                objCompupe = new ComparabancosUpeu
                                {
                                    FechaOpe = listaduplicado[0].FechaOperacion,
                                    Descripcion = listaduplicado[0].Descripcion,
                                    NroOpe = listaduplicado[0].NroOpe,
                                    Importe = listaduplicado[0].Importe,
                                    Dh = listaduplicado[0].Dh,
                                    Terminal = listaduplicado[0].CodigoPos,
                                    Observacion = "no hay en mayor",
                                    Whoyo = listaduplicado[0].Whoyo,
                                    Mb = listaduplicado[0].ReferenciaVoucher
                                };
                                ListComparaupeu.Add(objCompupe);
                                objCompupe = new ComparabancosUpeu
                                {
                                    FechaOpe = listaduplicado[1].FechaOperacion,
                                    Descripcion = listaduplicado[1].Descripcion,
                                    NroOpe = listaduplicado[1].NroOpe,
                                    Importe = listaduplicado[1].Importe,
                                    Dh = listaduplicado[1].Dh,
                                    Terminal = listaduplicado[1].CodigoPos,
                                    Observacion = "no hay en mayor",
                                    Whoyo = listaduplicado[1].Whoyo,
                                    Mb = listaduplicado[1].ReferenciaVoucher
                                };
                                ListComparaupeu.Add(objCompupe);
                            }
                        }
                        else
                        {
                            string empty14 = string.Empty;
                        }
                    }
                    else
                    {
                        if (listaduplicado.Count < 3 || listaduplicado.First().NroOpe.Equals("VISANET") || listaduplicado.First().NroOpe.Equals("DETRAC") || CompruebaSiExiste(listaduplicado.First().NroOpe, listaduplicado.First().CodigoPos, listaduplicado.First().Importe))
                        {
                            continue;
                        }
                        List<BancoUpeu> list12 = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe.Equals(item.NroOpe) && w.CodigoPos == item.CodigoPos).ToList();
                        if (listaduplicado.Count == list12.Count)
                        {
                            string empty15 = string.Empty;
                        }
                        int num14 = 0;
                        foreach (BancoBCP item12 in listaduplicado)
                        {
                            objCompupe = new ComparabancosUpeu
                            {
                                Descripcion = item12.Descripcion,
                                FechaOpe = item12.FechaOperacion,
                                NroOpe = item12.NroOpe,
                                Importe = item12.Importe,
                                Dh = item12.Dh,
                                Terminal = item12.CodigoPos,
                                Observacion = "> 2 rows  " + item.Whoyo + " : " + num14,
                                Mb = item.ReferenciaVoucher,
                                Pintar = 100
                            };
                            ListComparaupeu.Add(objCompupe);
                            num14++;
                        }
                    }
                }
            }
        }

        public void ProcesaMayorUpeu()
        {
            foreach (BancoUpeu item in listasPendienteBancoYmayorUpeu)
            {
                if (item.NroOpe.Equals("COM-AE"))
                {
                    string empty = string.Empty;
                }
                if (item.NroOpe.Equals("1569"))
                {
                    var xx = string.Empty;
                }
                if (item.NroOpe.Equals("COM-VN") || item.NroOpe.Equals("COM-MC") || item.NroOpe.Equals("COM-AE"))
                {
                    string empty2 = string.Empty;
                    continue;
                }
                if (item.NroOpe.Equals("60770543"))
                {
                    string empty3 = string.Empty;
                }
                if (item.NroOpe.Equals("433123"))
                {
                    string empty4 = string.Empty;
                }
                List<BancoUpeu> list = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe == item.NroOpe && w.CodigoPos == item.CodigoPos).ToList();
                if (list.Count < 3)
                {
                    if (CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                    {
                        continue;
                    }
                    List<BancoUpeu> list2 = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe == item.NroOpe && w.CodigoPos == item.CodigoPos).ToList();
                    if (list2.Count == 1)
                    {
                        List<BancoBCP> list3 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == item.NroOpe && w.CodigoPos == item.CodigoPos).ToList();
                        if (list3.Count == 0)
                        {
                            if (!string.IsNullOrEmpty(item.FechaOperacion))
                            {
                                list3 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.FechaOperacion == DateTime.Parse(item.FechaOperacion) && w.CodigoPos == item.CodigoPos && w.Importe == item.Importe).ToList();
                                if (list3.Count == 1)
                                {
                                    if (!CompruebaSiExiste(list3.First().NroOpe, list3.First().CodigoPos, list3.First().Importe))
                                    {
                                        if (item.Dh == list3.First().Dh)
                                        {
                                            UnidosUpeuyBancos unidosUpeuyBancos = new UnidosUpeuyBancos();
                                            unidosUpeuyBancos.NroOpeU = item.NroOpe;
                                            unidosUpeuyBancos.FechaRegistroU = item.FechaOperacion;
                                            unidosUpeuyBancos.ReferenciaLibrosU = item.ReferenciaLibros;
                                            unidosUpeuyBancos.DescripcionU = item.Descripcion;
                                            unidosUpeuyBancos.FechaOperacionU = item.FechaOperacion;
                                            unidosUpeuyBancos.ImporteU = item.Importe;
                                            unidosUpeuyBancos.DhU = item.Dh;
                                            unidosUpeuyBancos.NroOpeB = list3.First().NroOpe;
                                            unidosUpeuyBancos.FechaOperacionB = list3.First().FechaOperacion.ToString("dd/MM/yyyy");
                                            unidosUpeuyBancos.ImporteB = list3.First().Importe;
                                            unidosUpeuyBancos.DescripcionB = list3.First().Descripcion;
                                            unidosUpeuyBancos.DhB = list3.First().Dh;
                                            unidosUpeuyBancos.Terminal = item.CodigoPos;
                                            unidosUpeuyBancos.Wherepath = "Con Fecha y terminal y importe";
                                            ListaUnidosOnly.Add(unidosUpeuyBancos);
                                            continue;
                                        }
                                        bool flag = CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe);
                                        bool flag2 = CompruebaSiExiste(list3.First().NroOpe, list3.First().CodigoPos, list3.First().Importe);
                                        if (!flag)
                                        {
                                            objCompbcp = new ComparabancosBcp
                                            {
                                                Referencialibros = item.ReferenciaLibros,
                                                FechaOpe = item.FechaRegistro,
                                                Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                                                NroOpe = item.NroOpe,
                                                Importe = item.Importe,
                                                Dh = ((item.Dh != 1) ? 1 : 2),
                                                Terminal = item.CodigoPos,
                                                Observacion = "No se ecuentra en Bancos dh dif:" + item.Whoyo
                                            };
                                            ListComparabcp.Add(objCompbcp);
                                        }
                                        if (!flag2)
                                        {
                                            objCompupe = new ComparabancosUpeu
                                            {
                                                FechaOpe = list3.First().FechaOperacion,
                                                Descripcion = list3.First().Descripcion,
                                                NroOpe = list3.First().NroOpe,
                                                Importe = list3.First().Importe,
                                                Dh = list3.First().Dh,
                                                Terminal = list3.First().CodigoPos,
                                                Observacion = "No se ecuentra en Bancos d y h dif:" + item.Whoyo,
                                                Mb = list3.First().ReferenciaVoucher
                                            };
                                            ListComparaupeu.Add(objCompupe);
                                        }
                                    }
                                    else if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                    {
                                        objCompbcp = new ComparabancosBcp
                                        {
                                            Referencialibros = item.ReferenciaLibros,
                                            FechaOpe = item.FechaRegistro,
                                            Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                                            NroOpe = item.NroOpe,
                                            Importe = item.Importe,
                                            Dh = ((item.Dh != 1) ? 1 : 2),
                                            Terminal = item.CodigoPos,
                                            Observacion = "MU en MB X Fecha y Monto :" + item.Whoyo
                                        };
                                        ListComparabcp.Add(objCompbcp);
                                    }
                                    continue;
                                }
                                if (list3.Count == 2)
                                {
                                    int num = 0;
                                    foreach (BancoBCP item2 in list3)
                                    {
                                        if (item2.NroOpe.Equals("561467054"))
                                        {
                                            string empty5 = string.Empty;
                                        }
                                        if (item2.Dh == item.Dh && item2.Importe == item.Importe)
                                        {
                                            if (!CompruebaSiExiste(item2.NroOpe, item2.CodigoPos, item2.Importe))
                                            {
                                                UnidosUpeuyBancos unidosUpeuyBancos2 = new UnidosUpeuyBancos();
                                                unidosUpeuyBancos2.NroOpeU = item.NroOpe;
                                                unidosUpeuyBancos2.FechaRegistroU = item.FechaRegistro.ToString("dd/MM/yyyy");
                                                unidosUpeuyBancos2.ReferenciaLibrosU = item.ReferenciaLibros;
                                                unidosUpeuyBancos2.DescripcionU = item.Descripcion;
                                                unidosUpeuyBancos2.FechaOperacionU = item.FechaOperacion;
                                                unidosUpeuyBancos2.ImporteU = item.Importe;
                                                unidosUpeuyBancos2.DhU = item.Dh;
                                                unidosUpeuyBancos2.NroOpeB = item2.NroOpe;
                                                unidosUpeuyBancos2.FechaOperacionB = item2.FechaOperacion.ToString("dd/MM/yyyy");
                                                unidosUpeuyBancos2.ImporteB = item2.Importe;
                                                unidosUpeuyBancos2.DescripcionB = item2.Descripcion;
                                                unidosUpeuyBancos2.DhB = item2.Dh;
                                                unidosUpeuyBancos2.Terminal = item2.CodigoPos;
                                                unidosUpeuyBancos2.Wherepath = "M-COM";
                                                if (item.NroOpe == item2.NroOpe)
                                                {
                                                    ListaUnidosOnly.Add(unidosUpeuyBancos2);
                                                }
                                                else
                                                {
                                                    ListaUnidosUyB.Add(unidosUpeuyBancos2);
                                                }
                                                num++;
                                            }
                                        }
                                        else if (!CompruebaSiExiste(item2.NroOpe, item2.CodigoPos, item2.Importe))
                                        {
                                            objCompupe = new ComparabancosUpeu
                                            {
                                                FechaOpe = item2.FechaOperacion,
                                                Descripcion = item2.Descripcion,
                                                NroOpe = item2.NroOpe,
                                                Importe = item2.Importe,
                                                Dh = item2.Dh,
                                                Terminal = item2.CodigoPos,
                                                Observacion = "c" + item2.Whoyo,
                                                Mb = item2.ReferenciaVoucher
                                            };
                                            ListComparaupeu.Add(objCompupe);
                                        }
                                    }
                                    if (num == 0 && !CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                    {
                                        objCompbcp = new ComparabancosBcp
                                        {
                                            Referencialibros = item.ReferenciaLibros,
                                            FechaOpe = item.FechaRegistro,
                                            Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                                            NroOpe = item.NroOpe,
                                            Importe = item.Importe,
                                            Dh = ((item.Dh != 1) ? 1 : 2),
                                            Terminal = item.CodigoPos,
                                            Observacion = "No:" + item.Whoyo
                                        };
                                        ListComparabcp.Add(objCompbcp);
                                    }
                                    continue;
                                }
                                if (list3.Count == 0)
                                {
                                    if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                    {
                                        objCompbcp = new ComparabancosBcp
                                        {
                                            Referencialibros = item.ReferenciaLibros,
                                            FechaOpe = item.FechaRegistro,
                                            Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                                            NroOpe = item.NroOpe,
                                            Importe = item.Importe,
                                            Dh = ((item.Dh != 1) ? 1 : 2),
                                            Terminal = item.CodigoPos,
                                            Observacion = "No se ecuentra en Bancos:" + item.Whoyo
                                        };
                                        ListComparabcp.Add(objCompbcp);
                                    }
                                    else
                                    {
                                        string empty6 = string.Empty;
                                    }
                                    continue;
                                }
                                foreach (BancoBCP itemb in list3)
                                {
                                    List<BancoUpeu> list4 = listasPendienteBancoYmayorUpeu.Where((BancoUpeu w) => w.NroOpe == itemb.NroOpe && w.CodigoPos == itemb.CodigoPos && w.Importe == itemb.Importe).ToList();
                                    if (list4.Count == 0 && !CompruebaSiExiste(itemb.NroOpe, itemb.CodigoPos, itemb.Importe))
                                    {
                                        objCompupe = new ComparabancosUpeu
                                        {
                                            FechaOpe = itemb.FechaOperacion,
                                            Descripcion = itemb.Descripcion,
                                            NroOpe = itemb.NroOpe,
                                            Importe = itemb.Importe,
                                            Dh = itemb.Dh,
                                            Terminal = itemb.CodigoPos,
                                            Observacion = "DETRAC BANCO A PEND" + itemb.Whoyo,
                                            Mb = itemb.ReferenciaVoucher
                                        };
                                        ListComparaupeu.Add(objCompupe);
                                    }
                                }
                                if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                                {
                                    objCompbcp = new ComparabancosBcp
                                    {
                                        Referencialibros = item.ReferenciaLibros,
                                        FechaOpe = item.FechaRegistro,
                                        Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                                        NroOpe = item.NroOpe,
                                        Importe = item.Importe,
                                        Dh = ((item.Dh != 1) ? 1 : 2),
                                        Terminal = item.CodigoPos,
                                        Observacion = "No se ecuentra en Bancos:" + item.Whoyo
                                    };
                                    ListComparabcp.Add(objCompbcp);
                                }
                            }
                            else if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                            {
                                objCompbcp = new ComparabancosBcp
                                {
                                    Referencialibros = item.ReferenciaLibros,
                                    FechaOpe = item.FechaRegistro,
                                    Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                                    NroOpe = item.NroOpe,
                                    Importe = item.Importe,
                                    Dh = ((item.Dh != 1) ? 1 : 2),
                                    Terminal = item.CodigoPos,
                                    Observacion = "No se ecuentra en Bancos:" + item.Whoyo
                                };
                                ListComparabcp.Add(objCompbcp);
                            }
                            else
                            {
                                string empty7 = string.Empty;
                            }
                        }
                        else
                        {
                            string empty8 = string.Empty;
                        }
                    }
                    else if (list2.Count == 2)
                    {
                        List<BancoBCP> list5 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == item.NroOpe && w.CodigoPos == item.CodigoPos && w.Importe == item.Importe).ToList();
                        if (list5.Count == 0)
                        {
                            if (list2[0].Dh != list2[1].Dh)
                            {
                                if (list2[0].Importe == list2[1].Importe)
                                {
                                    if (!CompruebaSiExiste(list2[0].NroOpe, list2[0].CodigoPos, list2[0].Importe))
                                    {
                                        UnidosUpeuyBancos unidosUpeuyBancos3 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos3.NroOpeU = list2[0].NroOpe;
                                        unidosUpeuyBancos3.FechaRegistroU = list2[0].FechaRegistro.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos3.ReferenciaLibrosU = list2[0].ReferenciaLibros;
                                        unidosUpeuyBancos3.DescripcionU = list2[0].Descripcion;
                                        unidosUpeuyBancos3.FechaOperacionU = list2[0].FechaOperacion;
                                        unidosUpeuyBancos3.ImporteU = list2[0].Importe;
                                        unidosUpeuyBancos3.DhU = list2[0].Dh;
                                        unidosUpeuyBancos3.NroOpeB = list2[0].NroOpe;
                                        unidosUpeuyBancos3.FechaOperacionB = string.Empty;
                                        unidosUpeuyBancos3.ImporteB = 0m;
                                        unidosUpeuyBancos3.DescripcionB = string.Empty;
                                        unidosUpeuyBancos3.DhB = list2[0].Dh;
                                        unidosUpeuyBancos3.Terminal = list2[0].CodigoPos;
                                        unidosUpeuyBancos3.Wherepath = "M-UPEU";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos3);
                                        UnidosUpeuyBancos unidosUpeuyBancos4 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos4.NroOpeU = list2[1].NroOpe;
                                        unidosUpeuyBancos4.FechaRegistroU = list2[1].FechaRegistro.ToString("dd/MM/yyyy");
                                        unidosUpeuyBancos4.ReferenciaLibrosU = list2[1].ReferenciaLibros;
                                        unidosUpeuyBancos4.DescripcionU = list2[1].Descripcion;
                                        unidosUpeuyBancos4.FechaOperacionU = list2[1].FechaOperacion;
                                        unidosUpeuyBancos4.ImporteU = list2[1].Importe;
                                        unidosUpeuyBancos4.DhU = list2[1].Dh;
                                        unidosUpeuyBancos4.NroOpeB = list2[1].NroOpe;
                                        unidosUpeuyBancos4.FechaOperacionB = string.Empty;
                                        unidosUpeuyBancos4.ImporteB = 0m;
                                        unidosUpeuyBancos4.DescripcionB = string.Empty;
                                        unidosUpeuyBancos4.DhB = list2[0].Dh;
                                        unidosUpeuyBancos4.Terminal = list2[1].CodigoPos;
                                        unidosUpeuyBancos4.Wherepath = "M-UPEU";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos4);
                                    }
                                }
                                else
                                {
                                    if (!CompruebaSiExiste(list2[0].NroOpe, list2[0].CodigoPos, list2[0].Importe))
                                    {
                                        objCompbcp = new ComparabancosBcp
                                        {
                                            Referencialibros = list2[0].ReferenciaLibros,
                                            FechaOpe = list2[0].FechaRegistro,
                                            Descripcion = list2[0].ReferenciaLibros + "-" + list2[0].Descripcion,
                                            NroOpe = list2[0].NroOpe,
                                            Importe = list2[0].Importe,
                                            Dh = ((list2[0].Dh != 1) ? 1 : 2),
                                            Terminal = list2[0].CodigoPos,
                                            Observacion = "No se enc. en mayor :" + list2[0].Whoyo,
                                            Whoyo = "M-UPEU"
                                        };
                                        ListComparabcp.Add(objCompbcp);
                                    }
                                    if (!CompruebaSiExiste(list2[1].NroOpe, list2[1].CodigoPos, list2[1].Importe))
                                    {
                                        objCompbcp = new ComparabancosBcp
                                        {
                                            Referencialibros = list2[1].ReferenciaLibros,
                                            FechaOpe = list2[1].FechaRegistro,
                                            Descripcion = list2[1].ReferenciaLibros + "-" + list2[1].Descripcion,
                                            NroOpe = list2[1].NroOpe,
                                            Importe = list2[1].Importe,
                                            Dh = ((list2[1].Dh != 1) ? 1 : 2),
                                            Terminal = list2[1].CodigoPos,
                                            Observacion = "No se enc. en mayor:" + list2[1].Whoyo,
                                            Whoyo = "M-UPEU"
                                        };
                                        ListComparabcp.Add(objCompbcp);
                                    }
                                }
                            }
                            else if (list2[0].Importe > 0m && list2[1].Importe > 0m)
                            {
                                if (list2[0].Importe == list2[1].Importe)
                                {
                                    objCompbcp = new ComparabancosBcp
                                    {
                                        Referencialibros = list2[0].ReferenciaLibros,
                                        FechaOpe = list2[0].FechaRegistro,
                                        Descripcion = list2[0].ReferenciaLibros + "-" + list2[0].Descripcion,
                                        NroOpe = list2[0].NroOpe,
                                        Importe = list2[0].Importe,
                                        Dh = ((list2[0].Dh != 1) ? 1 : 2),
                                        Terminal = list2[0].CodigoPos,
                                        Observacion = "No se enc. en mayor :" + list2[0].Whoyo,
                                        Whoyo = "M-UPEU"
                                    };
                                    ListComparabcp.Add(objCompbcp);
                                    objCompbcp = new ComparabancosBcp
                                    {
                                        Referencialibros = list2[1].ReferenciaLibros,
                                        FechaOpe = list2[1].FechaRegistro,
                                        Descripcion = list2[1].ReferenciaLibros + "-" + list2[1].Descripcion,
                                        NroOpe = list2[1].NroOpe,
                                        Importe = list2[1].Importe,
                                        Dh = ((list2[1].Dh != 1) ? 1 : 2),
                                        Terminal = list2[1].CodigoPos,
                                        Observacion = "No se enc. en mayor :" + list2[1].Whoyo,
                                        Whoyo = "M-UPEU"
                                    };
                                    ListComparabcp.Add(objCompbcp);
                                    continue;
                                }
                                if (!CompruebaSiExiste(list2[0].NroOpe, list2[0].CodigoPos, list2[0].Importe))
                                {

                                    decimal sumcheck = list2[0].Importe + list2[1].Importe;
                                    //busco en visa si este monto existe
                                    //modificado 04-03-2023
                                    List<BancoBCP> listvisa = listaVisaMCAE.Where((BancoBCP W) => W.NroOpe == item.NroOpe && W.CodigoPos == item.CodigoPos && W.Importe == sumcheck).ToList();
                                    if (listvisa.Count > 0)
                                    {
                                        UnidosUpeuyBancos unidosUpeuyBancos4 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos4.NroOpeU = list2[0].NroOpe;
                                        unidosUpeuyBancos4.FechaRegistroU = list2[0].FechaRegistro.ToShortDateString();
                                        unidosUpeuyBancos4.ReferenciaLibrosU = list2[0].ReferenciaLibros;
                                        unidosUpeuyBancos4.DescripcionU = list2[0].Descripcion;
                                        unidosUpeuyBancos4.FechaOperacionU = list2[0].FechaOperacion;
                                        unidosUpeuyBancos4.ImporteU = list2[0].Importe;
                                        unidosUpeuyBancos4.DhU = list2[0].Dh;
                                        unidosUpeuyBancos4.NroOpeB = string.Empty;
                                        unidosUpeuyBancos4.FechaOperacionB = string.Empty;
                                        unidosUpeuyBancos4.ImporteB = 0;
                                        unidosUpeuyBancos4.DescripcionB = string.Empty;
                                        unidosUpeuyBancos4.DhB = 0;
                                        unidosUpeuyBancos4.Terminal = list2[0].CodigoPos;
                                        unidosUpeuyBancos4.Wherepath = "Mayorx";
                                        ListaUnidosUyB.Add(unidosUpeuyBancos4);


                                        UnidosUpeuyBancos unidosUpeuyBancos5 = new UnidosUpeuyBancos();
                                        unidosUpeuyBancos5.NroOpeU = list2[1].NroOpe;
                                        unidosUpeuyBancos5.FechaRegistroU = list2[1].FechaRegistro.ToShortDateString();
                                        unidosUpeuyBancos5.ReferenciaLibrosU = list2[1].ReferenciaLibros;
                                        unidosUpeuyBancos5.DescripcionU = list2[1].Descripcion;
                                        unidosUpeuyBancos5.FechaOperacionU = list2[1].FechaOperacion;
                                        unidosUpeuyBancos5.ImporteU = list2[1].Importe;
                                    unidosUpeuyBancos5.DhU = list2[1].Dh;
                                    unidosUpeuyBancos5.NroOpeB = string.Empty;
                                    unidosUpeuyBancos5.FechaOperacionB = string.Empty;
                                    unidosUpeuyBancos5.ImporteB = 0;
                                    unidosUpeuyBancos5.DescripcionB = string.Empty;
                                    unidosUpeuyBancos5.Terminal = list2[1].CodigoPos;
                                    unidosUpeuyBancos5.DhB = 0;
                                    unidosUpeuyBancos5.Wherepath = "Mayorx";
                                    ListaUnidosUyB.Add(unidosUpeuyBancos5);

                                    UnidosUpeuyBancos unidosUpeuyBancos6 = new UnidosUpeuyBancos();
                                    unidosUpeuyBancos6.NroOpeU = listvisa.First().NroOpe;
                                    unidosUpeuyBancos6.FechaRegistroU = string.Empty;
                                    unidosUpeuyBancos6.ReferenciaLibrosU = string.Empty;
                                    unidosUpeuyBancos6.DescripcionU = string.Empty;
                                    unidosUpeuyBancos6.FechaOperacionU = string.Empty;
                                    unidosUpeuyBancos6.ImporteU = 0;
                                    unidosUpeuyBancos6.DhU = 0;
                                    unidosUpeuyBancos6.NroOpeB = listvisa.First().NroOpe;
                                    unidosUpeuyBancos6.FechaOperacionB = listvisa.First().FechaOperacion.ToShortDateString();
                                    unidosUpeuyBancos6.ImporteB = listvisa.First().Importe;
                                    unidosUpeuyBancos6.DescripcionB = listvisa.First().Descripcion;
                                    unidosUpeuyBancos6.Terminal = listvisa.First().CodigoPos;
                                    unidosUpeuyBancos6.DhB = 1;
                                    unidosUpeuyBancos6.Wherepath = "Visanetx";
                                    ListaUnidosUyB.Add(unidosUpeuyBancos6);


                                    //aqui falta los dos siguentes
                                }
                                else
                                {
                                    objCompbcp = new ComparabancosBcp
                                        {
                                            Referencialibros = list2[0].ReferenciaLibros,
                                            FechaOpe = list2[0].FechaRegistro,
                                            Descripcion = list2[0].ReferenciaLibros + "-" + list2[0].Descripcion,
                                            NroOpe = list2[0].NroOpe,
                                            Importe = list2[0].Importe,
                                            Dh = ((list2[0].Dh != 1) ? 1 : 2),
                                            Terminal = list2[0].CodigoPos,
                                            Observacion = "No se enc. en mayor :" + list2[0].Whoyo,
                                            Whoyo = "M-UPEU"
                                        };
                                        ListComparabcp.Add(objCompbcp);
                                    }
                                }
                                if (!CompruebaSiExiste(list2[1].NroOpe, list2[1].CodigoPos, list2[1].Importe))
                                {
                                    objCompbcp = new ComparabancosBcp
                                    {
                                        Referencialibros = list2[1].ReferenciaLibros,
                                        FechaOpe = list2[1].FechaRegistro,
                                        Descripcion = list2[1].ReferenciaLibros + "-" + list2[1].Descripcion,
                                        NroOpe = list2[1].NroOpe,
                                        Importe = list2[1].Importe,
                                        Dh = ((list2[1].Dh != 1) ? 1 : 2),
                                        Terminal = list2[1].CodigoPos,
                                        Observacion = "No se enc. en mayor :" + list2[1].Whoyo,
                                        Whoyo = "M-UPEU"
                                    };
                                    ListComparabcp.Add(objCompbcp);
                                }
                            }
                            else
                            {
                                string empty9 = string.Empty;
                            }
                        }
                        else
                        {
                            string empty10 = string.Empty;
                        }
                    }
                    else
                    {
                        if (list2.Count < 3)
                        {
                            continue;
                        }
                        string empty11 = string.Empty;
                        List<BancoBCP> list6 = listasPendienteUpeuYBancos.Where((BancoBCP w) => w.NroOpe == item.NroOpe && w.CodigoPos == item.CodigoPos).ToList();
                        if (list6.Count != 0 || CompruebaSiExiste(item.NroOpe, item.CodigoPos))
                        {
                            continue;
                        }
                        foreach (BancoUpeu item3 in list2)
                        {
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = item3.ReferenciaLibros,
                                FechaOpe = item3.FechaRegistro,
                                Descripcion = item3.ReferenciaLibros + "-" + item3.Descripcion,
                                NroOpe = item3.NroOpe,
                                Importe = item3.Importe,
                                Dh = ((item3.Dh != 1) ? 1 : 2),
                                Terminal = item3.CodigoPos,
                                Observacion = "TRIPICADO EN MAYOR UPEU :" + item3.Whoyo,
                                Whoyo = "M-UPEU"
                            };
                            ListComparabcp.Add(objCompbcp);
                        }
                    }
                }
                else if (item.NroOpe.Equals("VISANET"))
                {
                    string empty12 = string.Empty;
                    List<BancoBCP> list7 = listaPendUpeuVisanet.Where((BancoBCP w) => w.Importe == item.Importe).ToList();
                    if (list7.Count == 1)
                    {
                        string empty13 = string.Empty;
                        UnidosUpeuyBancos unidosUpeuyBancos5 = new UnidosUpeuyBancos();
                        unidosUpeuyBancos5.NroOpeU = item.NroOpe;
                        unidosUpeuyBancos5.FechaRegistroU = item.FechaOperacion;
                        unidosUpeuyBancos5.ReferenciaLibrosU = item.ReferenciaLibros;
                        unidosUpeuyBancos5.DescripcionU = item.Descripcion;
                        unidosUpeuyBancos5.FechaOperacionU = item.FechaOperacion;
                        unidosUpeuyBancos5.ImporteU = item.Importe;
                        unidosUpeuyBancos5.DhU = item.Dh;
                        unidosUpeuyBancos5.NroOpeB = list7.First().NroOpe;
                        unidosUpeuyBancos5.FechaOperacionB = list7.First().FechaOperacion.ToShortDateString();
                        unidosUpeuyBancos5.ImporteB = list7.First().Importe;
                        unidosUpeuyBancos5.DescripcionB = list7.First().Descripcion;
                        unidosUpeuyBancos5.DhB = list7.First().Dh;
                        unidosUpeuyBancos5.Terminal = list7.First().CodigoPos;
                        unidosUpeuyBancos5.Wherepath = "visanet mayor";
                        ListaUnidosOnly.Add(unidosUpeuyBancos5);
                    }
                    else
                    {
                        string empty14 = string.Empty;
                        objCompbcp = new ComparabancosBcp
                        {
                            Referencialibros = item.ReferenciaLibros,
                            FechaOpe = item.FechaRegistro,
                            Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                            NroOpe = item.NroOpe,
                            Importe = item.Importe,
                            Dh = ((item.Dh != 1) ? 1 : 2),
                            Terminal = item.CodigoPos,
                            Observacion = item.Whoyo,
                            Whoyo = "M-UPEU",
                            Mb = item.ReferenciaLibros
                        };
                        ListComparabcp.Add(objCompbcp);
                    }
                }
                else
                {
                    if (item.NroOpe.Equals("DETRAC") || CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.ReferenciaLibros))
                    {
                        continue;
                    }
                    int num2 = 0;
                    foreach (BancoUpeu item4 in list)
                    {
                        if (item4.ReferenciaLibros != null)
                        {
                            if (item4.NroOpe.Equals("305"))
                            {
                                string dcc = string.Empty;
                            }
                            if (!CompruebaSiExiste(item4.NroOpe, item4.CodigoPos, item4.ReferenciaLibros))
                            {
                                if (!CompruebaSiExisteExtra(item4.NroOpe, item4.CodigoPos, item4.ReferenciaLibros))
                                {
                                    objCompbcp = new ComparabancosBcp
                                {
                                    Referencialibros = item4.ReferenciaLibros,
                                    FechaOpe = item4.FechaRegistro,
                                    Descripcion = item4.ReferenciaLibros + "-" + item4.Descripcion,
                                    NroOpe = item4.NroOpe,
                                    Importe = item4.Importe,
                                    Dh = ((item4.Dh != 1) ? 1 : 2),
                                    Terminal = item4.CodigoPos,
                                    Observacion = "> a 2 :" + item4.Whoyo + ":" + num2,
                                    Whoyo = "M-UPEU",
                                    Mb = item4.ReferenciaLibros,
                                    Pintar = 101
                                };
                                ListComparabcp.Add(objCompbcp);
                                }
                            }
                        }
                        else
                        {
                            objCompbcp = new ComparabancosBcp
                            {
                                Referencialibros = item4.ReferenciaLibros,
                                FechaOpe = item4.FechaRegistro,
                                Descripcion = item4.ReferenciaLibros + "-" + item4.Descripcion,
                                NroOpe = item4.NroOpe,
                                Importe = item4.Importe,
                                Dh = ((item4.Dh != 1) ? 1 : 2),
                                Terminal = item4.CodigoPos,
                                Observacion = "> a 2 :" + item4.Whoyo + ":" + num2,
                                Whoyo = "M-UPEU",
                                Mb = item4.ReferenciaLibros,
                                Pintar = 101
                            };
                            ListComparabcp.Add(objCompbcp);
                        }
                        num2++;
                    }
                }
            }
        }

        public void RecorreMayoruUpeuT()
        {
            foreach (BancoUpeu item in listaMayorUpeuT)
            {
                if (item.NroOpe.Equals("13"))
                {
                    string empty = string.Empty;
                }
                if (item.NroOpe.Equals("1425"))
                {
                    var xx = string.Empty;
                }
                if (!CompruebaSiExiste(item.NroOpe, item.CodigoPos, item.Importe))
                {
                    objCompbcp = new ComparabancosBcp
                    {
                        Referencialibros = item.ReferenciaLibros,
                        FechaOpe = item.FechaRegistro,
                        Descripcion = item.ReferenciaLibros + "-" + item.Descripcion,
                        NroOpe = item.NroOpe,
                        Importe = item.Importe,
                        Dh = ((item.Dh != 1) ? 1 : 2),
                        Terminal = item.CodigoPos,
                        Observacion = "Recoorido Mayor T:" + item.Whoyo
                    };
                    ListComparabcp.Add(objCompbcp);
                }
            }
        }

        private void CargaExcel()
        {
            try
            {
                dtc = ExcelUtil.GetNameSheets(diag.FileName);
            }
            catch (Exception ex)
            {
                throw new Exception("Excel de trabajo Abierto: " + Environment.NewLine + ex.Message);
            }
        }

        private void DivujaExcel()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "Daniel Antazu - Movil 992144164";
            excelPackage.Workbook.Properties.Title = "Conciliacion";
            excelPackage.Workbook.Properties.Subject = "Se leen 4 hojas de excel visa, pendientes, mayor upeu,mayor banco en ese orden si no hay datos de uno de elllos dejarlo la hoja en blanco en ese orden";
            excelPackage.Workbook.Properties.Created = DateTime.Now;
            excelPackage.Workbook.CalcMode = ExcelCalcMode.Manual;
            ExcelWorksheet excelWorksheet = null;
            ExcelWorksheet excelWorksheet2 = null;
            ExcelWorksheet excelWorksheet3 = null;
            ExcelWorksheet excelWorksheet4 = null;
            ExcelWorksheet excelWorksheet5 = null;
            ExcelWorksheet excelWorksheet6 = excelPackage.Workbook.Worksheets.Add("Pendientes -" + CuentaContableUpeu);
            ExcelWorksheet excelWorksheet7 = excelPackage.Workbook.Worksheets.Add("Conciliados");
            ExcelWorksheet excelWorksheet8 = excelPackage.Workbook.Worksheets.Add("Conciliados-Otros");
            ExcelWorksheet excelWorksheet9 = null;
            if (ListComiNoAbo != null && ListComiNoAbo.Count > 0)
            {
                excelWorksheet2 = excelPackage.Workbook.Worksheets.Add("No-Abo");
            }
            if (ListComiAbo != null && ListComiAbo.Count > 0)
            {
                excelWorksheet3 = excelPackage.Workbook.Worksheets.Add("Abo");
            }
            if (ListComiTotAbo != null && ListComiTotAbo.Count > 0)
            {
                excelWorksheet4 = excelPackage.Workbook.Worksheets.Add("Tot-Abo");
            }
            excelWorksheet6.Cells["A1:A1"].Value = "No ignore los pendiente Solucionalo!";
            excelWorksheet6.Cells["F1:J1"].Value = "Valores no Contables";
            excelWorksheet6.Cells["F2:G2"].Value = "UPEU";
            excelWorksheet6.Cells["H2:I2"].Value = "BANCO";
            excelWorksheet6.Cells[3, 1].Value = "Fecha";
            excelWorksheet6.Cells[3, 2].Value = "Mb";
            excelWorksheet6.Cells[3, 3].Value = "Nro Doc";
            excelWorksheet6.Cells[3, 4].Value = "Detalle";
            excelWorksheet6.Cells[3, 5].Value = "Fecha Ope";
            excelWorksheet6.Cells[3, 6].Value = "Debito";
            excelWorksheet6.Cells[3, 7].Value = "Credito";
            excelWorksheet6.Cells[3, 8].Value = "Debito";
            excelWorksheet6.Cells[3, 9].Value = "Credito";
            excelWorksheet6.Cells[3, 10].Value = "Terminal";
            excelWorksheet6.Cells[3, 11].Value = "Anotaciones";
            excelWorksheet6.Column(1).Width = 15.0;
            excelWorksheet6.Column(2).Width = 12.0;
            excelWorksheet6.Column(3).Width = 12.0;
            excelWorksheet6.Column(4).Width = 50.0;
            excelWorksheet6.Column(5).Width = 13.0;
            excelWorksheet6.Column(6).Width = 13.0;
            excelWorksheet6.Column(7).Width = 13.0;
            excelWorksheet6.Column(8).Width = 13.0;
            excelWorksheet6.Column(9).Width = 13.0;
            excelWorksheet6.Column(10).Width = 13.0;
            excelWorksheet6.Column(1).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet6.Column(6).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet6.Column(7).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet6.Column(8).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet6.Column(9).Style.Numberformat.Format = "#,##0.00";
            using (ExcelRange excelRange = excelWorksheet6.Cells["A1:E1"])
            {
                excelRange.Merge = true;
                excelRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange.Style.Font.Bold = true;
                excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelRange.Style.Border.Top.Style = ExcelBorderStyle.Thick;
                excelRange.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                excelRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            using (ExcelRange excelRange2 = excelWorksheet6.Cells["F1:J1"])
            {
                excelRange2.Merge = true;
                excelRange2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange2.Style.Font.Bold = true;
                excelRange2.Style.Border.Top.Style = ExcelBorderStyle.Thick;
                excelRange2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelRange2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Border.Right.Style = ExcelBorderStyle.Thick;
            }
            using (ExcelRange excelRange3 = excelWorksheet6.Cells["A2:E2"])
            {
                excelRange3.Merge = true;
                excelRange3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                excelRange3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Font.Bold = true;
                excelRange3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange4 = excelWorksheet6.Cells["F2:G2"])
            {
                excelRange4.Merge = true;
                excelRange4.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange4.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange4.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange4.Style.Font.Bold = true;
                excelRange4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange5 = excelWorksheet6.Cells["H2:I2"])
            {
                excelRange5.Merge = true;
                excelRange5.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange5.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange5.Style.Border.Right.Style = ExcelBorderStyle.Thick;
                excelRange5.Style.Font.Bold = true;
                excelRange5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange6 = excelWorksheet6.Cells["A3:I3"])
            {
                excelRange6.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange6.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                excelRange6.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange6.Style.Font.Bold = true;
                excelRange6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange7 = excelWorksheet6.Cells["A3:A3"])
            {
                excelRange7.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange7.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                excelRange7.Style.Font.Bold = true;
                excelRange7.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange8 = excelWorksheet6.Cells["A3:k3"])
            {
                excelRange8.AutoFilter = true;
            }
            int num = 5;
            excelWorksheet6.Cells[4, 1].Value = string.Empty;
            excelWorksheet6.Cells[4, 2].Value = string.Empty;
            if (ListComparabcp.Count > 0)
            {
                excelWorksheet6.Cells[4, 4].Value = "Saldo al " + GetLastDayOf(ListComparabcp.First().FechaOpe);
            }
            excelWorksheet6.Cells[4, 6].Value = SaldoFiNUpeu;
            excelWorksheet6.Cells[4, 9].Value = SaldoFiNBanco;
            foreach (ComparabancosBcp item in ListComparabcp.OrderBy((ComparabancosBcp w) => w.NroOpe).ToList())
            {
                excelWorksheet6.Cells[num, 1].Value = item.FechaOpe;
                excelWorksheet6.Cells[num, 2].Value = item.Referencialibros;
                excelWorksheet6.Cells[num, 3].Value = item.NroOpe;
                excelWorksheet6.Cells[num, 4].Value = item.Descripcion;
                excelWorksheet6.Cells[num, 5].Value = string.Empty;
                if (item.Dh == 1)
                {
                    excelWorksheet6.Cells[num, 8].Value = item.Importe;
                }
                else
                {
                    excelWorksheet6.Cells[num, 9].Value = item.Importe;
                }
                excelWorksheet6.Cells[num, 10].Value = item.Terminal;
                excelWorksheet6.Cells[num, 11].Value = item.Observacion;
                using (ExcelRange excelRange9 = excelWorksheet6.Cells[num, 1, num, 11])
                {
                    excelRange9.Style.Font.Bold = false;
                    excelRange9.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange9.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange9.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange9.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange9.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                   // excelRange9.Style.Font.Color.SetColor(Color.Black);
                    
                    excelRange9.Style.Font.Color.SetColor(Color.Green);
                    //excelRange9.Style.ShrinkToFit = true;

                    
                }
                num++;
            }
            foreach (ComparabancosUpeu item2 in ListComparaupeu.OrderBy((ComparabancosUpeu w) => w.NroOpe).ToList())
            {
                excelWorksheet6.Cells[num, 1].Value = item2.FechaOpe;
                excelWorksheet6.Cells[num, 2].Value = string.Empty;
                excelWorksheet6.Cells[num, 3].Value = item2.NroOpe;
                excelWorksheet6.Cells[num, 4].Value = item2.Descripcion;
                excelWorksheet6.Cells[num, 5].Value = string.Empty;
                excelWorksheet6.Cells[num, 10].Value = item2.Terminal;
                excelWorksheet6.Cells[num, 11].Value = item2.Observacion;
                using (ExcelRange excelRange10 = excelWorksheet6.Cells[num, 3, num, 3])
                {
                    excelRange10.Style.Font.Bold = true;
                    excelRange10.Style.ShrinkToFit = true;
                }
                if (item2.Dh == 1)
                {
                    excelWorksheet6.Cells[num, 6].Value = item2.Importe;
                }
                else
                {
                    excelWorksheet6.Cells[num, 7].Value = item2.Importe;
                }
                if (item2.Pintar == 2)
                {
                    ExcelRange excelRange11 = excelWorksheet6.Cells[num, 1, num, 7];
                    excelRange11.Style.Font.Bold = true;
                    excelRange11.Style.Font.Color.SetColor(Color.OrangeRed);
                }
                if (item2.Pintar == 3)
                {
                    ExcelRange excelRange12 = excelWorksheet6.Cells[num, 1, num, 7];
                    excelRange12.Style.Font.Bold = true;
                    excelRange12.Style.Font.Color.SetColor(Color.Green);
                }
                if (item2.Whoyo == "1")
                {
                    ExcelRange excelRange13 = excelWorksheet6.Cells[num, 1, num, 7];
                    excelRange13.Style.Font.Bold = true;
                    excelRange13.Style.Font.Color.SetColor(Color.Firebrick);
                }
                if (item2.Whoyo == "MC")
                {
                    ExcelRange excelRange14 = excelWorksheet6.Cells[num, 1, num, 7];
                    excelRange14.Style.Font.Bold = true;
                    excelRange14.Style.Font.Color.SetColor(Color.Coral);
                }
                if (item2.Whoyo == "VN")
                {
                    ExcelRange excelRange15 = excelWorksheet6.Cells[num, 1, num, 7];
                    excelRange15.Style.Font.Bold = true;
                    excelRange15.Style.Font.Color.SetColor(Color.DarkViolet);
                }
                using (ExcelRange excelRange16 = excelWorksheet6.Cells[num, 1, num, 11])
                {
                    excelRange16.Style.Font.Bold = false;
                    excelRange16.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange16.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange16.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange16.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange16.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange16.Style.Font.Color.SetColor(Color.CornflowerBlue);
                }
                num++;
            }
            int num2 = num - 1;
            int num3 = num + 1;
            int num4 = num + 2;
            string address = "F" + num;
            string address2 = "G" + num;
            string address3 = "H" + num;
            string address4 = "I" + num;
            string text = "F" + num3;
            string text2 = "G" + num3;
            string text3 = "H" + num3;
            string text4 = "I" + num3;
            string address5 = "F" + num4;
            string address6 = "G" + num4;
            string address7 = "H" + num4;
            string address8 = "I" + num4;
            excelWorksheet6.Cells[address].Formula = "SUM(F4:F" + num2 + ")";
            excelWorksheet6.Cells[address].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet6.Cells[address2].Formula = "SUM(G4:G" + num2 + ")";
            excelWorksheet6.Cells[address2].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet6.Cells[address3].Formula = "SUM(H4:H" + num2 + ")";
            excelWorksheet6.Cells[address3].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet6.Cells[address4].Formula = "SUM(I4:I" + num2 + ")";
            excelWorksheet6.Cells[address4].Style.Numberformat.Format = "#,##0.00";
            string address9 = "F" + (num + 1);
            string address10 = "G" + (num + 1);
            string address11 = "H" + (num + 1);
            string address12 = "I" + (num + 1);
            excelWorksheet6.Cells[address].Calculate();
            excelWorksheet6.Cells[address2].Calculate();
            excelWorksheet6.Cells[address3].Calculate();
            excelWorksheet6.Cells[address4].Calculate();
            decimal num5 = decimal.Parse(excelWorksheet6.Cells[num, 6].Value.ToString());
            decimal num6 = decimal.Parse(excelWorksheet6.Cells[num, 7].Value.ToString());
            decimal num7 = decimal.Parse(excelWorksheet6.Cells[num, 8].Value.ToString());
            decimal num8 = decimal.Parse(excelWorksheet6.Cells[num, 9].Value.ToString());
            decimal num9 = num5 - num6;
            if (num9 < 0m)
            {
                excelWorksheet6.Cells[address9].Value = -1m * num9;
                excelWorksheet6.Cells[address10].Value = 0;
            }
            else
            {
                excelWorksheet6.Cells[address10].Value = num9;
                excelWorksheet6.Cells[address9].Value = 0;
            }
            decimal num10 = num7 - num8;
            if (num10 < 0m)
            {
                excelWorksheet6.Cells[address11].Value = -1m * num10;
                excelWorksheet6.Cells[address12].Value = 0;
            }
            else
            {
                excelWorksheet6.Cells[address12].Value = num10;
                excelWorksheet6.Cells[address11].Value = 0;
            }
            excelWorksheet6.Cells[address5].Formula = "SUM(F" + num + ":" + text + ")";
            excelWorksheet6.Cells[address6].Formula = "SUM(G" + num + ":" + text2 + ")";
            excelWorksheet6.Cells[address7].Formula = "SUM(H" + num + ":" + text3 + ")";
            excelWorksheet6.Cells[address8].Formula = "SUM(I" + num + ":" + text4 + ")";

            excelWorksheet6.Cells[num+5, 2].Value = msgPen;
            using (ExcelRange excelRange20 = excelWorksheet6.Cells[num+5, 1, num+5, 2])
            {
                excelRange20.Style.Font.Bold = false;
                excelRange20.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                //excelRange20.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                //excelRange20.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                //excelRange20.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                //excelRange20.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange20.Style.Font.Color.SetColor(Color.Red);
               
            }


            excelWorksheet7.Cells["B2:B2"].Value = "Fecha";
            excelWorksheet7.Cells["C2:C2"].Value = "Com";
            excelWorksheet7.Cells["D2:D2"].Value = "Descripcion";
            excelWorksheet7.Cells["E2:E2"].Value = "Nro Ope";
            excelWorksheet7.Cells["F2:F2"].Value = "FechaOpe";
            excelWorksheet7.Cells["G2:G2"].Value = "Debito";
            excelWorksheet7.Cells["H2:H2"].Value = "Credito";
            excelWorksheet7.Cells["K2:K2"].Value = "Fecha";
            excelWorksheet7.Cells["L2:L2"].Value = "Nro Ope";
            excelWorksheet7.Cells["M2:M2"].Value = "Referencia";
            excelWorksheet7.Cells["N2:N2"].Value = "Monto D";
            excelWorksheet7.Cells["O2:O2"].Value = "Monto H";
            excelWorksheet7.Cells["P2:P2"].Value = "Terminal";
            excelWorksheet7.Cells["Q2:Q2"].Value = "Obs";
            excelWorksheet7.Column(2).Width = 12.0;
            excelWorksheet7.Column(3).Width = 12.0;
            excelWorksheet7.Column(4).Width = 30.0;
            excelWorksheet7.Column(5).Width = 12.0;
            excelWorksheet7.Column(6).Width = 12.0;
            excelWorksheet7.Column(7).Width = 12.0;
            excelWorksheet7.Column(8).Width = 12.0;
            excelWorksheet7.Column(11).Width = 12.0;
            excelWorksheet7.Column(12).Width = 12.0;
            excelWorksheet7.Column(13).Width = 30.0;
            excelWorksheet7.Column(14).Width = 12.0;
            excelWorksheet7.Column(15).Width = 12.0;
            excelWorksheet7.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet7.Column(5).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet7.Column(7).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet7.Column(8).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet7.Column(11).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet7.Column(14).Style.Numberformat.Format = "#,##0.00";
            using (ExcelRange excelRange17 = excelWorksheet7.Cells["B2:H2"])
            {
                excelRange17.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange17.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange17.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange17.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange17.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange17.Style.Font.Bold = true;
                excelRange17.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange18 = excelWorksheet7.Cells["K2:O2"])
            {
                excelRange18.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange18.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange18.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange18.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange18.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange18.Style.Font.Bold = true;
                excelRange18.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange19 = excelWorksheet7.Cells["B2:Q2"])
            {
                excelRange19.AutoFilter = true;
            }
            int num11 = 3;
            foreach (UnidosUpeuyBancos item3 in ListaUnidosOnly)
            {
                excelWorksheet7.Cells[num11, 2].Value = item3.FechaRegistroU;
                excelWorksheet7.Cells[num11, 3].Value = item3.ReferenciaLibrosU;
                excelWorksheet7.Cells[num11, 4].Value = item3.DescripcionU;
                excelWorksheet7.Cells[num11, 5].Value = item3.NroOpeU;
                excelWorksheet7.Cells[num11, 6].Value = item3.FechaOperacionU;
                if (item3.DhU == 1)
                {
                    excelWorksheet7.Cells[num11, 7].Value = item3.ImporteU;
                }
                else
                {
                    excelWorksheet7.Cells[num11, 8].Value = item3.ImporteU;
                }
                excelWorksheet7.Cells[num11, 11].Value = item3.FechaOperacionB;
                excelWorksheet7.Cells[num11, 12].Value = item3.NroOpeB;
                excelWorksheet7.Cells[num11, 13].Value = item3.DescripcionB;
                if (item3.DhB == 1)
                {
                    excelWorksheet7.Cells[num11, 14].Value = item3.ImporteB;
                }
                else
                {
                    excelWorksheet7.Cells[num11, 15].Value = item3.ImporteB;
                }
                excelWorksheet7.Cells[num11, 16].Value = item3.Terminal;
                excelWorksheet7.Cells[num11, 17].Value = item3.Wherepath;
                num11++;
                ExcelRange excelRange20 = excelWorksheet8.Cells[num11, 3, num11, 17];
                excelRange20.Style.Font.Bold = false;
                excelRange20.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange20.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange20.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange20.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange20.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange20.Style.Font.Color.SetColor(Color.Black);
                excelRange20.Style.ShrinkToFit = true;
            }
            int num12 = num11 - 1;
            int num13 = num11 + 1;
            int num14 = num11 + 2;
            string address13 = "G" + num11;
            string address14 = "H" + num11;
            string address15 = "N" + num11;
            string address16 = "O" + num11;
            excelWorksheet7.Cells[address13].Calculate();
            excelWorksheet7.Cells[address14].Calculate();
            excelWorksheet7.Cells[address15].Calculate();
            excelWorksheet7.Cells[address16].Calculate();
            excelWorksheet7.Cells[address13].Formula = "SUM(G3:G" + num12 + ")";
            excelWorksheet7.Cells[address13].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet7.Cells[address14].Formula = "SUM(H3:H" + num12 + ")";
            excelWorksheet7.Cells[address14].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet7.Cells[address15].Formula = "SUM(N3:N" + num12 + ")";
            excelWorksheet7.Cells[address15].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet7.Cells[address16].Formula = "SUM(O3:O" + num12 + ")";
            excelWorksheet7.Cells[address16].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet8.Cells["B2:B2"].Value = "Fecha";
            excelWorksheet8.Cells["C2:C2"].Value = "Com";
            excelWorksheet8.Cells["D2:D2"].Value = "Descripcion";
            excelWorksheet8.Cells["E2:E2"].Value = "Nro Ope";
            excelWorksheet8.Cells["F2:F2"].Value = "FechaOpe";
            excelWorksheet8.Cells["G2:G2"].Value = "Debito";
            excelWorksheet8.Cells["H2:H2"].Value = "Credito";
            excelWorksheet8.Cells["K2:K2"].Value = "Fecha";
            excelWorksheet8.Cells["L2:L2"].Value = "Nro Ope";
            excelWorksheet8.Cells["M2:M2"].Value = "Referencia";
            excelWorksheet8.Cells["N2:N2"].Value = "Monto D";
            excelWorksheet8.Cells["O2:O2"].Value = "Monto H";
            excelWorksheet8.Cells["P2:P2"].Value = "Terminal";
            excelWorksheet8.Cells["Q2:Q2"].Value = "Obs";
            excelWorksheet8.Column(2).Width = 12.0;
            excelWorksheet8.Column(3).Width = 12.0;
            excelWorksheet8.Column(4).Width = 30.0;
            excelWorksheet8.Column(5).Width = 12.0;
            excelWorksheet8.Column(6).Width = 12.0;
            excelWorksheet8.Column(7).Width = 12.0;
            excelWorksheet8.Column(8).Width = 12.0;
            excelWorksheet8.Column(11).Width = 12.0;
            excelWorksheet8.Column(12).Width = 12.0;
            excelWorksheet8.Column(13).Width = 30.0;
            excelWorksheet8.Column(14).Width = 12.0;
            excelWorksheet8.Column(15).Width = 12.0;
            excelWorksheet8.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet8.Column(5).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet8.Column(7).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet8.Column(8).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet8.Column(11).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet8.Column(14).Style.Numberformat.Format = "#,##0.00";
            using (ExcelRange excelRange21 = excelWorksheet8.Cells["B2:H2"])
            {
                excelRange21.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange21.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange21.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange21.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange21.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange21.Style.Font.Bold = true;
                excelRange21.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange22 = excelWorksheet8.Cells["K2:O2"])
            {
                excelRange22.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange22.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange22.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange22.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange22.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange22.Style.Font.Bold = true;
                excelRange22.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange23 = excelWorksheet8.Cells["B2:Q2"])
            {
                excelRange23.AutoFilter = true;
            }
            int num15 = 3;
            foreach (UnidosUpeuyBancos item4 in ListaUnidosUyB)
            {
                excelWorksheet8.Cells[num15, 2].Value = item4.FechaRegistroU;
                excelWorksheet8.Cells[num15, 3].Value = item4.ReferenciaLibrosU;
                excelWorksheet8.Cells[num15, 4].Value = item4.DescripcionU;
                excelWorksheet8.Cells[num15, 5].Value = item4.NroOpeU;
                excelWorksheet8.Cells[num15, 6].Value = item4.FechaOperacionU;
                if (item4.DhU == 1)
                {
                    excelWorksheet8.Cells[num15, 7].Value = item4.ImporteU;
                }
                else
                {
                    excelWorksheet8.Cells[num15, 8].Value = item4.ImporteU;
                }
                excelWorksheet8.Cells[num15, 11].Value = item4.FechaOperacionB;
                excelWorksheet8.Cells[num15, 12].Value = item4.NroOpeB;
                excelWorksheet8.Cells[num15, 13].Value = item4.DescripcionB;
                if (item4.DhB == 1)
                {
                    excelWorksheet8.Cells[num15, 14].Value = item4.ImporteB;
                }
                else
                {
                    excelWorksheet8.Cells[num15, 15].Value = item4.ImporteB;
                }
                excelWorksheet8.Cells[num15, 16].Value = item4.Terminal;
                excelWorksheet8.Cells[num15, 17].Value = item4.Wherepath;
                num15++;
                ExcelRange excelRange24 = excelWorksheet8.Cells[num15, 3, num15, 17];
                excelRange24.Style.Font.Bold = false;
                excelRange24.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange24.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange24.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange24.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange24.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange24.Style.Font.Color.SetColor(Color.Black);
                excelRange24.Style.ShrinkToFit = true;
            }
            if (ExcelUtil.Listacom.Count > 0)
            {
                excelWorksheet5 = excelPackage.Workbook.Worksheets.Add("Comis");
                excelWorksheet5.Cells["B2:B2"].Value = "Fecha";
                excelWorksheet5.Cells["C2:C2"].Value = "Nro Ope";
                excelWorksheet5.Cells["D2:D2"].Value = "Descripcion";
                excelWorksheet5.Cells["E2:E2"].Value = "Importe";
                excelWorksheet5.Cells["F2:F2"].Value = "Terminal";
                excelWorksheet5.Column(2).Width = 12.0;
                excelWorksheet5.Column(3).Width = 12.0;
                excelWorksheet5.Column(4).Width = 30.0;
                excelWorksheet5.Column(5).Width = 12.0;
                excelWorksheet5.Column(6).Width = 12.0;
                excelWorksheet5.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
                excelWorksheet5.Column(5).Style.Numberformat.Format = "#,##0.00";
                using (ExcelRange excelRange25 = excelWorksheet5.Cells["B2:F2"])
                {
                    excelRange25.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange25.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange25.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange25.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange25.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange25.Style.Font.Bold = true;
                    excelRange25.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                using (ExcelRange excelRange26 = excelWorksheet5.Cells["B2:F2"])
                {
                    excelRange26.AutoFilter = true;
                }
                int num16 = 3;
                foreach (ComparabancosComision item5 in ExcelUtil.Listacom.OrderBy((ComparabancosComision w) => w.Descripcion))
                {
                    excelWorksheet5.Cells[num16, 2].Value = item5.FechaOpe;
                    excelWorksheet5.Cells[num16, 3].Value = item5.NroOpe;
                    excelWorksheet5.Cells[num16, 4].Value = item5.Descripcion;
                    excelWorksheet5.Cells[num16, 5].Value = item5.Importe;
                    excelWorksheet5.Cells[num16, 6].Value = string.Empty;
                    num16++;
                    ExcelRange excelRange27 = excelWorksheet5.Cells[num16, 3, num16, 6];
                    excelRange27.Style.Font.Bold = false;
                    excelRange27.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange27.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange27.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange27.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange27.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange27.Style.Font.Color.SetColor(Color.Black);
                    excelRange27.Style.ShrinkToFit = true;
                }
            }
            if (ExcelUtil.Listadet.Count > 0)
            {
                excelWorksheet = excelPackage.Workbook.Worksheets.Add("Detrac");
                excelWorksheet.Cells["B2:B2"].Value = "Fecha";
                excelWorksheet.Cells["C2:C2"].Value = "Nro Ope";
                excelWorksheet.Cells["D2:D2"].Value = "Descripcion";
                excelWorksheet.Cells["E2:E2"].Value = "Importe";
                excelWorksheet.Cells["F2:F2"].Value = "Terminal";
                excelWorksheet.Column(2).Width = 12.0;
                excelWorksheet.Column(3).Width = 12.0;
                excelWorksheet.Column(4).Width = 30.0;
                excelWorksheet.Column(5).Width = 12.0;
                excelWorksheet.Column(6).Width = 12.0;
                excelWorksheet.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
                excelWorksheet.Column(5).Style.Numberformat.Format = "#,##0.00";
                using (ExcelRange excelRange28 = excelWorksheet.Cells["B2:F2"])
                {
                    excelRange28.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange28.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange28.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange28.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange28.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange28.Style.Font.Bold = true;
                    excelRange28.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                using (ExcelRange excelRange29 = excelWorksheet.Cells["B2:F2"])
                {
                    excelRange29.AutoFilter = true;
                }
                int num17 = 3;
                foreach (ComparabancosComision item6 in ExcelUtil.Listadet.OrderBy((ComparabancosComision w) => w.Descripcion))
                {
                    excelWorksheet.Cells[num17, 2].Value = item6.FechaOpe;
                    excelWorksheet.Cells[num17, 3].Value = item6.NroOpe;
                    excelWorksheet.Cells[num17, 4].Value = item6.Descripcion;
                    excelWorksheet.Cells[num17, 5].Value = item6.Importe;
                    excelWorksheet.Cells[num17, 6].Value = string.Empty;
                    num17++;
                    ExcelRange excelRange30 = excelWorksheet.Cells[num17, 3, num17, 6];
                    excelRange30.Style.Font.Bold = false;
                    excelRange30.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange30.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange30.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange30.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange30.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange30.Style.Font.Color.SetColor(Color.Black);
                    excelRange30.Style.ShrinkToFit = true;
                }
            }
            if (ListComiTotAbo != null && ListComiTotAbo.Count > 0)
            {
                excelWorksheet4.Cells["B2:B2"].Value = "Terminal";
                excelWorksheet4.Cells["C2:C2"].Value = "Descripcion";
                excelWorksheet4.Cells["D2:D2"].Value = "Fecha Trans";
                excelWorksheet4.Cells["E2:E2"].Value = "Importe-tran";
                excelWorksheet4.Cells["F2:F2"].Value = "Importe-Abono";
                excelWorksheet4.Cells["G2:G2"].Value = "Fecha Abono";
                excelWorksheet4.Cells["H2:H2"].Value = "Nro Ope";
                excelWorksheet4.Cells["I2:I2"].Value = "Comision";
                excelWorksheet4.Column(2).Width = 12.0;
                excelWorksheet4.Column(3).Width = 30.0;
                excelWorksheet4.Column(4).Width = 12.0;
                excelWorksheet4.Column(5).Width = 12.0;
                excelWorksheet4.Column(6).Width = 12.0;
                excelWorksheet4.Column(7).Width = 12.0;
                excelWorksheet4.Column(8).Width = 12.0;
                using (ExcelRange excelRange31 = excelWorksheet4.Cells["B2:I2"])
                {
                    excelRange31.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange31.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange31.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange31.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange31.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange31.Style.Font.Bold = true;
                    excelRange31.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                using (ExcelRange excelRange32 = excelWorksheet4.Cells["B2:I2"])
                {
                    excelRange32.AutoFilter = true;
                }
                int num18 = 3;
                foreach (ComparabancosUpeu item7 in ListComiTotAbo)
                {
                    excelWorksheet4.Cells[num18, 2].Value = item7.Terminal;
                    excelWorksheet4.Cells[num18, 3].Value = item7.Descripcion;
                    excelWorksheet4.Cells[num18, 4].Value = item7.FechaOpe.ToShortDateString();
                    excelWorksheet4.Cells[num18, 5].Value = item7.ImporteTransac;
                    excelWorksheet4.Cells[num18, 6].Value = item7.ImporteAbono;
                    excelWorksheet4.Cells[num18, 7].Value = item7.FechaAbono;
                    excelWorksheet4.Cells[num18, 8].Value = item7.NroOpe;
                    excelWorksheet4.Cells[num18, 9].Value = item7.Importe;
                    using (ExcelRange excelRange33 = excelWorksheet4.Cells[num18, 2, num18, 9])
                    {
                        excelRange33.Style.Font.Bold = false;
                        excelRange33.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        excelRange33.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        excelRange33.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        excelRange33.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        excelRange33.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        excelRange33.Style.Font.Color.SetColor(Color.Black);
                        excelRange33.Style.ShrinkToFit = true;
                    }
                    if (item7.Whoyo == "MC")
                    {
                        ExcelRange excelRange34 = excelWorksheet4.Cells[num18, 2, num18, 9];
                        excelRange34.Style.Font.Bold = true;
                        excelRange34.Style.Font.Color.SetColor(Color.Coral);
                    }
                    if (item7.Descripcion.Contains("VISA AGRUPADOS"))
                    {
                        ExcelRange excelRange35 = excelWorksheet4.Cells[num18, 2, num18, 9];
                        excelRange35.Style.Font.Bold = true;
                        excelRange35.Style.Font.Color.SetColor(Color.Tomato);
                    }
                    if (item7.Whoyo == "VN")
                    {
                        ExcelRange excelRange36 = excelWorksheet4.Cells[num18, 2, num18, 9];
                        excelRange36.Style.Font.Bold = true;
                        excelRange36.Style.Font.Color.SetColor(Color.Violet);
                    }
                    num18++;
                }
            }
            if (ListComiNoAbo != null && ListComiNoAbo.Count > 0)
            {
                excelWorksheet2.Cells["B2:B2"].Value = "Terminal";
                excelWorksheet2.Cells["C2:C2"].Value = "Descripcion";
                excelWorksheet2.Cells["D2:D2"].Value = "Fecha Trans";
                excelWorksheet2.Cells["E2:E2"].Value = "Importe-tran";
                excelWorksheet2.Cells["F2:F2"].Value = "Importe-Abono";
                excelWorksheet2.Cells["G2:G2"].Value = "Fecha Abono";
                excelWorksheet2.Cells["H2:H2"].Value = "Nro Ope";
                excelWorksheet2.Cells["I2:I2"].Value = "Comision";
                excelWorksheet2.Column(2).Width = 12.0;
                excelWorksheet2.Column(3).Width = 30.0;
                excelWorksheet2.Column(4).Width = 12.0;
                excelWorksheet2.Column(5).Width = 12.0;
                excelWorksheet2.Column(6).Width = 12.0;
                excelWorksheet2.Column(7).Width = 12.0;
                excelWorksheet2.Column(8).Width = 12.0;
                using (ExcelRange excelRange37 = excelWorksheet2.Cells["B2:I2"])
                {
                    excelRange37.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange37.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange37.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange37.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange37.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange37.Style.Font.Bold = true;
                    excelRange37.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                using (ExcelRange excelRange38 = excelWorksheet2.Cells["B2:I2"])
                {
                    excelRange38.AutoFilter = true;
                }
                int num19 = 3;
                foreach (ComparabancosUpeu item8 in ListComiNoAbo)
                {
                    excelWorksheet2.Cells[num19, 3].Value = item8.Descripcion;
                    excelWorksheet2.Cells[num19, 2].Value = item8.Terminal;
                    excelWorksheet2.Cells[num19, 4].Value = item8.FechaOpe.ToShortDateString();
                    excelWorksheet2.Cells[num19, 5].Value = item8.ImporteTransac;
                    excelWorksheet2.Cells[num19, 6].Value = item8.ImporteAbono;
                    excelWorksheet2.Cells[num19, 7].Value = item8.FechaAbono;
                    excelWorksheet2.Cells[num19, 8].Value = item8.NroOpe;
                    excelWorksheet2.Cells[num19, 9].Value = item8.Importe;
                    using (ExcelRange excelRange39 = excelWorksheet2.Cells[num19, 2, num19, 9])
                    {
                        excelRange39.Style.Font.Bold = false;
                        excelRange39.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        excelRange39.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        excelRange39.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        excelRange39.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        excelRange39.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        excelRange39.Style.Font.Color.SetColor(Color.Black);
                        excelRange39.Style.ShrinkToFit = true;
                    }
                    if (item8.Whoyo == "MC")
                    {
                        ExcelRange excelRange40 = excelWorksheet2.Cells[num19, 2, num19, 9];
                        excelRange40.Style.Font.Bold = true;
                        excelRange40.Style.Font.Color.SetColor(Color.Coral);
                    }
                    if (item8.Descripcion.Contains("VISA AGRUPADOS"))
                    {
                        ExcelRange excelRange41 = excelWorksheet2.Cells[num19, 2, num19, 9];
                        excelRange41.Style.Font.Bold = true;
                        excelRange41.Style.Font.Color.SetColor(Color.Tomato);
                    }
                    if (item8.Whoyo == "VN")
                    {
                        ExcelRange excelRange42 = excelWorksheet2.Cells[num19, 2, num19, 9];
                        excelRange42.Style.Font.Bold = true;
                        excelRange42.Style.Font.Color.SetColor(Color.Violet);
                    }
                    num19++;
                }
                num19++;
                num19++;
                foreach (ComparabancosUpeu item9 in ListAgrupaNoAbo)
                {
                    excelWorksheet2.Cells[num19, 3].Value = item9.Descripcion;
                    excelWorksheet2.Cells[num19, 2].Value = item9.Terminal;
                    excelWorksheet2.Cells[num19, 4].Value = item9.FechaOpe.ToShortDateString();
                    excelWorksheet2.Cells[num19, 5].Value = item9.ImporteTransac;
                    excelWorksheet2.Cells[num19, 6].Value = item9.ImporteAbono;
                    excelWorksheet2.Cells[num19, 7].Value = item9.FechaAbono;
                    excelWorksheet2.Cells[num19, 8].Value = item9.NroOpe;
                    excelWorksheet2.Cells[num19, 9].Value = item9.Importe;
                    using (ExcelRange excelRange43 = excelWorksheet2.Cells[num19, 2, num19, 9])
                    {
                        excelRange43.Style.Font.Bold = false;
                        excelRange43.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        excelRange43.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        excelRange43.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        excelRange43.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        excelRange43.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        excelRange43.Style.Font.Color.SetColor(Color.Black);
                        excelRange43.Style.ShrinkToFit = true;
                    }
                    if (item9.Whoyo == "MC")
                    {
                        ExcelRange excelRange44 = excelWorksheet2.Cells[num19, 2, num19, 9];
                        excelRange44.Style.Font.Bold = true;
                        excelRange44.Style.Font.Color.SetColor(Color.Coral);
                    }
                    if (item9.Descripcion.Contains("VISA AGRUPADOS"))
                    {
                        ExcelRange excelRange45 = excelWorksheet2.Cells[num19, 2, num19, 9];
                        excelRange45.Style.Font.Bold = true;
                        excelRange45.Style.Font.Color.SetColor(Color.Tomato);
                    }
                    if (item9.Whoyo == "VN")
                    {
                        ExcelRange excelRange46 = excelWorksheet2.Cells[num19, 2, num19, 9];
                        excelRange46.Style.Font.Bold = true;
                        excelRange46.Style.Font.Color.SetColor(Color.Violet);
                    }
                    num19++;
                }
            }
            if (ListComiAbo != null && ListComiAbo.Count > 0)
            {
                excelWorksheet3.Cells["B2:B2"].Value = "Terminal";
                excelWorksheet3.Cells["C2:C2"].Value = "Descripcion";
                excelWorksheet3.Cells["D2:D2"].Value = "Fecha Trans";
                excelWorksheet3.Cells["E2:E2"].Value = "Importe-tran";
                excelWorksheet3.Cells["F2:F2"].Value = "Importe-Abono";
                excelWorksheet3.Cells["G2:G2"].Value = "Fecha Abono";
                excelWorksheet3.Cells["H2:H2"].Value = "Nro Ope";
                excelWorksheet3.Cells["I2:I2"].Value = "Comision";
                excelWorksheet3.Column(2).Width = 12.0;
                excelWorksheet3.Column(3).Width = 30.0;
                excelWorksheet3.Column(4).Width = 12.0;
                excelWorksheet3.Column(5).Width = 12.0;
                excelWorksheet3.Column(6).Width = 12.0;
                excelWorksheet3.Column(7).Width = 12.0;
                excelWorksheet3.Column(8).Width = 12.0;
                using (ExcelRange excelRange47 = excelWorksheet3.Cells["B2:I2"])
                {
                    excelRange47.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange47.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange47.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange47.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange47.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange47.Style.Font.Bold = true;
                    excelRange47.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                int num20 = 3;
                foreach (ComparabancosUpeu item10 in ListComiAbo)
                {
                    excelWorksheet3.Cells[num20, 3].Value = item10.Descripcion;
                    excelWorksheet3.Cells[num20, 2].Value = item10.Terminal;
                    excelWorksheet3.Cells[num20, 4].Value = item10.FechaOpe.ToShortDateString();
                    excelWorksheet3.Cells[num20, 5].Value = item10.ImporteTransac;
                    excelWorksheet3.Cells[num20, 6].Value = item10.ImporteAbono;
                    excelWorksheet3.Cells[num20, 7].Value = item10.FechaAbono;
                    excelWorksheet3.Cells[num20, 8].Value = item10.NroOpe;
                    excelWorksheet3.Cells[num20, 9].Value = item10.Importe;
                    using (ExcelRange excelRange48 = excelWorksheet3.Cells[num20, 2, num20, 9])
                    {
                        excelRange48.Style.Font.Bold = false;
                        excelRange48.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        excelRange48.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        excelRange48.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        excelRange48.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        excelRange48.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        excelRange48.Style.Font.Color.SetColor(Color.Black);
                        excelRange48.Style.ShrinkToFit = true;
                    }
                    if (item10.Whoyo == "MC")
                    {
                        ExcelRange excelRange49 = excelWorksheet3.Cells[num20, 2, num20, 9];
                        excelRange49.Style.Font.Bold = true;
                        excelRange49.Style.Font.Color.SetColor(Color.Coral);
                    }
                    if (item10.Descripcion.Contains("VISA AGRUPADOS"))
                    {
                        ExcelRange excelRange50 = excelWorksheet3.Cells[num20, 2, num20, 9];
                        excelRange50.Style.Font.Bold = true;
                        excelRange50.Style.Font.Color.SetColor(Color.Tomato);
                    }
                    if (item10.Whoyo == "VN")
                    {
                        ExcelRange excelRange51 = excelWorksheet3.Cells[num20, 2, num20, 9];
                        excelRange51.Style.Font.Bold = true;
                        excelRange51.Style.Font.Color.SetColor(Color.Violet);
                    }
                    num20++;
                }
            }
            excelWorksheet9 = excelPackage.Workbook.Worksheets.Add("DataUnion");
            excelWorksheet9.Cells["A1:A1"].Value = "No ignore los pendiente Solucionalo!";
            excelWorksheet9.Cells["F1:J1"].Value = "Valores no Contables";
            excelWorksheet9.Cells["F2:G2"].Value = "UPEU";
            excelWorksheet9.Cells["H2:I2"].Value = "BANCO";
            excelWorksheet9.Cells[3, 1].Value = "Fecha";
            excelWorksheet9.Cells[3, 2].Value = "Mb";
            excelWorksheet9.Cells[3, 3].Value = "Nro Doc";
            excelWorksheet9.Cells[3, 4].Value = "Detalle";
            excelWorksheet9.Cells[3, 5].Value = "Fecha Ope";
            excelWorksheet9.Cells[3, 6].Value = "Debito";
            excelWorksheet9.Cells[3, 7].Value = "Credito";
            excelWorksheet9.Cells[3, 8].Value = "Debito";
            excelWorksheet9.Cells[3, 9].Value = "Credito";
            excelWorksheet9.Cells[3, 10].Value = "Terminal";
            excelWorksheet9.Cells[3, 11].Value = "Origen";
            excelWorksheet9.Column(1).Width = 15.0;
            excelWorksheet9.Column(2).Width = 12.0;
            excelWorksheet9.Column(3).Width = 12.0;
            excelWorksheet9.Column(4).Width = 50.0;
            excelWorksheet9.Column(5).Width = 13.0;
            excelWorksheet9.Column(6).Width = 13.0;
            excelWorksheet9.Column(7).Width = 13.0;
            excelWorksheet9.Column(8).Width = 13.0;
            excelWorksheet9.Column(9).Width = 13.0;
            excelWorksheet9.Column(10).Width = 13.0;
            excelWorksheet9.Column(1).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet9.Column(6).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet9.Column(7).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet9.Column(8).Style.Numberformat.Format = "#,##0.00";
            excelWorksheet9.Column(9).Style.Numberformat.Format = "#,##0.00";
            using (ExcelRange excelRange52 = excelWorksheet9.Cells["A1:E1"])
            {
                excelRange52.Merge = true;
                excelRange52.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange52.Style.Font.Bold = true;
                excelRange52.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelRange52.Style.Border.Top.Style = ExcelBorderStyle.Thick;
                excelRange52.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                excelRange52.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }
            using (ExcelRange excelRange53 = excelWorksheet9.Cells["F1:J1"])
            {
                excelRange53.Merge = true;
                excelRange53.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange53.Style.Font.Bold = true;
                excelRange53.Style.Border.Top.Style = ExcelBorderStyle.Thick;
                excelRange53.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange53.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelRange53.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange53.Style.Border.Right.Style = ExcelBorderStyle.Thick;
            }
            using (ExcelRange excelRange54 = excelWorksheet9.Cells["A2:E2"])
            {
                excelRange54.Merge = true;
                excelRange54.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange54.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange54.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                excelRange54.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange54.Style.Font.Bold = true;
                excelRange54.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange55 = excelWorksheet9.Cells["F2:G2"])
            {
                excelRange55.Merge = true;
                excelRange55.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange55.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange55.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange55.Style.Font.Bold = true;
                excelRange55.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange56 = excelWorksheet9.Cells["H2:I2"])
            {
                excelRange56.Merge = true;
                excelRange56.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange56.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange56.Style.Border.Right.Style = ExcelBorderStyle.Thick;
                excelRange56.Style.Font.Bold = true;
                excelRange56.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange57 = excelWorksheet9.Cells["A3:I3"])
            {
                excelRange57.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange57.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                excelRange57.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange57.Style.Font.Bold = true;
                excelRange57.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange58 = excelWorksheet9.Cells["A3:A3"])
            {
                excelRange58.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange58.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                excelRange58.Style.Font.Bold = true;
                excelRange58.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (ExcelRange excelRange59 = excelWorksheet9.Cells["A3:k3"])
            {
                excelRange59.AutoFilter = true;
            }
            int num21 = 5;
            excelWorksheet9.Cells[4, 1].Value = string.Empty;
            excelWorksheet9.Cells[4, 2].Value = string.Empty;
            if (ListComparabcp.Count > 0)
            {
                excelWorksheet9.Cells[4, 4].Value = "Saldo al " + GetLastDayOf(ListComparabcp.First().FechaOpe);
            }
            excelWorksheet9.Cells[4, 6].Value = SaldoFiNUpeu;
            excelWorksheet9.Cells[4, 8].Value = SaldoFiNBanco;
            foreach (BancoUpeu item11 in listaPenBanco.OrderBy((BancoUpeu w) => w.NroOpe).ToList())
            {
                excelWorksheet9.Cells[num21, 1].Value = item11.FechaOperacion;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = item11.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item11.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                if (item11.Dh == 1)
                {
                    excelWorksheet9.Cells[num21, 8].Value = item11.Importe;
                }
                else
                {
                    excelWorksheet9.Cells[num21, 9].Value = item11.Importe;
                }
                excelWorksheet9.Cells[num21, 10].Value = item11.CodigoPos;
                excelWorksheet9.Cells[num21, 11].Value = "Pendiente";
                using (ExcelRange excelRange60 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange60.Style.Font.Bold = false;
                    excelRange60.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange60.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange60.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange60.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange60.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange60.Style.Font.Color.SetColor(Color.Black);
                    excelRange60.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (BancoBCP item12 in listaPenUpe.OrderBy((BancoBCP w) => w.NroOpe).ToList())
            {
                excelWorksheet9.Cells[num21, 1].Value = item12.FechaOperacion;
                excelWorksheet9.Cells[num21, 2].Value = item12.ReferenciaVoucher;
                excelWorksheet9.Cells[num21, 3].Value = item12.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item12.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                excelWorksheet9.Cells[num21, 10].Value = item12.CodigoPos;
                excelWorksheet9.Cells[num21, 11].Value = "Pendiente";
                using (ExcelRange excelRange61 = excelWorksheet9.Cells[num21, 3, num21, 3])
                {
                    excelRange61.Style.Font.Bold = true;
                    excelRange61.Style.ShrinkToFit = true;
                }
                if (item12.Dh == 1)
                {
                    excelWorksheet9.Cells[num21, 6].Value = item12.Importe;
                }
                else
                {
                    excelWorksheet9.Cells[num21, 7].Value = item12.Importe;
                }
                if (item12.Whoyo == "1")
                {
                    ExcelRange excelRange62 = excelWorksheet6.Cells[num21, 1, num21, 7];
                    excelRange62.Style.Font.Bold = true;
                    excelRange62.Style.Font.Color.SetColor(Color.Firebrick);
                }
                if (item12.Whoyo == "MC")
                {
                    ExcelRange excelRange63 = excelWorksheet6.Cells[num21, 1, num21, 7];
                    excelRange63.Style.Font.Bold = true;
                    excelRange63.Style.Font.Color.SetColor(Color.Coral);
                }
                if (item12.Whoyo == "VN")
                {
                    ExcelRange excelRange64 = excelWorksheet6.Cells[num21, 1, num21, 7];
                    excelRange64.Style.Font.Bold = true;
                    excelRange64.Style.Font.Color.SetColor(Color.DarkViolet);
                }
                using (ExcelRange excelRange65 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange65.Style.Font.Bold = false;
                    excelRange65.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange65.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange65.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange65.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange65.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                }
                num21++;
            }
            foreach (BancoUpeu item13 in listaMayorUpeu)
            {
                excelWorksheet9.Cells[num21, 1].Value = item13.FechaRegistro;
                excelWorksheet9.Cells[num21, 2].Value = item13.ReferenciaLibros;
                excelWorksheet9.Cells[num21, 3].Value = item13.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item13.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = item13.FechaOperacion;
                if (item13.Dh == 1)
                {
                    excelWorksheet9.Cells[num21, 8].Value = item13.Importe;
                }
                else
                {
                    excelWorksheet9.Cells[num21, 9].Value = item13.Importe;
                }
                excelWorksheet9.Cells[num21, 10].Value = item13.CodigoPos;
                excelWorksheet9.Cells[num21, 11].Value = "Mayor";
                using (ExcelRange excelRange66 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange66.Style.Font.Bold = false;
                    excelRange66.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange66.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange66.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange66.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange66.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange66.Style.Font.Color.SetColor(Color.Black);
                    excelRange66.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (BancoBCP item14 in listaBancoGeneral)
            {
                excelWorksheet9.Cells[num21, 1].Value = item14.FechaOperacion;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = item14.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item14.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                if (item14.Dh == 1)
                {
                    excelWorksheet9.Cells[num21, 6].Value = item14.Importe;
                }
                else
                {
                    excelWorksheet9.Cells[num21, 7].Value = item14.Importe;
                }
                excelWorksheet9.Cells[num21, 10].Value = item14.CodigoPos;
                excelWorksheet9.Cells[num21, 11].Value = "Bancos";
                using (ExcelRange excelRange67 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange67.Style.Font.Bold = false;
                    excelRange67.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange67.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange67.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange67.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange67.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange67.Style.Font.Color.SetColor(Color.Black);
                    excelRange67.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (BancoBCP item15 in listaCafetinAux.Where((BancoBCP w) => w.NroOpe != string.Empty))
            {
                excelWorksheet9.Cells[num21, 1].Value = item15.FechaOperacion;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = item15.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item15.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                if (item15.Dh == 1)
                {
                    excelWorksheet9.Cells[num21, 6].Value = item15.Importe;
                }
                else
                {
                    excelWorksheet9.Cells[num21, 7].Value = item15.Importe;
                }
                excelWorksheet9.Cells[num21, 10].Value = item15.CodigoPos;
                excelWorksheet9.Cells[num21, 11].Value = "Visa";
                using (ExcelRange excelRange68 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange68.Style.Font.Bold = false;
                    excelRange68.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange68.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange68.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange68.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange68.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange68.Style.Font.Color.SetColor(Color.Black);
                    excelRange68.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (BancoBCP item16 in listaMasterC.Where((BancoBCP w) => w.NroOpe != string.Empty))
            {
                excelWorksheet9.Cells[num21, 1].Value = item16.FechaOperacion;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = item16.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item16.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                if (item16.Dh == 1)
                {
                    excelWorksheet9.Cells[num21, 6].Value = item16.Importe;
                }
                else
                {
                    excelWorksheet9.Cells[num21, 7].Value = item16.Importe;
                }
                excelWorksheet9.Cells[num21, 10].Value = item16.CodigoPos;
                excelWorksheet9.Cells[num21, 11].Value = "Mastercard";
                using (ExcelRange excelRange69 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange69.Style.Font.Bold = false;
                    excelRange69.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange69.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange69.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange69.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange69.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange69.Style.Font.Color.SetColor(Color.Black);
                    excelRange69.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (BancoBCP item17 in listaMasterAEx.Where((BancoBCP w) => w.NroOpe != string.Empty))
            {
                excelWorksheet9.Cells[num21, 1].Value = item17.FechaOperacion;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = item17.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item17.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                if (item17.Dh == 1)
                {
                    excelWorksheet9.Cells[num21, 6].Value = item17.Importe;
                }
                else
                {
                    excelWorksheet9.Cells[num21, 7].Value = item17.Importe;
                }
                excelWorksheet9.Cells[num21, 10].Value = item17.CodigoPos;
                excelWorksheet9.Cells[num21, 11].Value = "Ameri.Express";
                using (ExcelRange excelRange70 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange70.Style.Font.Bold = false;
                    excelRange70.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange70.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange70.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange70.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange70.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange70.Style.Font.Color.SetColor(Color.Black);
                    excelRange70.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (ComparabancosUpeu item18 in ListComiAbo)
            {
                excelWorksheet9.Cells[num21, 1].Value = item18.FechaAbono;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = "VISANET";
                excelWorksheet9.Cells[num21, 4].Value = item18.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                excelWorksheet9.Cells[num21, 6].Value = string.Empty;
                excelWorksheet9.Cells[num21, 7].Value = item18.ImporteAbono;
                excelWorksheet9.Cells[num21, 10].Value = item18.Terminal;
                excelWorksheet9.Cells[num21, 11].Value = "Abonado";
                using (ExcelRange excelRange71 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange71.Style.Font.Bold = false;
                    excelRange71.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange71.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange71.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange71.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange71.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange71.Style.Font.Color.SetColor(Color.Black);
                    excelRange71.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (var item19 in (from p in ListComiNoAbo
                                    group p by new { p.Whoyo }).ToList())
            {
                objCompupe = new ComparabancosUpeu
                {
                    Descripcion = "Comision no Abono-" + item19.Key.Whoyo,
                    FechaOpe = DateTime.Parse(item19.First().FechaAbono),
                    NroOpe = item19.First().NroOpe,
                    Importe = item19.Sum((ComparabancosUpeu w) => w.Importe),
                    ImporteAbono = 0m,
                    ImporteTransac = 0m,
                    Dh = 2,
                    Terminal = item19.First().Terminal,
                    Observacion = "Comision-No Abonado",
                    Whoyo = item19.Key.Whoyo,
                    FechaAbono = item19.First().FechaAbono,
                    Mb = string.Empty
                };
                excelWorksheet9.Cells[num21, 1].Value = objCompupe.FechaAbono;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = objCompupe.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = objCompupe.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                excelWorksheet9.Cells[num21, 6].Value = string.Empty;
                excelWorksheet9.Cells[num21, 7].Value = objCompupe.Importe;
                excelWorksheet9.Cells[num21, 10].Value = objCompupe.Terminal;
                excelWorksheet9.Cells[num21, 11].Value = "Comision No Abono";
                using (ExcelRange excelRange72 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange72.Style.Font.Bold = false;
                    excelRange72.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange72.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange72.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange72.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange72.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange72.Style.Font.Color.SetColor(Color.Black);
                    excelRange72.Style.ShrinkToFit = true;
                }
                num21++;
            }
            foreach (ComparabancosUpeu item20 in ListAgrupaNoAbo)
            {
                excelWorksheet9.Cells[num21, 1].Value = item20.FechaOpe;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = item20.NroOpe;
                excelWorksheet9.Cells[num21, 4].Value = item20.Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                excelWorksheet9.Cells[num21, 6].Value = string.Empty;
                excelWorksheet9.Cells[num21, 7].Value = item20.Importe;
                excelWorksheet9.Cells[num21, 10].Value = item20.Terminal;
                excelWorksheet9.Cells[num21, 11].Value = "No Abonado";
                using (ExcelRange excelRange73 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange73.Style.Font.Bold = false;
                    excelRange73.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange73.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange73.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange73.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange73.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange73.Style.Font.Color.SetColor(Color.Black);
                    excelRange73.Style.ShrinkToFit = true;
                }
                num21++;
            }
            var list = (from p in ListComiAbo
                        group p by new { p.Whoyo }).ToList();
            foreach (var item21 in list)
            {
                decimal num22 = item21.Sum((ComparabancosUpeu w) => w.Importe);
                string value = string.Empty;
                if (item21.Key.Whoyo.Equals("V-NET"))
                {
                    value = "COM-VN";
                }
                else if (item21.Key.Whoyo.Equals("M-CARD"))
                {
                    value = "COM-MC";
                }
                else if (item21.Key.Whoyo.Equals("A-EXP"))
                {
                    value = "COM-AE";
                }
                excelWorksheet9.Cells[num21, 1].Value = item21.First().FechaOpe;
                excelWorksheet9.Cells[num21, 2].Value = string.Empty;
                excelWorksheet9.Cells[num21, 3].Value = value;
                excelWorksheet9.Cells[num21, 4].Value = item21.First().Descripcion;
                excelWorksheet9.Cells[num21, 5].Value = string.Empty;
                excelWorksheet9.Cells[num21, 6].Value = string.Empty;
                excelWorksheet9.Cells[num21, 7].Value = num22;
                excelWorksheet9.Cells[num21, 10].Value = "0";
                excelWorksheet9.Cells[num21, 11].Value = "ComiAbono";
                using (ExcelRange excelRange74 = excelWorksheet9.Cells[num21, 1, num21, 11])
                {
                    excelRange74.Style.Font.Bold = false;
                    excelRange74.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    excelRange74.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    excelRange74.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    excelRange74.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    excelRange74.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    excelRange74.Style.Font.Color.SetColor(Color.Black);
                    excelRange74.Style.ShrinkToFit = true;
                }
                num21++;
            }
            string cuentaContableBanco = CuentaContableBanco;
            FileInfo file = new FileInfo("C:\\Excel\\Conci_" + NombreBanco + "_" + MonthName(MesTrabajo.Month) + "_" + cuentaContableBanco + ".xlsx");
            try
            {
                excelPackage.Workbook.Calculate();
                excelPackage.SaveAs(file);
                excelPackage.Dispose();
            }
            catch (Exception)
            {
                MessageBox.Show("Crea una carpeta Excel en la Unidad C: ", "Informacion");
            }
        }

        public string MonthName(int month)
        {
            DateTimeFormatInfo dateTimeFormat = new CultureInfo("es-ES", useUserOverride: false).DateTimeFormat;
            return dateTimeFormat.GetMonthName(month);
        }

        private void DivujaExcelAudit()
        {
            ExcelPackage excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "Daniel Antazu - Movil 992144164";
            excelPackage.Workbook.Properties.Title = "Conciliacion";
            excelPackage.Workbook.Properties.Subject = "Se leen 4 hojas de excel visa, pendientes, mayor upeu,mayor banco en ese orden si no hay datos de uno de elllos dejarlo la hoja en blanco en ese orden";
            excelPackage.Workbook.Properties.Created = DateTime.Now;
            excelPackage.Workbook.CalcMode = ExcelCalcMode.Manual;
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("Upeu");
            ExcelWorksheet excelWorksheet2 = excelPackage.Workbook.Worksheets.Add("Banco");
            ExcelWorksheet excelWorksheet3 = excelPackage.Workbook.Worksheets.Add("Noconcilia");
            ExcelWorksheet excelWorksheet4 = excelPackage.Workbook.Worksheets.Add("Concilia");
            excelWorksheet.Cells["B2:B2"].Value = "Fecha";
            excelWorksheet.Cells["C2:C2"].Value = "Nro Ope";
            excelWorksheet.Cells["D2:D2"].Value = "Descripcion";
            excelWorksheet.Cells["E2:E2"].Value = "D";
            excelWorksheet.Cells["F2:F2"].Value = "H";
            excelWorksheet.Cells["G2:G2"].Value = "Terminal";
            excelWorksheet.Column(2).Width = 12.0;
            excelWorksheet.Column(3).Width = 12.0;
            excelWorksheet.Column(4).Width = 30.0;
            excelWorksheet.Column(5).Width = 12.0;
            excelWorksheet.Column(6).Width = 12.0;
            excelWorksheet.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet.Column(5).Style.Numberformat.Format = "#,##0.00";
            int num = 3;
            int num2 = 3;
            foreach (BancoUpeu item in listasPendienteBancoYmayorUpeu.OrderByDescending((BancoUpeu w) => w.NroOpe).ToList())
            {
                excelWorksheet.Cells[num, 2].Value = item.FechaOperacion;
                excelWorksheet.Cells[num, 3].Value = item.NroOpe;
                excelWorksheet.Cells[num, 4].Value = item.ReferenciaLibros + "-" + item.Descripcion;
                if (item.Dh == 1)
                {
                    excelWorksheet.Cells[num, 5].Value = item.Importe;
                    excelWorksheet.Cells[num, 6].Value = 0;
                }
                else
                {
                    excelWorksheet.Cells[num, 6].Value = item.Importe;
                    excelWorksheet.Cells[num, 5].Value = 0;
                }
                excelWorksheet.Cells[num, 7].Value = item.CodigoPos;
                num++;
                ExcelRange excelRange = excelWorksheet.Cells[num, 3, num, 7];
                excelRange.Style.Font.Bold = false;
                excelRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Font.Color.SetColor(Color.Black);
                excelRange.Style.ShrinkToFit = true;
            }
            int num3 = num - 1;
            string address = "E" + num;
            string address2 = "F" + num;
            excelWorksheet.Cells[address].Formula = "SUM(E3:E" + num3 + ")";
            excelWorksheet.Cells[address].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet.Cells[address2].Formula = "SUM(F3:F" + num3 + ")";
            excelWorksheet.Cells[address2].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet.Cells[address].Calculate();
            excelWorksheet.Cells[address2].Calculate();
            decimal num4 = decimal.Parse(excelWorksheet.Cells[address].Value.ToString());
            decimal num5 = decimal.Parse(excelWorksheet.Cells[address2].Value.ToString());
            decimal num6 = SaldoFiNUpeu - num4;
            excelWorksheet.Cells[num + 1, 5].Value = SaldoFiNUpeu;
            excelWorksheet.Cells[num + 1, 4].Value = "Saldo Fin (-)";
            excelWorksheet.Cells[num + 2, 5].Value = ((num6 > 0m) ? num6 : (-1m * num6));
            SumadebeUpeu = num5 - ((num6 > 0m) ? num6 : (-1m * num6));
            excelWorksheet.Cells[num + 2, 6].Value = SumadebeUpeu;
            excelWorksheet2.Cells["B2:B2"].Value = "Fecha";
            excelWorksheet2.Cells["C2:C2"].Value = "Nro Ope";
            excelWorksheet2.Cells["D2:D2"].Value = "Descripcion";
            excelWorksheet2.Cells["E2:E2"].Value = "D";
            excelWorksheet2.Cells["F2:F2"].Value = "H";
            excelWorksheet2.Cells["G2:G2"].Value = "Terminal";
            excelWorksheet2.Column(2).Width = 12.0;
            excelWorksheet2.Column(3).Width = 12.0;
            excelWorksheet2.Column(4).Width = 30.0;
            excelWorksheet2.Column(5).Width = 12.0;
            excelWorksheet2.Column(6).Width = 12.0;
            excelWorksheet2.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet2.Column(5).Style.Numberformat.Format = "#,##0.00";
            foreach (BancoBCP item2 in listasPendienteUpeuYBancos.OrderByDescending((BancoBCP w) => w.NroOpe).ToList())
            {
                excelWorksheet2.Cells[num2, 2].Value = item2.FechaOperacion;
                excelWorksheet2.Cells[num2, 3].Value = item2.NroOpe;
                excelWorksheet2.Cells[num2, 4].Value = item2.Descripcion;
                if (item2.Dh == 1)
                {
                    excelWorksheet2.Cells[num2, 5].Value = item2.Importe;
                }
                else
                {
                    excelWorksheet2.Cells[num2, 6].Value = item2.Importe;
                }
                excelWorksheet2.Cells[num2, 7].Value = item2.CodigoPos;
                num2++;
                ExcelRange excelRange2 = excelWorksheet2.Cells[num2, 3, num2, 7];
                excelRange2.Style.Font.Bold = false;
                excelRange2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Font.Color.SetColor(Color.Black);
                excelRange2.Style.ShrinkToFit = true;
            }
            int num7 = num2 - 1;
            string address3 = "E" + num2;
            string address4 = "F" + num2;
            excelWorksheet2.Cells[address3].Formula = "SUM(E3:E" + num7 + ")";
            excelWorksheet2.Cells[address3].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet2.Cells[address4].Formula = "SUM(F3:F" + num7 + ")";
            excelWorksheet2.Cells[address4].Style.Numberformat.Format = "#,##0.00";
            excelWorksheet2.Cells[address3].Calculate();
            excelWorksheet2.Cells[address4].Calculate();
            decimal num8 = decimal.Parse(excelWorksheet2.Cells[address3].Value.ToString());
            decimal num9 = decimal.Parse(excelWorksheet2.Cells[address4].Value.ToString());
            decimal num10 = SaldoFiNBanco - num8;
            excelWorksheet2.Cells[num2 + 1, 5].Value = SaldoFiNBanco;
            excelWorksheet2.Cells[num2 + 1, 4].Value = "Saldo Fin(-)";
            excelWorksheet2.Cells[num2 + 2, 5].Value = ((num10 > 0m) ? num10 : (-1m * num10));
            SumadebeBanco = num9 - ((num10 > 0m) ? num10 : (-1m * num10));
            excelWorksheet2.Cells[num2 + 2, 6].Value = SumadebeBanco;
            excelWorksheet3.Cells["B2:B2"].Value = "Fecha";
            excelWorksheet3.Cells["C2:C2"].Value = "Nro Ope";
            excelWorksheet3.Cells["D2:D2"].Value = "Descripcion";
            excelWorksheet3.Cells["E2:E2"].Value = "Importe-tran";
            excelWorksheet3.Cells["F2:F2"].Value = "dh";
            excelWorksheet3.Cells["G2:G2"].Value = "Terminal";
            excelWorksheet3.Column(2).Width = 12.0;
            excelWorksheet3.Column(3).Width = 12.0;
            excelWorksheet3.Column(4).Width = 30.0;
            excelWorksheet3.Column(5).Width = 12.0;
            excelWorksheet3.Column(6).Width = 12.0;
            excelWorksheet3.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet3.Column(5).Style.Numberformat.Format = "#,##0.00";
            int num11 = 3;
            foreach (BancoBCP item3 in listaNoexisten)
            {
                excelWorksheet3.Cells[num11, 2].Value = item3.FechaOperacion;
                excelWorksheet3.Cells[num11, 3].Value = item3.NroOpe;
                excelWorksheet3.Cells[num11, 4].Value = item3.Descripcion;
                excelWorksheet3.Cells[num11, 5].Value = item3.Importe;
                excelWorksheet3.Cells[num11, 6].Value = item3.Dh;
                excelWorksheet3.Cells[num11, 7].Value = item3.CodigoPos;
                num11++;
                ExcelRange excelRange3 = excelWorksheet3.Cells[num11, 3, num11, 7];
                excelRange3.Style.Font.Bold = false;
                excelRange3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Font.Color.SetColor(Color.Black);
                excelRange3.Style.ShrinkToFit = true;
            }
            excelWorksheet4.Cells["C2:C2"].Value = "Nro Ope";
            excelWorksheet4.Cells["D2:D2"].Value = "Descripcion";
            excelWorksheet4.Cells["E2:E2"].Value = "Importe-tran";
            excelWorksheet4.Cells["F2:F2"].Value = "dh";
            excelWorksheet4.Cells["G2:G2"].Value = "Terminal";
            excelWorksheet4.Cells["B2:B2"].Value = "Fecha";
            excelWorksheet4.Column(2).Width = 12.0;
            excelWorksheet4.Column(3).Width = 12.0;
            excelWorksheet4.Column(4).Width = 30.0;
            excelWorksheet4.Column(5).Width = 12.0;
            excelWorksheet4.Column(6).Width = 12.0;
            excelWorksheet4.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet4.Column(5).Style.Numberformat.Format = "#,##0.00";
            int num12 = 3;
            foreach (BancoBCP item4 in listaFechaMonto)
            {
                excelWorksheet4.Cells[num12, 2].Value = item4.FechaOperacion;
                excelWorksheet4.Cells[num12, 3].Value = item4.NroOpe;
                excelWorksheet4.Cells[num12, 4].Value = item4.Descripcion;
                excelWorksheet4.Cells[num12, 5].Value = item4.Importe;
                excelWorksheet4.Cells[num12, 6].Value = item4.Dh;
                excelWorksheet4.Cells[num12, 7].Value = item4.CodigoPos;
                num12++;
                ExcelRange excelRange4 = excelWorksheet4.Cells[num12, 3, num12, 7];
                excelRange4.Style.Font.Bold = false;
                excelRange4.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange4.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange4.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange4.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange4.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange4.Style.Font.Color.SetColor(Color.Black);
                excelRange4.Style.ShrinkToFit = true;
            }
            FileInfo file = new FileInfo("C:\\Excel\\AUDIT-" + CuentaContableBanco + ".xlsx");
            excelPackage.Workbook.Calculate();
            excelPackage.SaveAs(file);
            excelPackage.Dispose();
        }

        private void DivujaExcelOnly()
        {
            ExcelPackage excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "Daniel Antazu - Movil 992144164";
            excelPackage.Workbook.Properties.Title = "Conciliacion";
            excelPackage.Workbook.Properties.Subject = "Se leen 4 hojas de excel visa, pendientes, mayor upeu,mayor banco en ese orden si no hay datos de uno de elllos dejarlo la hoja en blanco en ese orden";
            excelPackage.Workbook.Properties.Created = DateTime.Now;
            excelPackage.Workbook.CalcMode = ExcelCalcMode.Manual;
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("Extras");
            excelWorksheet.Cells["B2:B2"].Value = "Fecha";
            excelWorksheet.Cells["C2:C2"].Value = "Nro Ope";
            excelWorksheet.Cells["D2:D2"].Value = "Descripcion";
            excelWorksheet.Cells["E2:E2"].Value = "Importe-tran";
            excelWorksheet.Cells["F2:F2"].Value = "Importe-Abono";
            excelWorksheet.Cells["G2:G2"].Value = "Comision";
            excelWorksheet.Cells["H2:H2"].Value = "Terminal";
            excelWorksheet.Column(2).Width = 12.0;
            excelWorksheet.Column(3).Width = 12.0;
            excelWorksheet.Column(4).Width = 30.0;
            excelWorksheet.Column(5).Width = 12.0;
            excelWorksheet.Column(6).Width = 12.0;
            excelWorksheet.Column(2).Style.Numberformat.Format = "dd/MM/yyyy";
            excelWorksheet.Column(5).Style.Numberformat.Format = "#,##0.00";
            using (ExcelRange excelRange = excelWorksheet.Cells["B2:G2"])
            {
                excelRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange.Style.Font.Bold = true;
                excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            int num = 3;
            foreach (BancoBCP item in listaBancoGeneral)
            {
            }
            foreach (ComparabancosBcp item2 in listAuxbcp)
            {
                excelWorksheet.Cells[num, 2].Value = item2.FechaOpe;
                excelWorksheet.Cells[num, 3].Value = item2.NroOpe;
                excelWorksheet.Cells[num, 4].Value = item2.Descripcion;
                excelWorksheet.Cells[num, 5].Value = item2.Importe;
                excelWorksheet.Cells[num, 6].Value = 0;
                excelWorksheet.Cells[num, 7].Value = 0;
                excelWorksheet.Cells[num, 8].Value = item2.Terminal;
                num++;
                ExcelRange excelRange2 = excelWorksheet.Cells[num, 3, num, 8];
                excelRange2.Style.Font.Bold = false;
                excelRange2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange2.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange2.Style.Font.Color.SetColor(Color.Black);
                excelRange2.Style.ShrinkToFit = true;
            }
            foreach (ComparabancosUpeu item3 in listAuxupe)
            {
                excelWorksheet.Cells[num, 2].Value = item3.FechaOpe;
                excelWorksheet.Cells[num, 3].Value = item3.NroOpe;
                excelWorksheet.Cells[num, 4].Value = item3.Descripcion;
                excelWorksheet.Cells[num, 5].Value = item3.Importe;
                excelWorksheet.Cells[num, 6].Value = 0;
                excelWorksheet.Cells[num, 7].Value = 0;
                excelWorksheet.Cells[num, 8].Value = item3.Terminal;
                num++;
                ExcelRange excelRange3 = excelWorksheet.Cells[num, 3, num, 8];
                excelRange3.Style.Font.Bold = false;
                excelRange3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                excelRange3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelRange3.Style.Font.Color.SetColor(Color.Black);
                excelRange3.Style.ShrinkToFit = true;
            }
            FileInfo file = new FileInfo("C:\\Excel\\Estra.xlsx");
            excelPackage.Workbook.Calculate();
            excelPackage.SaveAs(file);
            excelPackage.Dispose();
        }

        private bool CompruebaSiExiste(string pNroOpe, string pTerminal)
        {
            bool result = true;
            try
            {
                List<UnidosUpeuyBancos> list = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeU == pNroOpe && w.Terminal == pTerminal).ToList();
                List<UnidosUpeuyBancos> list2 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeU == pNroOpe && w.Terminal == pTerminal).ToList();
                List<UnidosUpeuyBancos> list3 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeB == pNroOpe && w.Terminal == pTerminal).ToList();
                List<UnidosUpeuyBancos> list4 = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeB == pNroOpe && w.Terminal == pTerminal).ToList();
                List<ComparabancosBcp> list5 = ListComparabcp.Where((ComparabancosBcp w) => w.NroOpe == pNroOpe && w.Terminal == pTerminal).ToList();
                List<ComparabancosUpeu> list6 = ListComparaupeu.Where((ComparabancosUpeu w) => w.NroOpe == pNroOpe && w.Terminal == pTerminal).ToList();
                if (list.Count == 0 && list4.Count == 0 && list5.Count == 0 && list6.Count == 0 && list2.Count == 0 && list3.Count == 0)
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }

        private bool CompruebaSiExiste(string pNroOpe, string pTerminal, decimal pImporte)
        {
            bool result = true;
            try
            {
                List<UnidosUpeuyBancos> list = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeU == pNroOpe && w.Terminal == pTerminal && w.ImporteU == pImporte).ToList();
                List<UnidosUpeuyBancos> list2 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeU == pNroOpe && w.Terminal == pTerminal && w.ImporteU == pImporte).ToList();
                List<UnidosUpeuyBancos> list3 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeB == pNroOpe && w.Terminal == pTerminal && w.ImporteB == pImporte).ToList();
                List<UnidosUpeuyBancos> list4 = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeB == pNroOpe && w.Terminal == pTerminal && w.ImporteB == pImporte).ToList();
                List<ComparabancosBcp> list5 = ListComparabcp.Where((ComparabancosBcp w) => w.NroOpe == pNroOpe && w.Terminal == pTerminal && w.Importe == pImporte).ToList();
                List<ComparabancosUpeu> list6 = ListComparaupeu.Where((ComparabancosUpeu w) => w.NroOpe == pNroOpe && w.Terminal == pTerminal && w.Importe == pImporte).ToList();
                if (list.Count == 0 && list4.Count == 0 && list5.Count == 0 && list6.Count == 0 && list2.Count == 0 && list3.Count == 0)
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }

        private bool CompruebaSiExisteDup(decimal pImporte, string pTerminal, string pFechaOpe)
        {
            bool result = true;
            try
            {
                List<UnidosUpeuyBancos> list = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.ImporteU == pImporte && w.Terminal == pTerminal && w.FechaOperacionU == pFechaOpe).ToList();
                List<UnidosUpeuyBancos> list2 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.ImporteU == pImporte && w.Terminal == pTerminal && w.FechaOperacionU == pFechaOpe).ToList();
                List<UnidosUpeuyBancos> list3 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.ImporteU == pImporte && w.Terminal == pTerminal && w.FechaOperacionU == pFechaOpe).ToList();
                List<UnidosUpeuyBancos> list4 = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.ImporteU == pImporte && w.Terminal == pTerminal && w.FechaOperacionU == pFechaOpe).ToList();
                List<ComparabancosBcp> list5 = ListComparabcp.Where((ComparabancosBcp w) => w.Importe == pImporte && w.Terminal == pTerminal && w.FechaOpe == DateTime.Parse(pFechaOpe)).ToList();
                List<ComparabancosUpeu> list6 = ListComparaupeu.Where((ComparabancosUpeu w) => w.Importe == pImporte && w.Terminal == pTerminal && w.FechaOpe == DateTime.Parse(pFechaOpe)).ToList();
                if (list.Count == 0 && list4.Count == 0 && list5.Count == 0 && list6.Count == 0 && list2.Count == 0 && list3.Count == 0)
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }

        private bool CompruebaSiExisteVisanet(string pNroOpe, string pTerminal, decimal pImporte, string pDescripcion)
        {
            bool result = true;
            try
            {
                List<UnidosUpeuyBancos> list = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeU == pNroOpe && w.ImporteU == pImporte && w.DescripcionU == pDescripcion).ToList();
                List<UnidosUpeuyBancos> list2 = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeB == pNroOpe && w.ImporteB == pImporte && w.DescripcionB == pDescripcion).ToList();
                List<ComparabancosBcp> list3 = ListComparabcp.Where((ComparabancosBcp w) => w.NroOpe == pNroOpe && w.Importe == pImporte && w.Descripcion == pDescripcion).ToList();
                List<ComparabancosUpeu> list4 = ListComparaupeu.Where((ComparabancosUpeu w) => w.NroOpe == pNroOpe && w.Importe == pImporte && w.Descripcion == pDescripcion).ToList();
                if (list.Count == 0 && list2.Count == 0 && list3.Count == 0 && list4.Count == 0)
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }

        private bool CompruebaSiExiste(string pNroOpe, string pTerminal, string pMb)
        {
            bool result = true;
            try
            {
                List<UnidosUpeuyBancos> list = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.NroOpeU == pNroOpe && w.Terminal == pTerminal && w.ReferenciaLibrosU == pMb).ToList();
                List<ComparabancosBcp> list2 = ListComparabcp.Where((ComparabancosBcp w) => w.NroOpe == pNroOpe && w.Terminal == pTerminal && w.Referencialibros == pMb).ToList();

               
                if (list2.Count == 0 && list.Count == 0)
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }
        private bool CompruebaSiExisteExtra(string pNroOpe, string pTerminal, string pMExtra)
        {
            bool result = true;
            try
            {
                List<UnidosUpeuyBancos> list3 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.NroOpeU == pNroOpe && w.Terminal == pTerminal && w.ReferenciaLibrosU == pMExtra).ToList();
               
                
                if (list3.Count == 0 )
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }


        private bool CompruebaDatos(string pFecha, string pTerminal, decimal pImporte)
        {
            bool result = true;
            try
            {
                List<UnidosUpeuyBancos> list = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.FechaOperacionU == pFecha && w.ImporteU == pImporte).ToList();
                List<UnidosUpeuyBancos> list2 = ListaUnidosOnly.Where((UnidosUpeuyBancos w) => w.FechaOperacionB == pFecha && w.ImporteB == pImporte).ToList();
                List<UnidosUpeuyBancos> list3 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.FechaOperacionU == pFecha && w.ImporteU == pImporte).ToList();
                List<UnidosUpeuyBancos> list4 = ListaUnidosUyB.Where((UnidosUpeuyBancos w) => w.FechaOperacionB == pFecha && w.ImporteB == pImporte).ToList();
                List<ComparabancosBcp> list5 = ListComparabcp.Where((ComparabancosBcp w) => w.FechaOpe == DateTime.Parse(pFecha) && w.Importe == pImporte).ToList();
                List<ComparabancosUpeu> list6 = ListComparaupeu.Where((ComparabancosUpeu w) => w.FechaOpe == DateTime.Parse(pFecha) && w.Importe == pImporte).ToList();
                if (list.Count == 0 && list2.Count == 0 && list5.Count == 0 && list6.Count == 0 && list3.Count == 0 && list4.Count == 0)
                {
                    result = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return result;
        }
        public void ProcesaSolComisionGroup()
        {
            ListComiTotAbo = new List<ComparabancosUpeu>();
            List<TablaUpeu> descripcionTer = ExcelUtil.GetDescripcionTerFile();
            foreach (BancoBCP itemv in listaVisaMCAEGroup)
            {
                string descripcion = descripcionTer.Where((TablaUpeu w) => w.Terminal == itemv.CodigoPos).First().Descripcion;
                ComparabancosUpeu item = new ComparabancosUpeu
                {
                    FechaOpe = DateTime.Parse(itemv.FechaAbono),
                    Descripcion = descripcion,
                    NroOpe = "",
                    Importe = itemv.Diferencia,
                    ImporteAbono = itemv.NetoAbonar,
                    ImporteTransac = itemv.Importe,
                    Dh = 1,
                    Terminal = itemv.CodigoPos,
                    Observacion = "Diferencia",
                    Whoyo = itemv.Whoyo,
                    Mb = itemv.ReferenciaVoucher
                };
                ListComiTotAbo.Add(item);
            }
        }

        private void BtnImportar_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                ValidateNames = true,
                InitialDirectory = "C:\\\\"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                diag = openFileDialog;
                Label7.Visible = true;
                Label7.Text = openFileDialog.FileName;
                BtnLeerDatos.Visible = true;
                BtnLeerDatos.Focus();
            }
        }

        private void BtnLeerDatos_Click(object sender, EventArgs e)
        {

            timer1.Start();
            progressBar1.Visible = true;
            backgroundWorker2.RunWorkerAsync();
            BtnProcesar.Focus();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //if (!check)
            //{
            //    if (progressBar1.Value < progressBar1.Maximum)
            //    {
            //        progressBar1.Value++;
            //    }
            //    else
            //    {
            //        check = true;
            //    }
            //}
            //else if (progressBar1.Value > progressBar1.Minimum)
            //{
            //    progressBar1.Value--;
            //}
            //else
            //{
            //    check = false;
            //}
            if (check)
            {
                int newValue = this.progressBar1.Value + 10;
                this.progressBar1.Value = (newValue > 100 ? 0 : newValue);
            }
            else
            {
                this.progressBar1.Value = 100;
            }
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                CargaExcel();
                new Logconci(" CargaExcel ok ", true);
                ExceltoListNewv(dtc);
                new Logconci(" ExceltoListNewv ok ", true);
                ProcesoUnion();
                new Logconci(" ProcesoUnion ok ", true);
            }
            catch (Exception ex)
            {
                //BtnProcesar.Visible = false;
                termino = 1;
                MessageBox.Show(this, ex.Message);
            }
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            BtnProcesar.Visible = true;
            if (termino == 0)
            {
                BtnProcesar.Visible = true;
                listaae.Text = listaMasterAEx.Where((BancoBCP w) => w.NroOpe != string.Empty).ToList().Count.ToString();
                listamc.Text = listaMasterC.Where((BancoBCP w) => w.NroOpe != string.Empty).ToList().Count.ToString();
                listavn.Text = listaCafetinAux.Where((BancoBCP w) => w.NroOpe != string.Empty).ToList().Count.ToString();
                listamcaevnTOT.Text = listaVisaMCAE.Where((BancoBCP w) => w.NroOpe != string.Empty).ToList().Count.ToString();
                lblCantPU.Text = listaPendUpeu.Count.ToString();
                lblCantMB.Text = listaBancos.Count.ToString();
                lblCantUnion3.Text = listaVisaMCAE.Count.ToString();
                lblCanTotal1lado.Text = listasPendienteUpeuYBancos.Count.ToString();
                lblCantMU.Text = listaMayorUpeu.Count.ToString();
                lblCantPB.Text = listaPendBanco.Count.ToString();
                lblTOTALL2.Text = listasPendienteBancoYmayorUpeu.Count.ToString();
            }
            timer1.Stop();
            progressBar1.Visible = false;
        }

        private void BtnProcesar_Click(object sender, EventArgs e)
        {
            //timer1.Interval = (1000) * (1);
            //timer1.Enabled = true;
            //timer1.Start();
            progressBar1.Visible = true;
            backgroundWorker1.RunWorkerAsync();
        }

      
        private void clearTxtLbl()
        {
            timer1.Stop();
            progressBar1.Visible = false;
            BtnLeerDatos.Visible = false;
            lblCantPU.Text = string.Empty;
            lblCantMB.Text = string.Empty;
            lblCantUnion3.Text = string.Empty;
            lblCanTotal1lado.Text = string.Empty;
            lblCantMU.Text = string.Empty;
            lblCantPB.Text = string.Empty;
            lblTOTALL2.Text = string.Empty;
            BtnProcesar.Visible = false;
            listaae.Text = string.Empty;
            listamc.Text = string.Empty;
            listavn.Text = string.Empty;
            listamcaevnTOT.Text = string.Empty;
            Label7.Text = string.Empty;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            diag.Dispose();

            ClearLists();
            clearTxtLbl();


        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

       
    }
}
