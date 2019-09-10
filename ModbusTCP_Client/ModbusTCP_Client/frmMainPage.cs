using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Configuration;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using EasyModbus;
using ModbusTCP_Client;

namespace ModbusTCP_Client
{
    public partial class frmMainPage : Form
    {

        //Global ----
        DataSet ds;
        String PLCAddr;
        int PLCPort;
        bool EnableCyclicRead = true;
        // ------------------

        ModbusClient modbusTCP = new ModbusClient();
        public frmMainPage()
        {
            InitializeComponent();
            LabelStrip001.Text = "Aplicação Inicializada";
            // Leitura do arquivo de configuração ao iniciar o software.
           
        }
        private void btnCon_Click(object sender, EventArgs e)
        {
            PLCAddr = txtIpAddress.Text;
            PLCPort = Int32.Parse(txtPort.Text);
            Conectar();
        }
        private void btnReadCoil_Click(object sender, EventArgs e)
        {
            ReadCoils();
        }
        private void btnDisconnect_Click(object sender, EventArgs e)
        {
                Disconnect();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
                Disconnect();
        }
        private void btnWriteSingle_Click(object sender, EventArgs e)
        {
            bool WriteValue;
            if(modbusTCP.Connected == true)
            {
                try
                {
                    if (checkTrueWrite.Checked) {WriteValue = true;} else {WriteValue = false;}

                    int WriteSingleAddress = Int32.Parse(txtWriteSingleAddress.Text);
                    modbusTCP.WriteSingleCoil(WriteSingleAddress, WriteValue);
                }
                catch
                {
                    ;
                }
            }
            else
            {
                ;
            }
        }
        private void btnWriteMultCoils_Click(object sender, EventArgs e)
        {
            if (modbusTCP.Connected == true)
            {
                try
                {
                    bool[] QuantMult = new bool[10];
                    if (chkMult1.Checked)  { QuantMult[0] = true; } else { QuantMult[0] = false; }
                    if (chkMult2.Checked)  { QuantMult[1] = true; } else { QuantMult[1] = false; }
                    if (chkMult3.Checked)  { QuantMult[2] = true; } else { QuantMult[2] = false; }
                    if (chkMult4.Checked)  { QuantMult[3] = true; } else { QuantMult[3] = false; }
                    if (chkMult5.Checked)  { QuantMult[4] = true; } else { QuantMult[4] = false; }
                    if (chkMult6.Checked)  { QuantMult[5] = true; } else { QuantMult[5] = false; }
                    if (chkMult7.Checked)  { QuantMult[6] = true; } else { QuantMult[6] = false; }
                    if (chkMult8.Checked)  { QuantMult[7] = true; } else { QuantMult[7] = false; }
                    if (chkMult9.Checked)  { QuantMult[8] = true; } else { QuantMult[8] = false; }
                    if (chkMult10.Checked) { QuantMult[9] = true; } else { QuantMult[9] = false; }
                    int j = Int32.Parse(txtQuantityMultCoil.Text); //j é a quantidade de addresses
                    for (int i = 0; i < j; i++) {modbusTCP.WriteSingleCoil(Int32.Parse(txtStartingAddressMultCoil.Text) + i, QuantMult[i]);}
                    LabelStrip001.Text = "Escrita sem erros";
                }
                catch
                {
                    LabelStrip001.Text = "Escrita não foi possível";
                }
            }
            else
            {
                LabelStrip001.Text = " O sistema não está conectado ao PLC. ";
            }         
        }
        private void timerReadCoil_Tick(object sender, EventArgs e)
        {
 
            if (modbusTCP.Connected && EnableCyclicRead)
            {
                ReadCoils();               
            }
            else if (modbusTCP.Connected && !EnableCyclicRead)
            {
                LabelStrip001.Text = " O sistema está conectado ao PLC mas leitura cíclica não habilitada. ";
            }
            else
            {
                LabelStrip001.Text = " O sistema não está conectado ao PLC. ";
            }           
        }
        private void btnEnviarModelo_Click(object sender, EventArgs e)
        {
            if(txtTipoOleo.Text == "Ultracoolant")
                {
                modbusTCP.WriteSingleCoil(48, true);
                modbusTCP.WriteSingleCoil(49, false);
                txtCodigoProd.Text = "";
                txtCodigoOleo.Text = "";
                txtTipoOleo.Text = "";
            }
            else if(txtTipoOleo.Text == "UltraEL")
            {
                modbusTCP.WriteSingleCoil(48, false);
                modbusTCP.WriteSingleCoil(49, true);
                txtCodigoProd.Text = "";
                txtCodigoOleo.Text = "";
                txtTipoOleo.Text = "";
            }
            else
            {
                modbusTCP.WriteSingleCoil(48, false);
                modbusTCP.WriteSingleCoil(49, false);
                LabelStrip001.Text = "Nenhum modelo válido de óleo  selectionado";
                txtCodigoProd.Text = "";
                txtCodigoOleo.Text = "";
                txtTipoOleo.Text = "";
            }
        }
        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog ExcelDialog = new OpenFileDialog();
            ExcelDialog.Filter = "Excel Files (*.xls) | *.xls";
            ExcelDialog.RestoreDirectory = true;
            ExcelDialog.Title = "Selecione a planilha com os dados.";

            if (ExcelDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFileName.Text = ExcelDialog.FileName;
                txtFileName.ReadOnly = true;
                txtFileName.Click -= btnLoad_Click;
            }
        }
        private void txtSearchExpr_TextChanged(object sender, EventArgs e)
        {
            // Se existe algum código lido é realizada a filtragem do datagrid com o excel carregado.
            if (!string.IsNullOrEmpty(txtSearchExpr.Text))
            {
                try
                {
                    string SearchValue = txtSearchExpr.Text.ToString();
                    bool found = false;
                    dataGridEmpList.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    try
                    {
                        dataGridEmpList.ClearSelection();
                        foreach(DataGridViewRow row in dataGridEmpList.Rows)
                        {                            
                            if (row.Cells[0].Value.ToString().Contains(SearchValue.Replace(" ","")))
                            {
                                row.Selected = true;
                                row.Cells[0].Selected = true;
                                found = true;
                                txtCodigoOleo.Text = row.Cells[1].Value.ToString();
                                txtTipoOleo.Text = row.Cells[9].Value.ToString();
                                
                                break;
                            }
                        }
                    }
                    catch
                        {
                        found = false; 
                        }
                    if (!found)
                    {
                        txtCodigoOleo.Text = "";
                        txtTipoOleo.Text = "";
                    }
                }
                catch
                {
                    ;
                }
            }
            else
            {
                dataGridEmpList.ClearSelection();
                txtCodigoOleo.Text = "";
                txtTipoOleo.Text = "";
            }
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txtUser.Text == "admin" && txtPassword.Text == "WNG1156")
            {
                grpBoxConexao.Enabled = true;
                groupExcel.Enabled = true;
                btnLogin.Enabled = false;
                btnLogout.Enabled = true;
                btnCon.Enabled = true;
                txtIpAddress.Enabled = true;
                txtPort.Enabled = true;
                LabelStrip001.Text = " Login realizado com sucesso ";
                btnWriteMultCoils.Visible = true;
                btnWriteSingle.Visible = true;
                btnCyclicRead.Visible = true;
            }
            else if (txtUser.Text == "prod" && txtPassword.Text == "prod")
            {
                grpBoxConexao.Enabled = true;
                groupExcel.Enabled = true;
                btnLogin.Enabled = false;
                btnLogout.Enabled = true;
                btnCon.Enabled = true;
                txtIpAddress.Enabled = true;
                txtPort.Enabled = true;
                LabelStrip001.Text = " Login realizado com sucesso ";
                btnWriteMultCoils.Visible = false;
                btnWriteSingle.Visible = false;
                btnCyclicRead.Visible = false;
            }
            else
            {
                LabelStrip001.Text = " Usuário e/ou senha inválidos ";
                btnWriteMultCoils.Visible = false;
                btnWriteSingle.Visible = false;
                btnCyclicRead.Visible = false;
            }
            funcCarregaCfg();
        }
        private void btnLogout_Click(object sender, EventArgs e)
        {
            try
            {
                if (modbusTCP.Connected == true) { Disconnect(); } else {; }
                groupExcel.Enabled = false;
                btnLogin.Enabled = true;
                btnLogout.Enabled = false;
                btnCon.Enabled = false;
                txtIpAddress.Enabled = false;
                txtPort.Enabled = false;
                LabelStrip001.Text = " Logout realizado com sucesso.";
                btnWriteMultCoils.Visible = false;
                btnWriteSingle.Visible = false;
            }
            catch
            {
                LabelStrip001.Text = "Não foi possível fazer o logout. ";
            }
        }
   
        private void btnAbreDados_Click(object sender, EventArgs e)
        {
            try
            {
                DataInfo.Text = "Carregando Dados";
                Thread t = new Thread(() => funcAbreDados(txtFileName.Text.ToString()));
                t.Start();
                t.Join();
                DataInfo.Text = "Dados Carregados";
            }
            catch (Exception expn)
            {
                MessageBox.Show(expn.ToString());
                LabelStrip001.Text = " Ocorreu um erro ao carregar os dados. ";
            }
            dataGridEmpList.DataSource = ds.Tables[0];
        }

        private void funcCarregaCfg()
        {
            try
            {
                var Config = File.ReadAllLines(@"C:\Program Files (x86)\Weingartner Sistemas de Automação\Oil Selection\Config.txt").Select(l => l.Split(new[] { '=' })).ToDictionary(str => str[0].Trim(), str => str[1].Trim());
                string IPAddress = Config["IPAdress"];
                string Port = Config["Port"];
                string Path = Config["Data"];
                txtIpAddress.Text = IPAddress;
                txtPort.Text = Port;
                txtFileName.Text = Path;
                LabelStrip001.Text = " Arquivo de configuração carregado. ";
            }
            catch
            {
                LabelStrip001.Text = " Não foi possível carregar o arquivo de configuração. ";
                OpenFileDialog CfgDialog = new OpenFileDialog();
                CfgDialog.Filter = "Text|*.txt|All|*.*";
                CfgDialog.RestoreDirectory = true;
                CfgDialog.Title = "Selecione o arquivo de configuração";
                if (CfgDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var Config = File.ReadAllLines(CfgDialog.FileName).Select(l => l.Split(new[] { '=' })).ToDictionary(str => str[0].Trim(), str => str[1].Trim());
                    string IPAddress = Config["IPAdress"];
                    string Port = Config["Port"];
                    string Path = Config["Data"];
                    txtIpAddress.Text = IPAddress;
                    txtPort.Text = Port;
                    txtFileName.Text = Path;
                    LabelStrip001.Text = " Arquivo de configuração carregado. ";
                }
            }

            // Modo manual ao iniciar o software.
            if (timerReadCoil.Enabled == false)
            {
                txtModo.BackColor = Color.Blue;
                txtModo.Text = "  Manual";
            }
            else
            {
                txtModo.BackColor = Color.LightGreen;
                txtModo.Text = "  Automático";
            }
        }
        public void funcAbreDados(String fileName)
        {
            var connectionStringFormat = ConfigurationManager.AppSettings["Microsoft.ACE.OLEDB"].ToString();
            var excelNamePath = fileName;
            FileInfo finfo = new FileInfo(excelNamePath);
            string excelFileName = finfo.Name.ToString();
            var dataSet = new DataSet(excelFileName);
            var excelType = "Excel 11.0";
            if (excelFileName.Contains(".xls")) { excelType = "Excel 12.0 XML"; }
            var connectionString = string.Format(connectionStringFormat, excelNamePath, excelType);
            using (var oleDbConnection = new OleDbConnection(connectionString))
            {
                oleDbConnection.Open();
                var schemaDataTable = (DataTable)oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                oleDbConnection.Close();
                var sheetsName = GetSheetsName(schemaDataTable);

                OleDbCommand selectCommand = null;

                selectCommand = new OleDbCommand();
                selectCommand.CommandText = "SELECT * FROM [" + sheetsName[1] + "]";
                selectCommand.Connection = oleDbConnection;

                oleDbConnection.Open();
                using (var oleDbDataReader = selectCommand.ExecuteReader(CommandBehavior.CloseConnection))
                {
                    var dataTable = new DataTable(sheetsName[1].Replace("$", "").Replace("'", ""));
                    dataTable.Load(oleDbDataReader);
                    dataSet.Tables.Add(dataTable); //dataTable 

                }
            }
            ds = dataSet;
        }

        private void txtCodigoProd_TextChanged(object sender, EventArgs e)
        {
            txtSearchExpr.Text = txtCodigoProd.Text;
        }
        private void FrmMainPage_Load(object sender, EventArgs e)
        {

        }

        public void ReadCoils()
        {
            try
            {
                if (modbusTCP.Connected)
                {
                    bool[] readCoil;
                    txtRead.Text = "";
                    int StartAddress = 0;     //Int32.Parse(txtReadAddress.Text);
                    int Quantity = 74;        //Int32.Parse(txtQuantity.Text);
                    readCoil = modbusTCP.ReadCoils(StartAddress, Quantity);
                    for (int i = 0; i < readCoil.Length; i++)
                    {
                        txtRead.Text += readCoil[i].ToString() + "\r\n";
                    }
                    fill_HMI(readCoil);
                    LabelStrip001.Text = "Leitura sem erros";
                }
                else
                {
                    LabelStrip001.Text = " O sistema não está conectado ao PLC. ";
                }
            }
            catch
            {
                LabelStrip001.Text = "Leitura não foi possível POS 05";
            }
        }
        private void fill_HMI(bool[] rd_Coil)
        {
            EST_P00000.Text = rd_Coil[0] ? "TRUE" : "FALSE";
            EST_P00001.Text = rd_Coil[1] ? "TRUE" : "FALSE";
            EST_P00002.Text = rd_Coil[2] ? "TRUE" : "FALSE";
            EST_P00003.Text = rd_Coil[3] ? "TRUE" : "FALSE";
            EST_P00004.Text = rd_Coil[4] ? "TRUE" : "FALSE";
            EST_P00005.Text = rd_Coil[5] ? "TRUE" : "FALSE";
            EST_P00006.Text = rd_Coil[6] ? "TRUE" : "FALSE";
            EST_P00007.Text = rd_Coil[7] ? "TRUE" : "FALSE";
            EST_P00008.Text = rd_Coil[8] ? "TRUE" : "FALSE";
            EST_P00009.Text = rd_Coil[9] ? "TRUE" : "FALSE";

            EST_P00040.Text = rd_Coil[65] ? "TRUE" : "FALSE";
            EST_P00041.Text = rd_Coil[64] ? "TRUE" : "FALSE";
            EST_P00042.Text = rd_Coil[66] ? "TRUE" : "FALSE";
            EST_P00043.Text = rd_Coil[67] ? "TRUE" : "FALSE";
            EST_P00044.Text = rd_Coil[68] ? "TRUE" : "FALSE";
            EST_P00045.Text = rd_Coil[69] ? "TRUE" : "FALSE";
            EST_P00046.Text = rd_Coil[70] ? "TRUE" : "FALSE";
            EST_P00047.Text = rd_Coil[71] ? "TRUE" : "FALSE";
            EST_P00048.Text = rd_Coil[72] ? "TRUE" : "FALSE";
            EST_P00049.Text = rd_Coil[73] ? "TRUE" : "FALSE";
            if (!rd_Coil[4])
            {
                txtModo.BackColor = Color.Blue;
                txtModo.Text = "  Manual";
            }
            else
            {
                txtModo.BackColor = Color.LightGreen;
                txtModo.Text = "  Automático";
            }
        }
        public void Disconnect()
        {
            timerReadCoil.Enabled = false;
            if (modbusTCP.Connected == true)
            {
                try
                {
                    modbusTCP.Disconnect();
                    grpBoxCoils.Enabled = false;
                    grpBoxRegisters.Enabled = false;
                    btnCon.Enabled = true;
                    groupIHM.Enabled = false;
                    btnDisconnect.Enabled = false;

                    txtConexao.BackColor = Color.Red;
                    txtConexao.Text = "  Desconectado";
                    txtConexao.Enabled = true;

                    txtModo.BackColor = Color.Blue;
                    txtModo.Text = "  Manual";
                    txtModo.Enabled = true;
                    LabelStrip001.Text = " Desconectado com sucesso. ";
                }
                catch
                {
                    LabelStrip001.Text = " Erro ao desconectar. ";
                }
            }
            else
            {
                LabelStrip001.Text = " PLC não está conectado. ";
            }
        }
        public void Conectar()
        {
            try
            {
                modbusTCP.IPAddress = PLCAddr;
                modbusTCP.Port = PLCPort;
                modbusTCP.Connect();
                grpBoxCoils.Enabled = true;
                grpBoxRegisters.Enabled = true;
                btnDisconnect.Enabled = true;
                groupIHM.Enabled = true;
                txtConexao.BackColor = Color.LightGreen;
                txtConexao.Text = "  Conectado";
                LabelStrip001.Text = "Conexão realizada com sucesso";
                tabControl1.SelectedIndex = (tabControl1.SelectedIndex + 1) % tabControl1.TabCount;
                timerReadCoil.Enabled = true;
            }
            catch
            {
                txtConexao.BackColor = Color.Red;
                timerReadCoil.Enabled = false;
                txtConexao.Text = "  Desconectado";
            }
               


        }

        private List<string> GetSheetsName(DataTable schemaDataTable)
        {
            var sheets = new List<string>();
            foreach (var dataRow in schemaDataTable.AsEnumerable())
            {
                sheets.Add(dataRow.ItemArray[2].ToString());
            }
            return sheets;
        }

        private void CyclicRead_Click(object sender, EventArgs e)
        {
            if (EnableCyclicRead)
            {
                EnableCyclicRead = false;
            }
            else
                EnableCyclicRead = true;
            }
        }
    }
