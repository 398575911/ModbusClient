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
using EasyModbus;
using ModbusTCP_Client;

namespace ModbusTCP_Client
{
    public partial class frmMainPage : Form
    {
        ModbusClient modbusTCP = new ModbusClient();
        // Ao iniciar o software carrega-se algumas configurações.
        public frmMainPage()
        {
            InitializeComponent();

            // Leitura do arquivo de configuração ao iniciar o software.
            try
            {
                var Config = File.ReadAllLines(@"C:\Users\lubck\Documents\BR180014 - INGERSOLL -  Modernização Seleção de Óleo\Ingersoll\HEX_Automation_ModbusTCP_Client\Config.txt").Select(l => l.Split(new[] { '=' })).ToDictionary(str => str[0].Trim(), str => str[1].Trim());
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
                LabelStrip001.Text =" Não foi possível carregar o arquivo de configuração. ";
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

        // Botão para se conectar ao PLC via modbus.
        private void btnCon_Click(object sender, EventArgs e)
        {
                Conectar();
                if (modbusTCP.Connected)
                {
                    timerReadCoil.Enabled = true;
                }
                else
                {
                    timerReadCoil.Enabled = false;
                }
        }

        // Função para conexão com PLC via modbus.
        public void Conectar()
        {
            try
            {
                modbusTCP.Connect(txtIpAddress.Text, Int32.Parse(txtPort.Text));
                grpBoxCoils.Enabled = true;
                grpBoxRegisters.Enabled = true;
                btnCon.Enabled = false;
                btnDisconnect.Enabled = true;
                groupIHM.Enabled = true;
                txtConexao.BackColor = Color.LightGreen;
                txtConexao.Text = "  Conectado";
                LabelStrip001.Text = "Conexão realizada com sucesso";
                tabControl1.SelectedIndex = (tabControl1.SelectedIndex + 1) % tabControl1.TabCount;
            }
            catch
            {
                txtConexao.BackColor = Color.Red;
                txtConexao.Text = "  Desconectado";
                LabelStrip001.Text = " Não foi possível se conectar. ";
            }
        }
        // Botão para leitura das bobinas via modbus.
        private void btnReadCoil_Click(object sender, EventArgs e)
        {
            ReadCoils();
        }
        // Botão para se desconectar ao PLC.
        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            try
            {
                Disconnect();
            }
            catch
            {
                ;
            }
        }
        // Ao fechar o programa deve-se garantir que a conexão modbus seja encerrada, e o arquivo .xlsx também.
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                Disconnect();
            }
            catch
            {
                ;
            }    
        }
        // Botão para escrita de uma bobina via modbus.
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
        // Botão para escrita de múltiplas bobinas via modbus.
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
        // Função para leitura das bobinas via modbus.
        public void ReadCoils()
        {
            try
            {
                if (modbusTCP.Connected)
                {
                    bool[] readCoil;
                    txtRead.Text = "";
                    int StartAddress = Int32.Parse(txtReadAddress.Text);
                    int Quantity = Int32.Parse(txtQuantity.Text);
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

        // Quando conectado, a cada "tick" do timer é chamada a função de leitura das bobinas.
        private void timerReadCoil_Tick(object sender, EventArgs e)
        {
            if (modbusTCP.Connected)
            {
                ReadCoils();               
            }
            else
            {
                LabelStrip001.Text = " O sistema não está conectado ao PLC. "; ;
            }           
        }
 
        // Botão para entrar em envio do modelo.
        // NÃO IMPLEMENTADO!!
        private void btnEnviarModelo_Click(object sender, EventArgs e)
        {
            ;
        }

        // Botão para abrir caixa de diálogo de seleção para o arquivo .xlsx.
        private void btnLoad_Click(object sender, EventArgs e)
        {
            OpenFileDialog ExcelDialog = new OpenFileDialog();
            ExcelDialog.Filter = "Excel Files (*.xlsx) | *.xlsx";
            ExcelDialog.RestoreDirectory = true;
            ExcelDialog.Title = "Selecione a planilha com os dados.";

            if (ExcelDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFileName.Text = ExcelDialog.FileName;
                txtFileName.ReadOnly = true;
                txtFileName.Click -= btnLoad_Click;
            }
        }
        // Função para filtragem do arquivo .xlsx.
        private void txtSearchExpr_TextChanged(object sender, EventArgs e)
        {
            // Se existe algum código lido é realizada a filtragem do datagrid com o excel carregado.
            if (!string.IsNullOrEmpty(txtSearchExpr.Text))
            {
                //dataGridEmpList.DataSource = MyExcel.FilterEmpList(cmbSearch.Text.ToString(), txtSearchExpr.Text.ToLower());
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
                            if (row.Cells[0].Value.ToString().Equals(SearchValue))
                            {
                                row.Selected = true;
                                found = true;
                                break;
                            }
                        }
                    }
                    catch
                        {
                        found = false; ;
                        }
                    if (found)
                    {
                        // Se o código lido for único, o Código e o Tipo do óleo são lançados para a página da IHM.
                        txtCodigoOleo.Text = dataGridEmpList.Rows[dataGridEmpList.Rows.Count - 1].Cells[1].Value.ToString();
                        txtTipoOleo.Text = dataGridEmpList.Rows[dataGridEmpList.Rows.Count - 1].Cells[9].Value.ToString();
                        //ProdutoUnico = true;
                    }
                    else
                    {
                        // Se o código lido for NÃO único os campos ficam em branco.
                        txtCodigoOleo.Text = "";
                        txtTipoOleo.Text = "";
                        //ProdutoUnico = false;
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
            }
        }
        // Botão para login no sistema.
        private void btnLogin_Click(object sender, EventArgs e)
        {
            if (txtUser.Text == "admin" && txtPassword.Text == "admin")
            {
                grpBoxConexao.Enabled = true;
                groupExcel.Enabled = true;
                btnLogin.Enabled = false;
                btnLogout.Enabled = true;
                btnCon.Enabled = true;
                txtIpAddress.Enabled = true;
                txtPort.Enabled = true;
                LabelStrip001.Text = " Login realizado com sucesso ";
            }
            else
            {              
                LabelStrip001.Text = " Usuário e/ou senha inválidos ";
            }
        }

        // Botão para logout do sistema.
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
            }
            catch
            {
                LabelStrip001.Text = "Não foi possível fazer o logout. ";
            }
        }

        // Função para fechar a conexão Modbus com o PLC.
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
        private void txtCodigoProd_TextChanged(object sender, EventArgs e)
        {
            txtSearchExpr.Text = txtCodigoProd.Text;
        }
        // Botão para carrgar a planilha a partir do caminho na caixa de texto.
        private void btnAbreDados_Click(object sender, EventArgs e)
        {
            try
            {
                var connectionStringFormat = ConfigurationManager.AppSettings["Microsoft.ACE.OLEDB"].ToString();
 
                var excelNamePath = txtFileName.Text.ToString();
                FileInfo finfo = new FileInfo(excelNamePath);
                string excelFileName = finfo.Name.ToString();
                var dataSet = new DataSet(excelFileName);
                //Setup Connection string based on which excel file format we are using
                var excelType = "Excel 8.0";
                if (excelFileName.Contains(".xlsx")){excelType = "Excel 12.0 XML";}
                var connectionString = string.Format(connectionStringFormat, excelNamePath, excelType);
                //Create a connection to the excel file
                using (var oleDbConnection = new OleDbConnection(connectionString))
                {
                    oleDbConnection.Open();
                    var schemaDataTable = (DataTable)oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    oleDbConnection.Close();
                    var sheetsName = GetSheetsName(schemaDataTable);
                    //schemaDataTable.Columns[0].DataType = typeof(String);

                    //For each sheet name
                    OleDbCommand selectCommand = null;
                    for (var i = 0; i< sheetsName.Count;i++)
                    {
                        //Setup select command
                        selectCommand = new OleDbCommand();
                        selectCommand.CommandText = "SELECT * FROM [" + sheetsName[i] + "]";
                        selectCommand.Connection = oleDbConnection;

                        oleDbConnection.Open();
                        using(var oleDbDataReader = selectCommand.ExecuteReader(CommandBehavior.CloseConnection))
                        {
                            //Convert data to DataTable
                            var dataTable =new DataTable(sheetsName[i].Replace("$", "").Replace("'", ""));
                            dataTable.Load(oleDbDataReader);

                            //Add to Dataset
                            dataSet.Tables.Add(dataTable); //dataTable
                            
                           
                            dataGridEmpList.DataSource = dataSet.Tables[0];
                            //dataGridEmpList.AutoGenerateColumns = true;
                        }
                    }

                }
            }
            catch (Exception expn)
            {
                MessageBox.Show(expn.ToString());  //Remove "//" to get the exception when it occour
                LabelStrip001.Text = " Ocorreu um erro ao carregar os dados. ";
            }
        }

        private void FrmMainPage_Load(object sender, EventArgs e)
        {

        }
        private List<string> GetSheetsName(DataTable schemaDataTable)
        {
            var sheets = new List<string>();
            foreach(var dataRow in schemaDataTable.AsEnumerable())
            {
                sheets.Add(dataRow.ItemArray[2].ToString());
            }
            return sheets;
        }
    }
}
