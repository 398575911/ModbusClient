using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
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
        //bool ProdutoUnico = false;

        // Ao iniciar o software carrega-se algumas configurações.
        public frmMainPage()
        {
            InitializeComponent();

            // Leitura do arquivo de configuração ao iniciar o software.
            try
            {
                var Config = File.ReadAllLines(@"C:\Users\Rodrigo\OneDrive - HEX Automation Corp\BR180014 - INGERSOLL -  Modernização Seleção de Óleo\Ingersoll\HEX_Automation_ModbusTCP_Client\Config.txt").Select(l => l.Split(new[] { '=' })).ToDictionary(str => str[0].Trim(), str => str[1].Trim());
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
                //MessageBox.Show(" Não foi possível carregar o arquivo de configuração. ", "Erro!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                //MessageBox.Show(" Conexão realizada com sucesso. ", "Conexão", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LabelStrip001.Text = "Conexão realizada com sucesso";
                tabControl1.SelectedIndex = (tabControl1.SelectedIndex + 1) % tabControl1.TabCount;
            }
            catch
            {
                txtConexao.BackColor = Color.Red;
                txtConexao.Text = "  Desconectado";
                //MessageBox.Show(" Não foi possível se conectar. ", "Erro!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                modbusTCP.Disconnect();
                //if (!string.IsNullOrEmpty(MyExcel.DB_PATH))MyExcel.CloseExcel(); //CHECAR ESSA LINHA DE CÓDIGO.
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
                    bool WriteMult1, WriteMult2, WriteMult3, WriteMult4, WriteMult5, WriteMult6, WriteMult7, WriteMult8, WriteMult9, WriteMult10;
                    if (chkMult1.Checked) { WriteMult1 = true; } else { WriteMult1 = false; }
                    if (chkMult2.Checked) { WriteMult2 = true; } else { WriteMult2 = false; }
                    if (chkMult3.Checked) { WriteMult3 = true; } else { WriteMult3 = false; }
                    if (chkMult4.Checked) { WriteMult4 = true; } else { WriteMult4 = false; }
                    if (chkMult5.Checked) { WriteMult5 = true; } else { WriteMult5 = false; }
                    if (chkMult6.Checked) { WriteMult6 = true; } else { WriteMult6 = false; }
                    if (chkMult7.Checked) { WriteMult7 = true; } else { WriteMult7 = false; }
                    if (chkMult8.Checked) { WriteMult8 = true; } else { WriteMult8 = false; }
                    if (chkMult9.Checked) { WriteMult9 = true; } else { WriteMult9 = false; }
                    if (chkMult10.Checked) { WriteMult10 = true; } else { WriteMult10 = false; }
                    bool[] QuantMult = new bool[10];
                    QuantMult[0] = WriteMult1; QuantMult[1] = WriteMult2; QuantMult[2] = WriteMult3; QuantMult[3] = WriteMult4; QuantMult[4] = WriteMult5;
                    QuantMult[5] = WriteMult6; QuantMult[6] = WriteMult7; QuantMult[7] = WriteMult8; QuantMult[8] = WriteMult9; QuantMult[9] = WriteMult10;
                    int j = Int32.Parse(txtQuantityMultCoil.Text); //j é a quantidade de addresses
                    for (int i = 0; i < j; i++)
                    {
                        modbusTCP.WriteSingleCoil(Int32.Parse(txtStartingAddressMultCoil.Text) + i, QuantMult[i]);
                    }
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
                }
                else
                {
                    //MessageBox.Show(" O sistema não está conectado ao PLC. ", "Erro!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LabelStrip001.Text = " O sistema não está conectado ao PLC. ";
                }
            }
            catch
            {
                //MessageBox.Show("Erro ao realizar a leitura.");
            }
        }
        private void fill_HMI(bool[] rd_Coil)
        {
            EST_P00000.Text = rd_Coil[0].ToString();
            EST_P00001.Text = rd_Coil[1].ToString();
            EST_P00002.Text = rd_Coil[2].ToString();
            EST_P00003.Text = rd_Coil[3].ToString();
            EST_P00004.Text = rd_Coil[4].ToString();
            EST_P00005.Text = rd_Coil[5].ToString();
            EST_P00006.Text = rd_Coil[6].ToString();
            EST_P00007.Text = rd_Coil[7].ToString();
            EST_P00008.Text = rd_Coil[8].ToString();
            EST_P00009.Text = rd_Coil[9].ToString();
            EST_P00041.Text = rd_Coil[10].ToString();
            EST_P00040.Text = rd_Coil[11].ToString();
            EST_P00042.Text = rd_Coil[12].ToString();
            EST_P00043.Text = rd_Coil[13].ToString();
            EST_P00044.Text = rd_Coil[14].ToString();
            EST_P00045.Text = rd_Coil[15].ToString();
            EST_P00046.Text = rd_Coil[16].ToString();
            EST_P00047.Text = rd_Coil[17].ToString();
            EST_P00048.Text = rd_Coil[18].ToString();
            EST_P00049.Text = rd_Coil[19].ToString();
        }

        // Botão para entrar em modo Automático.
        private void btnAuto_Click(object sender, EventArgs e)
        {
            if (modbusTCP.Connected)
            {
                timerReadCoil.Enabled = true;
                txtModo.BackColor = Color.LightGreen;
                txtModo.Text = "  Automático";
            }
            else
            {
                timerReadCoil.Enabled = false;
                txtModo.BackColor = Color.Blue;
                txtModo.Text = "  Manual";
            }
        }

        // Em modo automático, a cada "tick" do timer é chamada a função de leitura das bobinas.
        private void timerReadCoil_Tick(object sender, EventArgs e)
        {
            ReadCoils();
        }

        // Botão para entrar em modo Manual.
        // NÃO IMPLEMENTADO!!
        private void btnManual_Click(object sender, EventArgs e)
        {
            timerReadCoil.Enabled = false;
            txtModo.BackColor = Color.Blue;
            txtModo.Text = " Manual";
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
                dataGridEmpList.DataSource = MyExcel.FilterEmpList(cmbSearch.Text.ToString(), txtSearchExpr.Text.ToLower());
                try
                {
                    int Rows = dataGridEmpList.Rows.Count;
                    if(Rows == 1)
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
                dataGridEmpList.DataSource = MyExcel.EmpList;
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
                    //MessageBox.Show(" Erro ao desconectar. ", "Erro!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LabelStrip001.Text = " Erro ao desconectar. ";
                }
            }
            else
            {
                //MessageBox.Show(" PLC não está conectado. ", "Erro!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LabelStrip001.Text = " PLC não está conectado. ";
            }
        }

        // O código de produto da página de IHM é reproduzido para a página dos DADOS a cada mudança na caixa de texto.
        private void txtCodigoProd_TextChanged(object sender, EventArgs e)
        {
            txtSearchExpr.Text = txtCodigoProd.Text;
        }

        // Botão para carrgar a planilha a partir do caminho na caixa de texto.
        private void btnAbreDados_Click(object sender, EventArgs e)
        {
            try
            {
                //MyExcel.DB_PATH = ExcelDialog.FileName;
                MyExcel.DB_PATH = txtFileName.Text;
                MyExcel.InitializeExcel();
                dataGridEmpList.DataSource = MyExcel.ReadMyExcel();              
                LabelStrip001.Text = " Dados carregados com sucesso. ";
                tabControl1.SelectedIndex = (tabControl1.SelectedIndex + 1) % tabControl1.TabCount;
                MyExcel.CloseExcel();
            }
            catch
            {
                //MessageBox.Show(" Ocorreu um erro ao carregar os dados. ", "Erro!!", MessageBoxButtons.OK, MessageBoxIcon.Error); ;
                LabelStrip001.Text = " Ocorreu um erro ao carregar os dados. ";
            }
        }

    }
}
