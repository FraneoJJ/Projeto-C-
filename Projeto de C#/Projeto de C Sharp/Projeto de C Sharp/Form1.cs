using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;// Para DataTable
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;// Cria um alias para Excel
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace Projeto_de_C_Sharp
{
    public partial class Form1 : Form
    {
        string caminho = "server=localhost; user=root; password=; database=excel;";
        //SslMode=None;usado para desativar a criptografia SSL/TLS na comunicação entre o aplicativo e o servidor de banco de dados.
        System.Data.DataTable tabela = new System.Data.DataTable();
        public Form1()
        {
            InitializeComponent();
        }
        private void arqivoExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog abrirDOC = new OpenFileDialog();

            try
            {
                abrirDOC.Filter = "Excel Files|*.xlsx";
                if (abrirDOC.ShowDialog() == DialogResult.OK)
                {
                    string arquivo = abrirDOC.FileName;
                    enviar(arquivo);
                }

            }
            catch (Exception erro)
            {
                MessageBox.Show($"Falaha ao carregar o arquivo; Erro: {erro}");
            }
        }
        private void enviar(string localAqui)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(localAqui);
                Excel.Worksheet worksheet = workbook.Sheets[1];
                Excel.Range range = worksheet.UsedRange;

                // Usado para DataTable


                // Criar colunas no DataTable
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    string colName = range.Cells[1, col].Value2?.ToString() ?? $"Coluna {col}";
                    tabela.Columns.Add(colName);
                }

                // Preencher as linhas do DataTable
                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    DataRow dataRow = tabela.NewRow();
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        dataRow[col - 1] = range.Cells[row, col].Value2?.ToString() ?? string.Empty;
                    }
                    tabela.Rows.Add(dataRow);
                }

                // Liberar recursos
                workbook.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                // Exibir no DataGridView
                dataGridView1.DataSource = tabela;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar o Excel: {ex.Message}");
            }
        }

        private void salvarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string nomeTabela = nome_table.Text;

            if(nomeTabela == ""){
                MessageBox.Show("Informe o nome da panilha");
                nome_table.Focus();
            }
            else
            {
                using (MySqlConnection conn = new MySqlConnection(caminho))
                {
                    try
                    {
                        conn.Open();

                        // 1. Cria a tabela dinamicamente com base nas colunas do DataGridView
                        var colunasParaTabela = new List<string>();
                        foreach (DataGridViewColumn coluna in dataGridView1.Columns)
                        {
                            if (coluna.Name != "id")
                            {
                                colunasParaTabela.Add($"{coluna.Name} VARCHAR(255)");
                            }
                        }

                        string createTableQuery = $"CREATE TABLE {nomeTabela} (id INT AUTO_INCREMENT PRIMARY KEY, {string.Join(", ", colunasParaTabela)})";
                        MySqlCommand creandoTBL = new MySqlCommand(createTableQuery, conn);
                        creandoTBL.ExecuteNonQuery();

                        // 2. Insere os dados do DataGridView na tabela
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                var colunas = new List<string>();
                                var parametros = new List<string>();
                                var valores = new Dictionary<string, object>();

                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    if (cell.Value != null && !string.IsNullOrEmpty(cell.OwningColumn.Name))
                                    {
                                        string coluna = cell.OwningColumn.Name;
                                        colunas.Add(coluna);
                                        parametros.Add($"@{coluna}");
                                        valores.Add($"@{coluna}", cell.Value);
                                    }
                                }

                                string insertQuery = $"INSERT INTO {nomeTabela} ({string.Join(", ", colunas)}) VALUES ({string.Join(", ", parametros)})";
                                MySqlCommand cmd = new MySqlCommand(insertQuery, conn);

                                foreach (var valor in valores)
                                {
                                    cmd.Parameters.AddWithValue(valor.Key, valor.Value);
                                }

                                cmd.ExecuteNonQuery();
                            }
                        }

                        MessageBox.Show("Dados enviados com sucesso!");
                    }
                    catch (MySqlException ex)
                    {
                        MessageBox.Show($"Erro MySQL: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro geral: {ex.Message}");
                    }
                }
            }

            
        }

        private void osDadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabela.Rows.Clear();
            nome_table.Text = "";
        }

        private void tudoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabela.Columns.Clear();
            tabela.Rows.Clear();
            nome_table.Text = "";
        }

        private void bancoDeDadosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                using (MySqlConnection conn = new MySqlConnection(caminho))
                {
                    conn.Open();
                    MySqlCommand ver_tabelas = new MySqlCommand("show tables", conn);
                    using (MySqlDataReader reader = ver_tabelas.ExecuteReader())
                    {
                        string tabelas = "Tabelas no banco de dados:\n";
                        while (reader.Read())
                        {
                            tabelas += reader.GetString(0) + "\n";
                        }
                        MessageBox.Show(tabelas);
                    }
                }
            }
            catch (Exception erro)
            {
                MessageBox.Show($"Erro: {erro.Message}");
            }
        }

        private void tabelaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string bancoExistente = nome_table.Text;
            if (bancoExistente == "")
            {
                MessageBox.Show("Informe o nome da tabela que deseja ver");
                nome_table.Focus();
            }
            else
            {
                try
                {
                    MySqlConnection conn = new MySqlConnection(caminho);
                    MySqlCommand mostrar = new MySqlCommand($"select * from {bancoExistente}", conn);
                    MySqlDataAdapter adaptar = new MySqlDataAdapter(mostrar);
                    adaptar.Fill(tabela);
                    dataGridView1.DataSource = tabela;
                }
                catch (Exception)
                {
                    MessageBox.Show($"Essa tabela não existe no banco de dados\nEscreva corretamente o nome da tabela");
                    nome_table.Focus();
                }
            }
        }

        private void salvarEdiçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string nomeTabela = nome_table.Text;

            if (string.IsNullOrEmpty(nomeTabela))
            {
                MessageBox.Show("Informe o nome da tabela");
                nome_table.Focus();
                return;
            }

            using (MySqlConnection conn = new MySqlConnection(caminho))
            {
                try
                {
                    conn.Open();

                    // Obter a estrutura da tabela do banco de dados
                    System.Data.DataTable schemaTable;
                    using (MySqlCommand schemaCmd = new MySqlCommand($"SELECT * FROM {nomeTabela} WHERE 1=0", conn))
                    using (MySqlDataAdapter schemaAdapter = new MySqlDataAdapter(schemaCmd))
                    {
                        schemaTable = new System.Data.DataTable();
                        schemaAdapter.Fill(schemaTable);
                    }

                    // Criar um DataAdapter com comandos de atualização
                    MySqlDataAdapter adapter = new MySqlDataAdapter($"SELECT * FROM {nomeTabela}", conn);
                    MySqlCommandBuilder commandBuilder = new MySqlCommandBuilder(adapter);

                    // Atualizar o banco de dados com as alterações do DataTable
                    System.Data.DataTable changedTable = tabela.GetChanges();
                    if (changedTable != null)
                    {
                        int rowsAffected = adapter.Update(changedTable);
                        MessageBox.Show($"{rowsAffected} registro(s) atualizado(s) com sucesso!");
                        tabela.AcceptChanges(); // Marcar as alterações como confirmadas
                    }
                    else
                    {
                        MessageBox.Show("Nenhuma alteração para salvar.");
                    }
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show($"Erro MySQL: {ex.Message}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro geral: {ex.Message}");
                }
            }
        }
    }
}
