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
        System.Data.DataTable tabela = new System.Data.DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_import_Click(object sender, EventArgs e)
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
            catch(Exception erro)
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
            MessageBox.Show($"Verifique se tem o pacote office instalado no seu computador; Erro ao carregar o Excel: {ex.Message}");
        }
    }

        private void InserirDadosNoBanco()
        {
            try
            {
                // String de conexão com o banco de dados SQL Server
                string conexaoString = @"Data Source=SEU_SERVIDOR;Initial Catalog=SEU_BANCO;Integrated Security=True";

                using (SqlConnection conexao = new SqlConnection(conexaoString))
                {
                    conexao.Open();

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)  // Ignora a linha de edição vazia
                        {
                            string comandoSQL = "INSERT INTO TabelaExemplo (Coluna1, Coluna2, Coluna3) VALUES (@valor1, @valor2, @valor3)";

                            using (SqlCommand comando = new SqlCommand(comandoSQL, conexao))
                            {
                                // Adicionando parâmetros para evitar SQL Injection
                                comando.Parameters.AddWithValue("@valor1", row.Cells[0].Value ?? DBNull.Value);
                                comando.Parameters.AddWithValue("@valor2", row.Cells[1].Value ?? DBNull.Value);
                                comando.Parameters.AddWithValue("@valor3", row.Cells[2].Value ?? DBNull.Value);

                                comando.ExecuteNonQuery();
                            }
                        }
                    }

                    MessageBox.Show("Dados inseridos com sucesso!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao inserir dados no banco: {ex.Message}");
            }
        }

        private void btnSalvar_Click(object sender, EventArgs e)
        {
            MySqlConnection conexao = new MySqlConnection(caminho);
            conexao.Open();
            MySqlCommand comando = new MySqlCommand("alter table aluno add column @nome_coluna", conexao);
            comando.Parameters.AddWithValue("@nome_coluna", dataGridView1.Rows);
            conexao.Close();
        }
    }
}
