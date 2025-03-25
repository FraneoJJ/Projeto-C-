namespace Projeto_de_C_Sharp
{
    partial class Form1
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.abrirToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.arqivoExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.arquivoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.salvarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.limparToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.osDadosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tudoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.verToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bancoDeDadosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabelaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.nome_table = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.salvarEdiçãoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(40, 156);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 28;
            this.dataGridView1.Size = new System.Drawing.Size(744, 323);
            this.dataGridView1.TabIndex = 0;
            // 
            // abrirToolStripMenuItem
            // 
            this.abrirToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.arqivoExcelToolStripMenuItem});
            this.abrirToolStripMenuItem.Name = "abrirToolStripMenuItem";
            this.abrirToolStripMenuItem.Size = new System.Drawing.Size(98, 29);
            this.abrirToolStripMenuItem.Text = "Importar";
            // 
            // arqivoExcelToolStripMenuItem
            // 
            this.arqivoExcelToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("arqivoExcelToolStripMenuItem.Image")));
            this.arqivoExcelToolStripMenuItem.Name = "arqivoExcelToolStripMenuItem";
            this.arqivoExcelToolStripMenuItem.Size = new System.Drawing.Size(210, 34);
            this.arqivoExcelToolStripMenuItem.Text = "Arqivo Excel";
            this.arqivoExcelToolStripMenuItem.Click += new System.EventHandler(this.arqivoExcelToolStripMenuItem_Click);
            // 
            // arquivoToolStripMenuItem
            // 
            this.arquivoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.salvarToolStripMenuItem,
            this.limparToolStripMenuItem,
            this.salvarEdiçãoToolStripMenuItem});
            this.arquivoToolStripMenuItem.Name = "arquivoToolStripMenuItem";
            this.arquivoToolStripMenuItem.Size = new System.Drawing.Size(91, 29);
            this.arquivoToolStripMenuItem.Text = "Arquivo";
            // 
            // salvarToolStripMenuItem
            // 
            this.salvarToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("salvarToolStripMenuItem.Image")));
            this.salvarToolStripMenuItem.Name = "salvarToolStripMenuItem";
            this.salvarToolStripMenuItem.Size = new System.Drawing.Size(322, 34);
            this.salvarToolStripMenuItem.Text = "Salvar no Banco de Dados";
            this.salvarToolStripMenuItem.Click += new System.EventHandler(this.salvarToolStripMenuItem_Click);
            // 
            // limparToolStripMenuItem
            // 
            this.limparToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.osDadosToolStripMenuItem,
            this.tudoToolStripMenuItem});
            this.limparToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("limparToolStripMenuItem.Image")));
            this.limparToolStripMenuItem.Name = "limparToolStripMenuItem";
            this.limparToolStripMenuItem.Size = new System.Drawing.Size(322, 34);
            this.limparToolStripMenuItem.Text = "Limpar";
            // 
            // osDadosToolStripMenuItem
            // 
            this.osDadosToolStripMenuItem.Name = "osDadosToolStripMenuItem";
            this.osDadosToolStripMenuItem.Size = new System.Drawing.Size(270, 34);
            this.osDadosToolStripMenuItem.Text = "Os dados";
            this.osDadosToolStripMenuItem.Click += new System.EventHandler(this.osDadosToolStripMenuItem_Click);
            // 
            // tudoToolStripMenuItem
            // 
            this.tudoToolStripMenuItem.Name = "tudoToolStripMenuItem";
            this.tudoToolStripMenuItem.Size = new System.Drawing.Size(270, 34);
            this.tudoToolStripMenuItem.Text = "Tudo";
            this.tudoToolStripMenuItem.Click += new System.EventHandler(this.tudoToolStripMenuItem_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.menuStrip1.GripMargin = new System.Windows.Forms.Padding(2, 2, 0, 2);
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.abrirToolStripMenuItem,
            this.arquivoToolStripMenuItem,
            this.verToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;
            this.menuStrip1.Size = new System.Drawing.Size(818, 33);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // verToolStripMenuItem
            // 
            this.verToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bancoDeDadosToolStripMenuItem,
            this.tabelaToolStripMenuItem});
            this.verToolStripMenuItem.Name = "verToolStripMenuItem";
            this.verToolStripMenuItem.Size = new System.Drawing.Size(53, 29);
            this.verToolStripMenuItem.Text = "Ver";
            // 
            // bancoDeDadosToolStripMenuItem
            // 
            this.bancoDeDadosToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("bancoDeDadosToolStripMenuItem.Image")));
            this.bancoDeDadosToolStripMenuItem.Name = "bancoDeDadosToolStripMenuItem";
            this.bancoDeDadosToolStripMenuItem.Size = new System.Drawing.Size(270, 34);
            this.bancoDeDadosToolStripMenuItem.Text = "Banco de Dado";
            this.bancoDeDadosToolStripMenuItem.Click += new System.EventHandler(this.bancoDeDadosToolStripMenuItem_Click);
            // 
            // tabelaToolStripMenuItem
            // 
            this.tabelaToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("tabelaToolStripMenuItem.Image")));
            this.tabelaToolStripMenuItem.Name = "tabelaToolStripMenuItem";
            this.tabelaToolStripMenuItem.Size = new System.Drawing.Size(270, 34);
            this.tabelaToolStripMenuItem.Text = "Tabela";
            this.tabelaToolStripMenuItem.Click += new System.EventHandler(this.tabelaToolStripMenuItem_Click);
            // 
            // nome_table
            // 
            this.nome_table.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nome_table.Location = new System.Drawing.Point(40, 107);
            this.nome_table.Multiline = true;
            this.nome_table.Name = "nome_table";
            this.nome_table.Size = new System.Drawing.Size(318, 43);
            this.nome_table.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(40, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(286, 26);
            this.label1.TabIndex = 7;
            this.label1.Text = "Nome da  panilha importada\r\n";
            // 
            // salvarEdiçãoToolStripMenuItem
            // 
            this.salvarEdiçãoToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("salvarEdiçãoToolStripMenuItem.Image")));
            this.salvarEdiçãoToolStripMenuItem.Name = "salvarEdiçãoToolStripMenuItem";
            this.salvarEdiçãoToolStripMenuItem.Size = new System.Drawing.Size(322, 34);
            this.salvarEdiçãoToolStripMenuItem.Text = "Salvar Edição";
            this.salvarEdiçãoToolStripMenuItem.Click += new System.EventHandler(this.salvarEdiçãoToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(818, 511);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.nome_table);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExcelToDB";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ToolStripMenuItem abrirToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem arqivoExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem arquivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salvarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem limparToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.TextBox nome_table;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem osDadosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tudoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem bancoDeDadosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tabelaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem salvarEdiçãoToolStripMenuItem;
    }
}

