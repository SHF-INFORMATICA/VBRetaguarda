namespace Macoratti.TesteEmail
{
    partial class frmEmail
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmEmail));
            this.grbDePara = new System.Windows.Forms.GroupBox();
            this.txtAssuntoTitulo = new System.Windows.Forms.TextBox();
            this.txtEnviadoPor = new System.Windows.Forms.TextBox();
            this.txtEnviarPara = new System.Windows.Forms.TextBox();
            this.lblSubjectLine = new System.Windows.Forms.Label();
            this.lblRemetente = new System.Windows.Forms.Label();
            this.lblDestinatario = new System.Windows.Forms.Label();
            this.grpMensagem = new System.Windows.Forms.GroupBox();
            this.txtMensagem = new System.Windows.Forms.TextBox();
            this.btnIncluir = new System.Windows.Forms.Button();
            this.btnEnviar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.ofd1 = new System.Windows.Forms.OpenFileDialog();
            this.cmbMesAno = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.svfdlg1 = new System.Windows.Forms.SaveFileDialog();
            this.txtLocal = new System.Windows.Forms.TextBox();
            this.txtAnexos = new System.Windows.Forms.TextBox();
            this.txtArquivoCompactado = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.grbDePara.SuspendLayout();
            this.grpMensagem.SuspendLayout();
            this.SuspendLayout();
            // 
            // grbDePara
            // 
            this.grbDePara.Controls.Add(this.txtAssuntoTitulo);
            this.grbDePara.Controls.Add(this.txtEnviadoPor);
            this.grbDePara.Controls.Add(this.txtEnviarPara);
            this.grbDePara.Controls.Add(this.lblSubjectLine);
            this.grbDePara.Controls.Add(this.lblRemetente);
            this.grbDePara.Controls.Add(this.lblDestinatario);
            this.grbDePara.Location = new System.Drawing.Point(17, 13);
            this.grbDePara.Margin = new System.Windows.Forms.Padding(4);
            this.grbDePara.Name = "grbDePara";
            this.grbDePara.Padding = new System.Windows.Forms.Padding(4);
            this.grbDePara.Size = new System.Drawing.Size(647, 113);
            this.grbDePara.TabIndex = 0;
            this.grbDePara.TabStop = false;
            // 
            // txtAssuntoTitulo
            // 
            this.txtAssuntoTitulo.Location = new System.Drawing.Point(77, 73);
            this.txtAssuntoTitulo.Margin = new System.Windows.Forms.Padding(4);
            this.txtAssuntoTitulo.Name = "txtAssuntoTitulo";
            this.txtAssuntoTitulo.Size = new System.Drawing.Size(531, 22);
            this.txtAssuntoTitulo.TabIndex = 5;
            this.txtAssuntoTitulo.Text = "XML";
            // 
            // txtEnviadoPor
            // 
            this.txtEnviadoPor.Location = new System.Drawing.Point(77, 44);
            this.txtEnviadoPor.Margin = new System.Windows.Forms.Padding(4);
            this.txtEnviadoPor.Name = "txtEnviadoPor";
            this.txtEnviadoPor.Size = new System.Drawing.Size(531, 22);
            this.txtEnviadoPor.TabIndex = 4;
            this.txtEnviadoPor.Text = "shfhoracio@gmail.com";
            // 
            // txtEnviarPara
            // 
            this.txtEnviarPara.Location = new System.Drawing.Point(77, 14);
            this.txtEnviarPara.Margin = new System.Windows.Forms.Padding(4);
            this.txtEnviarPara.Name = "txtEnviarPara";
            this.txtEnviarPara.Size = new System.Drawing.Size(531, 22);
            this.txtEnviarPara.TabIndex = 3;
            this.txtEnviarPara.Text = "shfhoracio@gmail.com";
            // 
            // lblSubjectLine
            // 
            this.lblSubjectLine.AutoSize = true;
            this.lblSubjectLine.Location = new System.Drawing.Point(8, 78);
            this.lblSubjectLine.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblSubjectLine.Name = "lblSubjectLine";
            this.lblSubjectLine.Size = new System.Drawing.Size(61, 16);
            this.lblSubjectLine.TabIndex = 2;
            this.lblSubjectLine.Text = "Assunto:";
            // 
            // lblRemetente
            // 
            this.lblRemetente.AutoSize = true;
            this.lblRemetente.Location = new System.Drawing.Point(40, 49);
            this.lblRemetente.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblRemetente.Name = "lblRemetente";
            this.lblRemetente.Size = new System.Drawing.Size(29, 16);
            this.lblRemetente.TabIndex = 1;
            this.lblRemetente.Text = "De:";
            // 
            // lblDestinatario
            // 
            this.lblDestinatario.AutoSize = true;
            this.lblDestinatario.Location = new System.Drawing.Point(30, 18);
            this.lblDestinatario.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDestinatario.Name = "lblDestinatario";
            this.lblDestinatario.Size = new System.Drawing.Size(42, 16);
            this.lblDestinatario.TabIndex = 0;
            this.lblDestinatario.Text = "Para:";
            // 
            // grpMensagem
            // 
            this.grpMensagem.Controls.Add(this.txtMensagem);
            this.grpMensagem.Location = new System.Drawing.Point(13, 134);
            this.grpMensagem.Margin = new System.Windows.Forms.Padding(4);
            this.grpMensagem.Name = "grpMensagem";
            this.grpMensagem.Padding = new System.Windows.Forms.Padding(4);
            this.grpMensagem.Size = new System.Drawing.Size(647, 113);
            this.grpMensagem.TabIndex = 1;
            this.grpMensagem.TabStop = false;
            this.grpMensagem.Text = "Mensagem";
            // 
            // txtMensagem
            // 
            this.txtMensagem.Location = new System.Drawing.Point(13, 25);
            this.txtMensagem.Margin = new System.Windows.Forms.Padding(4);
            this.txtMensagem.Multiline = true;
            this.txtMensagem.Name = "txtMensagem";
            this.txtMensagem.Size = new System.Drawing.Size(595, 77);
            this.txtMensagem.TabIndex = 0;
            this.txtMensagem.Text = "Bom dia. Segue anexo arquivos XML";
            // 
            // btnIncluir
            // 
            this.btnIncluir.Location = new System.Drawing.Point(242, 359);
            this.btnIncluir.Margin = new System.Windows.Forms.Padding(4);
            this.btnIncluir.Name = "btnIncluir";
            this.btnIncluir.Size = new System.Drawing.Size(100, 28);
            this.btnIncluir.TabIndex = 7;
            this.btnIncluir.Text = "Incluir";
            this.btnIncluir.UseVisualStyleBackColor = true;
            this.btnIncluir.Click += new System.EventHandler(this.btnIncluir_Click);
            // 
            // btnEnviar
            // 
            this.btnEnviar.Location = new System.Drawing.Point(350, 359);
            this.btnEnviar.Margin = new System.Windows.Forms.Padding(4);
            this.btnEnviar.Name = "btnEnviar";
            this.btnEnviar.Size = new System.Drawing.Size(100, 28);
            this.btnEnviar.TabIndex = 3;
            this.btnEnviar.Text = "Enviar";
            this.btnEnviar.UseVisualStyleBackColor = true;
            this.btnEnviar.Click += new System.EventHandler(this.btnEnviar_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.Location = new System.Drawing.Point(458, 359);
            this.btnCancelar.Margin = new System.Windows.Forms.Padding(4);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(100, 28);
            this.btnCancelar.TabIndex = 4;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // ofd1
            // 
            this.ofd1.FileName = "openFileDialog1";
            this.ofd1.Multiselect = true;
            this.ofd1.Title = "Add Attachment";
            // 
            // cmbMesAno
            // 
            this.cmbMesAno.FormattingEnabled = true;
            this.cmbMesAno.Location = new System.Drawing.Point(104, 248);
            this.cmbMesAno.Name = "cmbMesAno";
            this.cmbMesAno.Size = new System.Drawing.Size(121, 24);
            this.cmbMesAno.TabIndex = 5;
            this.cmbMesAno.SelectedIndexChanged += new System.EventHandler(this.cmbMesAno_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 251);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "ANO/MÊS:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(576, 359);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 28);
            this.button1.TabIndex = 8;
            this.button1.Text = "teste";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtLocal
            // 
            this.txtLocal.Location = new System.Drawing.Point(232, 250);
            this.txtLocal.Margin = new System.Windows.Forms.Padding(4);
            this.txtLocal.Name = "txtLocal";
            this.txtLocal.Size = new System.Drawing.Size(389, 22);
            this.txtLocal.TabIndex = 9;
            // 
            // txtAnexos
            // 
            this.txtAnexos.Location = new System.Drawing.Point(104, 319);
            this.txtAnexos.Margin = new System.Windows.Forms.Padding(4);
            this.txtAnexos.Name = "txtAnexos";
            this.txtAnexos.Size = new System.Drawing.Size(483, 22);
            this.txtAnexos.TabIndex = 10;
            // 
            // txtArquivoCompactado
            // 
            this.txtArquivoCompactado.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtArquivoCompactado.Location = new System.Drawing.Point(232, 281);
            this.txtArquivoCompactado.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtArquivoCompactado.Name = "txtArquivoCompactado";
            this.txtArquivoCompactado.Size = new System.Drawing.Size(389, 24);
            this.txtArquivoCompactado.TabIndex = 11;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 319);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 16);
            this.label2.TabIndex = 12;
            this.label2.Text = "Anexo(s):";
            // 
            // frmEmail
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ClientSize = new System.Drawing.Size(679, 400);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtArquivoCompactado);
            this.Controls.Add(this.txtAnexos);
            this.Controls.Add(this.txtLocal);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnIncluir);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbMesAno);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.grpMensagem);
            this.Controls.Add(this.btnEnviar);
            this.Controls.Add(this.grbDePara);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmEmail";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Enviar email XML Contador";
            this.Load += new System.EventHandler(this.frmEmail_Load);
            this.grbDePara.ResumeLayout(false);
            this.grbDePara.PerformLayout();
            this.grpMensagem.ResumeLayout(false);
            this.grpMensagem.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox grbDePara;
        private System.Windows.Forms.TextBox txtAssuntoTitulo;
        private System.Windows.Forms.TextBox txtEnviadoPor;
        private System.Windows.Forms.TextBox txtEnviarPara;
        private System.Windows.Forms.Label lblSubjectLine;
        private System.Windows.Forms.Label lblRemetente;
        private System.Windows.Forms.Label lblDestinatario;
        private System.Windows.Forms.GroupBox grpMensagem;
        private System.Windows.Forms.TextBox txtMensagem;
        private System.Windows.Forms.Button btnIncluir;
        private System.Windows.Forms.Button btnEnviar;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.OpenFileDialog ofd1;
        private System.Windows.Forms.ComboBox cmbMesAno;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.SaveFileDialog svfdlg1;
        private System.Windows.Forms.TextBox txtLocal;
        private System.Windows.Forms.TextBox txtAnexos;
        private System.Windows.Forms.TextBox txtArquivoCompactado;
        private System.Windows.Forms.Label label2;
    }
}

