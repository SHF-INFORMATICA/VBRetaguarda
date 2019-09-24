using Ionic.Zip;
using System;
using System.IO;
using System.IO.Compression;
using System.Collections;
using System.Windows.Forms;

namespace Macoratti.TesteEmail
{
    public partial class frmEmail : Form
    {
        /// <summary>
        /// Um array lista contento todos os anexos
        /// </summary>
        ArrayList aAnexosEmail;
        private string CaminhoNFE_A;
        private int Cont_N = 0;

        /// <summary>
        /// O construtor padrão
        /// </summary>
        public frmEmail()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Incluir arquivos a serem anexaso na mensagem
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnIncluir_Click(object sender, EventArgs e)
        {
            if (ofd1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string[] arr = ofd1.FileNames;
                    aAnexosEmail = new ArrayList();
                    txtAnexos.Text = string.Empty;
                    aAnexosEmail.AddRange(arr);

                    foreach (string s in aAnexosEmail)
                    {
                        txtAnexos.Text += s + "; ";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                }
            }
        }
        /// <summary>
        /// Sai da aplicação
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancelar_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        /// <summary>
        /// Envia uma mensagem de mail com ou sem anexos
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEnviar_Click(object sender, EventArgs e)
        {
                 
            if (String.IsNullOrEmpty(txtEnviarPara.Text))
            {
                MessageBox.Show("Endereço de email do destinatário inválido.", "Erro ");
                return;
            }
            if (String.IsNullOrEmpty(txtEnviadoPor.Text))
            {
                MessageBox.Show("Endereço de email do remetente inválido.", "Erro ");
                return;
            }
            if (String.IsNullOrEmpty(txtAssuntoTitulo.Text))
            {
                MessageBox.Show("Definição do assunto inválida.", "Erro ");
                return;
            }
            if (String.IsNullOrEmpty(txtMensagem.Text))
            {
                MessageBox.Show("Mensagem inválida.", "Erro ");
                return;
            }

            //separa os anexos em um array de string
            string[] arr = txtAnexos.Text.Split(';');
            //cria um novo arraylist
            aAnexosEmail = new ArrayList();
            //percorre o array de string e inclui os anexos
            for (int i = 0; i < arr.Length; i++)
            {
                if (!String.IsNullOrEmpty(arr[i].ToString().Trim()))
                {
                    aAnexosEmail.Add(arr[i].ToString().Trim());
                }
            }

            // Se existirem anexos , envia a mensagem com 
            // a chamada a EnviaMensagemComAnexos senão
            // usa o método enviaMensagemEmail
            if (aAnexosEmail.Count > 0)
            {
                string resultado = EnviaEmail.EnviaEmail.EnviaMensagemComAnexos(txtEnviarPara.Text,
                    txtEnviadoPor.Text, txtAssuntoTitulo.Text, txtMensagem.Text,
                    aAnexosEmail);

                MessageBox.Show(resultado, "Email enviado com sucesso");
            }
            else
            {
                string resultado = EnviaEmail.EnviaEmail.EnviaMensagemEmail(txtEnviarPara.Text,
                    txtEnviadoPor.Text, txtAssuntoTitulo.Text, txtMensagem.Text);

                MessageBox.Show(resultado, "Email enviado com sucesso");
            }

        }

        private void cmbMesAno_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbMesAno.SelectedItem.ToString() == "")
            {
                MessageBox.Show("Selecione um Ano e Mês", "MegaNfe", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }

            try
            {
                ZipFile.CreateFromDirectory(txtLocal.Text, txtArquivoCompactado.Text);

                MessageBox.Show($"Compactação da pasta : <<{txtLocal.Text}>> feita com sucesso...");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void frmEmail_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //C:\MEGASIM\RETAGUARDA\EnviarEmail
            //CaminhoNFE_A = "@" + Application.StartupPath;
            CaminhoNFE_A = Application.StartupPath;

            //Console.WriteLine("Original string: \"{0}\"", str);
            //Console.WriteLine("CSV string:      \"{0}\"", str.Replace(' ', ','));

            CaminhoNFE_A = CaminhoNFE_A.Replace ("'\'","'\\'") + "\\megasim.ini";

            //MessageBox.Show(CaminhoNFE_A, "app.path");

            string[] linhas = File.ReadAllLines(CaminhoNFE_A);
            Cont_N = 1;
            
            foreach (string linha in linhas)
            {
                //MessageBox.Show(CaminhoNFE_A, "LENDO ARQUIVO .INI");
                //CaminhoNFE_A = "@" + linha;
                if (Cont_N == 15) {
                   CaminhoNFE_A = linha;
                }
                
                Cont_N = Cont_N + 1;
            }

            CaminhoNFE_A = CaminhoNFE_A.Replace("'\\\\'", "'\\'");

            string[] filePaths = Directory.GetDirectories(CaminhoNFE_A);

            txtLocal.Text = CaminhoNFE_A;

            DirectoryInfo di = new DirectoryInfo(CaminhoNFE_A);
            DirectoryInfo[] directories = di.GetDirectories("*", SearchOption.TopDirectoryOnly);
            foreach (var _file in directories)
            {
                //pega o nome do arquivo

                cmbMesAno.Items.Add(_file.Name.ToString());
            }
        }
    }
}
