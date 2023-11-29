using System.Data.SqlClient;
using System;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Drawing;

namespace SisFaturasENotasFiscais
{
    #region Formulario
    public partial class FrmLogin : Form
    {
        #region Funcao
        private async void Login()
        {
            try
            {
                if (TxtUsuario.Text == "")
                {
                    MessageBox.Show("Digite o Usuário", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtUsuario.Focus();
                }
                else if (TxtSenha.Text == "")
                {
                    MessageBox.Show("Digite uma Senha!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtSenha.Focus();
                }
                else
                {
                    bool AcessoPermitido = false;

                    try
                    {
                        using (SqlConnection connection = new SqlConnection("Data Source=ALMEIDA;Initial Catalog=SisFaturasENotasFiscais;Integrated Security=True"))
                        {
                            connection.Open();

                            using (SqlCommand cmd = new SqlCommand("SELECT * FROM Usuario WHERE NomeUsuario = @NomeUsuario AND Senha = @Senha", connection))
                            {
                                cmd.Parameters.AddWithValue("@NomeUsuario", TxtUsuario.Text);
                                cmd.Parameters.AddWithValue("@Senha", TxtSenha.Text);
                                cmd.Parameters.AddWithValue("@Cargo", LblCargo.Text);

                                using (SqlDataReader dr = cmd.ExecuteReader())
                                {
                                    if (dr.Read())
                                    {
                                        LblIDUsuario.Text = dr["IDUsuario"].ToString();
                                        LblCargo.Text = dr["Cargo"].ToString();
                                        AcessoPermitido = true;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Falha ao conectar!" + Environment.NewLine + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    if (AcessoPermitido)
                    {
                        int idUsuario;
                        if (int.TryParse(LblIDUsuario.Text, out idUsuario))
                        {
                            Image imagemOriginal = CarregaImagemDoBanco(idUsuario);

                            if (imagemOriginal != null)
                            {
                                double fatorDeZoom = 0.4;
                                int novaLargura = (int)(imagemOriginal.Width * fatorDeZoom);
                                int novaAltura = (int)(imagemOriginal.Height * fatorDeZoom);

                                Image imagemRedimensionada = new Bitmap(imagemOriginal, novaLargura, novaAltura);

                                PicBoxImagemUser.SizeMode = PictureBoxSizeMode.CenterImage;
                                PicBoxImagemUser.Image = imagemRedimensionada;

                                LblBemVindo.Visible = true;
                                LblWelcome.Visible = true;
                                PgBarLogin.Visible = true;
                                LblWelcome.Text = TxtUsuario.Text;

                                Thread thread = new Thread(new ThreadStart(CarregaProgressBar));
                                thread.Start();

                                await Task.Delay(3000);

                                FrmMenuPrincipal frm = new FrmMenuPrincipal();
                                frm.UsuarioLogado = TxtUsuario.Text;
                                frm.CargoUsuario = LblCargo.Text;
                                frm.Show();
                                this.Visible = false;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Usuário ou Senha inválidos!", "Sistema", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        TxtUsuario.Focus();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private Image CarregaImagemDoBanco(int idUsuario)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection("Data Source=ALMEIDA;Initial Catalog=SisFaturasENotasFiscais;Integrated Security=True"))
                {
                    using (SqlCommand cmd = new SqlCommand("SELECT Foto FROM Usuario WHERE IDUsuario = @IDUsuario", connection))
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@IDUsuario", idUsuario);

                        connection.Open();

                        byte[] imagemBytes = (byte[])cmd.ExecuteScalar();

                        if (imagemBytes != null)
                        {
                            using (MemoryStream ms = new MemoryStream(imagemBytes))
                            {
                                return Image.FromStream(ms);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Imagem não encontrada.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar imagem: " + ex.Message);
            }

            return null;
        }

        private void CarregaProgressBar()
        {
            PgBarLogin.Invoke(new Action(() => PgBarLogin.Maximum = 5));

            for (int i = 0; i <= 5; i++)
            {
                PgBarLogin.Invoke(new Action(() => PgBarLogin.Value = i));
                System.Threading.Thread.Sleep(500);
            }
            PgBarLogin.Invoke(new Action(() => PgBarLogin.Visible = false));
        }
        #endregion

        public FrmLogin()
        {
            InitializeComponent();
        }

        private void FrmLogin_Load(object sender, EventArgs e)
        {
            try
            {
                TxtUsuario.Select();
                LblOlhoAberto.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FrmLogin_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    SendKeys.Send("{TAB}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnCriarConta_Click(object sender, EventArgs e)
        {
            try
            {
                //Cadastros.FrmUsuarios frm = new Cadastros.FrmUsuarios();
                //frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LblOlhoFechado_Click(object sender, EventArgs e)
        {
            try
            {
                TxtSenha.PasswordChar = '\0';
                LblOlhoAberto.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LblOlhoAberto_Click(object sender, EventArgs e)
        {
            try
            {
                TxtSenha.PasswordChar = '*';
                LblOlhoAberto.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            try
            {
                Login();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LblSair_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void PicBoxImagemUser_Click(object sender, EventArgs e)
        {

        }
    }
    #endregion
}
