using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
//using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace SisFaturasENotasFiscais.Cadastros
{
    #region Formulario
    public partial class FrmUsuarios : Form
    {
        private SqlConnection conexao = new SqlConnection("Data Source=ALMEIDA;Initial Catalog=SisFaturasENotasFiscais;Integrated Security=True");
        string foto;

        #region Funcao
        private void ConfiguraLista()
        {
            try
            {
                DgUser.Columns["IDUsuario"].Visible = false;
                DgUser.Columns["NomeUsuario"].DisplayIndex = 1;
                DgUser.Columns["Senha"].DisplayIndex = 2;
                DgUser.Columns["DataCadastro"].DisplayIndex = 4;
                DgUser.Columns["Cargo"].DisplayIndex = 3;
                DgUser.Columns["Foto"].Visible = false;

                DgUser.Columns["NomeUsuario"].HeaderText = "Usuário";
                DgUser.Columns["Senha"].HeaderText = "Senha";
                DgUser.Columns["DataCadastro"].HeaderText = "Cadastro";
                DgUser.Columns["Cargo"].HeaderText = "Cargo";

                DgUser.Columns["NomeUsuario"].Width = 90;
                DgUser.Columns["Senha"].Width = 90;
                DgUser.Columns["DataCadastro"].Width = 50;
                DgUser.Columns["Cargo"].Width = 250;

                DgUser.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
                DgUser.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void AtualizaLista()
        {
            try
            {
                SqlCommand cmd = new SqlCommand("SELECT * FROM Usuario", conexao);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();

                da.Fill(dt);
                DgUser.DataSource = dt;
                ConfiguraLista();
                ContagemRegistros();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void CarregaComboCargo()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection("Data Source=ALMEIDA;Initial Catalog=SisFaturasENotasFiscais;Integrated Security=True"))
                {
                    connection.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT * FROM Cargo", connection))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        DataTable table = new DataTable();
                        adapter.Fill(table);

                        CmbCargo.DataSource = table;
                        CmbCargo.DisplayMember = "Cargo";
                        CmbCargo.ValueMember = "IDCargo";
                    }
                }
                CmbCargo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void LocalizaCampos()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection("Data Source=ALMEIDA;Initial Catalog=SisFaturasENotasFiscais;Integrated Security=True"))
                {
                    using (SqlCommand cmd = new SqlCommand("SELECT * FROM Usuario WHERE NomeUsuario LIKE @NomeUsuario ORDER BY NomeUsuario ASC", connection))
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@NomeUsuario", TxtLocalizar.Text + "%");

                        connection.Open();

                        using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);
                            DgUser.DataSource = dt;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private byte[] CarregaImagem()
        {
            byte[] imagem_byte = null;
            if (foto == "")
            {
                return null;
            }

            FileStream fs = new FileStream(foto, FileMode.Open, FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);

            imagem_byte = br.ReadBytes((int)fs.Length);

            return imagem_byte;
        }

        private void CarregaImagemDoBanco(int idUsuario)
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
                                PicBoxImagemUser.Image = Image.FromStream(ms);
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
        }

        private void ContagemRegistros()
        {
            try
            {
                int quantidadeRegistros = DgUser.RowCount;
                LblRegistros.Text = $"{quantidadeRegistros}";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void PreencheCamposComDadosUsuarioAtual()
        {
            try
            {
                TxtID.Text = DgUser.CurrentRow.Cells[0].Value?.ToString();
                TxtUsuario.Text = DgUser.CurrentRow.Cells[1].Value?.ToString();
                TxtSenha.Text = DgUser.CurrentRow.Cells[2].Value?.ToString();
                CmbCargo.Text = DgUser.CurrentRow.Cells[4].Value?.ToString();

                if (DgUser.CurrentRow.Index >= 0)
                {
                    int idUsuario = Convert.ToInt32(DgUser.Rows[DgUser.CurrentRow.Index].Cells["IDUsuario"].Value);
                    CarregaImagemDoBanco(idUsuario);
                    PicBoxImagemUser.SizeMode = PictureBoxSizeMode.CenterImage;
                    PicBoxImagemUser.SizeMode = PictureBoxSizeMode.Zoom;
                }

                LblCarregarImagem.Visible = false;
                BtnIncluir.Enabled = false;
                PicBoxImagemUser.Enabled = false;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void GeraRelatorioUsuarios()
        {
            try
            {
                if (MessageBox.Show("Deseja importar dados Usuários ?", "ATENÇÃO", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = true;
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                    int columnIndex = 1;
                    for (int i = 0; i < DgUser.Columns.Count; i++)
                    {
                        if (DgUser.Columns[i].HeaderText != "IDUsuario" && DgUser.Columns[i].HeaderText != "Foto")
                        {
                            worksheet.Cells[1, columnIndex] = DgUser.Columns[i].HeaderText;
                            columnIndex++;
                        }
                    }

                    for (int i = 0; i < DgUser.Rows.Count; i++)
                    {
                        columnIndex = 1;
                        for (int j = 0; j < DgUser.Columns.Count; j++)
                        {
                            if (DgUser.Columns[j].HeaderText != "IDUsuario" && DgUser.Columns[j].HeaderText != "Foto")
                            {
                                if (DgUser.Columns[j].HeaderText == "Cadastro" && DgUser.Rows[i].Cells[j].Value is DateTime dataCadastro)
                                {
                                    worksheet.Cells[i + 2, columnIndex] = dataCadastro.ToShortDateString();
                                }
                                else
                                {
                                    worksheet.Cells[i + 2, columnIndex] = DgUser.Rows[i].Cells[j].Value.ToString();
                                }

                                columnIndex++;
                            }
                        }
                    }

                    worksheet.Columns.AutoFit();
                    worksheet.Columns["IDUsuario"].EntireColumn.Hidden = true;
                    worksheet.Columns["Foto"].EntireColumn.Hidden = true;

                    string downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
                    workbook.SaveAs(downloadPath + @"\Usuarios.xlsx");
                    excelApp.Quit();
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void LimpaCampos()
        {
            try
            {
                TxtUsuario.Text = string.Empty;
                TxtSenha.Text = string.Empty;
                CmbCargo.Text = string.Empty;
                TxtLocalizar.Text = string.Empty;

                PicBoxImagemUser.Enabled = true;
                LblCarregarImagem.Visible = true;
                BtnIncluir.Enabled = true;

                PicBoxImagemUser.SizeMode = PictureBoxSizeMode.CenterImage;
                PicBoxImagemUser.Image = null;
                foto = "Imagem/user.png";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        #endregion

        public FrmUsuarios()
        {
            InitializeComponent();
        }

        private void FrmUsuarios_Load(object sender, EventArgs e)
        {
            try
            {
                AtualizaLista();
                DtData.Value = DateTime.Now;
                TxtLocalizar.TextChanged += TxtLocalizar_TextChanged;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void FrmUsuarios_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    SendKeys.Send("{TAB}");
                }
                else if (e.KeyCode == Keys.Delete)
                {
                    BtnDeletar.PerformClick();
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    this.Dispose();
                }

                if (DgUser.SelectedRows.Count > 0)
                {
                    if (DgUser.CurrentRow != null)
                    {
                        // seta para cima
                        if (e.KeyCode == Keys.Up)
                        {
                            int currentIndex = DgUser.CurrentRow.Index;

                            if (currentIndex > 0)
                            {
                                int previousRowIndex = currentIndex - 1;
                                DgUser.CurrentCell = DgUser.Rows[previousRowIndex].Cells[DgUser.CurrentCell.ColumnIndex];
                                PreencheCamposComDadosUsuarioAtual();

                                e.Handled = true;
                            }
                        }
                        // seta para baixo
                        else if (e.KeyCode == Keys.Down)
                        {
                            int currentIndex = DgUser.CurrentRow.Index;

                            if (currentIndex < DgUser.Rows.Count - 1)
                            {
                                int nextRowIndex = currentIndex + 1;
                                DgUser.CurrentCell = DgUser.Rows[nextRowIndex].Cells[DgUser.CurrentCell.ColumnIndex];
                                PreencheCamposComDadosUsuarioAtual();

                                e.Handled = true;
                            }
                        }
                    }
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void CmbCargo_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                CarregaComboCargo();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void TxtLocalizar_TextChanged(object sender, EventArgs e)
        {
            try
            {
                LocalizaCampos();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void PicBoxImagemUser_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();

                dialog.Filter = "Imagens (*.jpg; *.jpeg; *.png) | *.jpg; *.jpeg; *.png";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    foto = dialog.FileName.ToString();
                    PicBoxImagemUser.ImageLocation = foto;
                    PicBoxImagemUser.SizeMode = PictureBoxSizeMode.CenterImage;
                    PicBoxImagemUser.SizeMode = PictureBoxSizeMode.Zoom;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void DgUser_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                PreencheCamposComDadosUsuarioAtual();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnAtualizar_Click(object sender, EventArgs e)
        {
            try
            {
                AtualizaLista();
                LimpaCampos();
                BtnIncluir.Enabled = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnIncluir_Click(object sender, EventArgs e)
        {
            try
            {
                if (TxtUsuario.Text == "")
                {
                    MessageBox.Show("Digite um nome de Usuário!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtUsuario.Focus();
                    return;
                }
                else if (TxtSenha.Text == "")
                {
                    MessageBox.Show("Digite uma senha para o Usuário!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtSenha.Focus();
                    return;
                }
                else if (CmbCargo.Text == "")
                {
                    MessageBox.Show("Selecione o Cargo do Usuário!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbCargo.Focus();
                    return;
                }
                else if (PicBoxImagemUser.Image == null)
                {
                    MessageBox.Show("Selecione uma imagem para o Usuário!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    PicBoxImagemUser.Focus();
                    return;
                }
                else
                {
                    using (SqlConnection connection = new SqlConnection("Data Source = ALMEIDA; Initial Catalog = SisFaturasENotasFiscais; Integrated Security = True"))
                    {
                        using (SqlCommand cmd = new SqlCommand("InserirUsuario", connection))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.AddWithValue("@NomeUsuario", TxtUsuario.Text.Trim());
                            cmd.Parameters.AddWithValue("@Senha", TxtSenha.Text.Trim());
                            string dataTexto = DtData.Text;
                            string dataFormatada;
                            if (DateTime.TryParse(dataTexto, out _))
                            {
                                dataFormatada = DateTime.Parse(dataTexto).ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                MessageBox.Show("A data não está em um formato válido.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            cmd.Parameters.AddWithValue("@DataCadastro", dataFormatada);
                            cmd.Parameters.AddWithValue("@Cargo", CmbCargo.Text.Trim());
                            cmd.Parameters.AddWithValue("@Foto", CarregaImagem());

                            connection.Open();
                            cmd.ExecuteNonQuery();

                            MessageBox.Show("Cadastro realizado com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            AtualizaLista();
                            LimpaCampos();
                            TxtUsuario.Focus();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnAlterar_Click(object sender, EventArgs e)
        {
            try
            {
                if (TxtUsuario.Text == "")
                {
                    MessageBox.Show("Digite um nome de Usuário!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtUsuario.Focus();
                    return;
                }
                else if (TxtSenha.Text == "")
                {
                    MessageBox.Show("Digite uma senha para o Usuário!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TxtSenha.Focus();
                    return;
                }
                else if (CmbCargo.Text == "")
                {
                    MessageBox.Show("Selecione o Cargo do Usuário!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CmbCargo.Focus();
                    return;
                }
                else
                {
                    using (SqlConnection connection = new SqlConnection("Data Source = ALMEIDA; Initial Catalog = SisFaturasENotasFiscais; Integrated Security = True"))
                    {
                        connection.Open();

                        using (SqlCommand cmd = new SqlCommand("AlterarUsuario", connection))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;

                            cmd.Parameters.AddWithValue("@IDUsuario", TxtID.Text.Trim());
                            cmd.Parameters.AddWithValue("@NomeUsuario", TxtUsuario.Text.Trim());
                            cmd.Parameters.AddWithValue("@Senha", TxtSenha.Text.Trim());
                            string dataTexto = DtData.Text;
                            if (DateTime.TryParse(dataTexto, out DateTime dataFormatada))
                            {
                                cmd.Parameters.AddWithValue("@DataCadastro", dataFormatada.ToString("yyyy-MM-dd"));
                            }
                            else
                            {
                                MessageBox.Show("A data não está em um formato válido.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            cmd.Parameters.AddWithValue("@Cargo", CmbCargo.Text.Trim());

                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Registro alterado com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            AtualizaLista();
                            LimpaCampos();
                            TxtUsuario.Focus();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnDeletar_Click(object sender, EventArgs e)
        {
            try
            {
                if (DgUser.SelectedRows.Count > 0)
                {
                    int usuarioID = Convert.ToInt32(DgUser.SelectedRows[0].Cells["IDUsuario"].Value);
                    string usuario = DgUser.SelectedRows[0].Cells["NomeUsuario"].Value.ToString();

                    if (MessageBox.Show("Deseja realmente deletar o Usuário: " + usuario + " ?", "EXCLUSÃO", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        using (SqlConnection connection = new SqlConnection("Data Source = ALMEIDA; Initial Catalog = SisFaturasENotasFiscais; Integrated Security = True"))
                        {
                            connection.Open();

                            using (SqlCommand command = new SqlCommand("DeletarUsuario", connection))
                            {
                                command.CommandType = CommandType.StoredProcedure;
                                command.Parameters.AddWithValue("@IDUsuario", usuarioID);
                                command.ExecuteNonQuery();
                            }
                        }

                        MessageBox.Show("Usuário deletado com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        AtualizaLista();
                        LimpaCampos();
                        TxtUsuario.Focus();
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Selecione um Usuário na lista para excluir!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnRelatorio_Click(object sender, EventArgs e)
        {
            try
            {
                GeraRelatorioUsuarios();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnLimpar_Click(object sender, EventArgs e)
        {
            try
            {
                LimpaCampos();
                TxtUsuario.Focus();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void BtnVoltar_Click(object sender, EventArgs e)
        {
            try
            {
                Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
    #endregion
}
