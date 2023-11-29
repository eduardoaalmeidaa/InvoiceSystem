using System;
using System.Threading;
using System.Windows.Forms;

namespace SisFaturasENotasFiscais
{
    #region Formulario
    public partial class FrmMenuPrincipal : Form
    {
        public string UsuarioLogado { get; set; }
        public string CargoUsuario { get; set; }

        #region Funcao
        #endregion
        public FrmMenuPrincipal()
        {
            InitializeComponent();
        }

        private void FrmMenuPrincipal_Load(object sender, EventArgs e)
        {
            try
            {
                LblUsuario.Text = UsuarioLogado;
                LblCargo.Text = CargoUsuario;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void LblProdutos_Click(object sender, EventArgs e)
        {
            try
            {
                Cadastros.FrmProdutos frm = new Cadastros.FrmProdutos();
                frm.ShowDialog();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LblUsuarios_Click(object sender, EventArgs e)
        {
            try
            {
                Cadastros.FrmUsuarios frm = new Cadastros.FrmUsuarios();
                frm.ShowDialog();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void MnuTrocarDeUsuario_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Trocar de Usuário?", "ATENÇÃO", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    this.Close();

                    FrmLogin frm = new FrmLogin();
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        FrmMenuPrincipal myFrmPrincipal = new FrmMenuPrincipal();
                        myFrmPrincipal.Show();
                    }
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
    }
    #endregion
}
