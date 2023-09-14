namespace ExemploExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ManipulaExcel ex = new ManipulaExcel();
                ex.Salvar(textBox1.Text, textBox2.Text);
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
        }
    }
}