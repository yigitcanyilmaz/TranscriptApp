using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace YigitCanYilmazHW
{
    public partial class Form1 : Form
    {
        private int currentPage = 1;
        private int recordsPerPage = 5;
        private int totalRecords = 0;
        private int totalPages = 0;
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\YGYDB.accdb";

        public Form1()
        {
            InitializeComponent();
            this.button1.Click += new System.EventHandler(this.btnNext_Click);
            this.button2.Click += new System.EventHandler(this.btnPrevious_Click);
            dataGridView1.CellFormatting += new DataGridViewCellFormattingEventHandler(dataGridView1_CellFormatting);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadPage();
        }

        private void LoadPage()
        {
            totalRecords = GetTotalRecordCount();
            totalPages = (int)Math.Ceiling((double)totalRecords / recordsPerPage);
            dataGridView1.DataSource = GetRecordsForPage(currentPage, recordsPerPage);
            UpdatePageInfo();
            dataGridView1.AllowUserToAddRows = false;
            DataGridViewButtonColumn btnColumn = new DataGridViewButtonColumn();
            btnColumn.Name = "viewButtonColumn";
            btnColumn.HeaderText = "View";
            btnColumn.Text = "View";
            btnColumn.UseColumnTextForButtonValue = true;

            dataGridView1.Columns.Add(btnColumn);

            dataGridView1.RowHeadersVisible = false;

        }

        private DataTable GetRecordsForPage(int page, int recordsPerPage)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                int startRecord = (page - 1) * recordsPerPage;
                string commandText = $@"
    SELECT studentid AS [Student Number],Name,Birth,Department,City,status FROM 
        (SELECT TOP {recordsPerPage} * FROM 
            (SELECT TOP {recordsPerPage + startRecord} * FROM Student ORDER BY studentid ASC) 
         AS Tmp ORDER BY studentid DESC)
    AS Tmp2 ORDER BY studentid ASC";

                OleDbCommand cmd = new OleDbCommand(commandText, con);
                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);

                try
                {
                    con.Open();
                    adapter.Fill(dataTable);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while retrieving data: " + ex.Message);
                }
            }
            return dataTable;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dataGridView1.Columns["status"] != null && e.ColumnIndex == dataGridView1.Columns["status"].Index)
            {
                var statusValue = e.Value;

                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                if (statusValue != null && statusValue is bool)
                {
                    if ((bool)statusValue)
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.Red;
                    }
                }
            }
        }

        private int GetTotalRecordCount()
        {
            int count = 0;
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                string commandText = "SELECT COUNT(*) FROM Student";

                OleDbCommand cmd = new OleDbCommand(commandText, con);

                try
                {
                    con.Open();
                    count = Convert.ToInt32(cmd.ExecuteScalar());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while retrieving data: " + ex.Message);
                }
            }
            return count;
        }

        private void UpdatePageInfo()
        {
            label1.Text = $"Page {currentPage} / {totalPages}";
            button2.Enabled = currentPage > 1;
            button1.Enabled = currentPage < totalPages;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (currentPage < totalPages)
            {
                currentPage++;
                LoadPage();
            }
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (currentPage > 1)
            {
                currentPage--;
                LoadPage();
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["viewButtonColumn"].Index && e.RowIndex >= 0)
            {
                int studentIdColumnIndex = dataGridView1.Columns["Student Number"].Index;

                string studentId = dataGridView1.Rows[e.RowIndex].Cells[studentIdColumnIndex].Value.ToString();

                using (Form2 form2 = new Form2(studentId))
                {
                    form2.ShowDialog(this);

                    RefreshData();
                }
            }
        }
        private void RefreshData()
        {
            LoadPage();
        }
    }
}
