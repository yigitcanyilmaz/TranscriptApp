using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Reflection.Emit;

namespace YigitCanYilmazHW
{
    public partial class Form2 : Form
    {
        private int currentCoursePage = 1;
        private int recordsPerCoursePage = 5;
        private int totalCourseRecords = 0;
        private int totalCoursePages = 0;
        private string studentId;
        private string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\YGYDB.accdb";
        public string StudentName { get; private set; }
        public DateTime StudentBirth { get; private set; }
        public string StudentDepartment { get; private set; }
        public string StudentCity { get; private set; }
        public string StudentPicture { get; private set; }
        public bool StudentStatus { get; private set; }


        public Form2(string studentId)
        {
            InitializeComponent();
            this.studentId = studentId;
            label8.Text = studentId;
            this.button1.Click += new System.EventHandler(this.btnNext_Click);
            this.button2.Click += new System.EventHandler(this.btnPrevious_Click);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            LoadCoursePage();

        }
        private void LoadCoursePage()
        {
            LoadStudentDetails(Convert.ToInt32(studentId));
            totalCourseRecords = GetTotalCourseRecordCount();
            totalCoursePages = (int)Math.Ceiling((double)totalCourseRecords / recordsPerCoursePage);
            dataGridView1.DataSource = GetCoursesForPage(currentCoursePage, recordsPerCoursePage);
            UpdateCoursePageInfo();
            CalculateAverageGrade();
            dataGridView1.AllowUserToAddRows = false;
        }
        private void LoadStudentDetails(int studentId)
        {
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                string commandText = "SELECT Name, Birth, Department, City ,status ,Picture FROM Student WHERE studentid = @studentId";

                OleDbCommand cmd = new OleDbCommand(commandText, con);
                cmd.Parameters.AddWithValue("@studentId", studentId);

                try
                {
                    con.Open();
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            StudentName = reader["Name"].ToString();
                            StudentBirth = Convert.ToDateTime(reader["Birth"]);
                            StudentDepartment = reader["Department"].ToString();
                            StudentCity = reader["City"].ToString();
                            StudentPicture = reader["Picture"].ToString();
                            StudentStatus = (bool)reader["status"];
                            Labels();
                            LoadStudentPicture();
                        }
                        else
                        {
                            MessageBox.Show("Student not found.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while retrieving student information: " + ex.Message);
                }
            }
        }

        private void LoadStudentPicture()
        {
            string picturePath = StudentPicture;

            try
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                pictureBox1.Image = System.Drawing.Image.FromFile(picturePath);

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while loading the picture: " + ex.Message);
            }
        }




        private void CalculateAverageGrade()
        {
            int sumOfGrades = 0;
            int numberOfGrades = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["Grade"].Value != null)
                {
                    string grade = row.Cells["Grade"].Value.ToString().Trim();
                    switch (grade)
                    {
                        case "A":
                            sumOfGrades += 10;
                            break;
                        case "B":
                            sumOfGrades += 8;
                            break;
                        case "C":
                            sumOfGrades += 6;
                            break;
                        case "D":
                            sumOfGrades += 5;
                            break;
                        case "E":
                            sumOfGrades += 4;
                            break;
                        case "F":
                            sumOfGrades += 3;
                            break;
                        case "G":
                            sumOfGrades += 0;
                            break;
                    }
                    numberOfGrades++;
                }
            }

            double average = numberOfGrades > 0 ? (double)sumOfGrades / numberOfGrades : 0;

            label10.Text = average.ToString("0.00");
        }

        private DataTable GetCoursesForPage(int page, int recordsPerPage)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                int startRecord = (page - 1) * recordsPerPage;
                string commandText = $@"
SELECT Course,Year,Grade FROM 
  (SELECT TOP {recordsPerPage} * FROM 
    (SELECT TOP {startRecord + recordsPerPage} * FROM Course 
     WHERE studentid = {studentId} 
     ORDER BY courseid ASC) AS TmpOrdered 
   ORDER BY courseid DESC) AS TmpOrderedReverse
ORDER BY courseid ASC";


                OleDbCommand cmd = new OleDbCommand(commandText, con);
                OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);

                try
                {
                    con.Open();
                    adapter.Fill(dataTable);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error retrieving data: " + ex.Message);
                }
            }
            return dataTable;
        }
        public void Labels()
        {
            if (StudentStatus == true) { label11.Text = "Active"; button1.BackColor = Color.White; button1.Text = "Deactive"; }
            else { label11.Text = "Deactive"; button1.BackColor = Color.Red; button1.Text = "Active"; }
            this.Text = "" + StudentName.ToString() + "'s Details";
            label7.Text = StudentName.ToString();
            label8.Text = studentId.ToString();
            label9.Text = StudentBirth.ToString("yyyy-MM-dd");

        }

        private int GetTotalCourseRecordCount()
        {
            int count = 0;
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                string commandText = $"SELECT COUNT(*) FROM Course WHERE studentid = {studentId}";

                OleDbCommand cmd = new OleDbCommand(commandText, con);

                try
                {
                    con.Open();
                    count = Convert.ToInt32(cmd.ExecuteScalar());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error retrieving total record count: " + ex.Message);
                }
            }
            return count;
        }

        private void UpdateCoursePageInfo()
        {
            label14.Text = $"Page {currentCoursePage} / {totalCoursePages}";
            button4.Enabled = currentCoursePage > 1;
            button3.Enabled = currentCoursePage < totalCoursePages;
        }
        private void buttonNextCoursePage_Click(object sender, EventArgs e)
        {
            if (currentCoursePage < totalCoursePages)
            {
                currentCoursePage++;
                LoadCoursePage();
            }
        }

        private void buttonPreviousCoursePage_Click(object sender, EventArgs e)
        {
            if (currentCoursePage > 1)
            {
                currentCoursePage--;
                LoadCoursePage();
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (currentCoursePage < totalCoursePages)
            {
                currentCoursePage++;
                LoadCoursePage();
            }
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            if (currentCoursePage > 1)
            {
                currentCoursePage--;
                LoadCoursePage();
            }
        }

        private void UpdateStudentStatus()
        {
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                string commandText = "UPDATE Student SET status = NOT status WHERE studentid = @StudentId";

                OleDbCommand cmd = new OleDbCommand(commandText, con);
                cmd.Parameters.AddWithValue("@StudentId", studentId);

                try
                {
                    con.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Veri güncellenirken bir hata oluştu: " + ex.Message);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UpdateStudentStatus();
            LoadCoursePage(); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";
            saveFileDialog.DefaultExt = "pdf";
            saveFileDialog.AddExtension = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filename = saveFileDialog.FileName;
                PdfDocument document = new PdfDocument();
                document.Info.Title = "Created with PDFsharp";

                PdfPage page = document.AddPage();
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XFont titleFont = new XFont("Verdana", 20, XFontStyle.BoldItalic);
                XFont contentFont = new XFont("Verdana", 12, XFontStyle.Regular);

                gfx.DrawString("Student Details", titleFont, XBrushes.Black,
                    new XRect(0, 0, page.Width, page.Height),
                    XStringFormats.TopCenter);

                double yPos = 100;
                gfx.DrawString($"Name: {StudentName}", contentFont, XBrushes.Black, new XPoint(40, yPos));
                yPos += 30;
                gfx.DrawString($"Number: {studentId}", contentFont, XBrushes.Black, new XPoint(40, yPos));
                yPos += 30;
                gfx.DrawString($"Birthdate: {StudentBirth:yyyy-MM-dd}", contentFont, XBrushes.Black, new XPoint(40, yPos));
                yPos += 30;
                gfx.DrawString($"Department: {StudentDepartment}", contentFont, XBrushes.Black, new XPoint(40, yPos));
                yPos += 30;
                gfx.DrawString($"City: {StudentCity}", contentFont, XBrushes.Black, new XPoint(40, yPos));
                yPos += 30;
                gfx.DrawString($"GPA: {label10.Text}", contentFont, XBrushes.Black, new XPoint(40, yPos));
                yPos += 40;

                if (dataGridView1.Rows.Count > 0)
                {
                    double xPos = 40;
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        gfx.DrawString(col.HeaderText, contentFont, XBrushes.Black, new XPoint(xPos, yPos));
                        xPos += col.Width; 
                    }
                    yPos += 30;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        xPos = 40;
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            gfx.DrawString(Convert.ToString(cell.Value), contentFont, XBrushes.Black, new XPoint(xPos, yPos));
                            xPos += dataGridView1.Columns[cell.ColumnIndex].Width;
                        }
                        yPos += 20;
                    }
                }

                document.Save(filename);
                document.Close();

                MessageBox.Show($"PDF saved to: {filename}");
            }
        }




    }
}

