using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace idattend_to_oliver
{
    public class Student
    {
        public string EQID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string RollClass { get; set; }
        public string Gender { get; set; }
        public string Year { get; set; }
        public string EmailAddress { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<string, Student> students = new Dictionary<string, Student>();
            students = createDictionary();

            createCsvFile(students);
        }

        private static Dictionary<string, Student> createDictionary()
        {
            // Initialize HashSet of type Student.
            Dictionary<string, Student> studentDictionary = new Dictionary<string, Student>();

            /* Read through each line of the SQL query and create a new student object.
               Each new student is then added to the dictionary (C#'s answer to a HashSet in Java), eliminating duplicates. */
            SqlConnection sqlConnectionString = new SqlConnection("Server= eqmtn7123002\\IDattend; Database= IDAttend2016; Integrated Security=True;");
            SqlCommand cmd = openSqlConnection(sqlConnectionString);

            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    var student = new Student();
                    student.EQID = reader["ID"].ToString();
                    student.FirstName = reader["PreferredName"].ToString();
                    student.LastName = reader["PreferredLastName"].ToString();
                    student.RollClass = reader["HomeGroup"].ToString();
                    student.Gender = reader["Sex"].ToString();

                    /* Special case for Prep students. We want the year level
                       to read 'PY' rather than zero. */
                    if (reader["Year"].ToString() == "0") {
                        student.Year = "PY";
                    }
                    else {
                        student.Year = reader["Year"].ToString();
                    }
                   
                    student.EmailAddress = reader["Email"].ToString();

                    // Check to make sure we are not inserting duplicate keys.
                    if (!studentDictionary.ContainsKey(student.EQID))
                    {
                        studentDictionary.Add(student.EQID, student);
                    }
                    else
                    {
                        // Skip this record.
                    }
                }
            }

            // Close SQL connection.
            sqlConnectionString.Close();

            return studentDictionary;
        }

        private static SqlCommand openSqlConnection(SqlConnection connectionString)
        {
            try
            {
                connectionString.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            // SQL query to create a join with the desired columns.
            SqlCommand cmd = new SqlCommand("SELECT dbo.tblStudents.ID, dbo.tblStudents.PreferredName, dbo.tblStudents.PreferredLastName," +
                                    " dbo.tblStudents.Year, dbo.tblStudents.HomeGroup, dbo.tblStudents.Sex, dbo.tblContact.Email, dbo.tblStudents.Active" +
                                    " FROM dbo.tblContact" +
                                    " LEFT OUTER JOIN dbo.tblStudents ON dbo.tblContact.StudentID = dbo.tblStudents.ID" +
                                    " WHERE dbo.tblContact.Email IS NOT NULL AND DATALENGTH(dbo.tblContact.Email) > 0 AND dbo.tblStudents.Active = 1", connectionString);

            return cmd;
        }

        private static void closeSqlConnection(SqlConnection connectionString)
        {
            connectionString.Close();
        }

        private static int createCsvFile(Dictionary<string, Student> students)
        {
            var csv = new StringBuilder();
            var newLine = "";
            var filePath = @"d:/STUDENT_IMPORT.csv";

            foreach (KeyValuePair<string, Student> s in students)
            {
                string fullName = s.Value.LastName + " " + s.Value.FirstName;
                fullName = "" + fullName + "";
                newLine = string.Format("{0},{1},{2},{2},{3},{4},{5}", s.Value.EQID, fullName, s.Value.Year, s.Value.RollClass, 
                s.Value.Gender, s.Value.EmailAddress);
                csv.AppendLine(newLine);
            }

            // Check if the CSV file exists
            if (!File.Exists(filePath))
            {
                // Create the file.
                using (FileStream fs = File.Create(filePath)){ /* Add some file info in here if you want */ }
            }

            // Check if the file is open 
            try
            {
                File.WriteAllText(filePath, csv.ToString());
            }
            catch (IOException e)
            {
                MessageBox.Show("Please close the student.csv file and Card Exchange before running this.");
                return -1;
            }
            
            return 0;
        }
    }
}
