using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HR
{
    public partial class UpdateVersion : Form
    {
        public bool Update { get; set; }//האם נדרש לעדכן
        DbServiceSQL dbAdmins = new DbServiceSQL();
        public int EmpNum { get; set; }
        public string Password  { get; set; }
        public UpdateVersion()
        {
            InitializeComponent();
            Update = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int EmpNum;
            DataTable TableAdmins = new DataTable();
            bool number = Int32.TryParse(txt_NumEmpUp.Text, out EmpNum);
            this.EmpNum = EmpNum;
            Password = txt_Password.Text;
            string qry = $@"SELECT * 
                          FROM AccessUsers
                          WHERE UserNum={EmpNum} and UserPassword='{Password}' ";
           TableAdmins=dbAdmins.executeSelectQueryNoParam(qry);
            if(TableAdmins.Rows.Count==0)
            {
                MessageBox.Show("מנהל מערכת לא קיים");
            }
            else
            {
                qry = $@"UPDATE AccessUsers
                         SET Change=1
                         WHERE UserNum={EmpNum}";
                dbAdmins.ExecuteQuery(qry);
                Update = true;
            }
        }

        /// <summary>
        /// אחרי שסוגר תוכנית אצל כל המשתמשים מעדכן חזרה change=false
        /// </summary>
        public void AfterClose()
        {
            string qry = $@"UPDATE AccessUsers
                         SET Change=0
                         WHERE UserNum={EmpNum}";
            dbAdmins.ExecuteQuery(qry);
        }
    }
}
