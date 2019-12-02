using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;

namespace HR
{
    public partial class Form1 : Form
    {
        List<Employee> employees;
        DBService DBS = new DBService();
        DataTable dataTableActive = new DataTable();//טבלת עובדים פעילים
        DataTable dataTableAll = new DataTable();//טבלת כלל העובדים כולל לא פעילים
        DataTable dataTableForShow = new DataTable();//הטבלה שמשמשת להצגה בהתאם לפעולה
        bool TableIsDesigned = false;//האם הטבלה נעשתה כבר?אם כן לא מעצב ומוסיף לה עמודות  בפעם השניה
        bool FilterDate = false;//קליטה עבור סינון תאריך
        bool FilterLeft = false;//עבור סינון תאריך עזיבה-תקף רק בהצג הכל
        List<int> ListTotalIndex = new List<int>();//רשימה של איפה נמצא המילה total
        int Middle_Index = 0;//מיזוג ומרכוז עבור מצבת כ"א
        int tab;//בשביל לדעת איזה אקסל לפתוח
        bool finishInit = false;
        UpdateVersion updateVersion = new UpdateVersion();//עדכון גרסה -סגירת תוכנית בכל המחשבים
        DataView TableForFilterLastname;//טבלה עבור קומבו בוקס שמסננת שמות משפחה רלונטים בלבד
        public Form1()
        {
            LogWave("App start");
            InitializeComponent();
            ShowEmployees();
            StartScreen();
            tabControlll.Selecting += new TabControlCancelEventHandler(tabControl1_Selecting);
            chooseMonthLabel.Location = rangeFromLabel.Location;
            MonthsCheckBox.Location = dateTimePickerFrom.Location;
            List<string> temp = DateTimeFormatInfo.CurrentInfo.MonthNames.ToList();
            temp.RemoveAt(12);
            MonthsCheckBox.DataSource = temp;
            employees = Employee.GetListOfActiveEmployees();
            buildAbsenceReport();
            buildSabonReport();

            // sendMailToMissingWorkers();
            //  WorkFinishForms();
            FinishWorkComboBox.Text = "--בחר עובד--";
            FinishWorkComboBox.DataSource = employees;
            cbo_StartWork.DataSource = employees;//כנל גם עבור טופס תאריך תחילת עובד
            cbo_StopWork.DataSource = employees;
            cbo_Subscription.DataSource = employees;
            timer1.Start();// כל 10 דקות-בדיקה עבור שחרור גרסה
        }

        private void LogWave(string LogStr)
        {
            StreamWriter SW;
            string LogName = DateTime.Now.Date.ToString("yyyy-MM-dd") + " LogFile";

            if (!File.Exists(@"P:\\HR\" + LogName + ".txt"))
            {
                SW = File.CreateText(@"P:\\HR\" + LogName + ".txt");
                SW.Close();
            }

            SW = File.AppendText(@"P:\\HR\" + LogName + ".txt");
            SW.WriteLine(DateTime.Now.ToString() + " - " + LogStr);
            SW.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LogWave("form load ");
            initSelectTabs();
            LogWave("initSelectTabs end");
            //this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.Fixed3D; ;
            this.WindowState = FormWindowState.Maximized;
            dataGrid_workStrengh.AutoGenerateColumns = false;//קשור לאיחוד תאים במצבת כ"א
            LogWave("form load End");
            finishInit = true;
        }

        private void RadioButtonSelectionChanged(object sender, EventArgs e)
        {
            chooseMonthLabel.Visible = !chooseMonthLabel.Visible;
            MonthsCheckBox.Visible = !MonthsCheckBox.Visible;
            rangeFromLabel.Visible = !rangeFromLabel.Visible;
            RangeToLabel.Visible = !RangeToLabel.Visible;
            dateTimePickerFrom.Visible = !dateTimePickerFrom.Visible;
            dateTimePickerTo.Visible = !dateTimePickerTo.Visible;

            Employee employee = new Employee();
            DataTable Employee_Data = employee.GetData();
            StartScreen();
        }

        private void StartScreen()
        {
            cbo_search.Items.Add("מספר עובד");
            cbo_search.Items.Add("שם פרטי");
            cbo_search.Items.Add("שם מלא");
            cbo_search.Items.Add("ת.ז");
            cbo_search.Items.Add("מחלקה");
            cbo_search.Items.Add("שם משפחה");
        }


        private void ShowEmployees()
        {

            Employee employee = new Employee(date_start1.Value, date_start2.Value);
            bool ShowAllEmployees=false;//בודק אם להציג גם עובדים לא פעילים
            dataTableActive = employee.GetData(ShowAllEmployees, FilterDate,false);//הצגת כל העובדים?סינון תאריכי קליטה,סינון תאריכי עזיבה תמיד יהיה פולס בעובדים פעילים


            ShowAllEmployees = true;
            employee.Date_Left = date_left1.Value;//במידה שיש סינון תאריכי עזיבה
            employee.Date_Left2 = date_left2.Value;
            dataTableAll = employee.GetData(ShowAllEmployees, FilterDate,FilterLeft);//הצגת כל העובדים?סינון תאריכי קליטה,סינון תאריכי עזיבה


            FilterDate = false;
            FilterLeft = false;//החזרה למצב רגיל של הסינונים
            CheckWhichTableInsert(cb_ShowAll.Checked);//איזה טבלה להכניס לדטה גריד

        }

        ///// <summary>
        ///// מוסיף עמודות לדטה של כ"א
        ///// </summary>
        //private void AddColumnsDataTable(DataTable dataTable)
        //{
        //    //dataTable.Columns.Add(new DataColumn("סוג עובד", typeof(string)));
        //    //dataTable.Columns.Add(new DataColumn("ותק", typeof(string)));
        //    //dataTable.Columns.Add(new DataColumn("גיל", typeof(string)));
        //    //dataTable.Columns.Add(new DataColumn("סוג הסכם", typeof(string)));
        //}

        /// <summary>
        /// בודק איזה טבלה להכניס לדטה גריד
        /// </summary>
        private void CheckWhichTableInsert(bool ShowAllEmployees)
        {
            if (!ShowAllEmployees)
            {
                InsertDataTable(dataTableActive);
            }
            else
            {
                InsertDataTable(dataTableAll);
            }
        }

        /// <summary>
        /// מכניס טבלת כ"א לדטה גריד מסדר,מוסיף ומעצב אותה
        /// </summary>
        /// <param name="dataTable"></param>
        private void InsertDataTable(DataTable dataTable)
        {
            dataTableForShow = dataTable;//ההכנסה של הטבלה בפועל
            DataGridView_employees.DataSource = dataTableForShow;
            count_total2.Text = dataTableForShow.Rows.Count.ToString();
            Change_Column_Header();
            TableForFilterLastname = new DataView(dataTableForShow);//שייך לקומבובוקס של סינון שמות משפחה רלוונטיים-לראות באיוונט של עזיבת קומבובוקס 

            //for (int i = 0; i < dataTableForShow.Rows.Count; i++)
            //{

            //    if (dataTableForShow.Rows[i]["DateStart"].ToString() != "000000" && dataTableForShow.Rows[i]["DateStart"].ToString() != "")
            //    {
            //        var timeSpan = DateTime.Today - DateTime.Parse(dataTableForShow.Rows[i]["DateStart"].ToString());
            //        var years = timeSpan.Days / 365;
            //        var months = (timeSpan.Days - years * 365) / 30;
            //        ////var days = timeSpan.Days - years * 365 - months * 30;
            //        DataGridView_employees.Rows[i].Cells["ותק"].Value = years + "." + months;
            //    }
            //    if (dataTableForShow.Rows[i]["BIRTHDAY"].ToString() != "000000" && dataTableForShow.Rows[i]["BIRTHDAY"].ToString() != "")
            //    {
            //        var timeSpan = DateTime.Today - DateTime.Parse(dataTableForShow.Rows[i]["BIRTHDAY"].ToString());
            //        var years = timeSpan.Days / 365;
            //        DataGridView_employees.Rows[i].Cells["גיל"].Value = years;
            //    }

            //    if (dataTableForShow.Rows[i]["NUMBER"].ToString().Substring(0, 2) == "19")
            //        DataGridView_employees.Rows[i].Cells["סוג הסכם"].Value = "אישי";
            //    else DataGridView_employees.Rows[i].Cells["סוג הסכם"].Value = "קיבוצי";


            //switch (int.Parse(DataGridView_employees.Rows[i].Cells["TypeTime"].Value.ToString())) //תיאור סוג עובד כתיאור ולא כמספר
            //{
            //    case 1:
            //        DataGridView_employees.Rows[i].Cells["סוג עובד"].Value = "ישיר";
            //        break;

            //    case 2:
            //        DataGridView_employees.Rows[i].Cells["סוג עובד"].Value = "עקיף";
            //        break;

            //    case 3:
            //        DataGridView_employees.Rows[i].Cells["סוג עובד"].Value = "מנהל";
            //        break;

            //    case 4:
            //        DataGridView_employees.Rows[i].Cells["סוג עובד"].Value = "עקיף חרושת";
            //        break;

            //    case 9:
            //        DataGridView_employees.Rows[i].Cells["סוג עובד"].Value = "פנסיה";
            //        break;

            //    default:
            //        break;
            //}
            //}
            //this.DataGridView_employees.Columns["TypeTime"].Visible = false;
        }


        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            TabPage current = (sender as TabControl).SelectedTab;
            tab = e.TabPageIndex;//לדעת איזה קובץ אקסל לפתוח
            if (e.TabPage == tabpage_Employees)//לטפללל
            {   
                if(!TableIsDesigned)//לא יקרא לפונקציה יותר אחרי שהטבלה נעשתה פעם אחת
                WorkStrengh();
                TableIsDesigned = true;
            }

        }


        /// משנה כותרות עבור טב1 כ"א
        /// </summary>
        private void Change_Column_Header()
        {
            DataGridView_employees.Columns["NUMBER"].HeaderText = "מס עובד";
            DataGridView_employees.Columns["FirstName"].HeaderText = "שם פרטי";
            DataGridView_employees.Columns["LastName"].HeaderText = "שם משפחה";
            DataGridView_employees.Columns["Id"].HeaderText = "תעודת זהות";
            DataGridView_employees.Columns["ADRESS"].HeaderText = "כתובת";
            DataGridView_employees.Columns["DateStart"].HeaderText = "תאריך קליטה";
            DataGridView_employees.Columns["Unit"].HeaderText = "מחלקה";
            DataGridView_employees.Columns["UNITNAME"].HeaderText = "שם מחלקה";
            DataGridView_employees.Columns["BIRTHDAY"].HeaderText = "תאריך לידה";
            DataGridView_employees.Columns["EFirstName"].HeaderText = "First Name";
            DataGridView_employees.Columns["ELastName"].HeaderText = "Last Name";
            DataGridView_employees.Columns["TYPE"].HeaderText = "סוג עובד";
            DataGridView_employees.Columns["VETEK"].HeaderText = "ותק";
            DataGridView_employees.Columns["temp"].HeaderText = "ארעיות";
            DataGridView_employees.Columns["TYPECONTRACT"].HeaderText = "סוג הסכם";
            DataGridView_employees.Columns["AGE"].HeaderText = "גיל";
            DataGridView_employees.Columns["MonthStart"].HeaderText = "חודש קליטה";
            DataGridView_employees.Columns["TELEFON"].HeaderText = "טלפון";

            
            if (cb_ShowAll.Checked)
                DataGridView_employees.Columns["Left"].HeaderText = "תאריך עזיבה";
        }

        /// משנה כותרות עבור מצבת כ"א
        /// </summary>
        private void Change_Column_Header2()
        {
            dataGrid_workStrengh.Columns["AGAF"].HeaderText = "שם אגף";
            dataGrid_workStrengh.Columns["UNIT"].HeaderText = "מספר מחלקה";
            dataGrid_workStrengh.Columns["COUNT"].HeaderText = "סך הכל";
            dataGrid_workStrengh.Columns["UNITNAME"].HeaderText = "שם מחלקה";
        }

        /// ממלא את הקומבו בוקס למטרת חיפוש
        /// </summary>
        private void Search_Func()
        {
            try
            {
                if (cb_ShowAll.Checked == false)
                {
                    cbo_search_result.DataSource = dataTableActive;
                    cbo_lastname.DataSource = TableForFilterLastname;
                }
                else
                {
                    cbo_search_result.DataSource = dataTableAll;
                    cbo_lastname.DataSource = TableForFilterLastname;
                }

                if (cbo_search.SelectedIndex == 0) { cbo_search_result.DisplayMember = "NUMBER"; cbo_lastname.Visible = false; }
                else if (cbo_search.SelectedIndex == 1) { cbo_search_result.DisplayMember = "FirstName"; }
                else if (cbo_search.SelectedIndex == 2) { cbo_lastname.Visible = true; cbo_search_result.DisplayMember = "FirstName"; cbo_lastname.DisplayMember = "LastName"; }
                else if (cbo_search.SelectedIndex == 3) { cbo_search_result.DisplayMember = "ID"; cbo_lastname.Visible = false; }
                else if (cbo_search.SelectedIndex == 4)
                {
                    //cbo_search_result.DisplayMember = "UNIT";
                    cbo_lastname.Visible = false;
                    var distinctRows = (from DataRow dRow in dataTableForShow.Rows
                                        select dRow["UNIT"]).Distinct();

                    cbo_search_result.DataSource = null;
                    foreach (object item in distinctRows)
                    {
                        if ((!cbo_search_result.Items.Cast<string>().Contains(item)) && (!(item is null)))
                        {
                            cbo_search_result.Items.Add(item);
                        }
                    }

                    //cbo_search_result.Items.Add ( distinctRows);
                }
                else if (cbo_search.SelectedIndex == 5) { cbo_search_result.DisplayMember = "LastName"; cbo_lastname.Visible = false; }
            }
            catch(Exception ex)
            {

            }

        }


        private void initSelectTabs()
        {
            foreach (TabPage p in tabControlll.TabPages)
            {
                tabControlll.SelectedTab = p;
            }
            tabControlll.SelectedTab = strengthTab;
            tabControlll.SelectedTab = tabpage_Employees;
            

        }

        /// <summary>
        /// הצג את כל העובדים
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cb_ShowAll_CheckedChanged(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;//כפתור המתנה
            CheckWhichTableInsert(cb_ShowAll.Checked);//בודק איזה טבלה להכניס בהתאם למה שמסומן
            Cursor.Current = Cursors.Default;
            if(cb_ShowAll.Checked)
            {
                btn_show_left.Visible = true;
                l1.Visible = true;
                l2.Visible = true;
                date_left1.Visible = true;
                date_left2.Visible = true;
            }
            else
            {
                btn_show_left.Visible = false;
                l1.Visible = false;
                l2.Visible = false;
                date_left1.Visible = false;
                date_left2.Visible = false;
            }
        }

        private void cbo_search_SelectedIndexChanged(object sender, EventArgs e)
        {
            Search_Func();
            btn_search.Enabled = true;
            lbl_ChooseBySearch.Visible = true;
            switch(cbo_search.SelectedIndex)
            {
                case 0:
                    lbl_ChooseBySearch.Text = "מספר עובד";
                    break;

                case 1:
                    lbl_ChooseBySearch.Text = "שם פרטי";                  
                    break;

                case 2:
                    lbl_ChooseBySearch.Text = "שם פרטי";
                    lbl_LastName.Text = "שם משפחה";
                    break;

                case 3:
                    lbl_ChooseBySearch.Text = "תעודת זהות";
                    break;

                case 4:
                    lbl_ChooseBySearch.Text = "מספר מחלקה";
                    break;

                case 5:
                    lbl_ChooseBySearch.Text = "שם משפחה";
                    break;

            }

            if (cbo_search.SelectedIndex == 2)
                lbl_LastName.Visible = true;
            else
                lbl_LastName.Visible = false;
        }

        /// כפתור חיפוש לפי שדה ספציפי
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_search_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dataSearch = new DataTable();//טבלה עבור תוצאות חיפוש
                Employee Employee_Result = new Employee();
                dataSearch = Employee_Result.TableAfterSearch(cb_ShowAll.Checked, cbo_search.SelectedIndex, cbo_search_result.Text, cbo_lastname.Text); //שולח האם להציג את כולם,שדה לפיו מחפשים,והcbshowAll עצמו בודק אם להציג גם עובדים לא פעילים
                InsertDataTable(dataSearch);
            }
            

            catch
            {
                MessageBox.Show("לא הקלדת ערך קיים");                
            }

            lbl_LastName.Visible = false;

        }

        private void button1_Click_1(object sender, EventArgs e)///כפתור ביטול חיפוש
        {
            Cursor.Current = Cursors.WaitCursor;//כפתור המתנה
            ShowEmployees();
            Cursor.Current = Cursors.Default;
        }

        private void DataGridView_employees_SortStringChanged(object sender, EventArgs e)
        {
            dataTableForShow.DefaultView.Sort = this.DataGridView_employees.SortString;
            count_total2.Text = (DataGridView_employees.Rows.Count - 1).ToString();
        }

        /// הצגת מצבת כ"א
        /// </summary>
        private void WorkStrengh()
        {
            Employee employee = new Employee();
            try
            {
                int last_column = 7;//עבור מיקום עמודת ספירה סופית
                DataTable MaindataTable = employee.GetWorkStrengh();//מקבל טבלת מצבת כ"א        
                MaindataTable.Columns.Add("עובד ישיר", typeof(int));
                MaindataTable.Columns.Add("עובד עקיף", typeof(int));
                MaindataTable.Columns.Add("מנהל", typeof(int));
                MaindataTable.Columns.Add("עקיף חרושת", typeof(int));



                ///חישובים עבור חישובי ביניים של עובדים ישירים/עקיפים/מנהלים 
                DataTable TableForEmployeeType = employee.GetSubTotalofEmployeeType();
                string department = TableForEmployeeType.Rows[0]["UNIT"].ToString();//  מקבל שם מחלקה מטבלת החישובים 
                int row_index = 0;//מיקום השורה בטבלה המקורית בה יושמו המספרים של עיקרי וישיר
                for (int k = 0; k < TableForEmployeeType.Rows.Count; k++)//טבלה מקורית מול טבלת החישובים
                {
                    if (!(MaindataTable.Rows[row_index]["UNIT"].ToString() == department))//עובר על רשומה בטבלה הראשית ובודק את המחלקה עם סוגי העובדים שלה, אם אין יותר עובר לשורה הבאה
                    {
                        row_index++;
                    }


                    switch (int.Parse(TableForEmployeeType.Rows[k]["TYPE"].ToString()))///הכנסה לשורה הרלוונטית
                    {
                        case 1:
                            MaindataTable.Rows[row_index]["עובד ישיר"] = TableForEmployeeType.Rows[k]["COUNT"].ToString();
                            break;

                        case 2:
                            MaindataTable.Rows[row_index]["עובד עקיף"] = TableForEmployeeType.Rows[k]["COUNT"].ToString();
                            break;

                        case 3:
                            MaindataTable.Rows[row_index]["מנהל"] = TableForEmployeeType.Rows[k]["COUNT"].ToString();
                            break;

                        case 4:
                            MaindataTable.Rows[row_index]["עקיף חרושת"] = TableForEmployeeType.Rows[k]["COUNT"].ToString();
                            break;

                        default:
                            break;


                    }
                    if (k + 1 != TableForEmployeeType.Rows.Count)
                        department = TableForEmployeeType.Rows[k + 1]["UNIT"].ToString();

                }

                ///יצירת הטבלה לפי אגפים עם סיכומי ביניים
                string division = MaindataTable.Rows[0]["AGAF"].ToString();//מקבל שם אגף
                int grand_total = 0, total_direct = 0, total_indirect = 0, total_manager = 0, total_indirect_indust = 0;

                ///עובר על כל השורות וסוכם כל טור לפי הסוג שלו
                for (int i = 0; i <= MaindataTable.Rows.Count; i++)//כל סיום שם אגף עושה סב טוטל
                {
                    if (i != MaindataTable.Rows.Count)
                    {
                        if (MaindataTable.Rows[i]["AGAF"].ToString() == division)
                        {
                            if (!string.IsNullOrEmpty(MaindataTable.Rows[i]["עובד ישיר"].ToString() as string)) total_direct += int.Parse(MaindataTable.Rows[i]["עובד ישיר"].ToString());
                            if (!string.IsNullOrEmpty(MaindataTable.Rows[i]["עובד עקיף"].ToString() as string)) total_indirect += int.Parse(MaindataTable.Rows[i]["עובד עקיף"].ToString());
                            if (!string.IsNullOrEmpty(MaindataTable.Rows[i]["מנהל"].ToString() as string)) total_manager += int.Parse(MaindataTable.Rows[i]["מנהל"].ToString());
                            if (!string.IsNullOrEmpty(MaindataTable.Rows[i]["עקיף חרושת"].ToString() as string)) total_indirect_indust += int.Parse(MaindataTable.Rows[i]["עקיף חרושת"].ToString());
                            grand_total += int.Parse(MaindataTable.Rows[i]["COUNT"].ToString());
                        }
                        else ///אם סיימנו לעבור על שם האגף ניצור שורה חדשה עם הסב טוטל שלו
                        {
                            CreateRow(MaindataTable, division, total_direct, total_indirect, total_manager, total_indirect_indust, grand_total, i);//יצירת שורה חדשה                            
                            division = MaindataTable.Rows[i + 1]["AGAF"].ToString();
                            grand_total = 0;
                            total_direct = 0;
                            total_indirect = 0;
                            total_manager = 0;
                            total_indirect_indust = 0;
                        }
                    }
                    else///שייך לסב טקסט האחרון,נועד בשביל לא לצאת מגבולות הטבלה
                    {
                        CreateRow(MaindataTable, division, total_direct, total_indirect, total_manager, total_indirect_indust, grand_total, i);//יצירת שורה חדשה                            
                        if (i + 1 < MaindataTable.Rows.Count)  ///כשהגענו לרשומה האחרונה בשביל לא ליצור שגיאה ולצאת מגבולות הטבלה
                            division = MaindataTable.Rows[i + 1]["AGAF"].ToString();
                        else
                        {
                            grand_total = 0;
                            total_direct = 0;
                            total_indirect = 0;
                            total_manager = 0;
                            total_indirect_indust = 0;
                            break;
                        }

                    }

                }

                //הכנסה לדטה גריד
                dataGrid_workStrengh.DataSource = MaindataTable;

                ///תצוגה
                dataGrid_workStrengh.Columns["AGAF"].Width = 150;
                dataGrid_workStrengh.Columns["UNITNAME"].Width = 150;
                dataGrid_workStrengh.Columns["UNIT"].Width = 80;
                dataGrid_workStrengh.Columns["עובד ישיר"].Width = 50;
                dataGrid_workStrengh.Columns["עובד עקיף"].Width = 50;
                dataGrid_workStrengh.Columns["מנהל"].Width = 50;
                dataGrid_workStrengh.Columns["עקיף חרושת"].Width = 50;
                dataGrid_workStrengh.Columns["COUNT"].Width = 100;
                dataGrid_workStrengh.Columns["COUNT"].DisplayIndex = last_column;

                //מבחין באיזה שורות קיימת המילה טוטל עבור עיצוב שורה
                foreach (DataGridViewRow rows in dataGrid_workStrengh.Rows)
                {
                    if (!string.IsNullOrEmpty(rows.Cells[0].Value as string))
                    {
                        int length = rows.Cells[0].Value.ToString().Length;
                        if (length >= 5)
                        {
                            if (rows.Cells[0].Value.ToString().Substring(0, 5) == "Total")
                            {
                                ListTotalIndex.Add(rows.Index);//מוסיף לרשימה את מספר השורה שנמצא טוטל
                            }
                            length = 0;
                        }
                    }
                }

                CalculateAllTotal(MaindataTable, total_direct, total_indirect, total_manager, total_indirect_indust, grand_total);//מחשב טוטל סופי
                MergeAndCenter();//שם במרכז את שם האגף
                Change_Column_Header2();//כותרות לדטה גריד
                CountPensionersAndNoPayment();//ספירת פנסיונרים וללא תשלום
            }



            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void CountPensionersAndNoPayment()
        {
            Employee employee = new Employee();
            int Counter_Pensioner = employee.CountPensioner();
            lbl_countOldEmployees.Text = Counter_Pensioner.ToString();
            int Counter_Nonpayment = employee.CountNonPayment();
            lbl_nonPayCount.Text = Counter_Nonpayment.ToString();
            lbl_nonPayCount.Text = Counter_Nonpayment.ToString();
        }

        /// create row for subtotal 
        /// </summary>
        private void CreateRow(DataTable MaindataTable, string division, int total_direct, int total_indirect, int total_manager, int total_indirect_indust, int grand_total, int i)
        {
            DataRow row = MaindataTable.NewRow();
            row[0] = "Total " + division;
            row["עובד ישיר"] = total_direct;
            row["עובד עקיף"] = total_indirect;
            row["מנהל"] = total_manager;
            row["עקיף חרושת"] = total_indirect_indust;
            row["COUNT"] = grand_total;
            MaindataTable.Rows.InsertAt(row, i);
        }
        /// חישוב כולל של כל הטבלה
        /// </summary>
        private void CalculateAllTotal(DataTable MaindataTable, int total_direct, int total_indirect, int total_manager, int total_indirect_indust, int grand_total)
        {
            DataRow row = MaindataTable.NewRow();
            for (int i = 0; i < ListTotalIndex.Count; i++)
            {
                int index = ListTotalIndex[i];
                total_direct += int.Parse(dataGrid_workStrengh.Rows[index].Cells["עובד ישיר"].Value.ToString());
                total_indirect += int.Parse(dataGrid_workStrengh.Rows[index].Cells["עובד עקיף"].Value.ToString());
                total_indirect_indust += int.Parse(dataGrid_workStrengh.Rows[index].Cells["עקיף חרושת"].Value.ToString());
                total_manager += int.Parse(dataGrid_workStrengh.Rows[index].Cells["מנהל"].Value.ToString());
                grand_total += int.Parse(dataGrid_workStrengh.Rows[index].Cells["COUNT"].Value.ToString());
            }
            string division = "All";
            CreateRow(MaindataTable, division, total_direct, total_indirect, total_manager, total_indirect_indust, grand_total, MaindataTable.Rows.Count);
            total_direct = 0; total_indirect = 0; total_indirect_indust = 0; total_manager = 0; grand_total = 0;
        }
        /// בודק אם לאחד בין התאים-מכילים אותו ערך
        /// </summary>
        bool IsTheSameCellValue(int column, int row, string PreviousValue)
        {
            DataGridViewCell cell1 = dataGrid_workStrengh[column, row];
            //DataGridViewCell cell2 = dataGrid_workStrengh[column, row - 1];
            if (cell1.Value == null || PreviousValue == null)
            {
                return false;
            }
            return cell1.Value.ToString() == PreviousValue.ToString();
        }

        /// מוריד גבולות עבור כותרות האגפים
        /// </summary>

        private void dataGrid_workStrengh_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            ///הורדת גבולות מכל הטבלה
            e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
            e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;

            ///אם מופיע שורה עם המילה טוטל תוחם גבולות ומעצב תאים
            if (ListTotalIndex.Contains(e.RowIndex))
            {
                e.AdvancedBorderStyle.Top = dataGrid_workStrengh.AdvancedCellBorderStyle.Top;
                e.AdvancedBorderStyle.Bottom = dataGrid_workStrengh.AdvancedCellBorderStyle.Bottom;
                e.CellStyle.Font = new Font("Tahoma", 12, FontStyle.Bold);
                e.CellStyle.BackColor = Color.DeepSkyBlue;
            }


            //עיצוב שורה אחרונה של טוטל כללי
            if (e.RowIndex == dataGrid_workStrengh.Rows.Count - 1)
            {
                e.AdvancedBorderStyle.Top = dataGrid_workStrengh.AdvancedCellBorderStyle.Top;
                e.AdvancedBorderStyle.Bottom = dataGrid_workStrengh.AdvancedCellBorderStyle.Bottom;
                e.CellStyle.Font = new Font("Tahoma", 20, FontStyle.Bold);
                e.CellStyle.BackColor = Color.Red;
                dataGrid_workStrengh.Rows[e.RowIndex].Height = 42;
            }



            //את טור שם האגף לא סוגר בגבולות
            if (e.RowIndex < 0 || e.ColumnIndex == 0)
                return;

            ///סגירה בגבולות של כל השאר
            else
            {
                if (!ListTotalIndex.Contains(e.RowIndex))
                {
                    e.AdvancedBorderStyle.Top = dataGrid_workStrengh.AdvancedCellBorderStyle.Top;
                    e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.Inset;
                    e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.Single;
                }

            }

        }


        ///ממרכז את שם האגף ע"י מחיקת שאר השורות
        private void MergeAndCenter()
        {
            try
            {
                for (int i = 1; i < dataGrid_workStrengh.Rows.Count - 1; i++)
                {

                    ///אם מדובר בשורה שרשום בה טוטל המילה לא תימחק
                    bool RowWithTotal = false;
                    for (int j = 0; j < ListTotalIndex.Count; j++)
                    {
                        if (i == ListTotalIndex[j]) RowWithTotal = true;
                    }


                    if (!RowWithTotal) ///יכנס רק אם לא מדובר במילה טוטל
                    {
                        string SavePreviousValue = dataGrid_workStrengh.Rows[i].Cells[0].Value.ToString();//שומר ערך קודם למטרת השוואה מול תא נוכחי
                        if (IsTheSameCellValue(0, i, SavePreviousValue))
                        {
                            if (Middle_Index <= ListTotalIndex.Count)
                            {
                                if (Middle_Index == ListTotalIndex.Count) Middle_Index = 1000;///מחיקה במידה ומדובר ברגף אחרון

                                switch (Middle_Index)
                                {
                                    case 0:
                                        if (i != (ListTotalIndex[Middle_Index] / 2) - 1)
                                        {
                                            dataGrid_workStrengh.Rows[i].Cells[0].Value = "";
                                        }
                                        else Middle_Index++;
                                        break;

                                    case 1000:
                                        if (i != (ListTotalIndex[ListTotalIndex.Count - 1] / 2) - 1)
                                        {
                                            dataGrid_workStrengh.Rows[i].Cells[0].Value = "";
                                        }
                                        break;

                                    default:
                                        int num, sub;
                                        sub = (ListTotalIndex[Middle_Index] - ListTotalIndex[Middle_Index - 1]) / 2;
                                        num = ListTotalIndex[Middle_Index] - sub;
                                        if (i != num)
                                        {
                                            dataGrid_workStrengh.Rows[i].Cells[0].Value = "";
                                        }
                                        else Middle_Index++;
                                        break;
                                }
                            }

                        }
                    }
                }
                dataGrid_workStrengh.Rows[0].Cells[0].Value = "";
                Middle_Index = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        /// כפתור הדפסת מצבת כ"א
        /// </summary>
        Bitmap bmp;
        private void button3_Click(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter();
            printer.Title = "דוח מצבת כא";//Header
            printer.SubTitle = string.Format("Date: {0}", DateTime.Now.Date.ToString("dd/MM/yyyy"));
            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
            printer.PageNumbers = true;
            printer.PageNumberInHeader = false;
            printer.PorportionalColumns = true;
            printer.HeaderCellAlignment = StringAlignment.Near;
            printer.Footer = "HR";//Footer
            printer.FooterSpacing = 15;
            //Print landscape mode
            printer.printDocument.DefaultPageSettings.Landscape = true;

            dataGrid_workStrengh.Size = new Size(650, 180);
            //printPreviewDialog1.ShowDialog();


            //int height = dataGrid_workStrengh.Height;
            //dataGrid_workStrengh.Height = dataGrid_workStrengh.RowCount * dataGrid_workStrengh.RowTemplate.Height * 2;
            //dataGrid_workStrengh.Width = dataGrid_workStrengh.Width / 2;
            //bmp = new Bitmap(dataGrid_workStrengh.Width, dataGrid_workStrengh.Height);
            //dataGrid_workStrengh.DrawToBitmap(bmp, new Rectangle(0, 0, dataGrid_workStrengh.Width, dataGrid_workStrengh.Height));
            //dataGrid_workStrengh.Height = height;
            //printPreviewDialog1.ShowDialog();

            printer.PrintDataGridView(dataGrid_workStrengh);

            dataGrid_workStrengh.Size = new Size(1615, 656);
            dataGrid_workStrengh.Anchor = (AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom | AnchorStyles.Top);
        }

        /// <summary>
        /// כפתור הצג סינון תאריכי קליטה
        /// </summary>
        private void btn_show_Click(object sender, EventArgs e)
        {
            try
            {
                FilterDate = true;
                ShowEmployees();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// כפתור סינון תאריכי עזיבה
        /// </summary>
        private void btn_show_left_Click(object sender, EventArgs e)
        {
            try
            {
                FilterLeft = true;
                ShowEmployees();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void showReportBtn_Click(object sender, EventArgs e)
        {

            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn { ColumnName = "EmpNum" });
            dt.Columns.Add(new DataColumn { ColumnName = "FirstName" });
            dt.Columns.Add(new DataColumn { ColumnName = "LastName" });
            dt.Columns.Add(new DataColumn { ColumnName = "DateOfBirth" });
            dt.Columns.Add(new DataColumn { ColumnName = "Unit" });
            if (dateRangeRB.Checked)
            {
                //סינון לפי טווח
                buildBdayReport(dateTimePickerFrom.Value, dateTimePickerTo.Value);
            }
            else
            {
                //סינון לפי חודש
                int month = DateTimeFormatInfo.CurrentInfo.MonthNames.ToList().IndexOf(MonthsCheckBox.Text) + 1;
                DateTime from = new DateTime(DateTime.Now.Year, month, 1);
                DateTime to = from.AddMonths(1).AddDays(-1);
                buildBdayReport(from, to);
            }

        }

        private void DataGridView_employees_FilterStringChanged(object sender, EventArgs e)
        {
            dataTableForShow.DefaultView.RowFilter = this.DataGridView_employees.FilterString;
            count_total2.Text = (DataGridView_employees.Rows.Count ).ToString();

        }

        private void buildBdayReport(DateTime from, DateTime to)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn { ColumnName = "EmpNum" });
            dt.Columns.Add(new DataColumn { ColumnName = "Name" });
            //dt.Columns.Add(new DataColumn { ColumnName = "LastName" });
            dt.Columns.Add(new DataColumn { ColumnName = "DateOfBirth" });
            dt.Columns.Add(new DataColumn { ColumnName = "DateOfBirthNoYear" });
            dt.Columns.Add(new DataColumn { ColumnName = "Unit" });

            var x = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day).CompareTo(new DateTime( DateTime.Now.Year + 1, DateTime.Now.Month, DateTime.Now.Day));

            //employees.Sort((a, b) => a.Employee_Birthday.CompareTo(b.Employee_Birthday));
            //employees.OrderBy(a => a.Employee_Birthday.Year);


          //  employees.Sort((a, b) => new DateTime(DateTime.Now.Year, a.Employee_Birthday.Month, a.Employee_Birthday.Day).CompareTo(new DateTime((a.Employee_Birthday.Month < b.Employee_Birthday.Month ? DateTime.Now.Year : DateTime.Now.Year + 1), b.Employee_Birthday.Month, b.Employee_Birthday.Day)));
            List<Employee> problematicEmp = new List<Employee>();
            foreach(Employee emp in employees)
            {
                if (emp.Employee_Birthday == DateTime.MinValue)
                    problematicEmp.Add(emp);
            }
            //   employees.RemoveAll((a) => problematicEmp.Contains(a));
            //employees.Sort((a,b)=>a.CompareTo(b));
            if (from.Month > to.Month)
                employees = sortWhenToGreaterThanFrom(from,to);
            else {
                employees.Sort((a, b) => new DateTime(DateTime.Now.Year, a.Employee_Birthday.Month, a.Employee_Birthday.Day).CompareTo(new DateTime(DateTime.Now.Year, b.Employee_Birthday.Month, b.Employee_Birthday.Day)));
            }
            


            foreach (Employee emp in employees)
            {
                if (!emp.checkIfHaveBdayInRangeOfMonths(from, to) || emp.Stop_Work.Year<5000)
                    continue;
                if (emp.Employee_Birthday == DateTime.MinValue) continue;
                var row = dt.NewRow();
                row["EmpNum"] = emp.Employee_Num;
                row["Name"] = emp.First_Name+ " " + emp.Last_Name;
                //row["LastName"] = emp.Last_Name;
                row["DateOfBirth"] = emp.Employee_Birthday.ToShortDateString();
                row["DateOfBirthNoYear"] = emp.Employee_Birthday.ToString("dd/MM");
                row["Unit"] = emp.Unit;
                dt.Rows.Add(row);
            }
            List<System.Data.DataTable> result = dt.AsEnumerable()
            .GroupBy(row => DateTime.ParseExact(row.Field<string>("DateOfBirth"), "dd/MM/yyyy", CultureInfo.CurrentCulture).Month)
            .Select(g => g.CopyToDataTable())
            .ToList();
            BdayGridView.DataSource = dt;
            bdayCountLbl.Text = "סה\"כ "+dt.Rows.Count;

            StyleBdayDGV();

        }

        private List<Employee> sortWhenToGreaterThanFrom(DateTime from, DateTime to)
        {
            List<Employee> Jan_To_To = new List<Employee>();
            List<Employee> From_To_Dec = new List<Employee>();
            foreach(Employee e in employees)
            {
                if (!e.checkIfHaveBdayInRangeOfMonths(from, to))
                    continue;
                if (new DateTime(DateTime.Now.Year, e.Employee_Birthday.Month, e.Employee_Birthday.Day).CompareTo(from) > 0)
                    From_To_Dec.Add(e);
                else Jan_To_To.Add(e);
            }
            Jan_To_To.Sort((a, b) => new DateTime(DateTime.Now.Year, a.Employee_Birthday.Month, a.Employee_Birthday.Day).CompareTo(new DateTime(DateTime.Now.Year, b.Employee_Birthday.Month, b.Employee_Birthday.Day)));
            From_To_Dec.Sort((a, b) => new DateTime(DateTime.Now.Year, a.Employee_Birthday.Month, a.Employee_Birthday.Day).CompareTo(new DateTime(DateTime.Now.Year, b.Employee_Birthday.Month, b.Employee_Birthday.Day)));
            From_To_Dec.AddRange(Jan_To_To);
            return From_To_Dec;

        }


        private void makeStrengthReport()
        {
            Cursor.Current = Cursors.WaitCursor;

            copyAlltoClipboard();
            Excel.Application xlexcel;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[3, 1];//טווח מילוי הטבלה
            Excel.Range chartRange, divisionRange, agafRange;//קביעת רוחב עמודות
            chartRange = xlWorkSheet.get_Range("A1", "M1");
            divisionRange = xlWorkSheet.get_Range("F1");
            agafRange = xlWorkSheet.get_Range("H1");
            chartRange.ColumnWidth = 15;
            divisionRange.ColumnWidth = 35;
            agafRange.ColumnWidth = 30;


            //כותרות ועיצוב כללי
            //xlWorkSheet.get_Range("A1", "H1").Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            //xlWorkSheet.get_Range("A1", "H1").Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            xlWorkSheet.Cells[1, 1] = "מצבת כא מעודכן לתאריך " + DateTime.Today.ToShortDateString();
            xlWorkSheet.Cells[2, 1] = "סך הכל";
            xlWorkSheet.Cells[2, 2] = "עקיף חרושת";
            xlWorkSheet.Cells[2, 3] = "מנהל";
            xlWorkSheet.Cells[2, 4] = "עובד עקיף";
            xlWorkSheet.Cells[2, 5] = "עובד ישיר";
            xlWorkSheet.Cells[2, 6] = "שם מחלקה";
            xlWorkSheet.Cells[2, 7] = "מספר מחלקה";
            xlWorkSheet.Cells[2, 8] = "שם אגף";
            xlWorkSheet.get_Range("A1:H1").Merge();//מיזוג תאים
            xlWorkSheet.get_Range("A1:H1").Font.Bold = true;
            xlWorkSheet.get_Range("A1:H1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            xlWorkSheet.get_Range("A2:H2").RowHeight = 30;
            xlWorkSheet.get_Range("A1:H1").RowHeight = 50;
            xlWorkSheet.get_Range("A1:H1").Font.Size = 32;
            Excel.Range workSheet_range = xlWorkSheet.get_Range("A:M");
            workSheet_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            workSheet_range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            xlWorkSheet.get_Range("H:H").Font.Bold = true;
            xlWorkSheet.get_Range("A2:H2").Font.Bold = true;
            xlWorkSheet.get_Range("A2:H2").Font.Size = 16;

        



            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

            Excel.Range UsedRange = xlWorkSheet.UsedRange;
            int lastUsedRow = UsedRange.Row + UsedRange.Rows.Count - 2;//שורה אחרונה עם טקסט באקסל
            chartRange = xlWorkSheet.get_Range("A1:H" + lastUsedRow);//עיצוב תאים
            chartRange.Borders.Color = System.Drawing.Color.Black.ToArgb();

            //מחיקת שורות ריקות
            for (int j = 2; j < lastUsedRow; j++)//לא כולל שורה אחרונה
            {
                if (xlWorkSheet.Cells[j, "A"].Value2 ==null)
                {
                    xlWorkSheet.Rows[j].Delete();
                }
            }
            lastUsedRow = UsedRange.Row + UsedRange.Rows.Count -1;//עדכון מספר רשומות בלי שורות מיותרות והפחות 1 בשביל שלא ימחק את הסיכום הכללי

            ///מזג ומרכז
            int i = 1;
            int FirstCellMerge = 3;//עבור מיזוג שמות אגפים
            for (i = 1; i < lastUsedRow-1; i++)
            {
                string rowCell = "";
                if (!string.IsNullOrEmpty(xlWorkSheet.Cells[i, 8].Value2 as string))
                {

                    rowCell = Convert.ToString(xlWorkSheet.Cells[i, 8].Value2);
                    if (rowCell.Substring(0, 3) == "Tot")//אם מתחיל בטוטל מעצב שורה
                    {
                        Excel.Range usedRange1 = xlWorkSheet.get_Range("A" + i + ":H" + i);
                        usedRange1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        usedRange1.RowHeight = 27;
                        usedRange1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble;
                        usedRange1.Font.Bold = true;
                        xlWorkSheet.get_Range("H" + FirstCellMerge + ":H" + (i - 1)).Merge();
                        FirstCellMerge = i + 1;
                    }
                }

            }
            xlWorkSheet.get_Range("A" + lastUsedRow + ":H" + lastUsedRow).Font.Bold = true;
            xlWorkSheet.get_Range("A" + lastUsedRow + ":H" + lastUsedRow).Font.Size = 20;
            xlWorkSheet.get_Range("A" + lastUsedRow + ":H" + lastUsedRow).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumVioletRed);

            xlexcel.Visible = true;
            Cursor.Current = Cursors.Default;
            //xlWorkBook.Close();
            releaseObject(xlexcel);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);


        }

        private void makeGeneralEmployeeReport()
        {
            Cursor.Current = Cursors.WaitCursor;//כפתור המתנה
            if (cb_ShowAll.Checked) makeGeneralEmployeeReportAll();//אם אלי רוצה דו"ח אקסל של הכל
            else
            {
                copyAlltoClipboard();//מעתיק את הדטה גריד
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                //כותרת
                Excel.Range First = xlWorkSheet.get_Range("A1:q1");
                First.Merge();//מיזוג תאים
                First.Font.Bold = true;
                First.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.PaleTurquoise);
                First.RowHeight = 50;
                First.Font.Size = 32;
                xlWorkSheet.get_Range("A2:q2").Font.Bold = true;
                xlWorkSheet.get_Range("A2:q2").Font.Size = 16;
                //מרכוז
                Excel.Range workSheet_range = xlWorkSheet.get_Range("A:q");
                workSheet_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
                workSheet_range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
                workSheet_range.ColumnWidth = 12;


                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[3, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                xlWorkSheet.Cells[1, 1] = "כוח אדם-אליאנס";
                xlWorkSheet.Cells[2, 17] = "מספר עובד";
                xlWorkSheet.Cells[2, 16] = "שם פרטי";
                xlWorkSheet.Cells[2, 15] = "שם משפחה";
                xlWorkSheet.Cells[2, 14] = "ת.ז";
                xlWorkSheet.Cells[2, 13] = "כתובת";
                xlWorkSheet.Cells[2, 12] = "תאריך קליטה";
                xlWorkSheet.Cells[2, 11] = "חודש קליטה";
                xlWorkSheet.Cells[2, 10] = "טלפון";
                xlWorkSheet.Cells[2, 9] = "מחלקה";
                xlWorkSheet.Cells[2, 8] = "שם מחלקה";
                xlWorkSheet.Cells[2, 7] = "תאריך לידה";
                xlWorkSheet.Cells[2, 6] = "גיל";
                xlWorkSheet.Cells[2, 5] = "סוג עובד";
                xlWorkSheet.Cells[2, 4] = "First Name";
                xlWorkSheet.Cells[2, 3] = "Last Name";
                xlWorkSheet.Cells[2, 2] = "ותק";
                xlWorkSheet.Cells[2, 1] = "סוג הסכם";

                //גבולות תא
                Excel.Range UsedRange, chartRange;
                UsedRange = xlWorkSheet.UsedRange;
                int lastUsedRow = UsedRange.Row + UsedRange.Rows.Count - 1;//שורה אחרונה עם טקסט באקסל
                chartRange = xlWorkSheet.get_Range("A1:q" + lastUsedRow);//עיצוב תאים
                chartRange.Borders.Color = System.Drawing.Color.Black.ToArgb();


                //מחיקת שורות ריקות
                for (int i = 2; i < lastUsedRow; i++)
                {
                    if (string.IsNullOrEmpty(xlWorkSheet.Cells[i, "P"].Value2 as string))//לפי שם פרטי
                    {
                        //if (string.IsNullOrEmpty(xlWorkSheet.Cells[i + 1, "P"].Value2 as string))
                            xlWorkSheet.Rows[i].Delete();

                    }

                }
                Excel.Range firstRow = (Excel.Range)xlWorkSheet.Rows[2];
                firstRow.Application.ActiveWindow.FreezePanes = true;
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                //xlWorkSheet.get_Range("A3:q3").AutoFilter();
                dataGrid_workStrengh.CurrentCell = dataGrid_workStrengh.Rows[2].Cells[3];
                dataGrid_workStrengh.Focus();
                dataGrid_workStrengh.BeginEdit(true);
                xlexcel.Visible = true;
                //xlWorkBook.Close();
                releaseObject(xlexcel);
                releaseObject(xlWorkBook);
                releaseObject(xlWorkSheet);
            }
            Cursor.Current = Cursors.Default;

        }


        /// <summary>
        /// ייצוא לאקסל של כל העובדים-גם הלא פעילים
        /// </summary>
        private void makeGeneralEmployeeReportAll()
        {
            copyAlltoClipboard();//מעתיק את הדטה גריד
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Excel.Application();
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //כותרת
            Excel.Range First = xlWorkSheet.get_Range("A1:R1");
            First.Merge();//מיזוג תאים
            First.Font.Bold = true;
            First.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.PaleTurquoise);
            First.RowHeight = 50;
            First.Font.Size = 32;
            xlWorkSheet.get_Range("A2:R2").Font.Bold = true;
            xlWorkSheet.get_Range("A2:R2").Font.Size = 16;
            //מרכוז
            Excel.Range workSheet_range = xlWorkSheet.get_Range("A:R");
            workSheet_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            workSheet_range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;//מרכוז טקסט
            workSheet_range.ColumnWidth = 12;


            Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[3, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            xlWorkSheet.Cells[1, 1] = "כוח אדם-אליאנס";
            xlWorkSheet.Cells[2, 18] = "מספר עובד";
            xlWorkSheet.Cells[2, 17] = "שם פרטי";
            xlWorkSheet.Cells[2, 16] = "שם משפחה";
            xlWorkSheet.Cells[2, 15] = "ת.ז";
            xlWorkSheet.Cells[2, 14] = "כתובת";
            xlWorkSheet.Cells[2, 13] = "תאריך קליטה";
            xlWorkSheet.Cells[2, 12] = "חודש קליטה";
            xlWorkSheet.Cells[2, 11] = "טלפון";
            xlWorkSheet.Cells[2, 10] = "מחלקה";
            xlWorkSheet.Cells[2, 9] = "שם מחלקה";
            xlWorkSheet.Cells[2, 8] = "תאריך לידה";
            xlWorkSheet.Cells[2, 7] = "גיל";
            xlWorkSheet.Cells[2, 6] = "סוג עובד";
            xlWorkSheet.Cells[2, 5] = "תאריך עזיבה";
            xlWorkSheet.Cells[2, 4] = "First Name";
            xlWorkSheet.Cells[2, 3] = "Last Name";
            xlWorkSheet.Cells[2, 2] = "ותק";
            xlWorkSheet.Cells[2, 1] = "סוג הסכם";

            //גבולות תא
            Excel.Range UsedRange, chartRange;
            UsedRange = xlWorkSheet.UsedRange;
            int lastUsedRow = UsedRange.Row + UsedRange.Rows.Count - 1;//שורה אחרונה עם טקסט באקסל
            chartRange = xlWorkSheet.get_Range("A1:R" + lastUsedRow);//עיצוב תאים
            chartRange.Borders.Color = System.Drawing.Color.Black.ToArgb();


            ////מחיקת שורות ריקות
            for (int i = 2; i < lastUsedRow; i++)
            {
                if (string.IsNullOrEmpty(xlWorkSheet.Cells[i, 14].Value2 as string))
                {
                    if (string.IsNullOrEmpty(xlWorkSheet.Cells[i + 1, 14].Value2 as string)) break;
                    xlWorkSheet.Rows[i].Delete();

                }

            }

            Excel.Range firstRow = (Excel.Range)xlWorkSheet.Rows[2];
            firstRow.Application.ActiveWindow.FreezePanes = true;
            firstRow.AutoFilter(1,
                                Type.Missing,
                                Excel.XlAutoFilterOperator.xlAnd,
                                Type.Missing,
                                true);
            dataGrid_workStrengh.CurrentCell = dataGrid_workStrengh.Rows[2].Cells[3];
            dataGrid_workStrengh.Focus();
            dataGrid_workStrengh.BeginEdit(true);
            xlexcel.Visible = true;
            //xlWorkBook.Close();
            releaseObject(xlexcel);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);
        }

        /// בודק אם להעתיק דטה של כ"א כללי או מצבת כ"א
        /// </summary>
        private void copyAlltoClipboard()
        {
            DataObject dataObj = null;
            switch (tab)
            {
                case 0:
                    DataGridView_employees.SelectAll();
                    dataObj = DataGridView_employees.GetClipboardContent();
                    break;

                case 1:
                    dataGrid_workStrengh.SelectAll();
                    dataObj = dataGrid_workStrengh.GetClipboardContent();
                    break;

                default:
                    break;

            }


            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        /// חסימת מיון מצבת כ"א
        /// </summary>

        private void dataGrid_workStrengh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewColumn column in dataGrid_workStrengh.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void StyleBdayDGV()
        {
            BdayGridView.Columns["EmpNum"].HeaderText = "מס' עובד";
            BdayGridView.Columns["Name"].HeaderText = "שם";
            //BdayGridView.Columns["LastName"].HeaderText = "שם משפחה";
            BdayGridView.Columns["DateOfBirth"].HeaderText = "תאריך לידה";
            BdayGridView.Columns["DateOfBirthNoYear"].HeaderText = "תאריך לידה ללא שנה";
            BdayGridView.Columns["Unit"].HeaderText = "מחלקה";
            foreach (DataGridViewColumn col in BdayGridView.Columns)
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
          //  BdayGridView.Columns["DateOfBirth"].ValueType = typeof(DateTime);
          //  BdayGridView.Columns["DateOfBirth"].DefaultCellStyle.Format = "dd/MM/yyyy";
          //  BdayGridView.Sort( BdayGridView.Columns["DateOfBirth"], ListSortDirection.Ascending);
            
        }


        private void makeBdayExcelRep()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            //  Excel.Worksheet xlWorkSheet2;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            xlWorkSheet.Cells[1, 1] = "מס' עובד";
            xlWorkSheet.Cells[1, 2] = "שם";
            //xlWorkSheet.Cells[1, 3] = "שם משפחה";
            xlWorkSheet.Cells[1, 3] = "תאריך לידה";
            xlWorkSheet.Cells[1, 4] = "תאריך לידה ללא שנה";
            xlWorkSheet.Cells[1, 5] = "מחלקה";

            chartRange = xlWorkSheet.get_Range("D2", "D" + BdayGridView.Rows.Count); 
            chartRange.NumberFormat = "@";
            Excel.Range chartRange2= xlWorkSheet.get_Range("C2", "C" + BdayGridView.Rows.Count);
            chartRange.NumberFormat = "@";



            //            xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
            // Now apply autofilter
            Excel.Window xlWnd1 = xlApp.ActiveWindow;
            chartRange = xlWorkSheet.get_Range("A1", "A1").get_Offset(1, 0).EntireRow;
            chartRange.Select();
            xlWnd1.FreezePanes = true;
            Excel.Range firstRow = (Excel.Range)xlWorkSheet.Rows[1];
            firstRow.Application.ActiveWindow.FreezePanes = true;
            firstRow.AutoFilter(1,
                                Type.Missing,
                                Excel.XlAutoFilterOperator.xlAnd,
                                Type.Missing,
                                true);

            for (int i = 0; i < BdayGridView.Rows.Count; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = BdayGridView.Rows[i].Cells["EmpNum"].Value;
                xlWorkSheet.Cells[i + 2, 2] = BdayGridView.Rows[i].Cells["Name"].Value;//+" "+ BdayGridView.Rows[i].Cells["LastName"].Value;
                xlWorkSheet.Columns[i + 2].ColumnWidth = 18;
                //xlWorkSheet.Cells[i + 2, 3] = BdayGridView.Rows[i].Cells["LastName"].Value;
                xlWorkSheet.Cells[i + 2, 3] = BdayGridView.Rows[i].Cells["DateOfBirth"].Value;
                xlWorkSheet.Cells[i + 2, 4] = BdayGridView.Rows[i].Cells["DateOfBirthNoYear"].Value;
                xlWorkSheet.Cells[i + 2, 5] = BdayGridView.Rows[i].Cells["Unit"].Value;
            }

            //      xlApp.Visible = true;
            chartRange = xlWorkSheet.get_Range("A1", "E1");
            chartRange.Interior.Color = Color.LightBlue;
            xlWorkSheet.Columns.AutoFit();

            chartRange = xlWorkSheet.UsedRange.Columns["C:D", Type.Missing];
            chartRange.EntireColumn.NumberFormat = "MM/DD/YYYY";


            xlApp.DisplayAlerts = false;

            xlApp.Visible = true;
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void MakeExcellRepBTN_Click(object sender, EventArgs e)
        {
            if (tabControlll.SelectedTab.Name == "BdayTab")
                makeBdayExcelRep();
            if (tabControlll.SelectedTab.Name == "AbsenceRepTab")
                makeAbsenceExcelRep();
            if (tabControlll.SelectedTab.Name == "SabonTabPage")
                makeSabonExcelRep();
            if (tabControlll.SelectedTab.Name == "tabpage_Employees")
                makeGeneralEmployeeReport();
            if (tabControlll.SelectedTab.Name == "strengthTab")
                makeStrengthReport();

        }


        private void makeSabonExcelRep()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            //  Excel.Worksheet xlWorkSheet2;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            xlWorkSheet.Cells[1, 1] = "מחלקה";
            xlWorkSheet.Cells[1, 2] = "כמות סבונים";
            //            xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
            // Now apply autofilter
            Excel.Window xlWnd1 = xlApp.ActiveWindow;
            chartRange = xlWorkSheet.get_Range("A1", "A1").get_Offset(1, 0).EntireRow;
            chartRange.Select();
            xlWnd1.FreezePanes = true;
            Excel.Range firstRow = (Excel.Range)xlWorkSheet.Rows[1];
            firstRow.Application.ActiveWindow.FreezePanes = true;
            firstRow.AutoFilter(1,
                                Type.Missing,
                                Excel.XlAutoFilterOperator.xlAnd,
                                Type.Missing,
                                true);

            for (int i = 0; i < sabonReportGrid.Rows.Count; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = sabonReportGrid.Rows[i].Cells["UNIT"].Value;
                xlWorkSheet.Cells[i + 2, 2] = sabonReportGrid.Rows[i].Cells["AMOUNT"].Value;
            }

            //      xlApp.Visible = true;
            chartRange = xlWorkSheet.get_Range("A1", "B1");
            chartRange.Interior.Color = Color.LightBlue;
            xlWorkSheet.Columns.AutoFit();

            //chartRange = xlWorkSheet.UsedRange.Columns["F:F", Type.Missing];
            //chartRange.EntireColumn.NumberFormat = "MM/dd/yyyy";

            xlApp.DisplayAlerts = false;
            xlApp.Visible = true;
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);

        }

        private void makeAbsenceExcelRep()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            //  Excel.Worksheet xlWorkSheet2;
            object misValue = System.Reflection.Missing.Value;
            Excel.Range chartRange;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            xlWorkSheet.Cells[1, 1] = "מס' עובד";
            xlWorkSheet.Cells[1, 2] = "שם פרטי";
            xlWorkSheet.Cells[1, 3] = "שם משפחה";
            xlWorkSheet.Cells[1, 4] = "מחלקה";
            xlWorkSheet.Cells[1, 5] = "טלפון";
            xlWorkSheet.Cells[1, 6] = "מס' ימים מאז נוכחות אחרונה";
            xlWorkSheet.Cells[1, 7] = "סיבת העדרות";
            xlWorkSheet.Cells[1, 8] = "נוכחות אחרונה בעבודה";
            //            xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
            // Now apply autofilter
            Excel.Window xlWnd1 = xlApp.ActiveWindow;
            chartRange = xlWorkSheet.get_Range("A1", "A1").get_Offset(1, 0).EntireRow;
            chartRange.Select();
            xlWnd1.FreezePanes = true;
            Excel.Range firstRow = (Excel.Range)xlWorkSheet.Rows[1];
            firstRow.Application.ActiveWindow.FreezePanes = true;
            firstRow.AutoFilter(1,
                                Type.Missing,
                                Excel.XlAutoFilterOperator.xlAnd,
                                Type.Missing,
                                true);

            chartRange = xlWorkSheet.get_Range("F2", "F" + absenceGridView.Rows.Count+10);
            chartRange.NumberFormat = "@";

            for (int i = 0; i < absenceGridView.Rows.Count; i++)
            {
                xlWorkSheet.Cells[i + 2, 1] = absenceGridView.Rows[i].Cells["EmpNum"].Value;
                xlWorkSheet.Cells[i + 2, 2] = absenceGridView.Rows[i].Cells["FirstName"].Value;
                xlWorkSheet.Columns[i + 2].ColumnWidth = 18;
                xlWorkSheet.Cells[i + 2, 3] = absenceGridView.Rows[i].Cells["LastName"].Value;
                xlWorkSheet.Cells[i + 2, 4] = absenceGridView.Rows[i].Cells["Unit"].Value;
                xlWorkSheet.Cells[i + 2, 5] = absenceGridView.Rows[i].Cells["Phone"].Value;
                xlWorkSheet.Cells[i + 2, 6] = absenceGridView.Rows[i].Cells["AbsenceInARow"].Value;
                xlWorkSheet.Cells[i + 2, 7] = absenceGridView.Rows[i].Cells["AbsenceReason"].Value;
                xlWorkSheet.Cells[i + 2, 8] = absenceGridView.Rows[i].Cells["LastDateAtWork"].Value.ToString().Trim();
            }

            //      xlApp.Visible = true;
          

            chartRange = xlWorkSheet.get_Range("A1", "H1");
            chartRange.Interior.Color = Color.LightBlue;
            xlWorkSheet.Columns.AutoFit();

         chartRange = xlWorkSheet.UsedRange.Columns["E:F", Type.Missing];
         chartRange.Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;





            xlApp.DisplayAlerts = false;

            xlApp.Visible = true;
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
            releaseObject(xlWorkSheet);


        }

        private void getWorkReportFromHilanFiles()
        {
            string prevEmpNum = "";
            Employee currentEmp = null;
            var reader = new StreamReader(File.OpenRead(@"\\172.16.1.39\\WendimuTransfer\ZMIYOMI3.csv"));
            reader.ReadLine();

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine().Split(',');
                if (line[16] == "0340" && line[18].Trim() != "8.60")
                    continue;
                if (line[4] != prevEmpNum)
                {

                    foreach (Employee em in employees)
                    {
                        if (em.checkIfParamIsEmpNum(line[4]))
                        {
                            prevEmpNum = line[4];
                            currentEmp = em;
                            break;
                        }
                    }
                }
                currentEmp.addWorkDay(new WorkDay
                {
                    arriveToWork = line[16] == "0000" && double.Parse(line[18].Trim()) > 0,
                    date = DateTime.ParseExact(line[3], "dd/MM/yyyy",
                                       System.Globalization.CultureInfo.InvariantCulture),
                    absenceReason = line[16] != "0000" ? line[17].Trim() : ""
                });
            }

            reader = new StreamReader(File.OpenRead(@"\\172.16.1.39\\WendimuTransfer\ZMIYOMI2.csv"), Encoding.GetEncoding("windows-1255"));
            reader.ReadLine();
            prevEmpNum = "";
            currentEmp = null;

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine().Split(',');
                if (line[4] != prevEmpNum)
                {

                    foreach (Employee em in employees)
                    {
                        if (em.checkIfParamIsEmpNum(line[4]))
                        {
                            prevEmpNum = line[4];
                            currentEmp = em;
                            break;
                        }
                    }
                }
                if (currentEmp == null) continue;
                currentEmp.addWorkDay(new WorkDay
                {
                    arriveToWork = line[16] == "0000" && double.Parse(line[18].Trim()) > 0,
                    date = DateTime.ParseExact(line[3], "dd/MM/yyyy",
                                       System.Globalization.CultureInfo.InvariantCulture),
                    absenceReason = line[16] != "0000" ||!(line[16].Trim() == "0340" && line[18].Trim() != "8.6") ? line[17] : ""
                });

            }
        }


        private void buildAbsenceReport()
        {
            getWorkReportFromHilanFiles();
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn { ColumnName = "EmpNum" });
            dt.Columns.Add(new DataColumn { ColumnName = "FirstName" });
            dt.Columns.Add(new DataColumn { ColumnName = "LastName" });
            dt.Columns.Add(new DataColumn { ColumnName = "Unit" });
            dt.Columns.Add(new DataColumn { ColumnName = "Phone" });
            dt.Columns.Add(new DataColumn { ColumnName = "AbsenceInARow" });
            dt.Columns.Add(new DataColumn { ColumnName = "AbsenceReason" });
            dt.Columns.Add(new DataColumn { ColumnName = "LastDateAtWork" });
           // dt.Columns["AbsenceInARow"].DataType = typeof(System.Int32);

            foreach (Employee emp in employees)
            {
                if (emp.Stop_Work.Year < 5000) continue;
                Tuple<int,DateTime> t = emp.howManyAbsenceInARow();
                if (t.Item1 < 10) continue;
                var row = dt.NewRow();
                row["EmpNum"] = emp.Employee_Num;
                row["FirstName"] = emp.First_Name;
                row["LastName"] = emp.Last_Name;
                row["Phone"] = emp.phoneNum;
                row["LastDateAtWork"] = t.Item2.Year>1800?t.Item2.ToShortDateString():"לא ידוע";
                row["Unit"] = emp.Unit;
                row["AbsenceInARow"] = t.Item1+(t.Item2.Year > 1800 ?"":"  +");
                row["AbsenceReason"] = emp.getLastAbsenceReason();
                dt.Rows.Add(row);
            }
            absenceGridView.DataSource = dt;
            StyleAbsenceDGV();
        }

        public List<Tuple<int, int>> calculateHowManySabonForUnit()
        {
            List<Tuple<int, int>> howManySabonPerUnit = new List<Tuple<int, int>>();
            // foreach (Employee e in employees)
            //     if (!units.Contains(e.Unit))
            //         units.Add(e.Unit);
            List<int> units = new List<int> { 22,23,28,31,35,36,37,38,39,40,42,44,45,51,52,54,55,57,66,71,72,87,88,160,161,178,142,50,15};//רשימת המחלקות שאלי העביר
            units.Sort();
            foreach(int unit in units)
            {
                //2 סבונים לאדם, חישוב עבור רבעון
                howManySabonPerUnit.Add(new Tuple<int, int>(unit, countHowManyFromUnit(unit) * 2*3));
            }
            return howManySabonPerUnit;
        }

        private int countHowManyFromUnit(int unit)
        {
            int count = 0;
            foreach (Employee e in employees)
                if (e.Unit == unit && e.Stop_Work.Year>5000)
                    count++;
            return count;
        }

        private void StyleAbsenceDGV()
        {
            absenceGridView.Columns["EmpNum"].HeaderText = "מס' עובד";
            absenceGridView.Columns["FirstName"].HeaderText = "שם פרטי";
            absenceGridView.Columns["LastName"].HeaderText = "שם משפחה";
            absenceGridView.Columns["Phone"].HeaderText = "טלפון";
            absenceGridView.Columns["LastDateAtWork"].HeaderText = "נוכחות אחרונה בעבודה";
            absenceGridView.Columns["Unit"].HeaderText = "מחלקה";
            absenceGridView.Columns["AbsenceReason"].HeaderText = "סיבת העדרות";
            absenceGridView.Columns["AbsenceInARow"].HeaderText = "מס' ימים מאז נוכחות אחרונה";
            foreach (DataGridViewColumn col in absenceGridView.Columns)
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            absenceGridView.Sort(absenceGridView.Columns["AbsenceInARow"], ListSortDirection.Descending);
        }

        private string GetCurrentDate()
        {
            DateTime T = DateTime.Now;
            return T.ToString("yyMMdd");
        }

        private void buildSabonReport()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn { ColumnName = "UNIT" });
            dt.Columns.Add(new DataColumn { ColumnName = "AMOUNT" });
             foreach (Tuple<int,int> tup in calculateHowManySabonForUnit())
             {

                var row = dt.NewRow();
                row["UNIT"] = tup.Item1;
                row["AMOUNT"] = tup.Item2;
                 dt.Rows.Add(row);
             }
            sabonReportGrid.DataSource = dt;
            StyleSabonDGV();
        }

        private void StyleSabonDGV()
        {
            sabonReportGrid.Columns["AMOUNT"].HeaderText = "כמות סבונים";
            sabonReportGrid.Columns["UNIT"].HeaderText = "מחלקה";
            foreach (DataGridViewColumn col in sabonReportGrid.Columns)
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }


        private void sendMailToMissingWorkers()
        {
            getWorkReportFromHilanFiles();
            List<Employee> missingEmp = new List<Employee>();
            foreach (Employee emp in employees)
            {
                Tuple<int, DateTime> t = emp.howManyAbsenceInARow();
                if (t.Item1 != 10) continue;
                missingEmp.Add(emp);
            }
            foreach(Employee emp in missingEmp)
            {
                MailMessage message = new MailMessage();
                message.Subject = "עובד נעדר-"+emp.First_Name+" "+emp.Last_Name;
                message.Priority = MailPriority.High;
                message.Body = $@"שלום ,
לידיעתך , העובד {emp.First_Name + " " + emp.Last_Name} לא הגיע לעבודה מזה 10 ימים.
נא בדיקתך, תודה
";
                message.To.Add(new MailAddress(emp.getManagerMail()));
                message.From = new MailAddress("HR@atgtire.com");
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.SubjectEncoding = System.Text.Encoding.UTF8;
                SmtpClient client = new SmtpClient();
                client.Host = "almail";// ServerIP;
                client.Send(message);
                client.Dispose();
            }

        }
        private void finishWorkButton_Click(object sender, EventArgs e)
        {
            ((Employee)FinishWorkComboBox.SelectedItem).WorkFinishForms();
        }

        private void FinishWorkComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            detailsAboutWorkerLabel.Visible = true;
            
            detailsAboutWorkerLabel.Text = ((Employee)FinishWorkComboBox.SelectedItem).detailsForEndWork();

        }

        private void dataGrid_workStrengh_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        private void setSize()
        {
            MakeExcellRepBTN.Location = new Point(tabControlll.Location.X, tabControlll.Location.Y + tabControlll.Height);
            finishWorkLabel.Location = new Point((int)(tabControlll.Location.X + 0.28 * tabControlll.Width), tabControlll.Location.Y);
            detailsAboutWorkerLabel.Location = new Point(finishWorkLabel.Location.X + (int)(0.19 * tabControlll.Height), finishWorkLabel.Location.Y + (int)(0.06 * tabControlll.Width));
            FinishWorkComboBox.Location = new Point(finishWorkLabel.Location.X + (int)(0.25 * tabControlll.Height), finishWorkLabel.Location.Y + (int)(0.10 * tabControlll.Width));
            finishWorkButton.Location = new Point(finishWorkLabel.Location.X + (int)(0.25 * tabControlll.Height), finishWorkLabel.Location.Y + (int)(0.14 * tabControlll.Width));
        }

        private void Form1_Resize_1(object sender, EventArgs e)
        {
            setSize();
        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void date_start1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void date_start2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControlll.SelectedIndex == 5 && finishInit)
            {
                 FinishWorkComboBox.Focus();         
            }

            if(tabControlll.SelectedTab == SubscriptionNotes && finishInit)//לשונית כתבי מינוי
            {
                try
                {
                    SubscriptionNotesFunc();
                }

                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
           
            }

        }



        private void FinishWorkComboBox_TextChanged(object sender, EventArgs e)
        {
            if (tabControlll.SelectedIndex == 5 && finishInit)
            {
                string temp = FinishWorkComboBox.Text;
                FinishWorkComboBox.DroppedDown = true;
                Cursor.Current = Cursors.Default;
                if (temp.Length == 1)
                {
                    FinishWorkComboBox.Text = temp;
                    FinishWorkComboBox.Select(1, 1);
                }
            }
        }

        private void DataGridView_employees_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string NumberPicture = DataGridView_employees.Rows[e.RowIndex].Cells["NUMBER"].Value.ToString();
                Process.Start($@"T:\M14\אירה_אלי\הודפסו\{NumberPicture}.jpg");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
       
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void btn_CreateFileStart_Click(object sender, EventArgs e)
        {
            ((Employee)FinishWorkComboBox.SelectedItem).WorkStartFile();
        }
        private void btn_CreateFileStop_Click(object sender, EventArgs e)
        {
            ((Employee)cbo_StopWork.SelectedItem).MakeFileStopWork("T:\\Application\\C#\\HR\\שבלונה הפסקת עבודה\\STOPWORKFORMAT.docx");
        }
        private void Cbo_TextChanged(ComboBox comboBox,int indexTab)
        {
            if (tabControlll.SelectedIndex == indexTab && finishInit)
            {
                string temp = comboBox.Text;
                comboBox.DroppedDown = true;
                Cursor.Current = Cursors.Default;
                if (temp.Length == 1)
                {
                    comboBox.Text = temp;
                    comboBox.Select(1, 1);
                }
            }
        }

        private void Cbo_SelectedIndexChanged(ComboBox comboBox, Label lbl)
        {
            lbl.Visible = true;

            lbl.Text = ((Employee)comboBox.SelectedItem).detailsForEndWork();
        }

        private void cbo_StartWork_TextChanged(object sender, EventArgs e)
        {
            Cbo_TextChanged(cbo_StartWork, 6);
            //if (tabcControl.SelectedIndex == 6 && finishInit)
            //{
            //    string temp = cbo_StartWork.Text;
            //    cbo_StartWork.DroppedDown = true;
            //    Cursor.Current = Cursors.Default;
            //    if (temp.Length == 1)
            //    {
            //        cbo_StartWork.Text = temp;
            //        cbo_StartWork.Select(1, 1);
            //    }
            //}
        }

        private void cbo_StopWork_TextChanged(object sender, EventArgs e)
        {
            Cbo_TextChanged(cbo_StopWork, 7);
        }

        private void cbo_Subscription_TextChanged(object sender, EventArgs e)
        {
            Cbo_TextChanged(cbo_Subscription, 8);
        }
        private void cbo_StartWork_SelectedIndexChanged(object sender, EventArgs e)
        {
            //lbl_DetailsStartWork.Visible = true;

            //lbl_DetailsStartWork.Text = ((Employee)cbo_StartWork.SelectedItem).detailsForEndWork();
            Cbo_SelectedIndexChanged(cbo_StartWork, lbl_DetailsStartWork);

        }


        private void cbo_StopWork_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cbo_SelectedIndexChanged(cbo_StopWork, lbl_DetailsStoppWork);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //DbServiceSQL DBnewVesion = new DbServiceSQL();//הבעיה בפרוביידר
            //DataTable dataTable = new DataTable();
            //var Name = typeof(Form1).Namespace;
            //string qry = $@"SELECT Change
            //                FROM AppsForUpdate
            //                WHERE Apps='{Name.ToString()}'";
            //dataTable = DBnewVesion.executeSelectQueryNoParam(qry);
            //if (dataTable.Rows[0][0].ToString() == "True")
            //{
            //    Application.Exit();
            //}

        }





        private void cbo_search_result_Leave(object sender, EventArgs e)
        {
          
            try
            {
                if (cbo_search_result.SelectedIndex != 0 && cbo_search.Text == "שם מלא")
                {
                    TableForFilterLastname = new DataView(dataTableForShow);//שייך לקומבובוקס של סינון שמות משפחה רלוונטיים-לראות באיוונט של עזיבת קומבובוקס 
                    TableForFilterLastname.RowFilter = $@"FIRSTNAME = '{cbo_search_result.Text}' ";
                    cbo_lastname.DataSource = TableForFilterLastname;
                }
            }

            catch(Exception ex)
            {
                MessageBox.Show("הקלדת שם פרטי שגוי");
            }
        }

        /// <summary>
        /// לשונית כתבי מינוי שמפיקה קובץ וורד
        /// </summary>
        private void SubscriptionNotesFunc()
        {
            Employee employee = new Employee();
            DataTable dataTable = new DataTable();
            dataTable = employee.GetSubscriptionNotes();
            DataGrid_Subscription.DataSource = dataTable;
        }

        /// <summary>
        /// הוספת שם כתב מינוי
        /// </summary>
        private void btn_AddSub_Click(object sender, EventArgs e)
        {
            //if (txt_sub.Text != "")
            //{
            //    string qry = $@"INSERT into SubscriptionNotes_Hr values(?,?,?)";
            //    SqlParameter Subject = new SqlParameter("s", txt_sub.Text);
            //    //SqlParameter 

            //}
        }

    }
}


