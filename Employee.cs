using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Data;
using System.IO;

enum EmployeeType { Direct = 0, Indirect = 1, Manger = 2 }
enum WorkType { Temporary = 0, Permanent = 1, HA = 2 }
namespace HR
{
    class Employee
    {//itsik
        public int Employee_Num { get; set; }
        public string First_Name { get; set; }
        public string Last_Name { get; set; }
        public DateTime Date_Start { get; set; }
        public DateTime Stop_Work { get; set; }
        public string Id { get; set; } //ת.ז
        public EmployeeType Employee_Type { get; set; }//סוג עובד ישיר,עקיף,מנהל
        public DateTime Employee_Birthday { get; set; }
        public WorkType workType { get; set; }//קבוע/זמני/ח.א
        public string English_First_Name { get; set; }
        public string English_Last_Name { get; set; }
        public int seniority { get; set; } //ותק
        public bool Status { get; set; }//פעיל/לא פעיל
        public int Unit { get; set; }//מחלקה
        public List<WorkDay> workDays { get; set; } //רשימה שתחזיק את הנוכחות שלו בחודשיים האחרונים

        public DateTime Date_Start2 { get; set; }//הוספתי בשביל סינונים של תאריכי קליטה
        public DateTime Date_Left { get; set; }//עבור סינון תאריך עזיבה
        public DateTime Date_Left2 { get; set; }//2עבור סינון תאריך עזיבה
        public string phoneNum { get; set; }

        public Employee()
        {
            workDays = new List<WorkDay>();
        }
        public Employee(DateTime date_start, DateTime date_start2)
        {
            Date_Start = date_start;
            Date_Start2 = date_start2;
        }
        public static List<Employee> GetListOfActiveEmployees()
        {
            DBService DBS = new DBService();
            DataTable DataTable = new DataTable();
          //string Employee_Table = $@"SELECT  T.OVED as Number ,L.PRATI as FirstName ,L.FAMILY as LastName,T.mifal,
          //                         L.AVODA as DateStart, T.SUGOVD as TypeTime, L.MAHLAKA as Unit,T.BIRTHDAY  ,TEUDTZHUI AS ID,        (CASE  WHEN LENGTH(trim(t.kpa))<7 //THEN trim(t.kpa) else  substring(T.kpa,7,6 )end) as leavingDate
          //                        FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED 
          //                        WHERE T.mifal='01'";

            string Employee_Table= $@"SELECT  right(trim(T.OVED),5) as Number , name1||name2||name3||name4||name5||name6 as FirstName,                                      fmly1||fmly2||fmly3||fmly4||fmly5||fmly6||fmly7||fmly8||fmly9||fmly10 as LastName,T.TEUDTZHUI as Id,T.TELEFON as phone,
                                     T.ADRESS,T.AVODA as DateStart, H.HOFIL3 as TypeTime,right(trim(L.MAHLAKA),3) as Unit,T.BIRTHDAY,case when right(trim(T.KPA),6) = '000000' then '' else right(trim(T.KPA),6) end leavingDate                                    , H.HOMACH as EFirstName,H.HOMACTP as ELastName ,B.CDESC as UnitName
                                     FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED left join hsali.hovl02 as H on T.OVED=H.HOOVD left join  BPCSFV30. CDPL01  as B on right(trim(L.MAHLAKA),2)=B.CDEPT
                                     WHERE   (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129')  and (substring(T.kpa,11,2)  > '08' and (substring(T.kpa,11,2) < '80') or substring(T.kpa,7,6 )= '000000') and             	                     H.HOMIF in ('01', '09')";

            DataTable = DBS.executeSelectQueryNoParam(Employee_Table);
            List<Employee> toReturn = new List<Employee>();
            foreach (DataRow row in DataTable.Rows)
            {

                Employee temp = new Employee
                {
                    Employee_Num = int.Parse(row["NUMBER"].ToString()),
                    First_Name = row["FIRSTNAME"].ToString(),
                    Last_Name = row["LASTNAME"].ToString(),
                    workType = (WorkType)int.Parse(row["TYPETIME"].ToString()),
                    Unit = row["UNIT"].ToString() != "" ? int.Parse(row["UNIT"].ToString()) : 0,
                    Id = row["ID"].ToString(),
                    phoneNum = row["PHONE"].ToString()
            };
                if(row["leavingDate"].ToString()!="000000" && row["leavingDate"].ToString() != "")
                    temp.Stop_Work = DateTime.ParseExact(row["leavingDate"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture);
                else
                {
                    temp.Stop_Work = DateTime.MaxValue;
                }
                if ((row["DATESTART"].ToString() != "000000")&& row["DATESTART"].ToString() != "")
                    temp.Date_Start = DateTime.ParseExact(row["DATESTART"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture);
                else
                {
                    
                    temp.Date_Start = DateTime.MinValue;
                }
                if ((row["BIRTHDAY"].ToString() != "000000") && row["BIRTHDAY"].ToString() != "")
                    temp.Employee_Birthday = DateTime.ParseExact(row["BIRTHDAY"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture);
                else
                {
                    temp.Employee_Birthday = DateTime.MinValue;
                }
                toReturn.Add(temp);
            }
            return toReturn;
        }
        public bool checkIfHaveBdayInRangeOfMonths(DateTime from, DateTime to)
        {
            DateTime fromTemp = new DateTime(2018, from.Month, from.Day);
            DateTime toTemp = new DateTime(2018, to.Month, to.Day);
            DateTime bDayTemp = new DateTime(2018, Employee_Birthday.Month, Employee_Birthday.Day);
            if (from.Month > to.Month) toTemp = toTemp.AddYears(1);
            if (Employee_Birthday.Month < from.Month) bDayTemp = bDayTemp.AddYears(1);
            return bDayTemp >= fromTemp && bDayTemp <= toTemp;
        }
        override public string ToString()
        {
            return Employee_Num+" "+First_Name + " " + Last_Name;
        }
        public string detailsForEndWork()
        {
            return $@"מס' עובד:{Employee_Num} שם :{First_Name + " " + Last_Name} מחלקה:{Unit} ";
        }
        
        public DataTable GetData()
        {
            DBService DBS = new DBService();
            DataTable dataTable = new DataTable();
            string Employee_Table = $@"SELECT  T.OVED as Number ,L.PRATI as FirstName ,L.FAMILY as LastName,
                                     L.AVODA as DateStart, T.SUGOVD as TypeTime, L.MAHLAKA as Unit,T.BIRTHDAY 
                                    FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED 
                                    WHERE T.mifal='01'";
            try
            {
                DateTime bday = DateTime.MinValue;
                bday = DateTime.ParseExact(DBS.executeSelectQueryNoParam(Employee_Table).Rows[0]["BIRTHDAY"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception EX) { }
            dataTable = DBS.executeSelectQueryNoParam(Employee_Table);
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                if (dataTable.Rows[i]["BIRTHDAY"].ToString() != "000000")
                {
                    dataTable.Rows[i]["DateStart"] = DateTime.ParseExact(dataTable.Rows[i]["DateStart"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture);
                    dataTable.Rows[i]["BIRTHDAY"] = DateTime.ParseExact(dataTable.Rows[i]["BIRTHDAY"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            return dataTable;
            //DateTime.ParseExact(dbAS400.executeSelectQueryNoParam(StrSqlAS400).Rows[0]["BIRTHDAY"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture)
        }

        /// <summary>
        /// קבלת נתונים כלליים של כלל העובדים או פעילים או לא פעילים,עם אופציית סינונים
        public DataTable GetData(bool ShowAllEmployees, bool Filter_Date_Start,bool FilterDateLeft)
        {
            DBService DBS = new DBService();
            DataTable dataTable = new DataTable();
            string Employee_Table;
            if (ShowAllEmployees == false)
            {
                Employee_Table = $@"SELECT  right(trim(T.OVED),5) as Number ,L.PRATI as FirstName ,L.FAMILY as LastName,T.TEUDTZHUI as Id
                                    ,T.ADRESS,L.AVODA as DateStart,substring(L.AVODA,3,2) as MonthStart,T.TELEFON, case when substring(L.MAHLAKA,2,1)=1 then right(trim(L.MAHLAKA),3) when substring(L.MAHLAKA,2,1)=0 then  right(trim(L.MAHLAKA),2)
 	                                 end Unit, B.CDESC as UnitName,T.BIRTHDAY,case when substring(t.birthday,5,2)>40 then TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '19'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' ))) else TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '20'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' )))  end Age,                                      case when H.HOFIL3=2 then 'עקיף' when H.HOFIL3=1 then 'ישיר'  when H.HOFIL3=3 then 'מנהל' when H.HOFIL3=4 then 'עקיף חרושת' when H.HOFIL3=9 then 'פנסיה'  end type,                                     H.HOMACH as EFirstName,H.HOMACTP as ELastName ,
                                     case when substring(L.AVODA,5,2)>{DateTime.Now.Year.ToString().Substring(2,2)} then  dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('19'||substring(L.AVODA,5,2)||'-'||substring(L.AVODA,3,2)||'-'||substring(L.AVODA,1,2)))/100) , 5,0))/ 100,5,2)                                     when substring(L.AVODA,5,2)<={DateTime.Now.Year.ToString().Substring(2, 2)} then dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('20'||substring(L.AVODA,5,2)||'-'||substring(L.AVODA,3,2)||'-'||substring(L.AVODA,1,2))) /100) , 5,0)) / 100 , 5, 2) end Vetek,
                                    case when substring(T.OVED,3,2)=19 then 'אישי' else 'קיבוצי' end typeContract,case when h.hofil4=7 then 'זמני' else 'קבוע' end temp                                    FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED left join hsali.hovl02 as H on T.OVED=H.HOOVD left join  BPCSFV30. CDPL01  as B on right(trim(L.MAHLAKA),3)=B.CDEPT
                                    WHERE T.mifal='01' and (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129') and substring(T.kpa,7,6 )= '000000' and H.HOMIF in ('01', '09')";
            }
            else
            {
                Employee_Table = $@"SELECT  right(trim(T.OVED),5) as Number , name1||name2||name3||name4||name5||name6 as FirstName,                                      fmly1||fmly2||fmly3||fmly4||fmly5||fmly6||fmly7||fmly8||fmly9||fmly10 as LastName,T.TEUDTZHUI as Id,
                                     T.ADRESS,T.AVODA as DateStart,substring(L.AVODA,3,2) as MonthStart,T.TELEFON ,case when substring(L.MAHLAKA,2,1)=1 then right(trim(L.MAHLAKA),3) when substring(L.MAHLAKA,2,1)=0 then  right(trim(L.MAHLAKA),2)
 	                                 end Unit,T.BIRTHDAY,case when right(trim(T.KPA),6) = '000000' then '' else right(trim(T.KPA),6) end left                                    , H.HOMACH as EFirstName,H.HOMACTP as ELastName ,B.CDESC as UnitName,case when substring(t.birthday,5,2)>40 then TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '19'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' ))) else TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '20'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' )))  end Age,
                                     case when substring(T.AVODA,5,2)>{DateTime.Now.Year.ToString().Substring(2, 2)} then  dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('19'||substring(T.AVODA,5,2)||'-'||substring(T.AVODA,3,2)||'-'||substring(T.AVODA,1,2)))/100) , 5,0))/ 100,5,2)                                     when substring(T.AVODA,5,2)<={DateTime.Now.Year.ToString().Substring(2, 2)} then dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('20'||substring(T.AVODA,5,2)||'-'||substring(T.AVODA,3,2)||'-'||substring(T.AVODA,1,2))) /100) , 5,0)) / 100 , 5, 2) end Vetek,
                                     case when substring(T.OVED,3,2)=19 then 'אישי' else 'קיבוצי' end typeContract, case when H.HOFIL3=2 then 'עקיף' when H.HOFIL3=1 then 'ישיר'  when H.HOFIL3=3 then 'מנהל' when H.HOFIL3=4 then 'עקיף חרושת' when H.HOFIL3=9 then 'פנסיה'  end type,case when h.hofil4=7 then 'זמני' else 'קבוע' end temp 
                                     FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED left join hsali.hovl02 as H on T.OVED=H.HOOVD left join  BPCSFV30. CDPL01  as B on right(trim(L.MAHLAKA),3)=B.CDEPT
                                     WHERE   (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129')  and (substring(T.kpa,11,2)  > '08' and (substring(T.kpa,11,2) < '80') or substring(T.kpa,7,6 )= '000000') and             	                     H.HOMIF in ('01', '09')";
            }

            try
            {
                dataTable = DBS.executeSelectQueryNoParam(Employee_Table);
                ChangeDateFormat(dataTable, ShowAllEmployees);
                if (Filter_Date_Start) dataTable = FilterDateTable(dataTable);
                if (FilterDateLeft) dataTable = FilterDateLeftTable(dataTable);
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);
            }

            return dataTable;
        }



        private void ChangeDateFormat(DataTable dataTable, bool ShowAllEmployees)
        {
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                if (dataTable.Rows[i]["BIRTHDAY"].ToString() != "000000" && dataTable.Rows[i]["BIRTHDAY"].ToString() != "")
                    dataTable.Rows[i]["BIRTHDAY"] = DateTime.ParseExact(dataTable.Rows[i]["BIRTHDAY"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture).ToShortDateString();
                if (dataTable.Rows[i]["DateStart"].ToString() != "000000" && dataTable.Rows[i]["DateStart"].ToString() != "")
                    dataTable.Rows[i]["DateStart"] = DateTime.ParseExact(dataTable.Rows[i]["DateStart"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture).ToShortDateString();
                if (ShowAllEmployees == true)
                {
                    if (dataTable.Rows[i]["Left"].ToString() != "000000" && dataTable.Rows[i]["Left"].ToString() != "")
                        dataTable.Rows[i]["Left"] = DateTime.ParseExact(dataTable.Rows[i]["Left"].ToString(), "ddMMyy", System.Globalization.CultureInfo.InvariantCulture).ToShortDateString();
                }

            }
        }

        /// <summary>
        /// סינון תאריכי קליטה
        /// </summary>     
        private DataTable FilterDateTable(DataTable dataTable)
        {

            DataTable Filter_Date_Table = new DataTable();
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
            {
                if (dataTable.Rows[i]["DateStart"].ToString() != "000000" && dataTable.Rows[i]["DateStart"].ToString() != "")
                {
                    DataRow dr = dataTable.Rows[i];
                    if (DateTime.Parse(dataTable.Rows[i]["DateStart"].ToString()) <=DateTime.Parse( Date_Start2.ToShortDateString()) && DateTime.Parse(dataTable.Rows[i]["DateStart"].ToString()) >= DateTime.Parse(Date_Start.ToShortDateString())) continue;//עשיתי שינוי עם קיצור התאריך בשביל שלא יחשב עם שעות
                    else dr.Delete();
                }
                else { dataTable.Rows.Remove(dataTable.Rows[i]); }
            }
            dataTable.AcceptChanges();


            //DateTime t = new DateTime(2018, 10, 01);
            //ArrayList Filt = new ArrayList();
            //רשימת שורות שצריך למחוק
            //for (int i = 0; i < dataTable.Rows.Count; i++)
            //{
            //    if (dataTable.Rows[i]["DateStart"].ToString() != "000000" && dataTable.Rows[i]["DateStart"].ToString() != "")
            //    {
            //        if (DateTime.Parse(dataTable.Rows[i]["DateStart"].ToString()) < t)
            //            Filt.Add(i);

            //    }
            //    //else { dataTable.Rows.Remove(dataTable.Rows[i]); }
            //}
            //for (int j = 0; j < Filt.Count; j++)
            //{
            //    dataTable.Rows.Remove(dataTable.Rows[int.Parse(Filt[j].ToString())]);
            //    Filt[j + 1] = int.Parse(Filt[j + 1].ToString()) - 1;
            //}

            Filter_Date_Table = dataTable;
            return Filter_Date_Table;
        }

        /// <summary>
        /// סינון תאריכי עזיבה
        /// </summary>
        private DataTable FilterDateLeftTable(DataTable dataTable)
        {
            DataTable Filter_Date_Table = new DataTable();
                for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                {
                    if (dataTable.Rows[i]["Left"].ToString() != "000000" && dataTable.Rows[i]["Left"].ToString() != "")
                    {
                        DataRow dr = dataTable.Rows[i];
                        if (DateTime.Parse(dataTable.Rows[i]["Left"].ToString()) <= DateTime.Parse(Date_Left2.ToShortDateString()) && DateTime.Parse(dataTable.Rows[i]["Left"].ToString()) >= DateTime.Parse(Date_Left.ToShortDateString())) continue;//עשיתי שינוי עם קיצור התאריך בשביל שלא יחשב עם שעות
                        else dr.Delete();
                    }
                    else { dataTable.Rows.Remove(dataTable.Rows[i]); }
                }
            dataTable.AcceptChanges();
            Filter_Date_Table = dataTable;
            return Filter_Date_Table;
        }

        public DataTable TableAfterSearch(bool ShowAllEmployees, int Select_Search_Field, string TextSearch,string LastNameSearch)
        {
            DBService DBS = new DBService();
            DataTable dataTable = new DataTable();
            string Employee_Result = $@"  SELECT  right(trim(T.OVED),5) as Number ,L.PRATI as FirstName ,L.FAMILY as LastName,T.TEUDTZHUI as Id
                                         ,T.ADRESS,L.AVODA as DateStart,substring(L.AVODA,3,2) as MonthStart,T.TELEFON , case when substring(L.MAHLAKA,2,1)=1 then right(trim(L.MAHLAKA),3) when substring(L.MAHLAKA,2,1)=0 then  right(trim(L.MAHLAKA),2)
 	                                      end Unit,T.BIRTHDAY, H.HOMACH as EFirstName,H.HOMACTP as ELastName ,B.CDESC as UnitName ,case when substring(t.birthday,5,2)>40 then TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '19'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' ))) else TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '20'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' )))  end Age,
                                          case when substring(L.AVODA,5,2)>{DateTime.Now.Year.ToString().Substring(2, 2)} then  dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('19'||substring(L.AVODA,5,2)||'-'||substring(L.AVODA,3,2)||'-'||substring(L.AVODA,1,2)))/100) , 5,0))/ 100,5,2)                                          when substring(L.AVODA,5,2)<={DateTime.Now.Year.ToString().Substring(2, 2)} then dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('20'||substring(L.AVODA,5,2)||'-'||substring(L.AVODA,3,2)||'-'||substring(L.AVODA,1,2))) /100) , 5,0)) / 100 , 5, 2) end Vetek,
                                          case when substring(T.OVED,3,2)=19 then 'אישי' else 'קיבוצי' end typeContract, case when H.HOFIL3=2 then 'עקיף' when H.HOFIL3=1 then 'ישיר'  when H.HOFIL3=3 then 'מנהל' when H.HOFIL3=4 then 'עקיף חרושת' when H.HOFIL3=9 then 'פנסיה'  end type,case when h.hofil4=7 then 'זמני' else 'קבוע' end temp                                          FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED left join hsali.hovl02 as H on T.OVED=H.HOOVD left join  BPCSFV30. CDPL01  as B on right(trim(L.MAHLAKA),3)=B.CDEPT
                                           WHERE T.mifal='01' and (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129') and substring(T.kpa,7,6 )= '000000' and H.HOMIF in ('01', '09') and ";

            string All_Employee_Result = $@"SELECT  right(trim(T.OVED),5) as Number , name1||name2||name3||name4||name5||name6 as FirstName,                                          fmly1||fmly2||fmly3||fmly4||fmly5||fmly6||fmly7||fmly8||fmly9||fmly10 as LastName,T.TEUDTZHUI as Id,
                                         T.ADRESS,T.AVODA as DateStart,substring(L.AVODA,3,2) as MonthStart,T.TELEFON,case when substring(L.MAHLAKA,2,1)=1 then right(trim(L.MAHLAKA),3) when substring(L.MAHLAKA,2,1)=0 then  right(trim(L.MAHLAKA),2)
 	                                     end Unit,T.BIRTHDAY,case when right(trim(T.KPA),6) = '000000' then '' else right(trim(T.KPA),6) end left                                        , H.HOMACH as EFirstName,H.HOMACTP as ELastName ,B.CDESC as UnitName,case when substring(t.birthday,5,2)>40 then TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '19'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' ))) else TIMESTAMPDIFF(256, char(timestamp(current timestamp) - 
                                     timestamp( '20'||substring(T.BIRTHDAY,5,2)||'-'||substring(T.BIRTHDAY,3,2)||'-'||substring(T.BIRTHDAY,1,2)||' '||'00:00:00.000000' )))  end Age,
                                        case when substring(T.AVODA,5,2)>{DateTime.Now.Year.ToString().Substring(2, 2)} then  dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('19'||substring(T.AVODA,5,2)||'-'||substring(T.AVODA,3,2)||'-'||substring(T.AVODA,1,2)))/100) , 5,0))/ 100,5,2)                                        when substring(T.AVODA,5,2)<={DateTime.Now.Year.ToString().Substring(2, 2)} then dec((dec((((SELECT current date FROM sysibm.sysdummy1) - date('20'||substring(T.AVODA,5,2)||'-'||substring(T.AVODA,3,2)||'-'||substring(T.AVODA,1,2))) /100) , 5,0)) / 100 , 5, 2) end Vetek,
                                         case when substring(T.OVED,3,2)=19 then 'אישי' else 'קיבוצי' end typeContract, case when H.HOFIL3=2 then 'עקיף' when H.HOFIL3=1 then 'ישיר'  when H.HOFIL3=3 then 'מנהל' when H.HOFIL3=4 then 'עקיף חרושת' when H.HOFIL3=9 then 'פנסיה'  end type,case when h.hofil4=7 then 'זמני' else 'קבוע' end temp                                         FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED left join hsali.hovl02 as H on T.OVED=H.HOOVD left join  BPCSFV30. CDPL01  as B on right(trim(L.MAHLAKA),3)=B.CDEPT
                                         WHERE   (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129')  and (substring(T.kpa                                        ,11,2)  > '08' and (substring(T.kpa,11,2) < '80') or substring(T.kpa,7,6 )= '000000') and             	                         H.HOMIF in ('01', '09') ";
            bool Which_Str_Changed = false;//Employee_Result-false,All_Employee_Result=true
            switch (Select_Search_Field)
            {

                case 0:

                    if (ShowAllEmployees == false) Employee_Result += $@"right(trim(T.OVED),5)={TextSearch} ";
                    if (ShowAllEmployees == true) { All_Employee_Result += $@"and right(trim(T.OVED),5) ={TextSearch}"; Which_Str_Changed = true; }


                    break;

                case 1:
                    if (ShowAllEmployees == false) Employee_Result += $@"trim(L.PRATI) ='{TextSearch}' ";
                    if (ShowAllEmployees == true) { All_Employee_Result += $@"and trim(name1||name2||name3||name4||name5||name6)='{TextSearch}'"; Which_Str_Changed = true; }
                    break;

                case 2:

                    if (ShowAllEmployees == false) Employee_Result += $@"trim(L.PRATI) ='{TextSearch}' and trim(L.FAMILY)='{LastNameSearch}'";
                    if (ShowAllEmployees == true) { All_Employee_Result += $@"and trim(name1||name2||name3||name4||name5||name6)='{TextSearch}' and  trim(fmly1||fmly2||fmly3||fmly4||fmly5||fmly6||fmly7||fmly8||fmly9||fmly10)='{LastNameSearch}'"; Which_Str_Changed = true; }

                    break;

                case 3:

                    if (ShowAllEmployees == false) Employee_Result += $@"T.TEUDTZHUI={TextSearch}";
                    if (ShowAllEmployees == true) { All_Employee_Result += $@"and T.TEUDTZHUI={TextSearch}"; Which_Str_Changed = true; }

                    break;

                case 4:
                    if (TextSearch.Length == 2) TextSearch = "0" + TextSearch;//חייב 3 ספרות
                    if (ShowAllEmployees == false)
                        //if (TextSearch.Length == 2) TextSearch = "0" + TextSearch;//חייב 3 ספרות
                        Employee_Result += $@"right(trim(L.MAHLAKA),3)={TextSearch}  group by T.OVED  ,L.PRATI  ,L.FAMILY ,   T.ADRESS ,     T.TEUDTZHUI ,L.AVODA ,T.TELEFON, H.HOFIL3 , L.MAHLAKA ,T.BIRTHDAY ,H.HOMACH,H.HOMACTP,B.CDESC";
                    if (ShowAllEmployees == true) { All_Employee_Result += $@"and right(trim(L.MAHLAKA),3)={TextSearch} "; Which_Str_Changed = true; }///group by T.OVED  ,L.PRATI  ,L.FAMILY ,   T.ADRESS ,     T.TEUDTZHUI ,L.AVODA ,  H.HOFIL3 , L.MAHLAKA ,T.BIRTHDAY ,H.HOMACH,H.HOMACTP,B.CDESC; לא צריך כנראה תשמור שיהיה לך

                    break;

                case 5:
                    if (ShowAllEmployees == false) Employee_Result += $@"trim(L.FAMILY)='{TextSearch}'";
                    if (ShowAllEmployees == true) {All_Employee_Result += $@"trim(fmly1||fmly2||fmly3||fmly4||fmly5||fmly6||fmly7||fmly8||fmly9||fmly10)='{TextSearch}'"; Which_Str_Changed = true; }
                    break;


            }
            if (Which_Str_Changed) dataTable = DBS.executeSelectQueryNoParam(All_Employee_Result);
            else dataTable = DBS.executeSelectQueryNoParam(Employee_Result);
            ChangeDateFormat(dataTable, ShowAllEmployees);
            return dataTable;
        }

        public DataTable GetWorkStrengh()
        {
            DBService DBS = new DBService();
            DataTable dataTable = new DataTable();
            string Table_Work_Strengh =
                $@"  SELECT case when substring(B.DPPRF,0,6)='Mengi' then 'Maintenance engineering' else substring(B.DPPRF,0,9) end agaf , T.MAHLAKA as Unit ,B.CDESC UnitName,count(*) as Count
                     FROM isufkv.isav as T  inner join BPCSFV30.CDPL01 as B on right(trim(T.MAHLAKA),3)= B.CDEPT
                     WHERE T.mifal = '01' and (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129') and substring(T.kpa,7,6 )= '000000' and T.mahlaka*1 not in (196,197,198)
                     GROUP By B.DPPRF, T.MAHLAKA ,B.CDESC ,B.CDEPT
                     ORDER BY  substring(B.DPPRF,0,9) asc, T.MAHLAKA asc ";
            //$@"SELECT distinct case  when T.AGAF=0001 then 'לשכת מנכל' when T.AGAF=0002 THEN 'אגף כספים' when T.AGAF=0003 then 'אבטחת איכות' when T.AGAF=0006 or T.AGAF=0009 then 'אגף תפעול'  when T.AGAF=0004 then 'הנדסה' when T.AGAF=0008 then 'מופ'  when T.AGAF=0005 then 'אגף רכש ולוגיסטיקה'  when T.AGAF=0000 then 'פנסיונרים' when T.AGAF=9000 then 'לא ידוע' else t.agaf end agaf ,            //   T.MAHLAKA as Unit ,B.CDESC UnitName,count(*) as count            //   FROM isufkv.isav as T  inner join BPCSFV30.CDPL01 as B on right(trim(T.MAHLAKA),2)= B.CDEPT            //   WHERE T.mifal = '01'            //   GROUP By T.AGAF, T.MAHLAKA ,B.CDESC ,B.CDEPT";
            dataTable = DBS.executeSelectQueryNoParam(Table_Work_Strengh);
            return dataTable;
        }
        public DataTable GetSubTotalofEmployeeType()
        {
            DBService DBS = new DBService();
            DataTable TableForEmployeeType = new DataTable();
            string Table_Work_Strengh =
                  $@"SELECT  substring(B.DPPRF,0,9)  AS AGAF,T.MAHLAKA as Unit , H.HOFIL3 AS Type,count(*) as count
                     FROM isufkv.isav as T  inner join BPCSFV30.CDPL01 as B on right(trim(T.MAHLAKA),3)= B.CDEPT
	                 Left Join hsali.hovl02 as H on T.OVED=H.HOOVD
                      WHERE T.mifal = '01' and (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129') and substring(T.kpa,7,6 )= '000000' and T.mahlaka*1 not in (196,197,198) and ( right(trim(T.MAHLAKA),2)=H.HOMAH or right(trim(T.MAHLAKA),3)=H.HOMAH)
                     GROUP BY substring(B.DPPRF,0,9),T.MAHLAKA,H.HOFIL3
                     ORDER BY substring(B.DPPRF,0,9) asc,T.MAHLAKA asc ,H.HOFIL3  asc";

            TableForEmployeeType = DBS.executeSelectQueryNoParam(Table_Work_Strengh);
            return TableForEmployeeType;
        }

        public int CountPensioner()
        {

            DBService DBS = new DBService();
            DataTable TableForCount = new DataTable();
            int count = 0;
            string CountString =
               $@"SELECT count(*)
                 FROM isufkv.isav as T
                 WHERE T.mifal = '01' and(substring(T.OVED, 3, 5) < '26000'  or  substring(T.OVED, 3, 5) > '27129') and substring(T.kpa,7,6 )= '000000' and T.mahlaka * 1 not in (196, 197, 198) and t.agaf = '0000'";
            TableForCount = DBS.executeSelectQueryNoParam(CountString);
            count = int.Parse(TableForCount.Rows[0][0].ToString());
            return count;
        }

        public int CountNonPayment()
        {
            DBService DBS = new DBService();
            DataTable TableForCount = new DataTable();
            int count = 0;
            string CountString =
              $@"SELECT count(*)
                FROM isufkv.isav as T
                WHERE T.mifal = '01' and(substring(T.OVED, 3, 5) < '26000'  or  substring(T.OVED, 3, 5) > '27129') and substring(T.kpa,7,6 )= '000000' and T.mahlaka * 1  in (196, 197, 198)";
            TableForCount = DBS.executeSelectQueryNoParam(CountString);
            count = int.Parse(TableForCount.Rows[0][0].ToString());
            return count;
        }
        public void addWorkDay(WorkDay w) { workDays.Add(w); }
        
        public bool checkIfParamIsEmpNum(string s)
        {
            try
            {
                return int.Parse(s.Trim()) == Employee_Num;
            }
            catch (Exception ex) { }
            return false;
        }
        public Tuple<int,DateTime> howManyAbsenceInARow()
        {
            workDays.Reverse();
            int countDays = 0;
            DateTime lastDateAtWork = DateTime.MinValue;
            if (workDays.Count == 0) countDays = 35;
            foreach (WorkDay day in workDays)
            {
                if (day.checkIfWeekEnd()) {
                    countDays++;
                    continue;
                }
                if (!day.arriveToWork) { countDays++;}
                else
                {
                    if(lastDateAtWork==DateTime.MinValue)
                    lastDateAtWork = day.date;
                    break;
                }
            }
            workDays.Reverse();
            Tuple<int, DateTime> ret = new Tuple<int, DateTime>(countDays,lastDateAtWork);
            return ret;
        }


        //מחכה לעדכון טבלת HR
        public string getManagerMail()
        {
            return "ielia@atgtire.com";
        }
        public int CompareTo(Employee b)
        {
            return new DateTime(DateTime.Now.Year, this.Employee_Birthday.Month, this.Employee_Birthday.Day).CompareTo(new DateTime (DateTime.Now.Year , b.Employee_Birthday.Month, b.Employee_Birthday.Day));
        }
        public void WorkFinishForms()
        {
            makeFinishFormByPath("T:\\Application\\C#\\HR\\שבלונה סיום עבודה\\FinishForm1.docx");
            makeFinishFormByPath("T:\\Application\\C#\\HR\\שבלונה סיום עבודה\\FinishForm2.docx");
            makeFinishFormByPath("T:\\Application\\C#\\HR\\שבלונה סיום עבודה\\FinishForm3.docx");
        }
        private void makeFinishFormByPath(string fileName)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document aDoc = null;
            //string fileName = "C:\\Users\\ielia\\Desktop\\אלי שבלונה סיום עבודה\\FinishForm1.docx";
            if (File.Exists(fileName))
            {
                aDoc = wordApp.Documents.Open(fileName, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Word.Range range = aDoc.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "<name>", ReplaceWith: First_Name + " " + Last_Name, Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<id>", ReplaceWith: Id, Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<endDate>", ReplaceWith: this.Stop_Work==DateTime.MaxValue? DateTime.Now.ToShortDateString():this.Stop_Work.ToShortDateString(), Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<today>", ReplaceWith: DateTime.Now.ToString("dd MMMM yyyy"), Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<startDate>", ReplaceWith: Date_Start.ToShortDateString(), Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<empNum>", ReplaceWith: Employee_Num, Replace: Word.WdReplace.wdReplaceAll);
                //aDoc.Save();
                wordApp.Visible = true;
                releaseObject(wordApp);
                releaseObject(aDoc);

            }
        }


        internal void WorkStartFile()
        {
           MakeFileStartWork("T:\\Application\\C#\\HR\\שבלונה תחילת עבודה\\Start Work Format.docx");
        }
       
        /// <summary>
        /// טופס תחילת עבודה
        /// </summary>
        private void MakeFileStartWork(string FileName)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document aDoc = null;
            if (File.Exists(FileName))
            {
                aDoc = wordApp.Documents.Open(FileName, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Word.Range range = aDoc.Content;
               range.Find.ClearFormatting();
               range.Find.Execute(FindText: "<name>", ReplaceWith: First_Name + " " + Last_Name, Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<id>", ReplaceWith: Id, Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<today>", ReplaceWith: DateTime.Now.ToString("dd MMMM yyyy"), Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<startDate>", ReplaceWith: Date_Start.ToShortDateString(), Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<empNum>", ReplaceWith: Employee_Num, Replace: Word.WdReplace.wdReplaceAll);

                //aDoc.Save();
                wordApp.Visible = true;
                releaseObject(wordApp);
                releaseObject(aDoc);

            }
        }

        /// <summary>
        /// טופס הפסקת עבודה
        /// </summary>
        public void MakeFileStopWork(string FileName)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document aDoc = null;
            if (File.Exists(FileName))
            {
                aDoc = wordApp.Documents.Open(FileName, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Microsoft.Office.Interop.Word.Range range = aDoc.Content;
                range.Find.ClearFormatting();
                range.Find.Execute(FindText: "<name>", ReplaceWith: First_Name + " " + Last_Name, Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<empNum>", ReplaceWith: Employee_Num, Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<Unit>", ReplaceWith: Unit, Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<startDate>", ReplaceWith: Date_Start.ToShortDateString(), Replace: Word.WdReplace.wdReplaceAll);
                range.Find.Execute(FindText: "<endDate>", ReplaceWith: this.Stop_Work == DateTime.MaxValue ? DateTime.Now.ToShortDateString() : this.Stop_Work.ToShortDateString(), Replace: Word.WdReplace.wdReplaceAll);

                //aDoc.Save();
                wordApp.Visible = true;
                releaseObject(wordApp);
                releaseObject(aDoc);

            }
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

        public string getLastAbsenceReason()
        {
            
            workDays.Reverse();
            foreach(WorkDay w in workDays)
            {
                if (!w.arriveToWork) return w.absenceReason.Trim()!="" && w.absenceReason.Trim()!="נכח"?w.absenceReason:"לא ידוע";
            }
            return "לא ידוע";
        }

        private void SendBirthdayWish()
        {
            string qry = $@"SELECT  right(trim(T.OVED),5) as Number ,L.PRATI as FirstName ,L.FAMILY as LastName,T.TEUDTZHUI as Id,T.TELEFON,T.BIRTHDAY 
                          FROM isufkv.isav as T left join isufkv.isavl10 as L on T.OVED=L.OVED left join hsali.hovl02 as H on T.OVED=H.HOOVD left join  BPCSFV30. CDPL01  as B on right(trim(L.MAHLAKA),3)=B.CDEPT                          WHERE T.mifal='01' and (substring(T.OVED,3,5) < '26000'  or  substring(T.OVED,3,5) > '27129') and substring(T.kpa,7,6 )= '000000' and H.HOMIF in ('01', '09') and substring(T.BIRTHDAY,1,4)={DateTime.Now.Day.ToString()+DateTime.Now.Month.ToString()}";
            DataTable dataTable = new DataTable();
            DBService dBService = new DBService();
            dataTable=dBService.executeSelectQueryNoParam(qry);
        }


        /// <summary>
        /// מקבל מטבלה צדדית פרטים על נושאי כתבי מינוי במפעל
        /// </summary>
        internal DataTable GetSubscriptionNotes()
        {
            string qry = $@"SELECT *
                          FROM SubscriptionNotes_Hr
                          ORDER BY Subject";
            DbServiceSQL dbServiceSQL = new DbServiceSQL();
            DataTable dataTable = new DataTable();
            dataTable = dbServiceSQL.executeSelectQueryNoParam(qry);
            return dataTable;
        }
    }
}
