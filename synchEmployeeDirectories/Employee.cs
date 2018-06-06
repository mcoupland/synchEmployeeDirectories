using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace synchEmployeeDirectories
{
    public class Employee
    {
        private string lname;
        public string Lname { get => lname; set => lname = value; }
        private string fname;
        public string Fname { get => fname; set => fname = value; }
        private string initial;
        public string Initial { get => initial; set => initial = value; }
        private string suffix;
        public string Suffix
        {
            get { return suffix; }
            set
            {
                suffix = string.IsNullOrEmpty(value) ? "(none)" : suffix;
            }
        }
        private string nickname;
        public string Nickname { get => nickname; set => nickname = value; }
        private string id;
        public string Id { get => id; set => id = value; }
        private string jobtitle;
        public string Jobtitle { get => jobtitle; set => jobtitle = value; }
        private string division;
        public string Division { get => division; set => division = value; }
        private string department;
        public string Department
        {
            get { return department; }
            set
            {
                if (value.Contains("-"))
                {
                    department = value.Split('-')[1];
                }
                else
                {
                    department = value;
                }
            }
        }
        private string workphone;
        public string Workphone { get => workphone; set => workphone = GetPhoneNumber(value); }
        private string workfax;
        public string Workfax { get => workfax; set => workfax = GetPhoneNumber(value); }
        private string workwireless;
        public string Workwireless { get => workwireless; set => workwireless = GetPhoneNumber(value); }
        private string employeestatus;
        public string Employeestatus { get { return employeestatus == "A" ? "Active" : employeestatus; } set => employeestatus = value; }
        private string mgrlname;
        public string Mgrlname { get => mgrlname; set => mgrlname = value; }
        private string mgrfname;
        public string Mgrfname { get => mgrfname; set => mgrfname = value; }
        private string mgrmname;
        public string Mgrmname { get => mgrmname; set => mgrmname = value; }
        private string homedepartment;
        public string Homedepartment {
            get { return homedepartment; }
            set
            {
                if (value.Contains("-"))
                {
                    homedepartment = value.Split('-')[0];
                }
                else
                {
                    homedepartment = value;
                }
            }
        }

        private string GetPhoneNumber(string input)
        {
            return input.Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "");
        }
    }
}
