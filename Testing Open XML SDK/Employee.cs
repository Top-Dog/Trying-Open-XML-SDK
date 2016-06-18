using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Testing_Open_XML_SDK
{
    /* Class to represent an employee
     * Every employee has: Id, Name, date of birth,
     * and salary.
     */
    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime DOB { get; set; }
        public decimal Salary { get; set; }
    }

    /* Wrapper class to initialise a list
     * of employees.
     */
    public sealed class Employees
    {
        static List<Employee> _employees;
        const int COUNT = 15;

        public static List<Employee> EmployeesList
        {
            private set {}
            get
            {
                return _employees;
            }
        }

        static Employees()
        {
            Initialize();
        }

        private static void Initialize()
        {
            _employees = new List<Employee>();
            Random random = new Random();

            for (int i=0; i < COUNT; i++)
            {
                _employees.Add(new Employee()
                    {
                        Id = 1,
                        Name = "Employee " + i,
                        DOB = new DateTime(1999, 1, 1).AddMonths(i),
                        Salary = random.Next(100, 10000)
                    });

            }
        }
    }
}
