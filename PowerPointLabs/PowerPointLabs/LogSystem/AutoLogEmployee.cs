using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Practices.Unity;

namespace AutoLog
{

    public interface IEmployee
    {
        void Work();
    }

    public class Employee : IEmployee
    {

        public Employee()
        {  
        }

        public string Name { get; set; }

        [AutoLogCallHandler()]
        public void Work()
        {
            Console.WriteLine("Now is {0},{1} is working hard!", DateTime.Now.ToShortTimeString(), Name);
            throw new Exception("Customer Exception");
        }

        [AutoLogCallHandler()]
        public override string ToString()
        {
            return string.Format("This is {0}.", Name);
        }
    }
}