using ClosedXML.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DemoClosedXML.Model
{
    public class Student
    {
       
            [XLColumn(Header = "Mã Sinh Viên ")]
            public String Id { get; set; }
             [XLColumn(Header = "Họ Tên")]
            public String Name { get; set; }

        [XLColumn(Header = "Tuổi")]
        public Int32 Age { get; set; }
        [XLColumn(Header = "Ngành")]
        public string Major { get; set; }
        [XLColumn(Header = "Gio ")]
        public DateTime Date { get { return DateTime.Now; }}


    }
}
