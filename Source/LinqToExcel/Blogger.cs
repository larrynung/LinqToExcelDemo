using System;

namespace ConsoleApplication1
{
    class Blogger
    {
        public int ID { get; set; }
        public String FirstName { get; set; }
        public String LastName { get; set; }
        public SexType Sex { get; set; }
        public int Age { get; set; }
        public String Blog { get; set; }

        public override string ToString()
        {
            return string.Join(",", new string[] { ID.ToString(), FirstName, LastName, Sex.ToString(), Age.ToString(), Blog });
        }
    }
}
