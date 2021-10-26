using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Linq;

namespace enrolleeSorter
{

    class FullEnrollee
    {
        public string ID { get; set; }
        public string First { get; set; }
        public string Last { get; set; }
        public string Version { get; set; }
        public string Company { get; set; }

    }    
     
    class Program
    {

        public void loadCsvFile(string filePath)
        {
         string[] csvlines = File.ReadAllLines(@filePath);

            //full data
            var query = from csvline in csvlines
            let data = csvline.Split(',')
            select new
            {
            ID = data[0],
            Fullname = data[1],
            Version = data[2],
            Company = data[3]
            };

            
            //get dupes
            var dupeGroup = from c in query 
            group c by c.ID into grp
            where grp.Count() > 1
            select new{grp.Key};

            //get highest version
            var groupedID = from u in query
            from c in dupeGroup 
            where u.ID == c.Key
            orderby u.Version descending
            select  u
            ;

            //limit result to 1 per group
            var sortedByVersion = groupedID.GroupBy(x => x.ID)
            .Select(g => g.OrderBy(x => x.ID).FirstOrDefault());
            
            //noDupe
            var noDupeGroup = query.Where(x=> !dupeGroup.Any(h => h.Key == x.ID));
            
            //merge the two
            var allGroup = sortedByVersion.Concat(noDupeGroup);

            
            //split first and last name
            List<FullEnrollee> lst = new List<FullEnrollee>();
            foreach(var item in allGroup)
            {
                FullEnrollee curr = new FullEnrollee();
                curr.ID = item.ID;
                string[] fullname = item.Fullname.Split(' ');
                curr.First = fullname[0];
                curr.Last = fullname[1];
                curr.Version = item.Version;
                curr.Company = item.Company;
                lst.Add(curr);
           
            }

            //sort by company, last name, first name
            List<FullEnrollee> sorted = lst.OrderBy(x => x.Company)
                                    .ThenBy(x => x.Last)
                                    .ThenBy(x => x.First)
                                    .ToList();

            var currCo = "";
            var writeFilePath = "";
            string stringForFile = "";
            

            foreach(FullEnrollee f in sorted)
            {
                if(currCo != f.Company)
                {
                    writeFilePath=f.Company + "-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm");
                }
                    stringForFile = f.ID + "," + f.First + " " + f.Last + "," + f.Version + "," + f.Company;
                    File.AppendAllText(writeFilePath+".csv", stringForFile + Environment.NewLine);

            }

 
        }
        static void Main(string[] args)
        {
            Program reader = new Program();
            reader.loadCsvFile("enrollees.csv");
            Console.WriteLine("File has been consumed!");
        }
    }
}
