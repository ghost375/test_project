using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Web.Script.Serialization;
using ExApp = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using System.Net;

namespace TestProject1
{ 
    class Program
    {

        static public List<Contact> GetContacts(long totalOffset)
        {
            bool hasMore = false;
            long timeOffset = 0;
            long vidOffset = 0;
            List<Contact> contacts = new List<Contact>();

            do
            {
                string url;
                if (timeOffset == 0)
                {
                    url = "https://api.hubapi.com/contacts/v1/lists/recently_updated/contacts/recent?hapikey=demo";
                }
                else
                {
                    url = String.Format("https://api.hubapi.com/contacts/v1/lists/recently_updated/contacts/recent?hapikey=demo&timeOffset={0}&vidOffset={1}", timeOffset, vidOffset);
                }

                try
                {
                    HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                    myHttpWebRequest.Timeout = 30000;
                    myHttpWebRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0";
                    myHttpWebRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                    myHttpWebRequest.Headers.Add("Accept-Language", "ru");
                    HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

                    using (var reader = new StreamReader(myHttpWebResponse.GetResponseStream()))
                    {
                        var content = reader.ReadToEnd();
                        var deserializer = new JavaScriptSerializer();
                        var res = deserializer.Deserialize<dynamic>(content);
                        hasMore = res["has-more"];
                        timeOffset = res["time-offset"];
                        vidOffset = res["vid-offset"];
                        var tempContacts = res["contacts"];
                        foreach (var contact in tempContacts)
                        {
                            Contact tempContact = new Contact
                            {
                                Vid = contact["vid"],
                                FirstName = (contact["properties"].ContainsKey("firstname")) ? contact["properties"]["firstname"]["value"] : null,
                                LastName = (contact["properties"].ContainsKey("lastname")) ? contact["properties"]["lastname"]["value"] : null
                            };

                            contacts.Add(tempContact);
                        }
                    }
                    myHttpWebResponse.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());

                }
                
            }
            while ((timeOffset > totalOffset));
            return contacts;
        }

        static Contact GetCompany(Contact contact)
        {
            string url = "https://api.hubapi.com/contacts/v1/contact/vid/"+contact.Vid+"/profile?hapikey=demo";
            try
            {
                HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                myHttpWebRequest.Timeout = 30000;
                myHttpWebRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0";
                myHttpWebRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                myHttpWebRequest.Headers.Add("Accept-Language", "ru");
                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

                using (var reader = new StreamReader(myHttpWebResponse.GetResponseStream()))
                {
                    var content = reader.ReadToEnd();
                    var deserializer = new JavaScriptSerializer();
                    var res = deserializer.Deserialize<dynamic>(content);

                    contact.LifeCycleStage =(res["properties"].ContainsKey("lifecyclestage")) ? res["properties"]["lifecyclestage"]["value"] : null;
                    if(res.ContainsKey("associated-company"))// ? contact["properties"]["firstname"]["value"] : null,
                    {
                        Company company = new Company
                        {
                            Id = (res["associated-company"].ContainsKey("company-id")) ? res["associated-company"]["company-id"] : -1,
                            Name = (res["associated-company"]["properties"].ContainsKey("name")) ? res["associated-company"]["properties"]["name"]["value"] : null,
                            Phone = (res["associated-company"]["properties"].ContainsKey("phone")) ? res["associated-company"]["properties"]["phone"]["value"] : null,
                            City = (res["associated-company"]["properties"].ContainsKey("city")) ? res["associated-company"]["properties"]["city"]["value"] : null,
                            State = (res["associated-company"]["properties"].ContainsKey("state")) ? res["associated-company"]["properties"]["state"]["value"] : null,
                            WebSite = (res["associated-company"]["properties"].ContainsKey("website")) ? res["associated-company"]["properties"]["website"]["value"] : null,
                            Zip = (res["associated-company"]["properties"].ContainsKey("zip")) ? res["associated-company"]["properties"]["zip"]["value"] : null
                        };
                        contact.Company = company;
                    }

                }
                myHttpWebResponse.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());

            }
            return contact; 


        }

        static List<Contact> A(DateTime modifiedOnOrAfter)
        {
            long totalOffset = (Int64)modifiedOnOrAfter.ToUniversalTime().Subtract(
            new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;

            var contacts = GetContacts(totalOffset);
            List<Contact> resultList = new List<Contact>();
            foreach (var contact in contacts)
            {
                resultList.Add(GetCompany(contact));
            }

            return resultList;


        }

        static void B(List<Contact> contacts)
        {
            ExApp exApp = new ExApp { DisplayAlerts = false };
            Workbook workBook;
            Worksheet worksheet;

            string template = "template.xlsm";
            
            workBook = exApp.Workbooks.Open(Path.Combine(Environment.CurrentDirectory, template));
            worksheet = workBook.ActiveSheet as Worksheet;
            worksheet.Range["A1"].Value = "Vid";
            worksheet.Range["B1"].Value = "First Name";
            worksheet.Range["C1"].Value = "Last Name";
            worksheet.Range["D1"].Value = "LifeCycleStage";
            worksheet.Range["E1"].Value = "Company ID";
            worksheet.Range["F1"].Value = "Company Name";
            worksheet.Range["G1"].Value = "Company Phone";
            worksheet.Range["H1"].Value = "Company City";
            worksheet.Range["I1"].Value = "Company State";
            worksheet.Range["J1"].Value = "Company Zip";
            worksheet.Range["K1"].Value = "Company WebSite";
            worksheet.Range["A1","K1"].ColumnWidth = 15;
            for (int i = 2; i <= contacts.Count+1; i++)
            {

                worksheet.Cells[i, 1] = contacts.ElementAt(i - 2).Vid;
                worksheet.Cells[i, 2] = (contacts.ElementAt(i - 2).FirstName != null) ? contacts.ElementAt(i - 2).FirstName : null;
                worksheet.Cells[i, 3] = (contacts.ElementAt(i - 2).LastName != null) ? contacts.ElementAt(i - 2).LastName : null;
                worksheet.Cells[i, 4] = (contacts.ElementAt(i - 2).LifeCycleStage != null) ? contacts.ElementAt(i - 2).LifeCycleStage : null;
                if (contacts.ElementAt(i - 2).Company != null)
                {
                    worksheet.Cells[i, 5] = contacts.ElementAt(i - 2).Company.Id;
                    worksheet.Cells[i, 6] = (contacts.ElementAt(i - 2).Company.Name != null) ? contacts.ElementAt(i - 2).Company.Name : null;
                    worksheet.Cells[i, 7] = (contacts.ElementAt(i - 2).Company.Phone != null) ? contacts.ElementAt(i - 2).Company.Phone : null;
                    worksheet.Cells[i, 8] = (contacts.ElementAt(i - 2).Company.City != null) ? contacts.ElementAt(i - 2).Company.City : null;
                    worksheet.Cells[i, 9] = (contacts.ElementAt(i - 2).Company.State != null) ? contacts.ElementAt(i - 2).Company.State : null;
                    worksheet.Cells[i, 10] = (contacts.ElementAt(i - 2).Company.Zip != null) ? contacts.ElementAt(i - 2).Company.Zip : null;
                    worksheet.Cells[i, 11] = (contacts.ElementAt(i - 2).Company.WebSite != null) ? contacts.ElementAt(i - 2).Company.WebSite : null;
                }
            }

            exApp.Visible = true;
        }

        static void Main(string[] args)
        {
            Console.Write("Введите дату(dd.MM.yyyy): ");
            string inputData = Console.ReadLine();
            Console.WriteLine("Подождите, идет выполнение...");
            var contacts = A(Convert.ToDateTime(inputData));
            
            B(contacts);
            Console.WriteLine("Готово!");
            Console.ReadKey();
        }
    }
}
