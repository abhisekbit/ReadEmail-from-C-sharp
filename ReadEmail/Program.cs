using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ReadEmail
{
    class Program
    {
        
        static void Main(string[] args)
        {
            Outlook.NameSpace outlookNameSpace;
            Outlook.MAPIFolder inbox;
            Outlook.Items items;
            Outlook.Application app = new Outlook.Application();
            outlookNameSpace = app.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            items = inbox.Items;
            Console.WriteLine("Waiting for new Message");
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
            Console.ReadKey();
        }
        public static void items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                string tempvariable = mail.To;
                Console.WriteLine("Mail To : " + tempvariable);
                tempvariable = mail.Subject;
                Console.WriteLine("Subject : " + tempvariable);
                tempvariable = mail.Body;
                Console.WriteLine("Body    : " + tempvariable);
                tempvariable = mail.SenderEmailAddress;
                Console.WriteLine("Sender Email : " + tempvariable);
                tempvariable = mail.SenderName.ToString();
                Console.WriteLine("Sender Name : " + tempvariable);
            }
        }
    }
}

