using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
//using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using FinalFrontierLearnLib;
using System.IO;

namespace FinalFrontierLearner
{
    class Program
    {
        //private Dictionary<string, int> DictSenderName = new Dictionary<string, int>();
        //private Dictionary<string, int> DictSenderEmail = new Dictionary<string, int>();
        //private Dictionary<string, int> DictSenderCombo = new Dictionary<string, int>();

        //private DictionaryTools dt = new DictionaryTools();

        //private List<string> FolderList = new List<string>();

        //private int cnt_folder = 0;
        //private int cnt_mails = 0;
        //private string userpath;

        //private string[] badfolders = { "Junk", "Unwanted", "Trash", "Spam", "Posteingang", "Inbox" };

        //public void GetFolders(Outlook.Folder folder)
        //{
        //    Outlook.Folders childFolders =
        //        folder.Folders;
        //    if (childFolders.Count > 0)
        //    {
        //        foreach (Outlook.Folder childFolder in childFolders)
        //        {
        //            //Console.WriteLine(childFolder.FolderPath);
        //            FolderList.Add(childFolder.FolderPath);
        //            cnt_folder++;
        //            GetFolders(childFolder);
        //        }
        //    }
        //}

        //public void LearnFolders(Outlook.Folder folder, Boolean learn, int folderid)
        //{
        //    Outlook.Folders childFolders = folder.Folders;
        //    Boolean learning = learn;
        //    if (childFolders.Count > 0)
        //    {
        //        foreach (Outlook.Folder childFolder in childFolders)
        //        {
        //            if (childFolder.FolderPath.Contains(FolderList[folderid]))
        //            {
        //                learning = true;
        //            }
        //            else
        //            {
        //                learning = false;
        //            }
        //            foreach (string badfolder in badfolders)
        //            {
        //                if (childFolder.FolderPath.Contains(badfolder))
        //                {
        //                    learning = false;
        //                }
        //            }
        //            cnt_folder++;
        //            if (learning == true)
        //            {
        //                Console.WriteLine("learning from " + childFolder.FolderPath);
        //                try
        //                {
        //                    Items mails = childFolder.Items;
        //                    foreach (object mail in mails)
        //                    {
        //                        try
        //                        {
        //                            Outlook.MailItem thismail = (mail as Outlook.MailItem);
        //                            string senderName = thismail.SenderName;
        //                            string senderEmailAddress = thismail.SenderEmailAddress;
        //                            string senderCombo = senderName + "/" + senderEmailAddress;
        //                            if (DictSenderName.ContainsKey(senderName))
        //                                DictSenderName[senderName] = DictSenderName[senderName] + 1;
        //                            else
        //                                DictSenderName.Add(senderName, 1);
        //                            if (DictSenderEmail.ContainsKey(senderEmailAddress))
        //                                DictSenderEmail[senderEmailAddress] = DictSenderEmail[senderEmailAddress] + 1;
        //                            else
        //                                DictSenderEmail.Add(senderEmailAddress, 1);
        //                            if (DictSenderCombo.ContainsKey(senderCombo))
        //                                DictSenderCombo[senderCombo] = DictSenderCombo[senderCombo] + 1;
        //                            else
        //                                DictSenderCombo.Add(senderCombo, 1);
        //                            cnt_mails++;
        //                        }
        //                        catch (System.Exception ex)
        //                        {
        //                            Debug.WriteLine(ex.Message);
        //                        }
        //                    }
        //                }
        //                catch (System.Exception)
        //                { }
        //            }
        //            else
        //            {
        //                Console.WriteLine("Skipping folder " + childFolder.FolderPath);
        //            }
                    
        //            LearnFolders(childFolder, false, folderid);
        //        }
        //    }
        //    userpath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        //    dt.Write(DictSenderName, userpath + "\\dict-sender-name.bin");
        //    dt.Write(DictSenderEmail, userpath + "\\dict-sender-email.bin");
        //    dt.Write(DictSenderCombo, userpath + "\\dict-sender-combo.bin");

        //}

        static void Main(string[] args)
        {
            Outlook.Application outlook = new Outlook.Application();
            Outlook.Folder root = outlook.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            Learn learn = new Learn();

            int folderid = 0;
            var userpath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\FinalFrontier";
            
            Console.WriteLine("---AVAILABLE FOLDERS---");
            learn.GetFolders(root);
            for (int i = 0; i < learn.FolderList.Count; i++)
            {
                Console.WriteLine(i + ". " + learn.FolderList[i]);
            }
            Console.Write("Please enter the number (without trailing .) of the folder you want to learn from recursively: ");
            try
            {
                folderid = short.Parse(Console.ReadLine());
            }
            catch (System.Exception)
            {
                Console.WriteLine("could not read the selected folder. exiting...");
                return;
            }

            Console.WriteLine();
            
            Console.WriteLine("---LEARNING MAIL HISTORY---");
            learn.LearnFolders(root, folderid);
            Console.WriteLine("learned " + " mails  recursively from, starting in " + learn.FolderList[folderid] + ".");

            Console.WriteLine("dictionary files have been written to " + "... keep these files where they are so that FinalFrontier can find them.");

            DictionaryTools dtLearn = new DictionaryTools();

            Console.WriteLine("---VERIFYING---");

            var result = new Dictionary<string, int>();

            foreach (var file in Directory.GetFiles(userpath))
            {
                if (file.EndsWith("-dict-sender-email.bin") ||
                    file.EndsWith("-dict-sender-name.bin") ||
                    file.EndsWith("-dict-sender-combo.bin")) { 
                    foreach (var values in dtLearn.Read(file))
                        if (!result.ContainsKey(values.Key))
                            result.Add(values.Key, values.Value);
                    
                    Console.WriteLine($"{Path.GetFileName(file)}: " + dtLearn.Read(file).Count() + " entries");
                }
            }
            
            //Console.WriteLine("dict-sender-name.bin: " + dtLearn.Read(userpath + "\\dict-sender-name.bin").Count() + " entries");
            //Console.WriteLine("dict-sender-email.bin: " + dtLearn.Read(userpath + "\\dict-sender-email.bin").Count() + " entries");
            //Console.WriteLine("dict-sender-combo.bin: " + dtLearn.Read(userpath + "\\dict-sender-combo.bin").Count() + " entries");

            Console.WriteLine("[hit key to exit]");

            Console.ReadKey();
        }
    }

}