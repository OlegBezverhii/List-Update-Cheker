using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Dism;
using WUApiLib;
using System.Xml;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.IO;
using System.Management;

namespace WindowsUpdate
{
    class Update
    {
        public static List<string> XMLlist()
        {
            List<string> result = new List<string>(220);

            //http://www.outsidethebox.ms/17988/

            //регулярное выражение для поиска по маске KB
            Regex regex = new Regex(@"KB[0-9]{6,7}");
            //Regex(@"(\w{2}\d{6,7}) ?");

            //SortedSet не поддерживает повторяющиеся элементы, поэтмоу повторяющиеся элементы мы "группируем" ещё на стадии добавления
            SortedSet<string> spisok = new SortedSet<string>();
            XmlDocument xDoc = new XmlDocument();

            try
            {
                string path = "C:\\Windows\\servicing\\Packages\\wuindex.xml"; //путь до нашего xml
                xDoc.Load(path);
            }

            catch (Exception ex)
            {
                result.Add("Что-то пошло не так: " + ex.Message);
            }

            int kol = 0; //кол-во компонентов
            int total2 = 0; //кол-во полученных обновлений

            XmlNodeList name = xDoc.GetElementsByTagName("Mappings");
            foreach (XmlNode xnode in name)
            {
                kol++;
                XmlNode attr = xnode.Attributes.GetNamedItem("UpdateId");

                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    XmlNode childattr = childnode.Attributes.GetNamedItem("Package");

                    MatchCollection matches = regex.Matches(childattr.Value);
                    if (matches.Count > 0)
                    {
                        foreach (Match match in matches)
                            //Console.WriteLine(match.Value);
                            spisok.Add(match.Value);
                    }
                    else
                    {
                        //Console.WriteLine("Совпадений не найдено");
                    }
                }

            }

            foreach (string element in spisok)
            {
                //Console.WriteLine(element);
                result.Add(element);
                total2++;
            }

            result.Add("Количество пакетов: " + kol);                        
            result.Add("Количество KB обновлений: " + total2);

            return result;
        }

        public static List<string> listUpdateHistory()
        {
            //WUApi
            List<string> result = new List<string>(200);

            try
            {
                UpdateSession uSession = new UpdateSession();
                IUpdateSearcher uSearcher = uSession.CreateUpdateSearcher();
                uSearcher.Online = false;
                ISearchResult sResult = uSearcher.Search("IsInstalled=1 And IsHidden=0");

                string sw = "Количество обновлений через WUApi: " + sResult.Updates.Count;
                result.Add(sw);
                foreach (IUpdate update in sResult.Updates)
                {
                    result.Add(update.Title);
                }
            }

            catch (Exception ex)
            {
                result.Add("Что-то пошло не так: " + ex.Message);
            }
            
            return result;
        }

        public static List<string> DISMlist()
        {
            List<string> result = new List<string>(220);

            try
            {
                DismApi.Initialize(DismLogLevel.LogErrors);
                var dismsession = DismApi.OpenOnlineSession();
                var listupdate = DismApi.GetPackages(dismsession);

                int ab = listupdate.Count;
                //Console.WriteLine("Количество обновлений через DISM: " + ab);
                string sw = "Количество обновлений через DISM: " + ab;
                result.Add(sw);

                foreach (DismPackage feature in listupdate)
                {
                    result.Add(feature.PackageName);
                    //result.Add($"[Имя пакета] {feature.PackageName}");
                    //result.Add($"[Дата установки] {feature.InstallTime}");
                    //result.Add($"[Тип обновления] {feature.ReleaseType}");
                }
            }

            catch (Exception ex)
            {
                result.Add("Что-то пошло не так: " + ex.Message);
            }

            return result;
        }

        public static List<string> Sessionlist(string pc)
        {
            List<string> result = new List<string>(50); //не забудь изменить количество

            object sess = null;
            object search = null;
            object coll = null;

            try
            {
                sess = Activator.CreateInstance(Type.GetTypeFromProgID("Microsoft.Update.Session", pc));
                search = (sess as dynamic).CreateUpdateSearcher();

                int n = (search as dynamic).GetTotalHistoryCount();
                int kol = 0;
                //coll = (search as dynamic).QueryHistory(1, n);
                coll = (search as dynamic).QueryHistory(0, n);

                result.Add("Количество через Update.Session: " + n);
                foreach (dynamic item in coll as dynamic)
                {
                    if (item.Operation == 1) result.Add(item.Title);
                    kol++;
                    //Console.WriteLine("Количество: " + kol);
                }
                result.Add("Количество в цикле: " + kol);
            }
            catch (Exception ex)
            {
                result.Add("Что-то пошло не так: " + ex.Message);
            }
            finally
            {
                if (sess != null) Marshal.ReleaseComObject(sess);
                if (search != null) Marshal.ReleaseComObject(search);
                if (coll != null) Marshal.ReleaseComObject(coll);
            }

            return result;
        }

        public static List<string> GetWMIlist(params string[] list)
        {
            List<string> result = new List<string>(200); //не забудь изменить количество

            ManagementScope Scope;

            string ComputerName = list[0];
            string Username = list[1];
            string Password = list[2];

            int kol = 0;

            if (!ComputerName.Equals("localhost", StringComparison.OrdinalIgnoreCase))
            {
                //     Возвращает или задает полномочия, которые используются для проверки подлинности
                //     указанного пользователя.
                ConnectionOptions Conn = new ConnectionOptions();
                Conn.Username = Username;
                Conn.Password = Password;
                //Если значение свойства начинается со строки «NTLMDOMAIN:» аутентификация NTLM будет использоваться, и свойство должно содержать доменное имя NTLM.
                Conn.Authority = "ntlmdomain:DOMAIN";
                Scope = new ManagementScope(String.Format("\\\\{0}\\root\\CIMV2", ComputerName), Conn);
            }
            else
                Scope = new ManagementScope(String.Format("\\\\{0}\\root\\CIMV2", ComputerName), null);

            try
            {
                Scope.Connect();
                ObjectQuery Query = new ObjectQuery("SELECT * FROM Win32_QuickFixEngineering");
                ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Scope, Query);

                foreach (ManagementObject WmiObject in Searcher.Get())
                {
                    result.Add(WmiObject["HotFixID"].ToString());
                    //Console.WriteLine("{0,-35} {1,-40}", "HotFixID", WmiObject["HotFixID"]);// String
                    //result.Add();
                    /*result.Add("{0,-17} {1}", "Тип обновления: ", WmiObject["Description"]);
                    result.Add("{0,-17} {1}", "Ссылка: ", WmiObject["Caption"]);
                    result.Add("{0,-17} {1}", "Дата установки: ", WmiObject["InstalledOn"]);*/
                    kol++;
                }
                result.Add("Количество равно " + kol);
            }

            catch (Exception ex)
            {
                result.Add("Что-то пошло не так: " + ex.Message);
            }
            
            return result;
        }

        public static List<string> GetWSUSlist(params string[] list)
        {
            List<string> result = new List<string>(200); //не забудь изменить количество

            string ComputerName = list[0];
            string Username = list[1];
            string Password = list[2];

            try
            {

            }

            catch (Exception ex)
            {
                result.Add("Что-то пошло не так: " + ex.Message);
            }

            return result;
        }


    }
}
