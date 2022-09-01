using System;
using System.Xml;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using System.Net.Mail;

namespace MorphoTrak_XML_Parser
{
    class Program
    {
        //Global Variables
        static string _connectionString = ConfigurationManager.AppSettings["Connection_String"].ToString();
        static string _processingPath = ConfigurationManager.AppSettings["Processing_Directory"].ToString();
        static string _errorPath = ConfigurationManager.AppSettings["Error_Directory"].ToString();
        static string _completePath = ConfigurationManager.AppSettings["Complete_Directory"].ToString();
        static string _unprocessedPath = ConfigurationManager.AppSettings["Unprocessed_Directory"].ToString();
        static string _emailRecipient = ConfigurationManager.AppSettings["Email_Recipient"].ToString();
        static string _emailSender = ConfigurationManager.AppSettings["Email_Sender"].ToString();
        static string _smtpServer = ConfigurationManager.AppSettings["SMTP_Server"].ToString();
        static string _subject = "Error with MorphoTrak XML Parser";
        static string _file = null;
        static string _fileName = null;
        static string _fileContents = null;

        //Load XML file into filestream then save into a string. The aforementioned string will be passed to ParseXML and the output of that will be passed to ConstructSQL
        public static void LoadFile()
        {
            //Variables for output of ParseXML
            string outNameFirst, outNameMiddle, outNameFamily, outNameSuffix, outCodeGender, outCodeEyeColor, outCodeHairColor, outCurrentWeight, outCurrentHeight;
            DateTime? outDateOfBirth;

            try
            {
                //Load XML doc from specified directory
                foreach (string f in Directory.EnumerateFiles(_unprocessedPath, "*.xml"))
                {
                    _file = _processingPath + "\\" + Path.GetFileName(f);
                    _fileName = Path.GetFileName(_file);
                    File.Move(f, _file);

                    //Load the XML doc into a filestream then save it into a string
                    using (FileStream stream = File.OpenRead(_file))
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            _fileContents = reader.ReadToEnd();
                        }
                    }
                    ParseXML(_fileContents, out outNameFirst, out outNameMiddle, out outNameFamily, out outNameSuffix, out outDateOfBirth, out outCodeGender, out outCodeEyeColor, out outCodeHairColor, out outCurrentWeight, out outCurrentHeight);
                    ConstructSQL(outNameFirst, outNameMiddle, outNameFamily, outNameSuffix, outDateOfBirth, outCodeGender, outCodeEyeColor, outCodeHairColor, outCurrentWeight, outCurrentHeight);
                }
            }
            catch (NullReferenceException NRE)
            {
                Console.WriteLine(NRE.StackTrace.ToString() + "\n" + NRE.Message.ToString() + "\n");
                if (_file != null)
                {
                    File.Move(_processingPath + "\\" + _fileName, _errorPath + "\\" + _fileName);
                    File.AppendAllText(_errorPath + "\\" + "log.txt", DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + NRE.Message.ToString() + "\n" + NRE.StackTrace.ToString() + "\n");
                    string body = DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + NRE.Message.ToString() + "\n" + "\n" + NRE.StackTrace.ToString() + "\n";
                    MailMessage message = new MailMessage(_emailSender, _emailRecipient, _subject, body);
                    SmtpClient client = new SmtpClient(_smtpServer);
                    client.Send(message);
                }
                LoadFile();
            }
            catch (UnauthorizedAccessException UA)
            {
                Console.WriteLine(UA.StackTrace.ToString() + "\n" + UA.Message.ToString() + "\n");
                if (_file != null)
                {
                    File.Move(_processingPath + "\\" + _fileName, _errorPath + "\\" + _fileName);
                    File.AppendAllText(_errorPath + "\\" + "log.txt", DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + UA.Message.ToString() + "\n" + UA.StackTrace.ToString() + "\n");
                    string body = DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + UA.Message.ToString() + "\n" + "\n" + UA.StackTrace.ToString() + "\n";
                    MailMessage message = new MailMessage(_emailSender, _emailRecipient, _subject, body);
                    SmtpClient client = new SmtpClient(_smtpServer);
                    client.Send(message);
                }
                LoadFile();
            }
            catch (NotSupportedException NS)
            {
                Console.WriteLine(NS.StackTrace.ToString() + "\n" + NS.Message.ToString() + "\n");
                if (_file != null)
                {
                    File.Move(_processingPath + "\\" + _fileName, _errorPath + "\\" + _fileName);
                    File.AppendAllText(_errorPath + "\\" + "log.txt", DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + NS.Message.ToString() + "\n" + NS.StackTrace.ToString() + "\n");
                    string body = DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + NS.Message.ToString() + "\n" + "\n" + NS.StackTrace.ToString() + "\n";
                    MailMessage message = new MailMessage(_emailSender, _emailRecipient, _subject, body);
                    SmtpClient client = new SmtpClient(_smtpServer);
                    client.Send(message);
                }
                LoadFile();
            }
            catch (DirectoryNotFoundException DNF)
            {
                Console.WriteLine(DNF.StackTrace.ToString() + "\n" + DNF.Message.ToString() + "\n");
                if (_file != null)
                {
                    File.AppendAllText(_errorPath + "\\" + "log.txt", DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + DNF.Message.ToString() + "\n" + DNF.StackTrace.ToString() + "\n");
                    string body = DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + DNF.Message.ToString() + "\n" + "\n" + DNF.StackTrace.ToString() + "\n";
                    MailMessage message = new MailMessage(_emailSender, _emailRecipient, _subject, body);
                    SmtpClient client = new SmtpClient(_smtpServer);
                    client.Send(message);
                    return;
                }
                LoadFile();
            }
            catch (IndexOutOfRangeException IOOR)
            {
                Console.WriteLine(IOOR.StackTrace.ToString() + "\n" + IOOR.Message.ToString() + "\n");
                if (_file != null)
                {
                    File.Move(_processingPath + "\\" + _fileName, _errorPath + "\\" + _fileName);
                    File.AppendAllText(_errorPath + "\\" + "log.txt", DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + IOOR.Message.ToString() + "\n" + IOOR.StackTrace.ToString() + "\n");
                    string body = DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + IOOR.Message.ToString() + "\n" + "\n" + IOOR.StackTrace.ToString() + "\n";
                    MailMessage message = new MailMessage(_emailSender, _emailRecipient, _subject, body);
                    SmtpClient client = new SmtpClient(_smtpServer);
                    client.Send(message);
                }
                LoadFile();
            }
            catch (Exception E)
            {
                Console.WriteLine(E.StackTrace.ToString() + "\n" + E.Message.ToString() + "\n");
                if (_file != null)
                {
                    File.Move(_processingPath + "\\" + _fileName, _errorPath + "\\" + _fileName);
                    File.AppendAllText(_errorPath + "\\" + "log.txt", DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + E.Message.ToString() + "\n" + E.StackTrace.ToString() + "\n");
                    string body = DateTime.Now + " " + "[Filename: " + _fileName + ']' + " " + E.Message.ToString() + "\n" + "\n" + E.StackTrace.ToString() + "\n";
                    MailMessage message = new MailMessage(_emailSender, _emailRecipient, _subject, body);
                    SmtpClient client = new SmtpClient(_smtpServer);
                    client.Send(message);
                }
                LoadFile();
            }
        }

        //Receives a string of the XML file, parses said string and returns the selected info to pass to ConstructSQL
        public static void ParseXML(string fileContents, out string nameFirst, out string nameMiddle, out string nameFamily, out string nameSuffix, out DateTime? dateOfBirth, out string codeGender, out string codeEyeColor, out string codeHairColor, out string currentWeight, out string currentHeight)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(fileContents);

            //Variables for output
            nameFirst = null;
            nameFamily = null;
            nameMiddle = null;
            nameSuffix = null;
            dateOfBirth = null;

            //Get First/Middle/Last/Suffix Name
            XmlNodeList NAME = xmlDoc.SelectNodes("/descriptors/nam-j/Group");
            foreach (XmlNode nameTags in NAME)
            {
                nameFamily = nameTags["V1"].InnerText;
                nameFirst = nameTags["V2"].InnerText;
                nameMiddle = nameTags["V3"].InnerText;
                nameSuffix = nameTags["V4"].InnerText;
            }

            //Get Date of Birth and format it for DateTime format
            XmlNodeList DOB = xmlDoc.SelectNodes("/descriptors/_2.022/Group");
            foreach (XmlNode dobTags in DOB)
            {
                string unformattedDate = dobTags["V1"].InnerText;
                if (!string.IsNullOrEmpty(unformattedDate))
                {
                    dateOfBirth = DateTime.ParseExact(unformattedDate, "yyyyMMdd", null);
                }
            }

            //Get Gender
            codeGender = xmlDoc.GetElementsByTagName("_2.024")[0].InnerText;
            //Get Eye Color
            codeEyeColor = xmlDoc.GetElementsByTagName("_2.031")[0].InnerText;
            //Get Hair Color
            codeHairColor = xmlDoc.GetElementsByTagName("_2.032")[0].InnerText;
            //Get Weight
            currentWeight = xmlDoc.GetElementsByTagName("_2.029")[0].InnerText;
            //Get Height
            currentHeight = xmlDoc.GetElementsByTagName("_2.027")[0].InnerText;
        }

        //Receives the parsed information and uses it to construct a SQL Query to enter the information into the specifed database
        public static void ConstructSQL(string nameFirst, string nameMiddle, string nameFamily, string nameSuffix, DateTime? dateOfBirth, string codeGender, string codeEyeColor, string codeHairColor, string currentWeight, string currentHeight)
        {
            //Constructing the SQL Query            
            string sql = "INSERT INTO dbo.CNLV_PERSON_INFO (NAME_FIRST, NAME_MIDDLE, NAME_FAMILY, NAME_SUFFIX, DATE_OF_BIRTH, CODE_GENDER, CODE_EYE_COLOR, CODE_HAIR_COLOR, CURRENT_WEIGHT, CURRENT_HEIGHT) VALUES (@nameFirst, @nameMiddle, @nameFamily, @nameSuffix, @dateOfBirth, @codeGender, @codeEyeColor, @codeHairColor, @currentWeight, @currentHeight)";
            SqlConnection connection = new SqlConnection(_connectionString);
            using (SqlCommand command = new SqlCommand(sql, connection))
            {
                command.Parameters.AddWithValue("@nameFirst", (string.IsNullOrEmpty(nameFirst) ? (object)DBNull.Value : nameFirst));
                command.Parameters.AddWithValue("@nameMiddle", (string.IsNullOrEmpty(nameMiddle) ? (object)DBNull.Value : nameMiddle));
                command.Parameters.AddWithValue("@nameFamily", (string.IsNullOrEmpty(nameFamily) ? (object)DBNull.Value : nameFamily));
                command.Parameters.AddWithValue("@nameSuffix", (string.IsNullOrEmpty(nameSuffix) ? (object)DBNull.Value : nameSuffix));
                command.Parameters.AddWithValue("@dateOfBirth", (!dateOfBirth.HasValue ? (object)DBNull.Value : dateOfBirth));
                command.Parameters.AddWithValue("@codeGender", (string.IsNullOrEmpty(codeGender) ? (object)DBNull.Value : codeGender));
                command.Parameters.AddWithValue("@codeEyeColor", (string.IsNullOrEmpty(codeEyeColor) ? (object)DBNull.Value : codeEyeColor));
                command.Parameters.AddWithValue("@codeHairColor", (string.IsNullOrEmpty(codeHairColor) ? (object)DBNull.Value : codeHairColor));
                command.Parameters.AddWithValue("@currentWeight", (string.IsNullOrEmpty(currentWeight) ? (object)DBNull.Value : currentWeight));
                command.Parameters.AddWithValue("@currentHeight", (string.IsNullOrEmpty(currentHeight) ? (object)DBNull.Value : currentHeight));
                connection.Open();
                command.ExecuteNonQuery();
                connection.Close();
                string start_location = _processingPath + "\\" + _fileName;
                string complete_location = _completePath + "\\" + _fileName;
                File.Move(start_location, complete_location);
            }
        }
           
        static void Main(string[] args)
        {
            LoadFile(); //LoadFile calls both ParseXML and ConstructSQL during execution
        }
    }
}
