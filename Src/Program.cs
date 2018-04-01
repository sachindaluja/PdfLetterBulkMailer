using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;

namespace PdfLetterBulkMailer
{
    class Program
    {
        enum CmdArg
        {
            None,
            Unknown,
            EmailSenderId,
            EmailSenderPass,
            EmailSenderSmtp,
            EmailNoSSL,
        }

        enum ErrorCode
        {
            Success = 0,
            InvalidRecord,
            LatexError,
            EmailError,
            FileWriteError,
            UnspecifiedError,
        }

        static string emailSenderId = "";
        static SmtpClient smtpClient = null;
        static string texBase = "";
        static string emailMsgBase = "";

        static int Main(string[] args)
        {
            string emailSenderPass = "";
            string emailSenderSmtp = "";
            bool noSSL = false;

            bool bInvalidArgs = false;

            CmdArg lastCmdSwitch = CmdArg.None;

            foreach (string arg in args)
            {
                switch (lastCmdSwitch)
                {
                    case CmdArg.None:
                    case CmdArg.EmailNoSSL:
                        if (arg.Length > 1 && (arg[0] == '/' || arg[0] == '-'))
                        {
                            string option = arg.Substring(1);

                            if (option == "username")
                            {
                                lastCmdSwitch = CmdArg.EmailSenderId;
                            }
                            else if (option == "pass")
                            {
                                lastCmdSwitch = CmdArg.EmailSenderPass;
                            }
                            else if (option == "smtp")
                            {
                                lastCmdSwitch = CmdArg.EmailSenderSmtp;
                            }
                            else if (option == "nossl")
                            {
                                lastCmdSwitch = CmdArg.EmailNoSSL;

                                noSSL = true;
                            }
                            else
                            {
                                bInvalidArgs = true;
                            }
                        }
                        else
                        {
                            bInvalidArgs = true;
                        }

                        break;
                    case CmdArg.EmailSenderId:
                        emailSenderId = arg;
                        lastCmdSwitch = CmdArg.None;
                        break;
                    case CmdArg.EmailSenderPass:
                        emailSenderPass = arg;
                        lastCmdSwitch = CmdArg.None;
                        break;
                    case CmdArg.EmailSenderSmtp:
                        emailSenderSmtp = arg;
                        lastCmdSwitch = CmdArg.None;
                        break;
                    default:
                        bInvalidArgs = true;
                        break;
                }

                if (bInvalidArgs)
                    break;
            }

            if (!bInvalidArgs)
                bInvalidArgs = (emailSenderId == "") || (emailSenderSmtp == "");

            if (!bInvalidArgs && (emailSenderPass == ""))
            {
                Console.WriteLine("Type sender's email password for " + emailSenderId + ":");

                ConsoleKeyInfo key;

                do
                {
                    key = Console.ReadKey(true);

                    if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                    {
                        emailSenderPass += key.KeyChar;
                        Console.Write("*");
                    }
                    else
                    {
                        if (key.Key == ConsoleKey.Backspace && emailSenderPass.Length > 0)
                        {
                            emailSenderPass = emailSenderPass.Substring(0, (emailSenderPass.Length - 1));
                            Console.Write("\b \b");
                        }
                    }
                }
                while (key.Key != ConsoleKey.Enter);

                if(emailSenderPass == "")
                    bInvalidArgs = true;

                Console.WriteLine("\n");
            }

            if (bInvalidArgs)
            {
                Console.WriteLine("Invalid command line arguments. Usage: PdfLetterBulkMailer.exe -username <Sender's email Id> -smtp <Sender's smpt server[:port]>. Optional arguments: -pass <Sender's email password> -nossl");
                return 1;
            }

            String record = null;

            try
            {
                using (StreamReader sr = new StreamReader("Letter.tex"))
                    texBase = sr.ReadToEnd();
            }
            catch (Exception)
            {
                Console.WriteLine("The source tex file could not be read.");
                return 2;
            }

            try
            {
                using (StreamReader sr = new StreamReader("EmailMessage.txt"))
                    emailMsgBase = sr.ReadToEnd();
            }
            catch (Exception)
            {
                Console.WriteLine("The source email message text file could not be read.");
                return 2;
            }

            smtpClient = new SmtpClient();

            int colonIndex = emailSenderSmtp.IndexOf(':');
            if (colonIndex != -1)
            {
                smtpClient.Host = emailSenderSmtp.Substring(0, colonIndex);

                if (emailSenderSmtp.Length > (colonIndex + 1))
                {
                    string smtpPort = emailSenderSmtp.Substring(colonIndex + 1);

                    try
                    {
                        smtpClient.Port = Convert.ToInt32(smtpPort);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("The smtp server argument is invalid.");
                        return 1;
                    }
                }
            }
            else
            {
                smtpClient.Host = emailSenderSmtp;
            }

            NetworkCredential credentials = new NetworkCredential(emailSenderId, emailSenderPass);

            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = credentials;
            smtpClient.EnableSsl = !noSSL;
            //smtpClient.Timeout = 10000; //10,000 milli-seconds

            StreamReader donorDb = null;

            try
            {
                donorDb = new StreamReader("MailingListDb.csv");
            }
            catch (Exception)
            {
                Console.WriteLine("The mailing list database file could not be read.");
                return 3;
            }

            try
            {
                StreamWriter successLog = new StreamWriter("Output\\Processed.log", true);
                StreamWriter failLog = new StreamWriter("Output\\Failed_" + DateTime.Now.ToString("HHmmss") + ".log", false);
                
                record = donorDb.ReadLine();

                int recordNum = 0;
                while (record != null)
                {
                    recordNum++;

                    string recordCorrected = null;

                    string [] quotedTokens = record.Split(new Char[] {'"'});
                    if(quotedTokens.Length > 2)
                    {
                        StringBuilder sb = new StringBuilder(""); ;

                        sb.Append(quotedTokens[0]);

                        for (int i = 1; i < (quotedTokens.Length - 1); ++i)
                        {
                            string token = quotedTokens[i];

                            if ((i % 2) != 0)
                                token = token.Replace(",", "^");

                            sb.Append(token);
                        }
                        sb.Append(quotedTokens[quotedTokens.Length - 1]);

                        recordCorrected = sb.ToString();
                    }
                    else
                    {
                        recordCorrected = record;
                    }

                    string[] recordFields = recordCorrected.Split(new Char[] { ',' });

                    for (int i = 0; i < recordFields.Length; ++i)
                    {
                        if(i != 2)
                            recordFields[i] = recordFields[i].Replace("^", "");
                        else
                            recordFields[i] = recordFields[i].Replace("^", ",");
                    }

                    Console.WriteLine("Processing record: " + record);

                    ErrorCode err;

                    try
                    {
                        err = ProcessOneRecord(recordFields);

                        switch (err)
                        {
                            case ErrorCode.Success:
                                {
                                    break;
                                }
                            case ErrorCode.InvalidRecord:
                                {
                                    Console.WriteLine("Skipped record as the required fields were not found.");
                                    break;
                                }
                            case ErrorCode.LatexError:
                                {
                                    Console.WriteLine("The latex processor failed to create a pdf. Do you want to skip this record and continue?");

                                    string consoleInput = Console.ReadLine();
                                    consoleInput.ToUpper();

                                    if (!((consoleInput == "Y") || (consoleInput == "Yes") || (consoleInput == "y") || (consoleInput == "yes")))
                                        return 5;

                                    break;
                                }
                            case ErrorCode.EmailError:
                                {
                                    Console.WriteLine("Could not send email.");
                                    break;
                                }
                            default:
                                {
                                    Console.WriteLine("An unspecified error occurred while processing the record.");
                                    return 8;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("An error occurred while processing. Details: " + ex.Message);
                        return 9;
                    }

                    if (err == ErrorCode.Success)
                    {
                        successLog.WriteLine(record);
                        successLog.Flush();
                    }
                    else
                    {
                        failLog.WriteLine(record);
                        failLog.Flush();
                    }

                    record = donorDb.ReadLine();
                }

                successLog.Close();
                failLog.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred. Details: " + ex.Message);
                return 3;
            }

            return 0;
        }

        static ErrorCode ProcessOneRecord(string[] recordFields)
        {
            if (recordFields.Length != 7)
                return ErrorCode.InvalidRecord;

            string emailAddress = recordFields[6];

            string exampleNumber = recordFields[2];

            try
            {
                Convert.ToDouble(exampleNumber);
            }
            catch (Exception)
            {
                return ErrorCode.InvalidRecord;
            }

            string nameOrig = recordFields[1];

            for (int i = 0; i < recordFields.Length; ++i)
            {
                recordFields[i] = recordFields[i].Replace("&", @"\&");
                recordFields[i] = recordFields[i].Replace("_", @"\_");
                recordFields[i] = recordFields[i].Replace("$", @"\$");
                recordFields[i] = recordFields[i].Replace("#", @"\#");
                recordFields[i] = recordFields[i].Replace("%", @"\%");
                
                if(recordFields[i] == "")
                    recordFields[i] = "~";
            }

            string name = recordFields[1];
            string addrLine1 = recordFields[3];
            string addrTown = recordFields[4];
            string addrStateZip = recordFields[5];

            string tex = texBase;

            tex = tex.Replace("$$NAME$$", name);
            tex = tex.Replace("$$EXAMPLE_NUMBER$$", exampleNumber);
            tex = tex.Replace("$$ADDRESS_LINE1$$", addrLine1);
            tex = tex.Replace("$$TOWN$$", addrTown);
            tex = tex.Replace("$$STATE_AND_ZIP$$", addrStateZip);

            string texFileName = @"Temp\Letter_processed.tex";

            StreamWriter texfile = new StreamWriter(texFileName);
            texfile.Write(tex);
            texfile.Close();

            string pdfLatexCmdArgs = "-quiet -halt-on-error -output-directory=\"Temp\" \"" + texFileName + "\"";    /*-job-name=\"" + tempPdfFileName + "\"*/

            ProcessStartInfo psi = new ProcessStartInfo("pdflatex.exe", pdfLatexCmdArgs);
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;

            Process pdflatexProc = new Process();
            pdflatexProc.StartInfo = psi;

            bool bStartSuccess = pdflatexProc.Start();

            pdflatexProc.WaitForExit();

            if (pdflatexProc.ExitCode != 0)
                return ErrorCode.LatexError;

            string pdfPath = "Output\\" + nameOrig + ".pdf";

            System.IO.File.Copy("Temp\\Letter_processed.pdf", pdfPath, true);

            if (emailAddress.Contains('@'))
                SendEmail(emailAddress, nameOrig, pdfPath);
            else
                return ErrorCode.EmailError;

            return ErrorCode.Success;
        }

        static void SendEmail(string recipientEmail, string recipientName, string pdfPath)
        {
            string pdfFileName = pdfPath;
            int nPathSepIndex = pdfPath.LastIndexOfAny(new char[] {'\\', '/'});
            if((nPathSepIndex != -1) && (pdfPath.Length > (nPathSepIndex + 1)))
                pdfFileName = pdfPath.Substring(nPathSepIndex + 1);

            MailMessage message = new MailMessage(emailSenderId, recipientEmail);

            message.Subject = "Example message subject";

            message.Body = emailMsgBase;
            message.Body = message.Body.Replace("$$NAME$$", recipientName);

            FileStream fs = new FileStream(pdfPath, FileMode.Open, FileAccess.Read);

            ContentType ct = new ContentType();
            ct.MediaType = MediaTypeNames.Application.Octet;
            ct.Name = pdfFileName;
            Attachment data = new Attachment(fs, ct);
            
            message.Attachments.Add(data);

            smtpClient.Send(message);

            data.Dispose();
            fs.Close();
        }
    }
}
