using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using app.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Edm.Csdl;
using Microsoft.Extensions.Options;

namespace app.Controllers
{
    public class RenovacionesController : Controller
    {
        private readonly MassiveMailAppSetting _massiveMailAppSetting;

        public RenovacionesController(IOptions<MassiveMailAppSetting> massiveMailAppSetting)
        {
            _massiveMailAppSetting = massiveMailAppSetting.Value;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost("Renovaciones")]
        public IActionResult EnvioMailMasivoRenovacion(MailData mailData, List<IFormFile> files)
        {
            ThreadPool.QueueUserWorkItem(o => EnvioMailMasivoRenovacionAsync(mailData, files));
            return View("Finalizado");
        }

        public bool EnvioMailMasivoRenovacionAsync(MailData mailData, List<IFormFile> files)
        {
            try
            {
                IFormFile excelFile = null;
                string mailsError = string.Empty;

                if (!(files != null && files[0].Length > 0))
                {
                    return false;
                }

                excelFile = files[0];

                //ARCHIVOS-------------------------------------//
                FileInfo file = new FileInfo(_massiveMailAppSetting.FilesPath);
                file.Directory.Create();

                string excelPath = Path.Combine(_massiveMailAppSetting.FilesPath, excelFile.FileName);

                using (var stream = new FileStream(excelPath, FileMode.Create))
                {
                    excelFile.CopyTo(stream);
                }

                //EXCEL*************************************//

                DocumentosOpenXml openXmlDocuments = new DocumentosOpenXml();

                DataTable dtExcel = new DataTable();

                try
                {
                    if (Path.GetExtension(excelPath).ToLower() == ".csv")
                        dtExcel = openXmlDocuments.ConvertCSVtoDataTable(excelPath);
                    else
                        dtExcel = openXmlDocuments.ConvertExceltoDataTable(excelPath);
                }
                catch (Exception ex)
                {
                    string mensajeError = "No se pudo realizar el envío de mails porque se encontró una columna repetida en el Excel (" + ex.Message + ").";
                    EnviarMailResultado("MailMasivoDinamico@gcgestion.com.ar", mailData.Sender, "Resultado del envio de mails", mensajeError);
                    return false;
                }

                int posicionEmail;

                if (dtExcel.Columns["EMAIL"] != null)
                    posicionEmail = dtExcel.Columns["EMAIL"].Ordinal;
                else
                {
                    string mensajeError = "No se pudo realizar el envío de mails porque no se encontró la columna \"Email\" en el Excel.";
                    EnviarMailResultado("MailMasivoDinamico@gcgestion.com.ar", mailData.Sender, "Resultado del envio de mails", mensajeError);
                    return false;
                }

                string keyBuscar;
                int filaIndex = 1;
                foreach (DataRow excelItemsRow in dtExcel.Rows)
                {
                    filaIndex++;
                    string subjectCopy = mailData.Subject;
                    string messageCopy = mailData.Message;
                    string poliza = String.Empty;
                    string certificado = String.Empty;

                    bool campoVacio = false;
                    int pos = -1;
                    foreach (var item in excelItemsRow.ItemArray)
                    {
                        pos++;
                        string excelItemColumn = String.Empty;
                        

                        excelItemColumn = RemoverTildes(dtExcel.Columns[pos].ToString().Trim().ToLower()).Replace(" ", string.Empty);

                        keyBuscar = "key" + excelItemColumn;

                        //Guardamos el nombre del archivo
                        if (excelItemColumn == "poliza")
                        {
                            poliza = item.ToString();
                        }
                        else
                        {
                            if (excelItemColumn == "certificado")
                            {
                                certificado = item.ToString();
                            }
                        }

                        //SI EL item.ToString() está vacío, se guarda como error y se sigue al próximo registro
                        if (item.ToString() == "")
                        {
                            campoVacio = true;
                            break;
                        }
                        subjectCopy = Regex.Replace(subjectCopy, keyBuscar, item.ToString(), RegexOptions.IgnoreCase);
                        messageCopy = Regex.Replace(messageCopy, keyBuscar, item.ToString(), RegexOptions.IgnoreCase);
                    }
                    if (campoVacio)
                    {
                        mailsError += "Hay campos no cargados en la fila " + filaIndex + "." + Environment.NewLine;
                        continue;
                    }
                    string pdfPath = _massiveMailAppSetting.PdfsFilesPath + poliza + "_" + certificado + ".pdf";

                    //ENVIAR MAIL
                    if (excelItemsRow.ItemArray[posicionEmail].ToString() != "" && excelItemsRow.ItemArray[posicionEmail].ToString() != "N/A")
                    {
                        try
                        {
                            EnviarMailAdjunto(mailData.Sender, excelItemsRow.ItemArray[posicionEmail].ToString(), subjectCopy, messageCopy, pdfPath, poliza + "_" + certificado + ".pdf");
                        }
                        catch (Exception ex)
                        {
                            mailsError += excelItemsRow.ItemArray[posicionEmail].ToString() + "(" + ex.Message + ")." + Environment.NewLine;
                        }
                    }
                    else
                    {
                        //NO HAY MAIL CARGADO
                        mailsError += "No hay mail cargado en fila " + filaIndex + "." + Environment.NewLine;
                    }
                }
                string mensaje = string.Empty;

                if (mailsError == string.Empty)
                {
                    mensaje = "Se enviaron todos los mails correctamente." + Environment.NewLine + Environment.NewLine + "Gracias por utilizar nuestro servicio.";
                    EnviarMailResultado("MailMasivoDinamico@gcgestion.com.ar", mailData.Sender, "Resultado del envio de mails", mensaje);
                }
                else
                {
                    string nombreLog = GenerarLog(_massiveMailAppSetting.LogFilesPath + mailData.Sender + "\\", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + mailData.Subject + ".txt", mailsError);

                    EnviarMailAdjunto("MailMasivoDinamico@gcgestion.com.ar", mailData.Sender, "Resultado del envio de mails", "Se adjunta el resultado de los mails erroneos.", nombreLog, mailData.Sender + ".txt");
                }
            }
            catch (Exception ex)
            {

                GenerarLog(_massiveMailAppSetting.LogFilesPath, _massiveMailAppSetting.LogFileName, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + mailData.Sender + " Error no contemplado: " + ex.Message + ". ");

                Response.StatusCode = (int)HttpStatusCode.BadRequest;
                return false;
            }


            return true;
        }

        public bool EnviarMailResultado(string remitente, string destinatario, string asunto, string mensaje)
        {
            var smtp = new SmtpClient
            {
                Host = _massiveMailAppSetting.Host,
                Port = _massiveMailAppSetting.Port,
                EnableSsl = false,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(_massiveMailAppSetting.User, _massiveMailAppSetting.Password)
            };

            using (var message = new MailMessage(remitente, destinatario))
            {
                message.Subject = asunto;
                message.Body = mensaje;
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                smtp.Send(message);
            }

            return true;
        }

        static string RemoverTildes(string text)
        {
            return string.Concat(Regex.Replace(text, @"(?i)[\p{L}-[ña-z]]+", m => m.Value.Normalize(NormalizationForm.FormD)).Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark));
        }

        public bool EnviarMailAdjunto(string remitente, string destinatario, string asunto, string mensaje, string urlAdjunto, string nombreAdjunto)
        {
            var smtp = new SmtpClient
            {
                Host = _massiveMailAppSetting.Host,
                Port = _massiveMailAppSetting.Port,
                EnableSsl = false,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(_massiveMailAppSetting.User, _massiveMailAppSetting.Password)
            };

            using (var message = new MailMessage(remitente, destinatario))
            {
                Attachment attachment;
                attachment = new Attachment(urlAdjunto);
                attachment.Name = nombreAdjunto;
                message.Subject = asunto;
                message.Body = mensaje;
                message.Attachments.Add(attachment);
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };

                try
                {
                    smtp.Send(message);
                }
                catch (SmtpFailedRecipientException exc)
                {
                    throw exc;
                }
                catch (Exception EX)
                {
                    throw EX;
                }

            }
            return true;
        }

        public string GenerarLog(string rutaLog, string nombreLog, string mensaje)
        {
            rutaLog = RemoverCaracteresInvalidosParaRuta(rutaLog);
            nombreLog = RemoverCaracteresInvalidosParaArchivo(nombreLog);

            nombreLog = rutaLog + nombreLog;

            FileInfo file = new FileInfo(rutaLog);
            file.Directory.Create();

            TextWriter tw = new StreamWriter(nombreLog, true);
            tw.WriteLine(mensaje);
            tw.Close();

            return nombreLog;
        }

        static string RemoverCaracteresInvalidosParaRuta(string pathname)
        {
            foreach (char c in Path.GetInvalidPathChars())
            {
                pathname = pathname.Replace(c, '_');
            }
            return pathname;
        }

        static string RemoverCaracteresInvalidosParaArchivo(string filename)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                filename = filename.Replace(c, '_');
            }
            return filename;
        }
    }
}