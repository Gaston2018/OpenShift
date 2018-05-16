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
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace app.Controllers
{
    public class PresupuestosController : Controller
    {
        private readonly MassiveMailAppSetting _massiveMailAppSetting;

        public PresupuestosController(IOptions<MassiveMailAppSetting> massiveMailAppSetting)
        {
            _massiveMailAppSetting = massiveMailAppSetting.Value;
        }

        public IActionResult Index()
        {
            return View();
        }

        #region MiniSosa

        [HttpPost("Presupuestos")]
        public IActionResult EnvioMailMasivoPresupuesto(MailData mailData, List<IFormFile> files)
        {
            ThreadPool.QueueUserWorkItem(o => EnvioMailMasivoAsync(mailData, files));
            return View("Finalizado");
        }

        public bool EnvioMailMasivoAsync(MailData mailData, List<IFormFile> files)
        {

            try
            {
                IFormFile archivoWord = files[0];
                IFormFile archivoExcel = files[1];

                string mailsErroneos = string.Empty;

                if (archivoExcel.Length > 0 && archivoWord.Length > 0)
                {
                    //ARCHIVOS-------------------------------------//
                    FileInfo file = new FileInfo(_massiveMailAppSetting.FilesPath);
                    file.Directory.Create();


                    // full path to file in location
                    string rutaExcel = Path.Combine(_massiveMailAppSetting.FilesPath, archivoExcel.FileName);
                    string rutaWordOriginal = Path.Combine(_massiveMailAppSetting.FilesPath, archivoWord.FileName);

                    using (var stream = new FileStream(rutaExcel, FileMode.Create))
                    {
                        archivoExcel.CopyTo(stream);
                    }

                    using (var stream = new FileStream(rutaWordOriginal, FileMode.Create))
                    {
                        archivoWord.CopyTo(stream);
                    }

                    //EXCEL*************************************//

                    DocumentosOpenXml documentosOpenXml = new DocumentosOpenXml();

                    DataTable dtExcel = new DataTable();

                    try
                    {
                        if (Path.GetExtension(rutaExcel).ToLower() == ".csv")
                            dtExcel = documentosOpenXml.ConvertCSVtoDataTable(rutaExcel);
                        else
                            dtExcel = documentosOpenXml.ConvertExceltoDataTable(rutaExcel);
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

                    //WORD****************************************************//

                    // full path to file in temp location
                    string rutaWord = Path.Combine(_massiveMailAppSetting.FilesPath, "TEMP_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + archivoWord.FileName);

                    string keyBuscar;
                    int filaIndex = 1;

                    foreach (DataRow filas in dtExcel.Rows)
                    {
                        filaIndex++;
                        string copiaAsunto = mailData.Subject;
                        string copiaMensaje = mailData.Message;

                        System.IO.File.Copy(rutaWordOriginal, rutaWord, true);

                        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(rutaWord, true))
                        {
                            string docText = null;
                            using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                            {
                                docText = sr.ReadToEnd();
                                sr.Close();
                                sr.Dispose();
                            }

                            bool campoVacio = false;
                            int pos = -1;

                            foreach (var item in filas.ItemArray)
                            {
                                pos++;

                                keyBuscar = "key" + RemoverTildes(dtExcel.Columns[pos].ToString().Trim().ToLower()).Replace(" ", string.Empty);

                                //SI EL item.ToString() está vacío, se guarda como error y se sigue al próximo registro
                                if (item.ToString() == "")
                                {
                                    campoVacio = true;
                                    break;
                                }

                                docText = Regex.Replace(docText, keyBuscar, item.ToString(), RegexOptions.IgnoreCase);
                                copiaAsunto = Regex.Replace(copiaAsunto, keyBuscar, item.ToString(), RegexOptions.IgnoreCase);
                                copiaMensaje = Regex.Replace(copiaMensaje, keyBuscar, item.ToString(), RegexOptions.IgnoreCase);
                            }

                            if (campoVacio)
                            {
                                mailsErroneos += "Hay campos no cargados en la fila " + filaIndex + "." + Environment.NewLine;
                                continue;
                            }

                            using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                            {
                                sw.Write(docText);
                                sw.Close();
                            }
                            wordDoc.Close();
                        }
                        //ENVIAR MAIL
                        if (filas.ItemArray[posicionEmail].ToString() != "" && filas.ItemArray[posicionEmail].ToString() != "N/A")
                        {
                            try
                            {
                                EnviarMailAdjunto(mailData.Sender, filas.ItemArray[posicionEmail].ToString(), copiaAsunto, copiaMensaje, rutaWord, archivoWord.FileName);
                            }
                            catch (Exception ex)
                            {
                                mailsErroneos += filas.ItemArray[posicionEmail].ToString() + "(" + ex.Message + ")." + Environment.NewLine;
                            }
                        }
                        else
                        {
                            //NO HAY MAIL CARGADO
                            mailsErroneos += "No hay mail cargado en fila " + filaIndex + "." + Environment.NewLine;
                        }
                    }

                }

                string mensaje = string.Empty;

                if (mailsErroneos == string.Empty)
                {
                    mensaje = "Se enviaron todos los mails correctamente." + Environment.NewLine + Environment.NewLine + "Gracias por utilizar nuestro servicio.";
                    EnviarMailResultado("MailMasivoDinamico@gcgestion.com.ar", mailData.Sender, "Resultado del envio de mails", mensaje);
                }
                else
                {
                    string nombreLog = GenerarLog(_massiveMailAppSetting.LogFilesPath + mailData.Sender + "\\", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + mailData.Subject + ".txt", mailsErroneos);

                    EnviarMailAdjunto("MailMasivoDinamico@gcgestion.com.ar", mailData.Sender, "Resultado del envio de mails", "Se adjunta el resultado de los mails erroneos.", nombreLog, mailData.Sender + ".txt");
                }

                // process uploaded files
                // Don't rely on or trust the FileName property without validation.
                return true;
            }
            catch (Exception ex)
            {
                GenerarLog(_massiveMailAppSetting.LogFilesPath, _massiveMailAppSetting.LogFileName, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + mailData.Sender + " Error no contemplado: " + ex.Message + ". ");

                Response.StatusCode = (int)HttpStatusCode.BadRequest;
                return false;
            }
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

        #endregion
    }
}