using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Xml4U;
using System.Drawing.Printing;
using System.Text.RegularExpressions;
using System.Linq;

namespace Helper4U
{
    /// <summary>
    ///SqlHelper4U helps to minimize the your code
    /// </summary> 
    public class SqlHelper4U
    {
        /// <summary>
        ///getConectionString() funcation helps to connect cons name connection from web.config
        /// </summary> 
        public string getConectionString()
        {
            return ConfigurationManager.ConnectionStrings["cons"].ConnectionString;
        }

        #region Convert Amount in Words
        /// <summary>
        /// Convert Amount in Words [e.g.  AmountInWords(100)]
        /// </summary> 
        /// <param name="Amount">enter amount.</param>
        /// <returns>return amount in word </returns>
        public string AmountInWords(long Amount)
        {
            string[] m10 = new string[20];
            string[] m100 = new string[10];
            string[] mx = new string[11];
            string mstr = "";
            long ml1, ml2, ml;
            int pos, cond;
            int tl1, tl2, tl3;
            if (Amount == 0) mstr = "Zero";
            else
            {
                m10[0] = ""; m10[1] = " One"; m10[2] = " Two"; m10[3] = " Three"; m10[4] = " Four"; m10[5] = " Five"; m10[6] = " Six"; m10[7] = " Seven"; m10[8] = " Eight"; m10[9] = " Nine"; m10[10] = " Ten"; m10[11] = " Eleven"; m10[12] = " Twelve"; m10[13] = " Thirteen"; m10[14] = " Fourteen"; m10[15] = " Fifteen"; m10[16] = " Sixteen"; m10[17] = " Seventeen"; m10[18] = " Eighteen"; m10[19] = " Nineteen";
                m100[0] = ""; m100[1] = ""; m100[2] = " Twenty"; m100[3] = " Thirty"; m100[4] = " Fourty"; m100[5] = " Fifty"; m100[6] = " Sixty"; m100[7] = " Seventy"; m100[8] = " Eighty"; m100[9] = " Ninety";
                mx[0] = ""; mx[1] = ""; mx[2] = ""; mx[3] = " Hundred"; mx[4] = " Thousand"; mx[5] = " Thousand"; mx[6] = " Lac"; mx[7] = " Lac"; mx[8] = " Crore"; mx[9] = " Crore"; mx[10] = " Arab";
                ml1 = Math.Abs(Amount);
                ml = 1000000000;
                pos = 10;
                cond = 0;
                while (cond != 1)
                {
                    if (ml1 >= ml) break;
                    ml = ml / 10;
                    pos = pos - 1;
                    while (ml1 > 0)
                    {
                        if (pos == 9 || pos == 7 || pos == 5 || pos == 2)
                        {
                            ml = ml / 10;
                            pos = pos - 1;
                        }
                        ml2 = (int)(ml1 / ml);
                        ml1 = ml1 - (ml2 * ml);
                        ml = ml / 10;
                        if (ml2 >= 20)
                        {
                            tl2 = (int)(ml2 / 10);
                            tl3 = tl2 * 10;
                            tl1 = (int)ml2 - tl3;
                            mstr = mstr + m100[tl2];
                            mstr = mstr + m10[tl1];
                        }
                        else mstr = mstr + m10[ml2];
                        if (ml2 > 0) mstr = mstr + mx[pos];
                        pos = pos - 1;
                    }
                }
            }
            return mstr + " Rupes Only";

        }
        /// <summary>
        /// Convert Amount in Words [e.g.  AmountInWords(100.50)]
        /// </summary> 
        /// <param name="numbers">enter amount in decimal.</param>
        /// <param name="paisaconversion">true for append currency such as Rupees,Paise</param>
        /// <returns>return amount in word with paisa </returns>
        public string AmountInWords(double? numbers, Boolean paisaconversion = false)
        {
            var pointindex = numbers.ToString().IndexOf(".");
            var paisaamt = 0;
            if (pointindex > 0)
                paisaamt = Convert.ToInt32(numbers.ToString().Substring(pointindex + 1, 2));

            int number = Convert.ToInt32(numbers);

            if (number == 0) return "Zero";
            if (number == -2147483648) return "Minus Two Hundred and Fourteen Crore Seventy Four Lakh Eighty Three Thousand Six Hundred and Forty Eight";
            int[] num = new int[4];
            int first = 0;
            int u, h, t;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (number < 0)
            {
                sb.Append("Minus ");
                number = -number;
            }
            string[] words0 = { "", "One ", "Two ", "Three ", "Four ", "Five ", "Six ", "Seven ", "Eight ", "Nine " };
            string[] words1 = { "Ten ", "Eleven ", "Twelve ", "Thirteen ", "Fourteen ", "Fifteen ", "Sixteen ", "Seventeen ", "Eighteen ", "Nineteen " };
            string[] words2 = { "Twenty ", "Thirty ", "Forty ", "Fifty ", "Sixty ", "Seventy ", "Eighty ", "Ninety " };
            string[] words3 = { "Thousand ", "Lakh ", "Crore " };
            num[0] = number % 1000; // units
            num[1] = number / 1000;
            num[2] = number / 100000;
            num[1] = num[1] - 100 * num[2]; // thousands
            num[3] = number / 10000000; // crores
            num[2] = num[2] - 100 * num[3]; // lakhs
            for (int i = 3; i > 0; i--)
            {
                if (num[i] != 0)
                {
                    first = i;
                    break;
                }
            }
            for (int i = first; i >= 0; i--)
            {
                if (num[i] == 0) continue;
                u = num[i] % 10; // ones
                t = num[i] / 10;
                h = num[i] / 100; // hundreds
                t = t - 10 * h; // tens
                if (h > 0) sb.Append(words0[h] + "Hundred ");
                if (u > 0 || t > 0)
                {
                    if (h > 0 || i == 0) sb.Append("and ");
                    if (t == 0)
                        sb.Append(words0[u]);
                    else if (t == 1)
                        sb.Append(words1[u]);
                    else
                        sb.Append(words2[t - 2] + words0[u]);
                }
                if (i != 0) sb.Append(words3[i - 1]);
            }

            if (paisaamt == 0 && paisaconversion == false)
            {
                sb.Append("ruppes only");
            }
            else if (paisaamt > 0)
            {
                var paisatext = AmountInWords(paisaamt, true);
                sb.AppendFormat("rupees {0} paise only", paisatext);
            }
            return sb.ToString().TrimEnd();
        }

        #endregion

        #region Convert Month Number to Month Name Hindi-English
        /// <summary>
        ///Get Month Name in English [e.g. MonthNameEng(2,""); output-February]
        ///Get Month Name in Hindi [e.g. MonthNameEng(1,"HN"); output-जनवरी]
        /// </summary> 
        /// <param name="month_number">enter month number (1 to 12).</param>
        /// <param name="lang">enter "hn" for Hindi or Default in English</param>

        public string MonthName(int month_number, string lang)
        {
            int i = Convert.ToInt32(month_number);
            string str = "";
            if (lang.ToString().ToLower() == "hn")
            {
                if (i == 1)
                    str = "जनवरी";
                else if (i == 2)
                    str = "फ़रवरी";
                else if (i == 3)
                    str = "मार्च";
                else if (i == 4)
                    str = "अप्रैल";
                else if (i == 5)
                    str = "मई";
                else if (i == 6)
                    str = "जून";
                else if (i == 7)
                    str = "जुलाई";
                else if (i == 8)
                    str = "अगस्त";
                else if (i == 9)
                    str = "सितंबर";
                else if (i == 10)
                    str = "अक्टूबर";
                else if (i == 11)
                    str = "नवंबर";
                else if (i == 12)
                    str = "दिसंबर";
                else
                    str = "N/A";
            }
            else
            {
                if (i == 1)
                    str = "January";
                else if (i == 2)
                    str = "February";
                else if (i == 3)
                    str = "March";
                else if (i == 4)
                    str = "April";

                else if (i == 5)
                    str = "May";
                else if (i == 6)
                    str = "June";
                else if (i == 7)
                    str = "July";
                else if (i == 8)
                    str = "August";
                else if (i == 9)
                    str = "September";
                else if (i == 10)
                    str = "October";
                else if (i == 11)
                    str = "November";
                else if (i == 12)
                    str = "December";
                else
                    str = "N/A";

            }
            return str;
        }
        #endregion

        #region Day Name in English-Hindi
        /// <summary>
        ///Get Day Name in English [e.g. DayName(2,""); output-Sunday]
        ///Get Day Name in Hindi [e.g. DayName(1,"HN"); output-रविवार]
        /// </summary> 
        /// <param name="day_number">enter month number (1 to 7).</param>
        /// <param name="lang">enter "hn" for Hindi or Default in English</param>

        public string DayName(int day_number, string lang)
        {
            int i = Convert.ToInt32(day_number);
            string str = "";
            if (lang.ToString().ToLower() == "hn")
            {
                if (i == 1)
                    str = "रविवार";
                else if (i == 2)
                    str = "सोमवार";
                else if (i == 3)
                    str = "मंगलवार";
                else if (i == 4)
                    str = "बुधवार";

                else if (i == 5)
                    str = "गुरुवार";
                else if (i == 6)
                    str = "शुक्रवार";
                else if (i == 7)
                    str = "शनिवार";

                else
                    str = "N/A";
            }
            else
            {
                if (i == 1)
                    str = "Sunday";
                else if (i == 2)
                    str = "Monday";
                else if (i == 3)
                    str = "Tuesday";
                else if (i == 4)
                    str = "Wednesday";

                else if (i == 5)
                    str = "Thursday";
                else if (i == 6)
                    str = "Friday";
                else if (i == 7)
                    str = "Saturday";

                else
                    str = "N/A";
            }


            return str;
        }


        #endregion

        #region Get MAC and IPv4 Address

        /// <summary>
        /// Get Current System IPv4 Address [e.g.  GetIP4Address()]
        /// </summary> 
        /// <returns>return System IPv4 Address </returns>
        public string GetIP4Address()
        {
            string IP4Address = string.Empty;
            if (HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"] != null)
            {
                IP4Address = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString();
            }
            else if (HttpContext.Current.Request.UserHostAddress.Length != 0)
            {
                IP4Address = HttpContext.Current.Request.UserHostAddress;
            }

            return IP4Address;

        }
        /// <summary>
        /// Get Current System MAC Address [e.g.  GetMACAddress()]
        /// </summary> 
        /// <returns>return System MAC Address </returns>
        public string GetMACAddress()
        {
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            String sMacAddress = string.Empty;
            foreach (NetworkInterface adapter in nics)
            {
                if (sMacAddress == String.Empty)// only return MAC Address from first card  
                {
                    IPInterfaceProperties properties = adapter.GetIPProperties();
                    sMacAddress = adapter.GetPhysicalAddress().ToString();
                }
            }
            return sMacAddress;
        }
        #endregion

        #region Random Numbers
        /// <summary>
        /// Convert Amount in Words [e.g.  AmountInWords(100)]
        /// </summary> 
        /// <param name="types">1 for AlphaNumeric and Default is Numeric  </param>
        /// <param name="length"> lengths </param>
        public string GenerateRandomNum(int types, int length)
        {
            string alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string small_alphabets = "abcdefghijklmnopqrstuvwxyz";
            string numbers = "1234567890";

            string characters = numbers;
            if (types == 1)
            {
                characters += alphabets + small_alphabets + numbers;
            }

            string otp = string.Empty;
            for (int i = 0; i < length; i++)
            {
                string character = string.Empty;
                do
                {
                    int index = new Random().Next(0, characters.Length);
                    character = characters.ToCharArray()[index].ToString();
                } while (otp.IndexOf(character) != -1);
                otp += character;
            }
            return otp;
        }
        #endregion

        #region Encrypt Decrypt
        ///<summary>
        /// <para>Encrypt Function For Encrypt Any String Value With User Define Key  </para>
        /// <param name="strText">enter value to Encrypt.</param>
        /// <param name="key">enter Encryption Key.</param>
        ///</summary>
        public string Encrypt(string strText, string key)
        {
            byte[] IV = { 12, 34, 56, 78, 01, 54, 70, 31 };
            try
            {
                byte[] bykey = Encoding.UTF8.GetBytes(key);
                byte[] InputByteArray = Encoding.UTF8.GetBytes(strText);
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(bykey, IV), CryptoStreamMode.Write);
                cs.Write(InputByteArray, 0, InputByteArray.Length);
                cs.FlushFinalBlock();
                return Convert.ToBase64String(ms.ToArray());
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        ///<summary>
        /// <para>Encrypt Function For Encrypt Any String Value  </para>
        /// <param name="strText">enter value to Encrypt.</param>
        ///</summary>
        public string Encrypt(string strText)
        {
            byte[] IV = { 12, 34, 56, 78, 01, 54, 70, 31 };
            string key = "&%#@?,:*";
            try
            {
                byte[] bykey = Encoding.UTF8.GetBytes(key);
                byte[] InputByteArray = Encoding.UTF8.GetBytes(strText);
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(bykey, IV), CryptoStreamMode.Write);
                cs.Write(InputByteArray, 0, InputByteArray.Length);
                cs.FlushFinalBlock();
                return Convert.ToBase64String(ms.ToArray());
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        ///<summary>
        /// <para>Decrypt Function For Encrypt Any String Value With User Define Key  </para>
        /// <param name="strText">enter value to Encrypt.</param>
        /// <param name="sDecrKey">enter Decryption Key.</param>
        ///</summary>
        public string Decrypt(string strText, string sDecrKey)
        {
            byte[] IV = { 12, 34, 56, 78, 01, 54, 70, 31 };
            byte[] inputByteArray = new byte[strText.Length + 1];
            try
            {
                byte[] byKey = System.Text.Encoding.UTF8.GetBytes(sDecrKey.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                inputByteArray = Convert.FromBase64String(strText);
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write);
                cs.Write(inputByteArray, 0, inputByteArray.Length);
                cs.FlushFinalBlock();
                System.Text.Encoding encoding = System.Text.Encoding.UTF8;
                return encoding.GetString(ms.ToArray());
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        ///<summary>
        /// <para>Decrypt Function For Encrypt Any String Value  </para>
        /// <param name="strText">enter value to Encrypt.</param>
        ///</summary>
        public string Decrypt(string strText)
        {
            byte[] IV = { 12, 34, 56, 78, 01, 54, 70, 31 };
            string sDecrKey = "&%#@?,:*";
            byte[] inputByteArray = new byte[strText.Length + 1];
            try
            {
                byte[] byKey = System.Text.Encoding.UTF8.GetBytes(sDecrKey.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                inputByteArray = Convert.FromBase64String(strText);
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write);
                cs.Write(inputByteArray, 0, inputByteArray.Length);
                cs.FlushFinalBlock();
                System.Text.Encoding encoding = System.Text.Encoding.UTF8;
                return encoding.GetString(ms.ToArray());
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        #endregion

        #region Export to PDF,Excel,Word,CSV 
        ///<summary>
        /// <para>Export Sql Data in Word Format</para>
        ///</summary>
        public void ExportToWord(DataTable DT, string outputFileName)
        {
            //Create a dummy GridView
            GridView GridView1 = new GridView();
            GridView1.AllowPaging = false;
            GridView1.DataSource = DT;
            GridView1.DataBind();
            for (int c = 0; c < DT.Columns.Count; c++)
            {
                GridView1.HeaderRow.Cells[c].BackColor = System.Drawing.ColorTranslator.FromHtml("#fffac3");
                GridView1.HeaderRow.BackColor = System.Drawing.ColorTranslator.FromHtml("#fffac3");
                GridView1.HeaderRow.ForeColor = System.Drawing.Color.Black;
                GridView1.HeaderRow.Font.Bold = true;
                GridView1.HeaderRow.Font.Size = 12;
                GridView1.HeaderRow.Font.Name = "Cambria";
            }
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" + outputFileName + ".doc");
            HttpContext.Current.Response.Charset = "";
            HttpContext.Current.Response.ContentType = "application/vnd.ms-word ";
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            for (int i = 0; i < GridView1.Rows.Count; i++)
            {

                for (int c = 0; c < DT.Columns.Count; c++)
                {
                    if (i % 2 == 0)
                    {
                        GridView1.Rows[i].Cells[c].BackColor = System.Drawing.ColorTranslator.FromHtml("#fff");

                    }
                    else
                    {
                        GridView1.Rows[i].Cells[c].BackColor = System.Drawing.Color.LightGray;
                    }
                }
            }
            GridView1.RenderControl(hw);
            HttpContext.Current.Response.Output.Write(sw.ToString());
            HttpContext.Current.Response.Flush();
            HttpContext.Current.Response.End();
        }

        ///<summary>
        /// <para>Export Sql Data in Excel Format</para>
        ///</summary>
        public void ExportToExcel(DataTable DT, string outputFileName)
        {
            //Create a dummy GridView
            GridView GridView1 = new GridView();
            GridView1.AllowPaging = false;
            GridView1.DataSource = DT;
            GridView1.DataBind();
            for (int c = 0; c < DT.Columns.Count; c++)
            {
                GridView1.HeaderRow.Cells[c].BackColor = System.Drawing.ColorTranslator.FromHtml("#fffac3");
                GridView1.HeaderRow.BackColor = System.Drawing.ColorTranslator.FromHtml("#fffac3");
                GridView1.HeaderRow.ForeColor = System.Drawing.Color.Black;
                GridView1.HeaderRow.Font.Bold = true;
                GridView1.HeaderRow.Font.Size = 12;
                GridView1.HeaderRow.Font.Name = "Cambria";
            }

            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.AddHeader("content-disposition",
             "attachment;filename=" + outputFileName + ".xls");
            HttpContext.Current.Response.Charset = "";
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);

            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                //Apply text style to each Row
                GridView1.Rows[i].Attributes.Add("class", "textmode");

                for (int c = 0; c < DT.Columns.Count; c++)
                {
                    if (i % 2 == 0)
                    {
                        GridView1.Rows[i].Cells[c].BackColor = System.Drawing.ColorTranslator.FromHtml("#fff");

                    }
                    else
                    {
                        GridView1.Rows[i].Cells[c].BackColor = System.Drawing.Color.LightGray;
                    }
                }
            }
            GridView1.RenderControl(hw);

            //style to format numbers to string
            string style = @"<style> .textmode { mso-number-format:\@; } </style>";
            HttpContext.Current.Response.Write(style);
            HttpContext.Current.Response.Output.Write(sw.ToString());
            HttpContext.Current.Response.Flush();
            HttpContext.Current.Response.End();
        }

        ///<summary>
        /// <para>Export Sql Data in PDF Format</para>
        ///</summary>
        public void ExportToPDF(DataTable DT, string outputFileName)
        {
            //Create a dummy GridView
            GridView GridView1 = new GridView();
            GridView1.AllowPaging = false;
            GridView1.DataSource = DT;
            GridView1.DataBind();
            for (int c = 0; c < DT.Columns.Count; c++)
            {
                GridView1.HeaderRow.Cells[c].BackColor = System.Drawing.ColorTranslator.FromHtml("#fffac3");
                GridView1.HeaderRow.BackColor = System.Drawing.ColorTranslator.FromHtml("#fffac3");
                GridView1.HeaderRow.ForeColor = System.Drawing.Color.Black;
                GridView1.HeaderRow.Font.Bold = true;
                GridView1.HeaderRow.Font.Size = 12;
                GridView1.HeaderRow.Font.Name = "Cambria";
            }

            HttpContext.Current.Response.ContentType = "application/pdf";
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" + outputFileName + ".pdf");
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);
            for (int i = 0; i < GridView1.Rows.Count; i++)
            {

                for (int c = 0; c < DT.Columns.Count; c++)
                {
                    if (i % 2 == 0)
                    {
                        GridView1.Rows[i].Cells[c].BackColor = System.Drawing.ColorTranslator.FromHtml("#fff");

                    }
                    else
                    {
                        GridView1.Rows[i].Cells[c].BackColor = System.Drawing.Color.LightGray;
                    }
                }
            }
            GridView1.RenderControl(hw);
            StringReader sr = new StringReader(sw.ToString());
            Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
            HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
            PdfWriter.GetInstance(pdfDoc, HttpContext.Current.Response.OutputStream);
            pdfDoc.Open();
            htmlparser.Parse(sr);
            pdfDoc.Close();
            HttpContext.Current.Response.Write(pdfDoc);
            HttpContext.Current.Response.End();
        }

        ///<summary>
        /// <para>Export Sql Data in CSV Format</para>
        ///</summary>
        public void ExportToCSV(DataTable DT, string outputFileName)
        {

            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" + outputFileName + ".csv");
            HttpContext.Current.Response.Charset = "";
            HttpContext.Current.Response.ContentType = "application/text";


            StringBuilder sb = new StringBuilder();
            for (int k = 0; k < DT.Columns.Count; k++)
            {
                //add separator
                sb.Append(DT.Columns[k].ColumnName + ',');
            }
            //append new line
            sb.Append("\r\n");
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                for (int k = 0; k < DT.Columns.Count; k++)
                {
                    //add separator
                    sb.Append(DT.Rows[i][k].ToString().Replace(",", ";") + ',');
                }
                //append new line
                sb.Append("\r\n");
            }
            HttpContext.Current.Response.Output.Write(sb.ToString());
            HttpContext.Current.Response.Flush();
            HttpContext.Current.Response.End();
        }


        /// <summary>
        /// Export Sql Data in Excel Format (Using Excel Template )
        /// </summary> 
        /// <param name="path">pre formated excel path with file name</param>
        /// <param name="dt">SQl Data table</param>
        /// <param name="fileName">Enter output file name without extension</param>
        /// <param name="headerNo">Enter Header Number start data printing  </param>
        /// <param name="sheetName">Enter excel template Sheet Name [e.g Sheet1]</param>
        public void ExportToExcel(string path, DataTable dt, string fileName, int headerNo, string sheetName)
        {
            try
            {
                FileInfo template = new FileInfo(HttpContext.Current.Server.MapPath(path));
                using (ExcelPackage xlPackage = new ExcelPackage(template, false))
                {
                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[sheetName];

                    worksheet.Name = sheetName;

                    AddExcelSheet(xlPackage, worksheet, dt, headerNo, sheetName);

                    xlPackage.Workbook.Worksheets.Delete(1);
                    HttpContext.Current.Response.Clear();
                    HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    HttpContext.Current.Response.AddHeader("content-disposition", "attachment;  filename = " + fileName + ".xlsx");
                    HttpContext.Current.Response.BinaryWrite(xlPackage.GetAsByteArray());
                    HttpContext.Current.Response.End();
                }

            }
            catch (Exception ex)
            {

            }
        }
        private void AddExcelSheet(ExcelPackage xlPackage, ExcelWorksheet oSheetTemplate, DataTable dt, int headerNo, string sheetName)
        {
            ExcelWorksheet oSheet;
            oSheet = xlPackage.Workbook.Worksheets.Copy(sheetName, sheetName + " -" + System.DateTime.Now.ToString("yyyyMMddHHmmssffff"));
            int count = headerNo;
            oSheet.Cells["A" + count.ToString()].LoadFromDataTable(dt, false);
        }
        #endregion

        #region Create Directory or Folder
        ///<summary>
        /// <para>Create Empty Directory or Folder to user define location.</para>
        /// <param name="path">enter path.</param>
        ///</summary>
        public void CreateDirectory(string path)
        {
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
        }
        #endregion

        #region take nth first position
        ///<summary>
        /// <para>Take First N'th Character From String Type Value.</para>
        /// <param name="s">enter string.</param>
        /// <param name="len">enter lenght you want.</param>
        /// <para>e.g.  input - ("Chandan",2)   output - ch</para>
        ///</summary>
        public string TakeNthFirstPosition(string s, int len)
        {
            if (string.IsNullOrEmpty(s))
            {
                return String.Empty;
            }
            else if (s.Length > len)
            {
                return string.Format("{0}", s.Substring(0, len));
            }
            else
            {
                return s;
            }
        }

        #endregion

        #region Bind Printer
        /// <summary>
        /// Bind Printer Installed in your Server
        /// </summary> 
        /// <param name="DropDownListName">yor drop down list</param>
        public void BindPrinter(DropDownList DropDownListName)
        {
            DropDownListName.Items.Clear();
            foreach (var Printer in PrinterSettings.InstalledPrinters)
            {
                DropDownListName.Items.Add(Printer.ToString());
            }
        }

        #endregion

        #region Fill DataTable

        /// <summary>
        /// FillDataTable
        /// </summary> 
        /// <param name="sqlQuery">Pass SqlQuery </param>
        public DataTable FillDataTable(string sqlQuery)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
                sda.SelectCommand.CommandTimeout = 0;
                DataTable DT = new DataTable();
                sda.Fill(DT);
                return DT;
            }
        }

        /// <summary>
        /// FillDataTable
        /// </summary> 
        /// <param name="sqlQuery">Pass SqlQuery </param>
        /// <param name="con">Pass SqlConnection </param>
        public DataTable FillDataTable(string sqlQuery, SqlConnection con)
        {
            SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
            sda.SelectCommand.CommandTimeout = 0;
            DataTable DT = new DataTable();
            sda.Fill(DT);
            return DT;
        }


        /// <summary>
        /// FillDataTable  
        /// </summary> 
        /// <param name="com">Pass SqlCommand </param>
        public DataTable FillDataTable(SqlCommand com)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                com.Connection = con;
                com.CommandTimeout = 0;
                SqlDataAdapter da = new SqlDataAdapter(com);
                DataTable dt = new DataTable();
                da.Fill(dt);
                return dt;
            }
        }

        /// <summary>
        /// FillDataTable
        /// </summary> 
        /// <param name="com">Pass SqlCommand </param>
        /// <param name="con">Pass SqlConnection </param>
        public DataTable FillDataTable(SqlCommand com, SqlConnection con)
        {
            com.Connection = con;
            com.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(com);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }





        #endregion

        #region Fill DataSet

        /// <summary>
        /// FillDataTable
        /// </summary> 
        /// <param name="sqlQuery">Pass SqlQuery </param>
        public DataSet FillDataSet(string sqlQuery)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
                sda.SelectCommand.CommandTimeout = 0;
                DataSet DS = new DataSet();
                sda.Fill(DS);
                return DS;
            }
        }

        /// <summary>
        /// FillDataSet
        /// </summary> 
        /// <param name="sqlQuery">Pass SqlQuery </param>
        /// <param name="con">Pass SqlConnection </param>
        public DataSet FillDataSet(string sqlQuery, SqlConnection con)
        {
            SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
            sda.SelectCommand.CommandTimeout = 0;
            DataSet DS = new DataSet();
            sda.Fill(DS);
            return DS;
        }


        /// <summary>
        /// FillDataSet
        /// </summary> 
        /// <param name="com">Pass SqlCommand </param>
        public DataSet FillDataSet(SqlCommand com)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                com.Connection = con;
                com.CommandTimeout = 0;
                SqlDataAdapter da = new SqlDataAdapter(com);
                DataSet ds = new DataSet();
                da.Fill(ds);
                return ds;
            }
        }

        /// <summary>
        /// FillDataSet
        /// </summary> 
        /// <param name="com">Pass SqlCommand </param>
        /// <param name="con">Pass SqlConnection </param>
        public DataSet FillDataSet(SqlCommand com, SqlConnection con)
        {
            com.Connection = con;
            com.CommandTimeout = 0;
            SqlDataAdapter da = new SqlDataAdapter(com);
            DataSet ds = new DataSet();
            da.Fill(ds);
            return ds;
        }





        #endregion

        #region Fill DropDownList

        /// <summary>
        /// Bind Drop Down List 
        /// </summary> 
        /// <param name="DropDownName">DropDownName name</param>
        /// <param name="sqlQuery">SQL Query </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>
        /// <param name="header">Header Text [if No Header use ""]</param>
        /// <param name="IsHeaderDisable">1 for disable header</param>
        public void BindDropDown(DropDownList DropDownName, string sqlQuery, string val, string text, string header, int IsHeaderDisable)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                DataTable dt = new DataTable();
                dt = FillDataTable(sqlQuery, con);
                if (dt.Rows.Count > 0)
                {
                    DropDownName.DataSource = dt;
                    DropDownName.DataTextField = text;
                    DropDownName.DataValueField = val;
                    DropDownName.DataBind();
                    if (header != "")
                    {
                        DropDownName.Items.Add(new System.Web.UI.WebControls.ListItem(header, "-1"));
                        DropDownName.Items.FindByText(header).Selected = true;
                        if (IsHeaderDisable == 1)
                        {
                            DropDownName.Items.FindByText(header).Attributes["disabled"] = "disabled";
                        }
                    }
                }
                else
                {
                    DropDownName.Items.Clear();
                }
            }
        }

        /// <summary>
        /// Bind Drop Down List 
        /// </summary> 
        /// <param name="DropDownName">DropDownName name</param>
        /// <param name="sqlQuery">SQL Query </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>
        /// <param name="con">SQL Connection</param>
        /// <param name="header">Header Text [if No Header use ""]</param>
        /// <param name="IsHeaderDisable">1 for disable header</param>
        public void BindDropDown(DropDownList DropDownName, string sqlQuery, string val, string text, SqlConnection con, string header, int IsHeaderDisable)
        {
            DataTable dt = new DataTable();
            dt = FillDataTable(sqlQuery, con);
            if (dt.Rows.Count > 0)
            {
                DropDownName.DataSource = dt;
                DropDownName.DataTextField = text;
                DropDownName.DataValueField = val;
                DropDownName.DataBind();
                if (header != "")
                {
                    DropDownName.Items.Add(new System.Web.UI.WebControls.ListItem(header, "-1"));
                    DropDownName.Items.FindByText(header).Selected = true;
                    if (IsHeaderDisable == 1)
                    {
                        DropDownName.Items.FindByText(header).Attributes["disabled"] = "disabled";
                    }
                }
            }
            else
            {
                DropDownName.Items.Clear();
            }
        }

        /// <summary>
        /// Bind Drop Down List 
        /// </summary> 
        /// <param name="DropDownName">DropDownName name</param>
        /// <param name="dt">Datat Table </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>
        /// <param name="header">Header Text [if No Header use ""]</param>
        /// <param name="IsHeaderDisable">1 for disable header</param>
        public void BindDropDown(DropDownList DropDownName, DataTable dt, string val, string text, string header, int IsHeaderDisable)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();

                if (dt.Rows.Count > 0)
                {
                    DropDownName.DataSource = dt;
                    DropDownName.DataTextField = text;
                    DropDownName.DataValueField = val;
                    DropDownName.DataBind();
                    if (header != "")
                    {
                        DropDownName.Items.Add(new System.Web.UI.WebControls.ListItem(header, "-1"));
                        DropDownName.Items.FindByText(header).Selected = true;
                        if (IsHeaderDisable == 1)
                        {
                            DropDownName.Items.FindByText(header).Attributes["disabled"] = "disabled";
                        }
                    }
                }
                else
                {
                    DropDownName.Items.Clear();
                }
            }
        }

        /// <summary>
        /// Bind Drop Down List 
        /// </summary> 
        /// <param name="DropDownName">DropDownName name</param>
        /// <param name="dt">Datat Table </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>
        /// <param name="con">SQL Connection</param>
        ///  <param name="header">Header Text [if No Header use ""]</param>
        /// <param name="IsHeaderDisable">1 for disable header</param>
        public void BindDropDown(DropDownList DropDownName, DataTable dt, string val, string text, SqlConnection con, string header, int IsHeaderDisable)
        {
            if (dt.Rows.Count > 0)
            {
                DropDownName.DataSource = dt;
                DropDownName.DataTextField = text;
                DropDownName.DataValueField = val;
                DropDownName.DataBind();
                if (header != "")
                {
                    DropDownName.Items.Add(new System.Web.UI.WebControls.ListItem(header, "-1"));
                    DropDownName.Items.FindByText(header).Selected = true;
                    if (IsHeaderDisable == 1)
                    {
                        DropDownName.Items.FindByText(header).Attributes["disabled"] = "disabled";
                    }
                }
            }
            else
            {
                DropDownName.Items.Clear();
            }
        }



        #endregion

        #region Bind ListBox

        /// <summary>
        ///   Bind ListBox
        /// </summary> 
        /// <param name="ListBoxName">ListBox name</param>
        /// <param name="sqlQuery">SQL Query </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>

        public void BindListBox(ListBox ListBoxName, string sqlQuery, string val, string text)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                DataTable dt = new DataTable();
                dt = FillDataTable(sqlQuery, con);
                if (dt.Rows.Count > 0)
                {
                    ListBoxName.DataSource = dt;
                    ListBoxName.DataTextField = text;
                    ListBoxName.DataValueField = val;
                    ListBoxName.DataBind();

                }
            }
        }

        /// <summary>
        /// Bind ListBox
        /// </summary> 
        /// <param name="ListBoxName">ListBox name</param>
        /// <param name="sqlQuery">SQL Query </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>
        /// <param name="con">SQL Connection</param>

        public void BindListBox(ListBox ListBoxName, string sqlQuery, string val, string text, SqlConnection con)
        {
            DataTable dt = new DataTable();
            dt = FillDataTable(sqlQuery, con);
            if (dt.Rows.Count > 0)
            {
                ListBoxName.DataSource = dt;
                ListBoxName.DataTextField = text;
                ListBoxName.DataValueField = val;
                ListBoxName.DataBind();
            }
        }

        /// <summary>
        /// Bind ListBox
        /// </summary> 
        /// <param name="ListBoxName">ListBoxName name</param>
        /// <param name="dt">Datat Table </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>

        public void BindListBox(DropDownList ListBoxName, DataTable dt, string val, string text)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();

                if (dt.Rows.Count > 0)
                {
                    ListBoxName.DataSource = dt;
                    ListBoxName.DataTextField = text;
                    ListBoxName.DataValueField = val;
                    ListBoxName.DataBind();

                }
            }
        }

        /// <summary>
        /// Bind ListBox
        /// </summary> 
        /// <param name="ListBoxName">ListBoxName name</param>
        /// <param name="dt">Datat Table </param>
        /// <param name="val">DropDownList Item Values</param>
        /// <param name="text">DropDownList Display Test</param>
        /// <param name="con">SQL Connection</param>

        public void BindListBox(DropDownList ListBoxName, DataTable dt, string val, string text, SqlConnection con)
        {
            if (dt.Rows.Count > 0)
            {
                ListBoxName.DataSource = dt;
                ListBoxName.DataTextField = text;
                ListBoxName.DataValueField = val;
                ListBoxName.DataBind();

            }
        }



        #endregion

        #region Bind GridView 

        /// <summary>
        /// BindGridView
        /// </summary>  
        /// <param name="gv">GridView Name </param>
        /// <param name="sqlQuery">Pass SqlQuery </param>
        public void BindGridView(GridView gv, string sqlQuery)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
                sda.SelectCommand.CommandTimeout = 0;
                DataTable DT = new DataTable();
                sda.Fill(DT);
                gv.DataSource = DT;
                gv.DataBind();

            }
        }

        /// <summary>
        /// BindGridView
        /// </summary> 
        /// <param name="gv">GridView Name </param>
        /// <param name="sqlQuery">Pass SqlQuery </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindGridView(GridView gv, string sqlQuery, SqlConnection con)
        {
            SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
            sda.SelectCommand.CommandTimeout = 0;
            DataTable DT = new DataTable();
            sda.Fill(DT);
            gv.DataSource = DT;
            gv.DataBind();
        }

        /// <summary>
        /// BindGridView
        /// </summary>  
        /// <param name="gv">GridView Name </param>
        /// <param name="dt">Pass DataTable </param>
        public void BindGridView(GridView gv, DataTable dt)
        {
            using (SqlConnection con = new SqlConnection())
            {
                gv.DataSource = dt;
                gv.DataBind();

            }
        }

        /// <summary>
        /// BindGridView
        /// </summary> 
        /// <param name="gv">GridView Name </param>
        /// <param name="dt">Pass DataTable </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindGridView(GridView gv, DataTable dt, SqlConnection con)
        {

            gv.DataSource = dt;
            gv.DataBind();
        }

        /// <summary>
        /// BindGridView
        /// </summary>  
        /// <param name="gv">GridView Name </param>
        /// <param name="ds">Pass DataSet </param>
        public void BindGridView(GridView gv, DataSet ds)
        {
            using (SqlConnection con = new SqlConnection())
            {
                gv.DataSource = ds;
                gv.DataBind();

            }
        }

        /// <summary>
        /// BindGridView
        /// </summary> 
        /// <param name="gv">GridView Name </param>
        /// <param name="ds">Pass DataSet </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindGridView(GridView gv, DataSet ds, SqlConnection con)
        {

            gv.DataSource = ds;
            gv.DataBind();
        }


        #endregion

        #region Bind RadioButtonList

        /// <summary>
        /// BindRadioButtonList
        /// </summary>  
        /// <param name="rdo">RadioButtonList Name </param>
        /// <param name="sqlQuery">Pass SqlQuery </param>
        /// <param name="TextField">Display Text Field </param>
        /// <param name="ValueField">Value Field </param>
        public void BindRadioButtonList(RadioButtonList rdo, string sqlQuery, string TextField, string ValueField)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
                sda.SelectCommand.CommandTimeout = 0;
                DataTable DT = new DataTable();
                sda.Fill(DT);
                rdo.DataSource = DT;
                rdo.DataTextField = TextField;
                rdo.DataValueField = ValueField;
                rdo.DataBind();
            }
        }

        /// <summary>
        /// BindRadioButtonList
        /// </summary> 
        /// <param name="rdo">RadioButtonList Name </param>
        /// <param name="sqlQuery">Pass SqlQuery </param>
        /// <param name="TextField">Display Text Field </param>
        /// <param name="ValueField">Value Field </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindRadioButtonList(RadioButtonList rdo, string sqlQuery, string TextField, string ValueField, SqlConnection con)
        {
            SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
            sda.SelectCommand.CommandTimeout = 0;
            DataTable DT = new DataTable();
            sda.Fill(DT);
            rdo.DataSource = DT;
            rdo.DataTextField = TextField;
            rdo.DataValueField = ValueField;
            rdo.DataBind();
        }

        /// <summary>
        /// BindRadioButtonList
        /// </summary>  
        /// <param name="rdo">RadioButtonList Name </param>
        /// <param name="TextField">Display Text Field </param>
        /// <param name="ValueField">Value Field </param>
        /// <param name="dt">Pass DataTable </param>
        public void BindRadioButtonList(RadioButtonList rdo, DataTable dt, string TextField, string ValueField)
        {
            using (SqlConnection con = new SqlConnection())
            {
                rdo.DataSource = dt;
                rdo.DataTextField = TextField;
                rdo.DataValueField = ValueField;
                rdo.DataBind();

            }
        }

        /// <summary>
        /// BindRadioButtonList
        /// </summary> 
        /// <param name="rdo">RadioButtonList Name </param>
        /// <param name="TextField">Display Text Field </param>
        /// <param name="ValueField">Value Field </param>
        /// <param name="dt">Pass DataTable </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindRadioButtonList(RadioButtonList rdo, DataTable dt, string TextField, string ValueField, SqlConnection con)
        {

            rdo.DataSource = dt;
            rdo.DataTextField = TextField;
            rdo.DataValueField = ValueField;
            rdo.DataBind();
        }

        /// <summary>
        /// BindRadioButtonList
        /// </summary>  
        /// <param name="rdo">RadioButtonList Name </param>
        /// <param name="TextField">Display Text Field </param>
        /// <param name="ValueField">Value Field </param>
        /// <param name="ds">Pass DataSet </param>
        public void BindRadioButtonList(RadioButtonList rdo, DataSet ds, string TextField, string ValueField)
        {
            using (SqlConnection con = new SqlConnection())
            {
                rdo.DataSource = ds;
                rdo.DataTextField = TextField;
                rdo.DataValueField = ValueField;
                rdo.DataBind();

            }
        }

        /// <summary>
        /// BindRadioButtonList
        /// </summary> 
        /// <param name="rdo">RadioButtonList Name </param>
        /// <param name="TextField">Display Text Field </param>
        /// <param name="ValueField">Value Field </param>
        /// <param name="ds">Pass DataSet </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindRadioButtonList(RadioButtonList rdo, DataSet ds, string TextField, string ValueField, SqlConnection con)
        {
            rdo.DataSource = ds;
            rdo.DataTextField = TextField;
            rdo.DataValueField = ValueField;
            rdo.DataBind();
        }


        #endregion

        #region Bind ListView 

        /// <summary>
        /// BindListView
        /// </summary>  
        /// <param name="lv">ListView Name </param>
        /// <param name="sqlQuery">Pass SqlQuery </param>
        public void BindListView(ListView lv, string sqlQuery)
        {
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
                sda.SelectCommand.CommandTimeout = 0;
                DataTable DT = new DataTable();
                sda.Fill(DT);
                lv.DataSource = DT;
                lv.DataBind();

            }
        }

        /// <summary>
        /// BindListView
        /// </summary> 
        /// <param name="lv">ListView Name </param>
        /// <param name="sqlQuery">Pass SqlQuery </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindListView(ListView lv, string sqlQuery, SqlConnection con)
        {
            SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
            sda.SelectCommand.CommandTimeout = 0;
            DataTable DT = new DataTable();
            sda.Fill(DT);
            lv.DataSource = DT;
            lv.DataBind();
        }

        /// <summary>
        /// BindListView
        /// </summary>  
        /// <param name="lv">ListView Name </param>
        /// <param name="dt">Pass DataTable </param>
        public void BindListView(ListView lv, DataTable dt)
        {
            using (SqlConnection con = new SqlConnection())
            {
                lv.DataSource = dt;
                lv.DataBind();

            }
        }

        /// <summary>
        /// BindGridView
        /// </summary> 
        /// <param name="gv">GridView Name </param>
        /// <param name="dt">Pass DataTable </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindListView(ListView lv, DataTable dt, SqlConnection con)
        {

            lv.DataSource = dt;
            lv.DataBind();
        }

        /// <summary>
        /// BindListView
        /// </summary>  
        /// <param name="lv">ListView Name </param>
        /// <param name="ds">Pass DataSet </param>
        public void BindListView(ListView lv, DataSet ds)
        {
            using (SqlConnection con = new SqlConnection())
            {
                lv.DataSource = ds;
                lv.DataBind();

            }
        }

        /// <summary>
        /// BindListView
        /// </summary> 
        /// <param name="lv">ListView Name </param>
        /// <param name="ds">Pass DataSet </param>
        /// <param name="con">Pass SqlConnection </param>
        public void BindListView(ListView lv, DataSet ds, SqlConnection con)
        {

            lv.DataSource = ds;
            lv.DataBind();
        }


        #endregion

        #region Validate Mobile,Email,IsDateFormat,IsNumeric
        /// <summary>
        /// Validate 10-Digit Mobile Number [if return 1 that is valid]
        /// </summary> 
        /// <param name="number">enter 10-digit mobile number.</param>
        public int ValidateMobile(string number)
        {
            Regex regex = new Regex(@"^[0-9]{10}$");
            Match match = regex.Match(number);
            if (match.Success)
                return 1;
            else
                return 0;
        }
        /// <summary>
        /// Validate Email Address  [if return 1 that is valid]
        /// </summary> 
        /// <param name="email">enter email id.</param>
        public int ValidateEmail(string email)
        {
            Regex regex = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
            Match match = regex.Match(email);
            if (match.Success)
                return 1;
            else
                return 0;
        }

        /// <summary>
        /// Validate Date Format  [if return 1 that is valid]
        /// </summary> 
        /// <param name="inputDate">enter date</param>
        public int IsDateFormat(string inputDate)
        {
            DateTime dDate;
            if (DateTime.TryParse(inputDate, out dDate))
            {
                String.Format("{0:d/MM/yyyy}", dDate);
                return 1;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// Validate Numeric Format  [if return 1 that is valid]
        /// </summary> 
        /// <param name="inputNumber">enter number</param>
        public int IsNumeric(string inputNumber)
        {
            try
            {
                decimal i = decimal.Parse(inputNumber.Trim());
                return 1;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }
        #endregion

        #region IsExist User

        /// <summary>
        /// IsUserExists for Check it exist or not in database [return 1 for Exist]
        /// </summary> 
        /// <param name="sqlQuery">Pass SqlQuery </param>
        public int IsUserExists(string sqlQuery)
        {
            string a;
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                SqlCommand cmd = new SqlCommand(sqlQuery, con);
                a = cmd.ExecuteScalar().ToString();
                if (Convert.ToInt16(a) != 0)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
        }
        /// <summary>
        /// IsUserExists for Check it exist or not in database  [return 1 for Exist]
        /// </summary> 
        /// <param name="sqlQuery">Pass SqlQuery </param>
        /// <param name="con">Pass SqlConnection </param>
        public int IsUserExists(string sqlQuery, SqlConnection con)
        {
            string a;

            SqlCommand cmd = new SqlCommand(sqlQuery, con);
            a = cmd.ExecuteScalar().ToString();
            if (Convert.ToInt16(a) != 0)
            {
                return 1;
            }
            else
            {
                return 0;
            }

        }
        #endregion

        #region Execute Scalar
        /// <summary>
        /// get execute scalar value
        /// </summary> 
        /// <param name="str">SQl Query</param>
        public string Scalar(string str)
        {
            
            SqlConnection cn = new SqlConnection(getConectionString());
            if (cn.State == ConnectionState.Closed) { cn.Open(); }
            SqlCommand com = new SqlCommand();
            com.Connection = cn;
            com.CommandText = str;
            string Str = (string)com.ExecuteScalar();
            return Str;
        }
        /// <summary>
        /// get execute scalar value
        /// </summary> 
        /// <param name="str">SQl Query</param>
        /// <param name="con">SQl Connection</param>
        public string Scalar(string str, SqlConnection con)
        {

            SqlCommand com = new SqlCommand();
            com.Connection = con;
            com.CommandText = str;
            string Str = (string)com.ExecuteScalar();
            return Str;
        }
        #endregion

        #region Insert,Update and Delete In Sql Server Database

        /// <summary>
        /// IUD (Insert,Update and Delete) data in sql server database using SqlCommand
        /// </summary> 
        /// <param name="SqlQuery">Pass Sql Query </param>
        /// <param name="Msg">Pass SqlCommand </param>
        public int IUD(string SqlQuery, out String Msg)
        {
            int tag = -1;
            SqlConnection con = new SqlConnection(getConectionString());
            SqlCommand com = new SqlCommand(SqlQuery, con);
            try
            {
                con.Open();
                tag = com.ExecuteNonQuery();
                Msg = "success";
            }
            catch (SqlException ex)
            {
                tag = ex.ErrorCode;
            }
            catch (Exception ex)
            {
                tag = -999;
            }
            finally
            {
                con.Close();
            }
            Msg = ExceptionMessage(tag);
            return tag;
        }

        /// <summary>
        /// IUD (Insert,Update and Delete)  data in sql server database using SqlCommand
        /// </summary> 
        /// <param name="com">Pass SqlCommand </param>
        /// <param name="Msg">Pass SqlCommand </param>
        public int IUD(SqlCommand com, out String Msg)
        {
            int tag = -1;

            SqlConnection con = new SqlConnection(getConectionString());
            com.Connection = con;
            try
            {
                con.Open();
                tag = com.ExecuteNonQuery();
                Msg = "success";
            }
            catch (SqlException ex)
            {
                tag = ex.ErrorCode;
            }
            catch (Exception ex)
            {
                tag = -999;
            }
            finally
            {
                con.Close();
            }
            Msg = ExceptionMessage(tag);
            return tag;
        }

        /// <summary>
        /// IUD (Insert,Update and Delete) data in sql server database using SqlCommand
        /// </summary> 
        /// <param name="SqlQuery">Pass Sql Query </param>
        /// <param name="con">SQL Connection </param>
        /// <param name="Msg">return error msg </param>
        public int IUD(string SqlQuery, SqlConnection con, out String Msg)
        {
            int tag = -1;

            SqlCommand com = new SqlCommand(SqlQuery, con);
            try
            {
                con.Open();
                tag = com.ExecuteNonQuery();
                Msg = "success";
            }
            catch (SqlException ex)
            {
                tag = ex.ErrorCode;
            }
            catch (Exception ex)
            {
                tag = -999;
            }
            finally
            {
                con.Close();
            }
            Msg = ExceptionMessage(tag);
            return tag;
        }

        /// <summary>
        /// IUD (Insert,Update and Delete)  data in sql server database using SqlCommand
        /// </summary> 
        /// <param name="com">Pass SqlCommand </param>
        /// <param name="con">Pass SqlConnection </param>
        /// <param name="Msg">Pass SqlCommand </param>
        public int IUD(SqlCommand com, SqlConnection con, out String Msg)
        {
            int tag = -1;
            com.Connection = con;
            try
            {
                con.Open();
                tag = com.ExecuteNonQuery();
                Msg = "success";
            }
            catch (SqlException ex)
            {
                tag = ex.ErrorCode;
            }
            catch (Exception ex)
            {
                tag = -999;
            }
            finally
            {
                con.Close();
            }
            Msg = ExceptionMessage(tag);
            return tag;
        }
        #endregion

        #region Message Box
        ///<summary>
        /// <para>ShowMessageBox(this,"Hello Chandan") </para>
        ///</summary>
        public void ShowMessageBox(Page page, string Message)
        {
            page.ClientScript.RegisterStartupScript(page.GetType(), "Msg", "alert('" + Message + "');", true);
        }

        ///<summary>
        /// <para>Show User Define Message and then redirect user define page.</para>
        ///</summary>
        public void ShowMessageBox(Page page, string Message, string navigateUrl)
        {
            string s = "alert('" + Message + "');var vers = navigator.appVersion;if(vers.indexOf('MSIE 7.0') != -1) { window.location.href='" + navigateUrl + "';} else{ window.location.href='" + navigateUrl + "';}";
            page.ClientScript.RegisterStartupScript(page.GetType(), "Information", s, true);
        }
        #endregion


        #region Search a particular word or string on specific database feild after where 
        /// <summary>
        /// Search a particular word or string on specific database feild after where 
        /// </summary> 
        /// <param name="fieldName">Database Field Name</param>
        /// <param name="searchingWord">Searching string seperated by comma or space</param>
        public string SearchValue(string fieldName, string searchingWord)
        {
            string s = " (";


            string[] strArrayOne = new string[] { "" };
            //somewhere in your code
            strArrayOne = searchingWord.Split(',', ' ');
            for (int i = 0; i < strArrayOne.Length; i++)
            {
                s = s + "[" + fieldName + "]" + "  like '%" + strArrayOne[i] + "%' or ";
            }
            s = string.Concat(s.Reverse().Skip(3).Reverse());
            return s + ")";
        }
        #endregion


        /// <summary>
        ///getIndianDateTime() funcation return Indian date Time
        /// </summary> 
        public DateTime getIndianDateTime()
        {
            TimeZoneInfo timeZoneInfo;
            DateTime dateTime;
            //Set the time zone information to US Mountain Standard Time 
            timeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
            //Get date and time in US Mountain Standard Time 
            dateTime = TimeZoneInfo.ConvertTime(DateTime.Now, timeZoneInfo);
            return dateTime;
        }

        /// <summary>
        /// GenerateHTMLTable
        /// </summary> 
        /// <param name="dt">DataTable</param>
        public string GenerateHTMLTable(DataTable dt)
        {
            string div = "";
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    div += "<tr>";
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j == 0)
                        {
                            div += "<td>" + (i + 1) + "</td>";
                        }
                        div += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                    }
                    div += "</tr>";
                }
            }
            return div;
        }
        /// <summary>
        /// GenerateHTMLTable
        /// </summary> 
        /// <param name="sqlQuery">SQL Query</param>
        public string GenerateHTMLTable(string sqlQuery)
        {
            string div = "";
            using (SqlConnection con = new SqlConnection())
            {
                con.ConnectionString = getConectionString();
                SqlDataAdapter sda = new SqlDataAdapter(sqlQuery, con);
                sda.SelectCommand.CommandTimeout = 0;
                DataTable dt = new DataTable();
                sda.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        div += "<tr>";
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            if (j == 0)
                            {
                                div += "<td>" + (i + 1) + "</td>";
                            }
                            div += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                        }
                        div += "</tr>";
                    }
                }
            }
            return div;
        }

        public String CreateProcedure(string TableName)
        {
            SqlConnection cn = new SqlConnection(getConectionString());
            SqlCommand com = new SqlCommand();
            com.CommandText = "GetColumn";
            com.Parameters.AddWithValue("@TableName", TableName);
            com.Connection = cn;
            SqlDataAdapter da = new SqlDataAdapter(com);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int count = dt.Rows.Count;
            string IColumn = "";
            string IColumnValue = "";
            string Insert = "";
            string Update = "";
            string UColumn = "";
            string StrCom = "SqlCommand com = new SqlCommand();<BR/>";
            StrCom += "com.CommandType = CommandType.StoredProcedure;<BR/>";
            StrCom += "com.CommandText =\"" + TableName + "Insert\" ;<BR/>";

            string str = "Create Procedure " + TableName + "Insert";
            str += "<BR/>(";
            for (int i = 0; i < count; i++)
            {
                if (i == 0)
                {
                    str += "<BR/>@" + dt.Rows[i][0].ToString() + " " + (dt.Rows[i][1].ToString() == "varchar" ? dt.Rows[i][1].ToString() + "(" + dt.Rows[i][2].ToString() + ")" : dt.Rows[i][1].ToString());
                    IColumn = dt.Rows[i][0].ToString();
                    IColumnValue = "@" + dt.Rows[i][0].ToString();
                    UColumn = " <BR/>" + dt.Rows[i][0].ToString() + "=" + "@" + dt.Rows[i][0].ToString();
                }
                else
                {
                    str += ",<BR/>@" + dt.Rows[i][0].ToString() + " " + (dt.Rows[i][1].ToString() == "varchar" ? dt.Rows[i][1].ToString() + "(" + dt.Rows[i][2].ToString() + ")" : dt.Rows[i][1].ToString());
                    IColumn += ", " + dt.Rows[i][0].ToString();
                    IColumnValue += ", @" + dt.Rows[i][0].ToString();
                    UColumn += ",<BR/>" + dt.Rows[i][0].ToString() + "=" + "@" + dt.Rows[i][0].ToString();
                }
                StrCom += "com.Parameters.AddWithValue(\"@" + dt.Rows[i][0].ToString() + "\", \"StrPar\");<BR/>";
            }
            Insert = str + ")<BR/>As<BR/>";
            Insert += "Insert Into " + TableName + "(";
            Insert += IColumn;
            Insert += ")<BR/>Values (";
            Insert += IColumnValue + ")";
            Update = str.Replace("Insert", "Update") + ")<BR/>As<BR/>";
            Update += "Update " + TableName + " Set ";
            Update += UColumn;
            return Insert + "<BR/>" + Update + "<BR/>" + StrCom;
        }






        #region Get File Extension
        ///<summary>
        /// <para>Check file extension (.pdf,.jpeg,.jpg,.png)</para>
        ///</summary>
        public bool CheckFileType(string FileName)
        {
            string Ext = Path.GetExtension(FileName);
            switch (Ext.ToUpper())
            {
                case ".PDF":
                    return true;
                    break;

                case ".JPEG":
                    return true;
                    break;

                case ".JPG":
                    return true;
                    break;

                case ".PNG":
                    return true;
                    break;

                default:
                    return false;
                    break;
            }
        }
        #endregion



        #region Mail
        //https://www.google.com/settings/security/lesssecureapps
        ///<summary>
        /// <para>Send Mail Without Any Attachements</para>
        ///</summary>
        public string SendMail(string toEmail, string fromEmail, string fromEmailPass, string titleMsg, string sub, string bodyMsg)
        {

            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            mail.To.Add(toEmail);
            mail.From = new MailAddress(fromEmail, titleMsg, System.Text.Encoding.UTF8);
            mail.Subject = (sub);
            mail.SubjectEncoding = System.Text.Encoding.UTF8;
            mail.Body = (bodyMsg);

            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;
            SmtpClient client = new SmtpClient();
            client.Credentials = new System.Net.NetworkCredential(fromEmail, fromEmailPass);
            client.Port = 587;
            client.Host = "smtp.gmail.com";
            client.EnableSsl = true;
            try
            {
                client.Send(mail);
                return "1";
            }
            catch (Exception ex)
            {
                Exception ex2 = ex;
                string errorMessage = string.Empty;
                while (ex2 != null)
                {
                    errorMessage += ex2.ToString();
                    ex2 = ex2.InnerException;
                }
                return ex2.Message;
            }

        }


        ///<summary>
        /// <para>Send Mail With Single Attachements</para>
        ///</summary>
        public string SendMail(string toEmail, string fromEmail, string fromEmailPass, string attachmentsFilePath, string titleMsg, string sub, string bodyMsg)
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            mail.To.Add(toEmail);
            mail.From = new MailAddress(fromEmail, titleMsg, System.Text.Encoding.UTF8);
            mail.Subject = (sub);
            mail.SubjectEncoding = System.Text.Encoding.UTF8;
            mail.Body = (bodyMsg);
            mail.Attachments.Add(new Attachment(HttpContext.Current.Server.MapPath(attachmentsFilePath)));
            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;
            SmtpClient client = new SmtpClient();
            client.Credentials = new System.Net.NetworkCredential(fromEmail, fromEmailPass);
            client.Port = 587;
            client.Host = "smtp.gmail.com";
            client.EnableSsl = true;
            try
            {
                client.Send(mail);
                return "1";
            }
            catch (Exception ex)
            {
                Exception ex2 = ex;
                string errorMessage = string.Empty;
                while (ex2 != null)
                {
                    errorMessage += ex2.ToString();
                    ex2 = ex2.InnerException;
                }
                return ex2.Message;
            }

        }

        ///<summary>
        /// <para>Send Mail With Multiple Attachements</para>
        ///</summary>
        public string SendMailWithFolder(string toEmail, string fromEmail, string fromEmailPass, string attachmentsFolderPath, string titleMsg, string sub, string bodyMsg)
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            mail.To.Add(toEmail);
            mail.From = new MailAddress(fromEmail, titleMsg, System.Text.Encoding.UTF8);
            mail.Subject = (sub);
            mail.SubjectEncoding = System.Text.Encoding.UTF8;
            mail.Body = (bodyMsg);
            DirectoryInfo dir = new DirectoryInfo(attachmentsFolderPath);
            foreach (FileInfo file in dir.GetFiles("*.*"))
            {
                if (file.Exists)
                {
                    mail.Attachments.Add(new Attachment(file.FullName));
                }
            }
            //mail.Attachments.Add(new Attachment(Server.MapPath(filePath)));
            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;
            SmtpClient client = new SmtpClient();
            client.Credentials = new System.Net.NetworkCredential(fromEmail, fromEmailPass);
            client.Port = 587;
            client.Host = "smtp.gmail.com";
            client.EnableSsl = true;
            try
            {
                client.Send(mail);
                return "1";
            }
            catch (Exception ex)
            {
                Exception ex2 = ex;
                string errorMessage = string.Empty;
                while (ex2 != null)
                {
                    errorMessage += ex2.ToString();
                    ex2 = ex2.InnerException;
                }
                return ex2.Message;
            }

        }


        #endregion








        /// <summary>
        /// Exception Message
        /// </summary> 
        /// <param name="id">Exeption Code</param>
        public string ExceptionMessage(int id)
        {
            string Message = "";
            if (id < 1)
            {

                switch (id)
                {
                    case -999:
                        Message = "error code 999";
                        break;
                    case -1:
                        Message = "";
                        break;
                    default:
                        Message = "something wrong";
                        break;
                }
            }
            return Message;
        }

        

    }




    #region Calculate Age
    /// <summary>  
    /// For calculating age  
    /// </summary>  

    public class Age
    {
        /// <summary>  
        /// Declare Year Variable as int type
        /// </summary>  
        public int Years;
        /// <summary>  
        /// Declare Months Variable as int type
        /// </summary>  
        public int Months;
        /// <summary>  
        /// Declare Days Variable as int type
        /// </summary>  
        public int Days;

        /// <param name="Bday">Enter Date of Birth in Date Format to Calculate the age</param>
        /// <returns> years, months,days</returns>  
        public Age(DateTime Bday)
        {
            this.Count(Bday);
        }
        /// <param name="Bday">Enter Date of Birth in Date Format to Calculate the age</param>
        /// <param name="Cday">Enter Current Date in Date Format to Calculate the age</param>
        /// <returns> years, months,days</returns>  
        public Age(DateTime Bday, DateTime Cday)
        {
            this.Count(Bday, Cday);
        }

        /// <param name="Bday">Enter Date of Birth in Date Format to Calculate the age</param>
        public Age Count(DateTime Bday)
        {
            return this.Count(Bday, DateTime.Today);
        }

        /// <param name="Bday">Enter Date of Birth in Date Format to Calculate the age</param>
        /// <param name="Cday">Enter Current Date in Date Format to Calculate the age</param>
        public Age Count(DateTime Bday, DateTime Cday)
        {
            if ((Cday.Year - Bday.Year) > 0 ||
                (((Cday.Year - Bday.Year) == 0) && ((Bday.Month < Cday.Month) ||
                  ((Bday.Month == Cday.Month) && (Bday.Day <= Cday.Day)))))
            {
                int DaysInBdayMonth = DateTime.DaysInMonth(Bday.Year, Bday.Month);
                int DaysRemain = Cday.Day + (DaysInBdayMonth - Bday.Day);

                if (Cday.Month > Bday.Month)
                {
                    this.Years = Cday.Year - Bday.Year;
                    this.Months = Cday.Month - (Bday.Month + 1) + Math.Abs(DaysRemain / DaysInBdayMonth);
                    this.Days = (DaysRemain % DaysInBdayMonth + DaysInBdayMonth) % DaysInBdayMonth;
                }
                else if (Cday.Month == Bday.Month)
                {
                    if (Cday.Day >= Bday.Day)
                    {
                        this.Years = Cday.Year - Bday.Year;
                        this.Months = 0;
                        this.Days = Cday.Day - Bday.Day;
                    }
                    else
                    {
                        this.Years = (Cday.Year - 1) - Bday.Year;
                        this.Months = 11;
                        this.Days = DateTime.DaysInMonth(Bday.Year, Bday.Month) - (Bday.Day - Cday.Day);
                    }
                }
                else
                {
                    this.Years = (Cday.Year - 1) - Bday.Year;
                    this.Months = Cday.Month + (11 - Bday.Month) + Math.Abs(DaysRemain / DaysInBdayMonth);
                    this.Days = (DaysRemain % DaysInBdayMonth + DaysInBdayMonth) % DaysInBdayMonth;
                }
            }
            else
            {
                throw new ArgumentException("Birthday date must be earlier than current date");
            }
            return this;
        }
    }

    /**
     * Usage example:
     * ==============
     * DateTime bday = new DateTime(1987, 11, 27);
     * DateTime cday = DateTime.Today;
     * Age age = new Age(bday, cday);
     * Console.WriteLine("It's been {0} years, {1} months, and {2} days since your birthday", age.Year, age.Month, age.Day);
     */


    #endregion

}
