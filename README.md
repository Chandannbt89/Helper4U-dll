# **Helper4U.dll** 
Created by : Chandan Kumar (+91 9905241078)

------------
##### **Helper4U.dll** को कैसे Use कैसे करे ?
Helper4U.dll को Use करने के पहले आप को इसे अपने प्रोजेक्ट में Import करना होगा |

Reference ->Add Reference ->Browse ->Select Helper4U.dll ->Add ->OK

> using Helper4U;

##### How to Convert Amount in Words
> SqlHelper4U obj = new SqlHelper4U();

> string amount = obj.AmountInWords(145);

Output :  One Hundred Fourty Five Rupes Only

> SqlHelper4U obj = new SqlHelper4U();

>string amount = obj.AmountInWords(145.457);

Output :  One Hundred and Forty Five rupees and     Forty Five paise only

##### How to Convert Month number to Name
> SqlHelper4U obj = new SqlHelper4U();

>string monthName = obj.MonthName(1, "");

Output : January

> SqlHelper4U obj = new SqlHelper4U();

>string monthName = obj.MonthName(1, "HN");

Output : जनवरी

##### How to Convert Day number to Name
> SqlHelper4U obj = new SqlHelper4U();

>string dayName = obj.DayName(1, "");

Output : Sunday

> SqlHelper4U obj = new SqlHelper4U();

>string dayName = obj.DayName(1, "hn");

Output : रविवार

##### How to get Ipv4 and MAC address
> SqlHelper4U obj = new SqlHelper4U();

>string ipAddress = obj.GetIP4Address();

Output : ::1

> SqlHelper4U obj = new SqlHelper4U();

>string MAC = obj.GetMACAddress();

Output : 001517B6EA19

##### How to generate Random Number
> SqlHelper4U obj = new SqlHelper4U();

>string radomNumber = obj.GenerateRandomNum(1,5);

Output : KkYym (generate 5-Digit random AlphaNumeric)

> SqlHelper4U obj = new SqlHelper4U();

>string radomNumber = obj.GenerateRandomNum(0,5);

Output : 82046 (generate 5-Digit random Numeric)

##### How to Encrypt
>  SqlHelper4U obj = new SqlHelper4U();
 
 >string encryptedString = obj.Encrypt("Chandan");
 
 Output : CHCpEUl18h8=
 
 With User Define Key(&%#@?,:*)
>  SqlHelper4U obj = new SqlHelper4U();
 
 >string encryptedString = obj.Encrypt("Chandan","&%#@?,:*");.
 
  Output : CHCpEUl18h8=
  
  ##### How to  Decrypt 
  >  SqlHelper4U obj = new SqlHelper4U();
 
 >string decryptString = obj.Decrypt("CHCpEUl18h8=");
 
 Output : Chandan
 
 With User Define Key(&%#@?,:*)
>  SqlHelper4U obj = new SqlHelper4U();
 
 >string decryptString = obj.Decrypt("CHCpEUl18h8=","&%#@?,:*");.
 
  Output : Chandan
  
  ##### How to  Export File in Word,Excel,Pdf and CSV Format
 
>  SqlHelper4U obj = new SqlHelper4U();
 
 >obj.ExportToPDF(dt, "myPDFFile");
 
   Output : myPDFFile.pdf
 
 >  SqlHelper4U obj = new SqlHelper4U();
 
 >obj.ExportToWord(dt, "myWordFile");
 
   Output : myWordFile.doc
 
 >  SqlHelper4U obj = new SqlHelper4U();
 
 >obj.ExportToExcel(dt, "myExcelFile");
 
 Output : myExcelFile.xls
 
  >  SqlHelper4U obj = new SqlHelper4U();
  
  >obj.ExportToExcel("formats/Book1.xlsx", dt, "myExcel", 3, "Sheet1");
 
 Output : myExcel.xlsx
 
 >  SqlHelper4U obj = new SqlHelper4U();
 
 >obj.ExportToCSV(dt, "myCSVFile");
 
  Output : myCSVFile.csv
  
  ##### Age Difference
  
  > Age age = new Age(Convert.ToDateTime("10/02/1989"), Convert.ToDateTime("01/01/2019"));
            
  > string ageDifference = age.Years + " years " + age.Months + " months" + age.Days +" days";
  
  Output : 29 years 10 months19 days
   
  ##### Create Directory or Folder

  > SqlHelper4U obj = new SqlHelper4U();
  
  > obj.CreateDirectory("~/formats/abc");

  ##### Create Directory or Folder

  > SqlHelper4U obj = new SqlHelper4U();
   
  > string firstNcharacter = obj.TakeNthFirstPosition("Chandan", 2);

  Output : Ch

  ##### Bind Printer Installed in your Server

  > SqlHelper4U obj = new SqlHelper4U();

  > obj.BindPrinter(ddlPrinter);

  Output : Show All Printer in DropDownList

 ##### Fill DataTable 

  > SqlHelper4U obj = new SqlHelper4U();

  > DataTable dt = new DataTable();
    
  > dt = obj.FillDataTable("SELECT DoctorTypeID as ID,DoctorType as [Doctor Type] FROM tblTypeOfDoctor");

  or (Execute Procedure)

  > dt= obj.FillDataTable("Exec ProcedureName @Parameter='" + parameterValue  + "'");

  Output : Fill DataTable Without Connection

  > SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["cons"].ConnectionString);

  > SqlHelper4U obj = new SqlHelper4U();

  > DataTable dt = new DataTable();
    
  > dt = obj.FillDataTable("SELECT DoctorTypeID as ID,DoctorType as [Doctor Type] FROM tblTypeOfDoctor",con);

  Output : Fill DataTable with Own Connection

   ##### Insert,Update and Delete In Sql Server Database

  > SqlHelper4U obj = new SqlHelper4U();

  > int a = obj.IUD("delete from test ",out string s);
