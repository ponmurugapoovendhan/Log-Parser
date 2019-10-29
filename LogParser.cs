using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Activities;
using System.ComponentModel;


namespace Log_Parser
{
    public  class Log_Parser : CodeActivity
    {


        [Category("Input")]
        [RequiredArgument]
        [Description("The Path of the output log file without extension.If the file already exists, it will be overwritten")]
        public InArgument<string> Path { get; set; }
        public enum ddEnum
        {
            html,
            xlsx,
            txt
        }

        [Category("Input")]
        [RequiredArgument]
        
        public ddEnum Format{ get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string Out_File_Path = OutputFilePath.Get(context);
            string File_Format = FileFormat.ToString();
            string Log_File_Path = "";
            string[] Log_files = Directory.GetFiles(Environment.ExpandEnvironmentVariables("%LocalAppData%/UiPath/Logs/"));
            StreamWriter Text_File = null;
            StreamWriter Html_File = null;
            Application Excel_App = null;
            Workbook wb=null;
            Worksheet ws=null;

            DateTime Mod_Time = DateTime.MinValue;
            foreach (string Log_File in Log_files)
            {
                if (File.GetLastWriteTime(Log_File) > Mod_Time && Log_File.Contains("_Execution"))
                {
                    Mod_Time = File.GetLastWriteTime(Log_File);
                    Log_File_Path = Log_File;
                }

            }
            //Console.WriteLine(Log_File_Path);
            System.IO.StreamReader file = new System.IO.StreamReader(Log_File_Path);
            if(File_Format == "txt")
            Text_File = new StreamWriter(Out_File_Path + ".txt");

            if (File_Format == "html")
            {
                Html_File = new StreamWriter(Out_File_Path + ".html");
                Html_File.WriteLine("<html>");
            }

            if (File_Format == "xlsx")
            {
                Excel_App = new Application();
                wb = Excel_App.Workbooks.Add();
                ws = wb.Sheets[1];
                ws.Columns[2].NumberFormat = "HH:MM:SS";
            }


            string Job_Id_Pattern = "\"jobId\":\"(.*)\",\"robotName\"";
            string line;
            string Job_Id = "";
            //string pattern = "(.*) (.*) {\"message\":\"(.*)\",\"level\":\"(.*)\",\"logType\":\"(.*)\",\"timeStamp\":\"(.*)\",\"fingerprint\":\"(.*)\",\"windowsIdentity\":\"(.*)\",\"machineName\":\"(.*)\",\"processName\":\"(.*)\",\"processVersion\":\"(.*)\",\"jobId\":\"(.*)\",\"robotName\":\"(.*)\",\"machineId\":(.*)";
            while ((line = file.ReadLine()) != null)
            {
                //Console.WriteLine(line);
                MatchCollection Job_Id_Matches = Regex.Matches(line, Job_Id_Pattern);
                Match match = Job_Id_Matches[0];

                Job_Id = match.Groups[1].ToString();

            }
            file.Close();
            file = new System.IO.StreamReader(Log_File_Path);
            int row = 1, column = 1;
            while ((line = file.ReadLine()) != null)
            {

                MatchCollection Job_Id_Matches = Regex.Matches(line, Job_Id_Pattern);
                Match match = Job_Id_Matches[0];


                if (Job_Id == match.Groups[1].ToString())
                {
                    string pattern = "(.*) (.*) {\"message\":\"(.*)\",\"level\".*\"timeStamp\":\"(.*)T.*,\"fingerprint\"";
                    MatchCollection Target_data_Matches = Regex.Matches(line, pattern);
                    Match Target_Data_Match = Target_data_Matches[0];
                    string Date = Target_Data_Match.Groups[4].ToString();
                    string Time = Target_Data_Match.Groups[1].ToString();
                    string Log_Level = Target_Data_Match.Groups[2].ToString();
                    string Message = Target_Data_Match.Groups[3].ToString();
                    string Result_Line = Date + "   " + Time + "   " + Log_Level + "   " + Message;
                    if (File_Format == "txt")
                        Text_File.WriteLine(Result_Line);

                    if (File_Format == "html")
                    {
                        if (Log_Level == "Error" || Log_Level == "Fatal")
                            Result_Line = "<p><font color=\"red\">" + Result_Line + "</Font></p>";
                        else if (Log_Level == "Warn")
                            Result_Line = "<p><font color=\"orange\">" + Result_Line + "</Font></p>";
                        else
                            Result_Line = "<p>" + Result_Line + "</p>";
                        Html_File.WriteLine(Result_Line);
                    }

                    if (File_Format == "xlsx")
                    {
                        ws.Cells[row, column] = Date;
                        ws.Cells[row, column + 1] = Time;
                        ws.Cells[row, column + 2] = Log_Level;
                        ws.Cells[row, column + 3] = Message;
                        row++;
                    }
                    //Console.WriteLine(Result_Line);

                }
            }

            if (File_Format == "html")
            {
                Html_File.WriteLine("</html>");
                Html_File.Close();
            }

            if (File_Format == "txt")
                Text_File.Close();

            if (File_Format == "xlsx")
            {
                Excel_App.DisplayAlerts = false;
                wb.SaveAs(Out_File_Path + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
                wb.Close();
                Excel_App.Quit();
            }
            // Console.Read();
            // System.Data.DataTable dt = new System.Data.DataTable();



        }
    }
}

