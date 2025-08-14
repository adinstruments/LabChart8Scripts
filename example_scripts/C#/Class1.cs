using System;
using System.Runtime.InteropServices;
using System.Reflection;

namespace DriveChart
{
   /// <summary>
   /// Summary description for Class1.
   /// </summary>
   class Class1
   {
      /// <summary>
      /// The main entry point for the application.
      /// </summary>
      [STAThread]
      static void Main(string[] args)
      {
         //ADIChart.Document doc = null;
         dynamic doc = null;

         //ADIChart.Application app = null;
         dynamic app = null;

         try
         {
            Type ADIChartDocType = Type.GetTypeFromProgID("ADIChart.Document");
            //doc = (ADIChart.Document)Activator.CreateInstance(ADIChartDocType); //Causes "Unable to cast COM object of type ..." exception in Windows 10

            //If Activator.CreateInstance(...) fails, LabChart may not be registered properly with Windows. 
            //To fix this, open an Administrator Command Prompt (Window Key | run | Run as adminstrator).
            //Then type:
            // cd C:\Program Files (x86)\ADInstruments\LabChart8 <Enter>
            // labchart8.exe -regserver <Enter>
            //
            doc = Activator.CreateInstance(ADIChartDocType);
            //doc = (ADIChart.Document)Marshal.BindToMoniker("D:/Test1.adicht"); // test LabChart document
            //Marshal.
            //app = (ADIChart.Application)Marshal.GetActiveObject("ADIChart.Application"); //Causes "Unable to cast COM object of type ..." exception in Windows 10
            app = Marshal.GetActiveObject("ADIChart.Application"); //Get running Chart Application object

            // start sampling
            doc.StartSampling();

            // get current ticks
            int oldTick = Environment.TickCount & Int32.MaxValue;

            int interval = 1000; // assume interval is 1000 ms (i.e. 1 second)

            // start looping forever
            while (true)
            {
               // get current ticks
               int newTick = Environment.TickCount & Int32.MaxValue;

               if ((newTick - oldTick) > interval) // if interval has passed
               {
                  oldTick = Environment.TickCount & Int32.MaxValue;

                  int docNumberOfRecords = doc.NumberOfRecords; // get current number of records
                  int docGetRecordLength = doc.GetRecordLength(docNumberOfRecords); // get current record length

                  // make a selection starting from (docGetRecordLength-1), ending at recordLength
                  doc.SetSelectionRange(docNumberOfRecords, docGetRecordLength - 1, docNumberOfRecords, docGetRecordLength);

                  doc.SelectChannel(0, true); //Select channel 1 (channel INDEX = 0)

                  object selectedValue = doc.GetSelectedValue(0, 1); // get selected value from channel 1 (channel NUMBER = 1)

                  double secsPerTick = doc.GetRecordSecsPerTick(docNumberOfRecords - 1); // get secs per tick for record INDEX (not record number) 

                  Console.WriteLine("ch1 comment value = "); // debugging info in console
                  Console.WriteLine(selectedValue); // current channel value
                  Console.WriteLine("ch2 comment time = "); // debugging info in console
                  Console.WriteLine(secsPerTick * docGetRecordLength); // current channel time

                  string commentStr1;

                  commentStr1 = "ch1 Comment value =";
                  commentStr1 += selectedValue;

                  doc.AddToDataPad();
                  doc.AppendComment(commentStr1, 0); // add new comment in document with current ch1 value at comment (use channel INDEX not number)

                  commentStr1 = "ch2 Comment time =";
                  commentStr1 += (secsPerTick * docGetRecordLength);

                  doc.AddToDataPad();
                  doc.AppendComment(commentStr1, 1); // add new comment in document with current ch2 time at comment (use channel INDEX not number)

                  Console.WriteLine((String)doc.Path);
                  Console.WriteLine((String)doc.Name);

                  string appName = app.Name;
                  Console.WriteLine("App name " + appName);
               }
            }
         }
         catch (Exception e)
         {
            Console.WriteLine("An exception occurred. The ProgID is wrong.");
            Console.WriteLine("Source: {0}", e.Source);
            Console.WriteLine("Message: {0}", e.Message);
         }
         finally
         {
            if (app != null)
               Marshal.ReleaseComObject(app);
            if (doc != null)
               Marshal.ReleaseComObject(doc);
         }
      }
   }
}


