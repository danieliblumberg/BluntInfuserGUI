using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security.AccessControl;
using System.Windows.Forms;
//using CodeProjectSerialComms;
using System.Xml;
using System.Threading;

namespace ControlGUI
{
    class FileIOLame
    {
        public static int count = 1;
        public static int countVHR = 1;
        public static string fullFileName { get; set; }
        public static string Ilokot { get; set; }
        public static string AlrmIn { get; set; }
        public static string Imode { get; set; }
        public static string Pmode { get; set; }
        public static string LASenb { get; set; }
        public static string PWMdty { get; set; }
        public static string PWMfrq { get; set; }
        public static string LASset { get; set; }
        public static string TEMPmn { get; set; }
        public static string LASact { get; set; }


        public void fileCreate()
        {
            try
            {
                var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                fullFileName = Path.Combine(desktopFolder, fullFileName + ".csv");
                var fs = new FileStream(fullFileName, FileMode.OpenOrCreate);
                fs.Close();

                try
                {

                    using (StreamWriter sw = new StreamWriter(fullFileName, true))
                    {
                        sw.WriteLine("," + "Date/Time" + "," + "Seconds" + "," + "Ilokot+" + "," + "AlrmIn" + "," + "Imode"
                             + "," + "Pmode+" + "," + "LASenb-" + "," + "PWMdty" + "," + "VCOMPWMfrq"  + "," + "LASset - " + "," + "TEMPmn" + "," + "LASact");
                    }
                    count = 1;
                }

                catch
                {
                    MessageBox.Show("Error Writing To File.");
                }

            }
            catch
            {
                MessageBox.Show("Error Writing To File.");
            }
        }

        public void fileWrite()
        {

            try
            {
                using (StreamWriter sw = new StreamWriter(fullFileName, true))
                {
                    sw.WriteLine(count.ToString() + "," + DateTime.Now.ToString() + "," + DateTime.Now.Second.ToString() + "," + Ilokot + "," + AlrmIn + "," + Imode + "," 
                        + Pmode + "," + LASenb + "," + PWMdty + "," + PWMfrq + "," + LASset + "," + TEMPmn + "," + LASact);
                }
                count++;
            }

            catch
            {
                MessageBox.Show("Error Writing To File.");
            }
        }

        public void fileWriteVHR(string[] readingVHR, string polarity)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(fullFileName, true))
                {
                    sw.WriteLine(DateTime.Now.ToString() + "," + polarity + "VHR Reading");
                    for (int i = 0; i < readingVHR.Length; i++)
                    {
                        sw.WriteLine(countVHR.ToString() + "," + readingVHR[i]);
                        countVHR++;
                        Thread.Sleep(10);
                    }
                }
                countVHR = 1;
            }

            catch
            {
                MessageBox.Show("Error Writing To File. (VHR)");
            }
        }
        public void fileHeaderWrite(string text)
        {

            try
            {

                using (StreamWriter sw = new StreamWriter(fullFileName, true))
                {
                    sw.WriteLine(",,,,,," + text);
                }
            }

            catch
            {
                MessageBox.Show("Error Writing To File.");
            }
        }
    }
}
