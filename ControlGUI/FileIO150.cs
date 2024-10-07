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
    class FileIO150
    {
        public static int count = 1;
        public static int countVHR = 1;
        public static string fullFileName { get; set; }
        public static string Dsyrpv { get; set; }
        public static string Dlmpbr { get; set; }
        public static string Dsyrch { get; set; }
        public static string Dcbump { get; set; }
        public static string boardid { get; set; }
        public static string LASset { get; set; }
        public static string PWMrms { get; set; }
        public static string dspvol { get; set; }
        public static string dsptim { get; set; }
        public static string Ddspgo { get; set; }
        public static string Ddpstr { get; set; }
        public static string Ddpgud { get; set; }
        public static string Dvslsv { get; set; }
        public static string Dsyrsv { get; set; }
        public static string Dvslspv { get; set; }
        public static string TECact { get; set; }


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
                        sw.WriteLine("," + "Date/Time" + "," + "Seconds" + "," + "Dsyrpv" + "," + "Dlmpbr" + "," + "Dsyrch"
                             + "," + "Dcbump" + "," + "boardid" + "," + "LASset" + "," + "PWMrms" + "," + "dspvol" + "," + "dsptim" + "," + "Ddspgo"
                             + "," + "Ddpstr" + "," + "Ddpgud" + "," + "Dvslsv" + "," + "Dsyrsv" + "," + "Dvslspv"+ "," + "TECact");
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
                    sw.WriteLine(count.ToString() + "," + DateTime.Now.ToString() + "," + DateTime.Now.Second.ToString() + "," + Dsyrpv + "," + Dlmpbr + "," + Dsyrch + "," 
                        + Dcbump + "," + boardid + "," + LASset + "," + PWMrms + "," + dspvol + "," + dsptim + "," + Ddspgo + ","
                        + Ddpstr + "," + Ddpgud + "," + Dvslsv + "," + Dsyrsv + "," + Dvslspv + "," + TECact);
                }
                count++;
            }

            catch
            {
                MessageBox.Show("Error Writing To File.");
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
