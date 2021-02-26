using Android.Content.Res;
using FrQSurvey.Droid;
using FrQSurvey.Services;
using Java.IO;
using System;
using System.IO;
using Xamarin.Forms;

[assembly: Dependency(typeof(FileService))]
namespace FrQSurvey.Droid
{
    public class FileService : IFileService
    {

        public void SaveAndView(string fileName, MemoryStream stream)
        {
            try
            {
                string root = null;
                if (Android.OS.Environment.IsExternalStorageEmulated)
                {
                    root = Android.OS.Environment.ExternalStorageDirectory.ToString();
                }
                else
                    root = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                Java.IO.File myDir = new Java.IO.File(root + "/AAAFrQSurvey");
                myDir.Mkdir();

                Java.IO.File file = new Java.IO.File(myDir, fileName);

                //Remove if the file exists
                if (file.Exists()) file.Delete();

                //Write the stream into the file
                FileOutputStream outs = new FileOutputStream(file);
                outs.Write(stream.ToArray());
                outs.Flush();
                outs.Close();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("ERROR:::::::::::::::::::::: " + ex.Message);
            }
        }
    }
}