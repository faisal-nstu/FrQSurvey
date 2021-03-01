using System.IO;

namespace FrQSurvey.Services
{
    public interface IFileService
    {
        void SaveAndView(string fileName, string folderName, MemoryStream stream);
    }
}
