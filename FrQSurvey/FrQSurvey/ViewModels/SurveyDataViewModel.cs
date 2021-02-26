using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FrQSurvey.Services;
using HtmlToOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Input;
using Xamarin.Forms;

namespace FrQSurvey.ViewModels
{
    public class SurveyDataViewModel : BaseViewModel
    {
        private string text;
        private string description;

        public string Text
        {
            get => text;
            set => SetProperty(ref text, value);
        }

        public string Description
        {
            get => description;
            set => SetProperty(ref description, value);
        }
        public SurveyDataViewModel()
        {
            Title = "Survey Data";
            SaveToDocCommand = new Command(() => SaveToDoc());
        }

        public ICommand SaveToDocCommand { get; }

        private void SaveToDoc()
        {
            string html = GetHtml();

            var replacables = new Dictionary<string, string>
            {
                { "initiatorName", text },
                { "initiatorAddress", description }
            };

            foreach (KeyValuePair<string, string> entry in replacables)
            {
                var dd = $"*|{entry.Key}|*";
                html = html.Replace(dd, entry.Value);
            }

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart);
                    converter.ParseHtml(html);

                    mainPart.Document.Save();
                }
                // save to file
                DependencyService.Get<IFileService>().SaveAndView("SurveyDoc.docx", generatedDocument);
            }
        }

        private string GetHtml()
        {
            string docFileName = "Doc.html";
            
            var assembly = typeof(SurveyDataViewModel).GetTypeInfo().Assembly;
            Stream stream = assembly.GetManifestResourceStream($"{assembly.GetName().Name}.ViewModels.{docFileName}");
            string docContent = null;
            using (var reader = new StreamReader(stream))
            {
                docContent = reader.ReadToEnd();
            }

            return docContent;
        }
    }
}