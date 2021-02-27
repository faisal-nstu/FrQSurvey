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
        private string valuationOfProperty;
        public string ValuationOfProperty
        {
            get => valuationOfProperty;
            set => SetProperty(ref valuationOfProperty, value);
        }
        private string presentUsageOfLand;
        public string PresentUsageOfLand
        {
            get => presentUsageOfLand;
            set => SetProperty(ref presentUsageOfLand, value);
        }
        private string approachRoad;
        public string ApproachRoad
        {
            get => approachRoad;
            set => SetProperty(ref approachRoad, value);
        }
        private string natureOfRoad;
        public string NatureOfRoad
        {
            get => natureOfRoad;
            set => SetProperty(ref natureOfRoad, value);
        }
        private string buildingFace;
        public string BuildingFace
        {
            get => buildingFace;
            set => SetProperty(ref buildingFace, value);
        }
        private string boundaryWall;
        public string BoundaryWall
        {
            get => boundaryWall;
            set => SetProperty(ref boundaryWall, value);
        }
        private string ownershipOfRoad;
        public string OwnershipOfRoad
        {
            get => ownershipOfRoad;
            set => SetProperty(ref ownershipOfRoad, value);
        }
        private string deedNoAndDate;
        public string DeedNoAndDate
        {
            get => deedNoAndDate;
            set => SetProperty(ref deedNoAndDate, value);
        }
        private string jLNo;
        public string JLNo
        {
            get => jLNo;
            set => SetProperty(ref jLNo, value);
        }
        private string khatianNo;
        public string KhatianNo
        {
            get => khatianNo;
            set => SetProperty(ref khatianNo, value);
        }
        private string dagNo;
        public string DagNo
        {
            get => dagNo;
            set => SetProperty(ref dagNo, value);
        }
        private string mutationKhatianNo;
        public string MutationKhatianNo
        {
            get => mutationKhatianNo;
            set => SetProperty(ref mutationKhatianNo, value);
        }
        private string mouza;
        public string Mouza
        {
            get => mouza;
            set => SetProperty(ref mouza, value);
        }
        private string pS;
        public string PS
        {
            get => pS;
            set => SetProperty(ref pS, value);
        }
        private string sRO;
        public string SRO
        {
            get => sRO;
            set => SetProperty(ref sRO, value);
        }
        private string district;
        public string District
        {
            get => district;
            set => SetProperty(ref district, value);
        }
        private string areaOfLandAsPerDeed;
        public string AreaOfLandAsPerDeed
        {
            get => areaOfLandAsPerDeed;
            set => SetProperty(ref areaOfLandAsPerDeed, value);
        }
        private string locationOfTheLand;
        public string LocationOfTheLand
        {
            get => locationOfTheLand;
            set => SetProperty(ref locationOfTheLand, value);
        }
        private string localAuthority;
        public string LocalAuthority
        {
            get => localAuthority;
            set => SetProperty(ref localAuthority, value);
        }
        private string importance;
        public string Importance
        {
            get => importance;
            set => SetProperty(ref importance, value);
        }
        private string onTheNorth;
        public string OnTheNorth
        {
            get => onTheNorth;
            set => SetProperty(ref onTheNorth, value);
        }
        private string onTheSouth;
        public string OnTheSouth
        {
            get => onTheSouth;
            set => SetProperty(ref onTheSouth, value);
        }
        private string onTheEast;
        public string OnTheEast
        {
            get => onTheEast;
            set => SetProperty(ref onTheEast, value);
        }
        private string onTheWest;
        public string OnTheWest
        {
            get => onTheWest;
            set => SetProperty(ref onTheWest, value);
        }
        private string nearestImportantPlace1Name;
        public string NearestImportantPlace1Name
        {
            get => nearestImportantPlace1Name;
            set => SetProperty(ref nearestImportantPlace1Name, value);
        }
        private string nearestImportantPlace1Value;
        public string NearestImportantPlace1Value
        {
            get => nearestImportantPlace1Value;
            set => SetProperty(ref nearestImportantPlace1Value, value);
        }
        private string nearestImportantPlace2Name;
        public string NearestImportantPlace2Name
        {
            get => nearestImportantPlace2Name;
            set => SetProperty(ref nearestImportantPlace2Name, value);
        }
        private string nearestImportantPlace2Value;
        public string NearestImportantPlace2Value
        {
            get => nearestImportantPlace2Value;
            set => SetProperty(ref nearestImportantPlace2Value, value);
        }
        private string nearestImportantPlace3Name;
        public string NearestImportantPlace3Name
        {
            get => nearestImportantPlace3Name;
            set => SetProperty(ref nearestImportantPlace3Name, value);
        }
        private string nearestImportantPlace3Value;
        public string NearestImportantPlace3Value
        {
            get => nearestImportantPlace3Value;
            set => SetProperty(ref nearestImportantPlace3Value, value);
        }
        private string waterSupply;
        public string WaterSupply
        {
            get => waterSupply;
            set => SetProperty(ref waterSupply, value);
        }
        private string gasSupply;
        public string GasSupply
        {
            get => gasSupply;
            set => SetProperty(ref gasSupply, value);
        }
        private string electricitySupply;
        public string ElectricitySupply
        {
            get => electricitySupply;
            set => SetProperty(ref electricitySupply, value);
        }
        private string communicationRoadTo;
        public string CommunicationRoadTo
        {
            get => communicationRoadTo;
            set => SetProperty(ref communicationRoadTo, value);
        }
        private string communicationRoadToDistance;
        public string CommunicationRoadToDistance
        {
            get => communicationRoadToDistance;
            set => SetProperty(ref communicationRoadToDistance, value);
        }
        private string communicationRoadFrom;
        public string CommunicationRoadFrom
        {
            get => communicationRoadFrom;
            set => SetProperty(ref communicationRoadFrom, value);
        }
        private string communicationRoadFromDistance;
        public string CommunicationRoadFromDistance
        {
            get => communicationRoadFromDistance;
            set => SetProperty(ref communicationRoadFromDistance, value);
        }
        private string totalAreaOfLand;
        public string TotalAreaOfLand
        {
            get => totalAreaOfLand;
            set => SetProperty(ref totalAreaOfLand, value);
        }
        private string presentRateOfLand;
        public string PresentRateOfLand
        {
            get => presentRateOfLand;
            set => SetProperty(ref presentRateOfLand, value);
        }
        private string presentValueOfLand;
        public string PresentValueOfLand
        {
            get => presentValueOfLand;
            set => SetProperty(ref presentValueOfLand, value);
        }
        private string description;
        public string Description
        {
            get => description;
            set => SetProperty(ref description, value);
        }
        private string floorIdentification;
        public string FloorIdentification
        {
            get => floorIdentification;
            set => SetProperty(ref floorIdentification, value);
        }
        private string plan;
        public string Plan
        {
            get => plan;
            set => SetProperty(ref plan, value);
        }
        private string physical;
        public string Physical
        {
            get => physical;
            set => SetProperty(ref physical, value);
        }
        private string floor;
        public string Floor
        {
            get => floor;
            set => SetProperty(ref floor, value);
        }
        private string roof;
        public string Roof
        {
            get => roof;
            set => SetProperty(ref roof, value);
        }
        private string door;
        public string Door
        {
            get => door;
            set => SetProperty(ref door, value);
        }
        private string window;
        public string Window
        {
            get => window;
            set => SetProperty(ref window, value);
        }

        private string text;
        public string Text
        {
            get => text;
            set => SetProperty(ref text, value);
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
                { "ValuationOfProperty", ValuationOfProperty },
                { "PresentUsageOfLand", PresentUsageOfLand },
                { "ApproachRoad", ApproachRoad },
                { "NatureOfRoad", NatureOfRoad },
                { "BuildingFace", BuildingFace },
                { "BoundaryWall", BoundaryWall },
                { "OwnershipOfRoad", OwnershipOfRoad },
                { "DeedNoAndDate", DeedNoAndDate },
                { "JLNo", JLNo },
                { "KhatianNo", KhatianNo },
                { "DagNo", DagNo },
                { "MutationKhatianNo", MutationKhatianNo },
                { "Mouza", Mouza },
                { "PS", PS },
                { "SRO", SRO },
                { "District", District },
                { "AreaOfLandAsPerDeed", AreaOfLandAsPerDeed },
                { "LocationOfTheLand", LocationOfTheLand },
                { "LocalAuthority", LocalAuthority },
                { "Importance", Importance },
                { "OnTheNorth", OnTheNorth },
                { "OnTheSouth", OnTheSouth },
                { "OnTheEast", OnTheEast },
                { "OnTheWest", OnTheWest },
                { "NearestImportantPlace1Name", NearestImportantPlace1Name },
                { "NearestImportantPlace1Value", NearestImportantPlace1Value },
                { "NearestImportantPlace2Name", NearestImportantPlace2Name },
                { "NearestImportantPlace2Value", NearestImportantPlace2Value },
                { "NearestImportantPlace3Name", NearestImportantPlace3Name },
                { "NearestImportantPlace3Value", NearestImportantPlace3Value },
                { "WaterSupply", WaterSupply },
                { "GasSupply", GasSupply },
                { "ElectricitySupply", ElectricitySupply },
                { "CommunicationRoadTo", CommunicationRoadTo },
                { "CommunicationRoadToDistance", CommunicationRoadToDistance },
                { "CommunicationRoadFrom", CommunicationRoadFrom },
                { "CommunicationRoadFromDistance", CommunicationRoadFromDistance },
                { "TotalAreaOfLand", TotalAreaOfLand },
                { "PresentRateOfLand", PresentRateOfLand },
                { "PresentValueOfLand", PresentValueOfLand },
                { "Description", Description },
                { "FloorIdentification", FloorIdentification },
                { "Plan", Plan },
                { "Physical", Physical },
                { "Floor", Floor },
                { "Roof", Roof },
                { "Door", Door },
                { "Window", Window }
            };

            foreach (KeyValuePair<string, string> replacable in replacables)
            {
                html = html.Replace($"*|{replacable.Key}|*", replacable.Value);
                System.Console.WriteLine("Key: " + replacable.Key + " :::: Value: " + replacable.Value);
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
            string docFileName = "template.html";

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