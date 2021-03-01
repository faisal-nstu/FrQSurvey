using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FrQSurvey.Models;
using FrQSurvey.Services;
using HtmlToOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Input;
using Xamarin.Forms;

namespace FrQSurvey.ViewModels
{
    public class SurveyDataViewModel : BaseViewModel
    {
        #region PROPS
        private string id;
        public string Id
        {
            get => id;
            set => SetProperty(ref id, value);
        }





        private string initiatorName;
        public string InitiatorName
        {
            get => initiatorName;
            set => SetProperty(ref initiatorName, value);
        }
        private string initiatorAddress;
        public string InitiatorAddress
        {
            get => initiatorAddress;
            set => SetProperty(ref initiatorAddress, value);
        }
        private string reference;
        public string Reference
        {
            get => reference;
            set => SetProperty(ref reference, value);
        }
        private string dateOfInitiation;
        public string DateOfInitiation
        {
            get => dateOfInitiation;
            set => SetProperty(ref dateOfInitiation, value);
        }
        private string dateOfSurvey;
        public string DateOfSurvey
        {
            get => dateOfSurvey;
            set => SetProperty(ref dateOfSurvey, value);
        }
        private string placeOfSurvey;
        public string PlaceOfSurvey
        {
            get => placeOfSurvey;
            set => SetProperty(ref placeOfSurvey, value);
        }
        private string nameOfBorrower;
        public string NameOfBorrower
        {
            get => nameOfBorrower;
            set => SetProperty(ref nameOfBorrower, value);
        }
        private string borrowerAddress;
        public string BorrowerAddress
        {
            get => borrowerAddress;
            set => SetProperty(ref borrowerAddress, value);
        }
        private string nameOfOwner;
        public string NameOfOwner
        {
            get => nameOfOwner;
            set => SetProperty(ref nameOfOwner, value);
        }
        private string presentAddress;
        public string PresentAddress
        {
            get => presentAddress;
            set => SetProperty(ref presentAddress, value);
        }
        private string permanentAddress;
        public string PermanentAddress
        {
            get => permanentAddress;
            set => SetProperty(ref permanentAddress, value);
        }
        private string relationshipWithBorrower;
        public string RelationshipWithBorrower
        {
            get => relationshipWithBorrower;
            set => SetProperty(ref relationshipWithBorrower, value);
        }
        














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
            set
            {
                SetProperty(ref totalAreaOfLand, value);
                SetPresentValueOfLand();
            }
        }
        private string presentRateOfLand;
        public string PresentRateOfLand
        {
            get => presentRateOfLand;
            set
            {
                SetProperty(ref presentRateOfLand, value);
                SetPresentValueOfLand();
            }
        }

        private void SetPresentValueOfLand()
        {
            if (totalAreaOfLand != null && presentRateOfLand != null)
            {
                decimal area;
                if (!decimal.TryParse(totalAreaOfLand, out area))
                {
                    area = 0;
                }
                decimal rate;
                if (!decimal.TryParse(presentRateOfLand, out rate))
                {
                    rate = 0;
                }

                PresentValueOfLand = (area * rate).ToString();
            }
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
        #endregion

        private ObservableCollection<Valuation> valuations;
        public ObservableCollection<Valuation> Valuations
        {
            get => valuations;
            set => SetProperty(ref valuations, value);
        }

        public SurveyDataViewModel()
        {
            Valuations = new ObservableCollection<Valuation>();
            SaveToDocCommand = new Command(async () =>
            {
                await Task.Run(async () =>
                {
                    IsBusy = true;
                    SaveToDoc();
                    IsBusy = false;
                });
            });
            AddValuationCommand = new Command(() => AddValuation());
        }

        private void AddValuation()
        {
            Valuations.Add(new Valuation());
        }

        public ICommand SaveToDocCommand { get; }
        public ICommand AddValuationCommand { get; }

        private void SaveToDoc()
        {
            if (string.IsNullOrEmpty(Id))
            {
                ShowError("Id cannot be empty");
                return;
            }

            var currentDate = DateTime.Now;
            var year = currentDate.Year.ToString();
            string html = GetTemplate("Document.html");

            var replacables = new Dictionary<string, string>
            {
                { "Id", Id },

                { "InitiatorName", InitiatorName },
                { "InitiatorAddress", InitiatorAddress },
                { "Reference", Reference },
                { "DateOfInitiation", DateOfInitiation },
                { "DateOfSurvey", DateOfSurvey },
                { "PlaceOfSurvey", PlaceOfSurvey },
                { "NameOfBorrower", NameOfBorrower },
                { "BorrowerAddress", BorrowerAddress },
                { "NameOfOwner", NameOfOwner },
                { "PresentAddress", PresentAddress },
                { "PermanentAddress", PermanentAddress },
                { "RelationshipWithBorrower", RelationshipWithBorrower },

                { "Year",  year },
                { "Date", currentDate.ToString("MMMM dd, yyyy") },
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
                // Console.WriteLine("Key: " + replacable.Key + " :::: Value: " + replacable.Value);
            }

            var valuationRows = "";
            double? totalAfterCompletion = 0;
            double? totalAtPresent = 0;
            foreach (var valuation in Valuations)
            {
                try
                {
                    valuation.Calculate();
                }
                catch (Exception ex)
                {
                    ShowError("Calculation failed. Maked sure you have entered valid numbers.");
                }

                totalAfterCompletion += valuation.TotalWhenCompleted;
                totalAtPresent += valuation.AfterDepreciation;

                var valuationRow = GetTemplate("ValuationRow.html");
                valuationRow = valuationRow.Replace($"*|Floor|*", valuation.Floor);
                valuationRow = valuationRow.Replace($"*|Area|*", valuation.Area.ToString());
                valuationRow = valuationRow.Replace($"*|ConstructionYear|*", valuation.ConstructionYear.ToString());
                valuationRow = valuationRow.Replace($"*|Completed|*", valuation.Completed.ToString());
                valuationRow = valuationRow.Replace($"*|CompleteRate|*", valuation.CompleteRate.ToString());
                valuationRow = valuationRow.Replace($"*|TotalWhenCompleted|*", valuation.TotalWhenCompleted.ToString());
                valuationRow = valuationRow.Replace($"*|PresentRate|*", valuation.PresentRate.ToString());
                valuationRow = valuationRow.Replace($"*|TotalPresentValue|*", valuation.TotalPresentValue.ToString());
                valuationRow = valuationRow.Replace($"*|Depreciation|*", valuation.Depreciation.ToString());
                valuationRow = valuationRow.Replace($"*|AfterDepreciation|*", valuation.AfterDepreciation.ToString());
                valuationRows += valuationRow;
            }
            html = html.Replace("*|ValuationRows|*", valuationRows);
            html = html.Replace("*|TotalAfterCompletion|*", totalAfterCompletion.ToString());
            html = html.Replace("*|TotalAtPresent|*", totalAtPresent.ToString());

            var folderName = "FrQSurveyDocs";
            var fileName = $"VR-{id}, {year}.docx";

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
                DependencyService.Get<IFileService>().SaveAndView(fileName, folderName, generatedDocument);
            }

            MessagingCenter.Send(this, "SUCCESS", $"Document \"{fileName}\" saved in \"{folderName}\"");
        }

        private void ShowError(string message)
        {
            MessagingCenter.Send(this, "ERROR", message);
        }

        private string GetTemplate(string templateName)
        {
            string docFileName = templateName;

            var assembly = typeof(SurveyDataViewModel).GetTypeInfo().Assembly;
            Stream stream = assembly.GetManifestResourceStream($"{assembly.GetName().Name}.Templates.{docFileName}");
            string docContent = null;
            using (var reader = new StreamReader(stream))
            {
                docContent = reader.ReadToEnd();
            }

            return docContent;
        }
    }
}