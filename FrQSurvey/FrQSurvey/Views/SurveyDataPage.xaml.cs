using FrQSurvey.ViewModels;
using Xamarin.Forms;

namespace FrQSurvey.Views
{
    public partial class SurveyDataPage : ContentPage
    {
        public SurveyDataPage()
        {
            InitializeComponent();

            MessagingCenter.Subscribe<SurveyDataViewModel, string>(this, "SUCCESS", (sender, arg) =>
            {
                Device.BeginInvokeOnMainThread(() =>
                {
                    DisplayAlert("Success!", arg, "OK");
                });
            });

            MessagingCenter.Subscribe<SurveyDataViewModel, string>(this, "ERROR", (sender, arg) =>
            {
                Device.BeginInvokeOnMainThread(() =>
                {
                    DisplayAlert("Error!", arg, "OK");
                });
            });
        }
    }
}