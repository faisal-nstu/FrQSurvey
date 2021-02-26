using FrQSurvey.Views;
using Xamarin.Forms;

namespace FrQSurvey
{
    public partial class App : Application
    {

        public App()
        {
            InitializeComponent();
            MainPage = new SurveyDataPage();
        }

        protected override void OnStart()
        {
        }

        protected override void OnSleep()
        {
        }

        protected override void OnResume()
        {
        }
    }
}
