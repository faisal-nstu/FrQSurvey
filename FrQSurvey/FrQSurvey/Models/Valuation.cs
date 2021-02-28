using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace FrQSurvey.Models
{
    public class Valuation : INotifyPropertyChanged
    {
        private string floor;
        public string Floor
        {
            get => floor;
            set => SetProperty(ref floor, value);
        }

        private string area;
        public string Area
        {
            get => area;
            set => SetProperty(ref area, value);
        }

        private string constructionYear;
        public string COnstructionYear
        {
            get => constructionYear;
            set => SetProperty(ref constructionYear, value);
        }

        private string completed;
        public string Completed
        {
            get => completed;
            set => SetProperty(ref completed, value);
        }

        private string completeRate;
        public string CompleteRate
        {
            get => completeRate;
            set => SetProperty(ref completeRate, value);
        }

        private string totalWhenCompleted;
        public string TotalWhenCompleted
        {
            get => totalWhenCompleted;
            set => SetProperty(ref totalWhenCompleted, value);
        }

        private string presentRate;
        public string PresentRate
        {
            get => presentRate;
            set => SetProperty(ref presentRate, value);
        }

        private string totalPresentValue;
        public string TotalPresentValue
        {
            get => totalPresentValue;
            set => SetProperty(ref totalPresentValue, value);
        }

        private string depereciation;
        public string Depreciation
        {
            get => depereciation;
            set => SetProperty(ref depereciation, value);
        }

        private string afterDepreciation;
        public string AfterDepreciation
        {
            get => afterDepreciation;
            set => SetProperty(ref afterDepreciation, value);
        }

        public Valuation(string floor, string area, string constructionYear, string completed, string completeRate, string presentRate)
        {
            Floor = floor;
            Area = area;
            COnstructionYear = constructionYear;
            Completed = completed;
            CompleteRate = completeRate;
            PresentRate = presentRate;
        }

        protected bool SetProperty<T>(ref T backingStore, T value,
            [CallerMemberName] string propertyName = "",
            Action onChanged = null)
        {
            if (EqualityComparer<T>.Default.Equals(backingStore, value))
                return false;

            backingStore = value;
            onChanged?.Invoke();
            OnPropertyChanged(propertyName);
            return true;
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            var changed = PropertyChanged;
            if (changed == null)
                return;

            changed.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
