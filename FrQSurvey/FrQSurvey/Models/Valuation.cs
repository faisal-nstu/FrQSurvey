using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace FrQSurvey.Models
{
    public class Valuation : INotifyPropertyChanged
    {
        const double DEPRECIATION_RATE = 1.25;
        private string floor;
        public string Floor
        {
            get => floor;
            set => SetProperty(ref floor, value);
        }

        private double? area;
        public double? Area
        {
            get => area;
            set => SetProperty(ref area, value);
        }

        private int? constructionYear;
        public int? ConstructionYear
        {
            get => constructionYear;
            set => SetProperty(ref constructionYear, value);
        }

        private double? completed;
        public double? Completed
        {
            get => completed;
            set => SetProperty(ref completed, value);
        }

        private double? completeRate;
        public double? CompleteRate
        {
            get => completeRate;
            set => SetProperty(ref completeRate, value);
        }

        private double? totalWhenCompleted;
        public double? TotalWhenCompleted
        {
            get => totalWhenCompleted;
            set => SetProperty(ref totalWhenCompleted, value);
        }

        private double? presentRate;
        public double? PresentRate
        {
            get => presentRate;
            set => SetProperty(ref presentRate, value);
        }

        private double? totalPresentValue;
        public double? TotalPresentValue
        {
            get => totalPresentValue;
            set => SetProperty(ref totalPresentValue, value);
        }

        private double? depereciation;
        public double? Depreciation
        {
            get => depereciation;
            set => SetProperty(ref depereciation, value);
        }

        private double? afterDepreciation;
        public double? AfterDepreciation
        {
            get => afterDepreciation;
            set => SetProperty(ref afterDepreciation, value);
        }

        public void Calculate()
        {
            TotalWhenCompleted = Area * CompleteRate;
            PresentRate = (Completed * CompleteRate) / 100;
            TotalPresentValue = PresentRate * Area;
            Depreciation = ConstructionYear != 0
                ? (DateTime.Now.Year - ConstructionYear) * DEPRECIATION_RATE
                : 0;
            AfterDepreciation = (TotalPresentValue * (100 - Depreciation)) / 100;
        }

        #region INotifyPropertyChanged
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
