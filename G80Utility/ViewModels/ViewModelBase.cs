using System.ComponentModel;

namespace G80Utility.ViewModels
{
    //INotifyPropertyChanged 這個介面是用來通知結繫端物件有屬性內容發生變更的，當結繫端收到此通知便會進行畫面更新的動作。
    //每個 ViewModel 基本上都會用到 INotifyPropertyChanged ，所以就乾脆寫一個基本類別，然後再給 各別 ViewModel 繼承
    class ViewModelBase : INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler PropertyChanged;


        protected void OnPropertyChanged(string PropertyName = null)
        {
            if (PropertyName != null)
            {
                PropertyChanged(this,
                    new PropertyChangedEventArgs(PropertyName));
            }
        }

    }
}
