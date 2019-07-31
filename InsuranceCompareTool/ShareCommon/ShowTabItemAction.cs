using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;
using Prism.Interactivity.InteractionRequest;
namespace InsuranceCompareTool.ShareCommon
{
    public class ShowTabItemAction : TriggerAction<TabControl>
    { 
        protected override void Invoke(object parameter)
        {
            var p = parameter as InteractionRequestedEventArgs;
            var c = p.Context as ShowTabItemNotification;
            this.AssociatedObject.SelectedIndex = c.Index;
        }
    }
}
