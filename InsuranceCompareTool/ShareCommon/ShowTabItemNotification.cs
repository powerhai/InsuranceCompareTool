using Prism.Interactivity.InteractionRequest;
namespace InsuranceCompareTool.ShareCommon {
    public class ShowTabItemNotification : INotification
    {
        public string Title { get; set; }
        public object Content { get; set; }
        public int Index { get; set; }
    }
}