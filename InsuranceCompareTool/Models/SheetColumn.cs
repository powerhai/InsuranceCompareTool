using NPOI.SS.UserModel;
namespace InsuranceCompareTool.Models
{
    public class SheetColumn
    {
        public int Index { get; set; } = -1;
        public string Title { get; set; }

        public int Width { get; set; } = 100;
        public CellType CellType { get; set; } = CellType.String;
        public HorizontalAlignment Alignment { get; set; } = HorizontalAlignment.Left;
    }

}
