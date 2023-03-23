namespace InventoryTask.Models
{
    public class ExcelModel
    {
        public string ProductCode { get; set; }
        public int EventType { get; set; }
        public int Quantity { get; set; }
        public double PricePerQty { get; set; }
        public DateTime Date { get; set; }
    }
}
