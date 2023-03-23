namespace InventoryTask.Models
{
	public class InventoryVM
	{
		public DateTime Date { get; set; }
		public string ProductCode { get; set; }
		public int EventType { get; set; }
		public double PurchasePrice { get; set; }
		public double SellingPrice { get; set; }
		public int TotalPurchaseQty { get; set; }
		public double TotalPurchaseAmt { get; set; }
		public int TotalSaleQty { get; set; }
		public double TotalSaleAmt { get; set; }
		public double ProfitLoss { get; set; }
		public int OpeningQty { get; set; }
		public int ClosingQty { get; set; }
	}
}
