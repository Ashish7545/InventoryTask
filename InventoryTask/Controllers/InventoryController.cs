using ExcelDataReader;
using InventoryTask.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.IO;

namespace InventoryTask.Controllers
{
    public class InventoryController : Controller
    {      
        public IActionResult Index()
        {
          
            var filePath = "C:\\Users\\Ashish.Kumar\\Downloads\\inevtory task.xlsx";
            var list = new List<ExcelModel>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                //To read sheet from Excel file such as Sheet1 as 0, Sheet2 as 1 ...so on.
                var worksheet = package.Workbook.Worksheets[0];

                //Total Rows & columns in the sheet
                var rowcount = worksheet.Dimension.Rows;
                var columncount = worksheet.Dimension.Columns;

                //Reading & adding the data from excel to list rows wise
                for (int row = 2; row <= rowcount; row++)
                {
                    list.Add(new ExcelModel
                    {
                        ProductCode = worksheet.Cells[row, 1].Value.ToString(),
                        EventType = int.Parse(worksheet.Cells[row, 2].Value.ToString()),
                        Quantity = int.Parse(worksheet.Cells[row, 3].Value.ToString()),
                        PricePerQty = double.Parse(worksheet.Cells[row, 4].Value.ToString()),
                        Date = DateTime.Parse(worksheet.Cells[row, 5].Value.ToString())
                    });
                }
                
                List<InventoryModel> inventoryList = new List<InventoryModel>();

                //Month Wise Sorting all data and adding into new inventoryList
                var monthWise = list.GroupBy(d => d.Date.Month).ToList();
				InventoryModel inv = new InventoryModel();
				foreach (var month in monthWise)
                {
                    var productWise = month.GroupBy(p => p.ProductCode).ToList();
                    foreach (var product in productWise)
                    {
                        inv.ClosingQty = 0;
                        inv.TotalPurchaseAmt = 0;
                        inv.TotalPurchaseQty = 0;
                        inv.TotalSaleAmt= 0;
                        inv.TotalSaleQty = 0;

                        foreach (var item in product)
                        {                     
                            if (item.EventType == 1)
                            {
                                inv.TotalPurchaseQty += item.Quantity;
                                inv.TotalPurchaseAmt += item.Quantity * item.PricePerQty;

                                //PurchasePrice PerQty
                                inv.PurchasePrice = inv.TotalPurchaseAmt / inv.TotalPurchaseQty;
                            }
                            else
                            {
                                inv.TotalSaleQty += item.Quantity;
                                inv.TotalSaleAmt += item.Quantity * item.PricePerQty;

                                //SellingPrice PerQty
                                inv.SellingPrice = inv.TotalSaleAmt / inv.TotalSaleQty;
                            }
                            inv.Date = item.Date.Date;
							inv.ProductCode = item.ProductCode;
                        }

                        //LastOrDefault return last element of a sequence(Month).
                        //And Used to add the previous month remaining product to next month 
						var co = inventoryList.LastOrDefault(a => a.ProductCode == inv.ProductCode);
						if(co != null)
                        {
                            if(inv.TotalPurchaseQty == 0 && inv.TotalPurchaseAmt == 0)
                            {
                                inv.PurchasePrice = co.ClosingAmt;
                            }
                            else if (inv.TotalPurchaseQty != 0 && inv.TotalSaleQty != 0)
                            {
                                inv.PurchasePrice = (inv.TotalPurchaseAmt + (co.ClosingAmt * co.ClosingQty)) / (inv.TotalPurchaseQty + co.ClosingQty);
                            }
                            else
                            {
                                inv.PurchasePrice = (inv.PurchasePrice + co.ClosingAmt) / 2;
                            }
                            inventoryList.Add(new InventoryModel
							{
								Date = inv.Date,
								ProductCode = inv.ProductCode,
								TotalPurchaseAmt = inv.TotalPurchaseAmt,
								TotalPurchaseQty = inv.TotalPurchaseQty,
								TotalSaleAmt = inv.TotalSaleAmt,
								TotalSaleQty = inv.TotalSaleQty,
								ProfitLoss = (inv.SellingPrice - inv.PurchasePrice) * inv.TotalSaleQty,
								ClosingQty = inv.TotalPurchaseQty + co.ClosingQty - inv.TotalSaleQty,
                                OpeningQty = co.ClosingQty,
                                ClosingAmt = (inv.PurchasePrice + co.ClosingAmt)/2,
                                OpeningAmt = co.ClosingAmt
                            });
						}
                        else
                        {
							inventoryList.Add(new InventoryModel
							{
								Date = inv.Date,
								ProductCode = inv.ProductCode,
								TotalPurchaseAmt = inv.TotalPurchaseAmt,
								TotalPurchaseQty = inv.TotalPurchaseQty,
								TotalSaleAmt = inv.TotalSaleAmt,
								TotalSaleQty = inv.TotalSaleQty,
								ProfitLoss = (inv.SellingPrice - inv.PurchasePrice) * inv.TotalSaleQty,
                                ClosingQty = inv.TotalPurchaseQty - inv.TotalSaleQty,
                                OpeningQty = 0,
                                ClosingAmt = inv.TotalPurchaseAmt / inv.TotalPurchaseQty,
                                OpeningAmt = 0,
                            });
						}
					}
                }
                return View(inventoryList);
            }
        }
    }
}
