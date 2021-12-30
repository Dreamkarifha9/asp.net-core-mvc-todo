using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AspnetCoreTODO.Data;
using AspnetCoreTODO.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AspnetCoreTODO.Controllers
{
    public class TodoController: Controller
    {
    private ApplicationDbContext _context;
    public TodoController(ApplicationDbContext context) {
      _context = context;
    }


    public async Task<IActionResult> Index() {
      var todos = await _context.Todos.OrderByDescending(x => x.Createdate).ToListAsync();
      return View(todos);
    }

    [HttpGet]
    public IActionResult Create() {
      return View();
    }

    [HttpPost]
    public async Task<IActionResult> Create(Todo todos) {
        if(ModelState.IsValid) {
            try
            {
          var todoData = new Todo()
          {
            Name = todos.Name,
            Createdate = System.DateTime.Now
          };
          
          _context.Add(todoData);
          await _context.SaveChangesAsync();
          return RedirectToAction("Index");
        }
          catch (System.Exception ex)
            {
          // TODO
          ModelState.AddModelError(string.Empty, $"Error {ex.Message}");
        }
        }
      ModelState.AddModelError(string.Empty, $"Error invalid model");
      return View(todos);
    }

    public async Task<IActionResult> Delete(int id) {
      var todos = await _context.Todos.FindAsync(id);
      _context.Todos.Remove(todos);
      await _context.SaveChangesAsync();
      return RedirectToAction("Index");
    }

  public async Task<IActionResult> ExportToExcel()
    {
        // Get the user list 
        var todos = await _context.Todos.OrderByDescending(x => x.Createdate).ToListAsync();

        var stream = new MemoryStream();
        using (var xlPackage = new ExcelPackage(stream))
        {
            var worksheet = xlPackage.Workbook.Worksheets.Add("Todos");
            var namedStyle = xlPackage.Workbook.Styles.CreateNamedStyle("HyperLink");
            namedStyle.Style.Font.UnderLine = true;
            namedStyle.Style.Font.Color.SetColor(Color.Blue);
            const int startRow = 5;
            var row = startRow;

            //Create Headers and format them
            worksheet.Cells["A1"].Value = "TodoList";
            using (var r = worksheet.Cells["A1:C1"])
            {
                r.Merge = true;
                r.Style.Font.Color.SetColor(Color.White);
                r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
            }

            worksheet.Cells["A4"].Value = "id";
            worksheet.Cells["B4"].Value = "Name";
            worksheet.Cells["C4"].Value = "Createdate";
            worksheet.Cells["A4:C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells["A4:C4"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            worksheet.Cells["A4:C4"].Style.Font.Bold = true;

            row = 5;
            foreach (var todo in todos)
            {
                    worksheet.Cells[row, 1].Value = todo.id;
                    worksheet.Cells[row, 2].Value = todo.Name;
                    worksheet.Cells[row, 3].Value = todo.Createdate;
                    row++;
            }

            // set some core property values
            xlPackage.Workbook.Properties.Title = "Todo List";
            xlPackage.Workbook.Properties.Author = "Mohamad Lawand";
            xlPackage.Workbook.Properties.Subject = "Todo List";
            // save the new spreadsheet
            xlPackage.Save();
            // Response.Clear();
        }
        stream.Position = 0;
        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "todos.xlsx");
    }
  }
}