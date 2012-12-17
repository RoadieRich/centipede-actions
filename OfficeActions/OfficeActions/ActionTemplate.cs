using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Centipede;

namespace Centipede.OfficeActions
{
    public abstract class WordAction : Action
    {
        protected WordAction(String name, Dictionary<String, Object> variables)
            : base(name, variables)
        { }

        protected Microsoft.Office.Interop.Word.Application WordApp;

    }

    public abstract class ExcelAction : Action
    {
        protected ExcelAction(String name, Dictionary<String, Object> variables)
            : base(name, variables)
        { }

        protected Excel.Workbook WorkBook;

        [ActionArgument]
        String WorksheetVarName = "Worksheet";

        protected Excel.Application ExcelApp;

        protected override void InitAction()
        {
            try
            {
                WorkBook = Variables[ParseStringForVariable(WorksheetVarName)] as Microsoft.Office.Interop.Excel.Workbook;
            }
            catch (KeyNotFoundException)
            {
                WorkBook = null;
                Variables[ParseStringForVariable(WorksheetVarName)] = null;
            }
            ExcelApp = ExcelApp.Application;
            
        }
        protected override void CleanupAction()
        {
            Variables[ParseStringForVariable(WorksheetVarName)] = WorkBook;
        }
    }

    [ActionCategory("Office", displayName = "Open Excel Document", iconName = "excel")]
    public class OpenExcelDocument : ExcelAction
    {
        public OpenExcelDocument(Dictionary<String, Object> variables)
            : base("Open Excel Document", variables)
        { }

        [ActionArgument]
        public String Filename = "";

        [ActionArgument(displayName = "Variable to store document")]
        public String ExcelDocumentVar = "WordDoc";

        protected override void DoAction()
        {
            WorkBook = ExcelApp.Workbooks.Open(ParseStringForVariable(Filename));
        }
    }

    [ActionCategory("Office", iconName="excel", displayName="Get Cell Value")]
    public class GetCellValue : ExcelAction
    {
        public GetCellValue(Dictionary<String, Object> v)
            : base("Get Value From Cell", v)
        { }

        [ActionArgument]
        public Int32 SheetNumber = 0;

        [ActionArgument]
        public String Address = "A1";
        protected override void DoAction()
        {
            Excel.Worksheet sheet = WorkBook.Worksheets.get_Item(SheetNumber) as Excel.Worksheet;

            Variables[ParseStringForVariable(ResultVarName)] = sheet.get_Range(Address).Value2;
        }
        [ActionArgument]
        public String ResultVarName;
    }

    [ActionCategory("Office", iconName="excel", displayName="Set Cell Value")]
    public class SetCellValue : ExcelAction
    {
        public SetCellValue(Dictionary<String, Object> v)
            : base("Get Value From Cell", v)
        { }

        [ActionArgument]
        public Int32 SheetNumber = 0;

        [ActionArgument]
        public String Address = "A1";

        [ActionArgument]
        public String Value = "";

        protected override void DoAction()
        {
            Excel.Worksheet sheet = WorkBook.Worksheets.get_Item(SheetNumber) as Excel.Worksheet;

            sheet.get_Range(Address).Value2 = ParseStringForVariable(Value);

            (sheet as Excel._Worksheet).Calculate();
        }
    }

    [ActionCategory("Office", iconName="excel", displayName="Save Workbook")]
    public class SaveWorkbook : ExcelAction
    {
        public SaveWorkbook(Dictionary<String, Object> v)
            : base("Save Workbook", v)
        { }
        
        [ActionArgument]
        public String Filename = "";

        [ActionArgument]
        public Boolean KeepOpen = false;

        protected override void DoAction()
        {
            WorkBook.SaveAs(Filename);
            if (!KeepOpen)
            {
                WorkBook.Close();
                WorkBook = null;
            }
        }
    }

    [ActionCategory("Office", iconName="excel", displayName="Show Worksheet")]
    public class ShowWorksheet : ExcelAction
    {

        ShowWorksheet(Dictionary<String, Object> v)
            : base("Show workbook", v)
        { }

        protected override void DoAction()
        {
            WorkBook.Application.Visible = true;
            (WorkBook as Excel._Workbook).Activate();
        }
    }
}
