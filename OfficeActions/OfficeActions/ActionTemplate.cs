using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Centipede;
using Action = Centipede.Action;


namespace OfficeActions
{
    public abstract class WordAction : Action
    {
        protected WordAction(String name, Dictionary<String, Object> variables)
            : base(name, variables)
        { }

        protected Word.Application WordApp;

    }

    public abstract class ExcelAction : Action
    {
        protected ExcelAction(String name, Dictionary<String, Object> variables)
            : base(name, variables)
        { }

        protected Excel.Workbook WorkBook = null;

        [ActionArgument]
        public String WorksheetVarName = "Worksheet";

        protected static Excel.Application ExcelApp = null;
        private Object _lockObject = new Object();

        protected override void InitAction()
        {
            if (ExcelApp == null)
            {
                lock (_lockObject)
                {
                    if (ExcelApp == null)
                    {
                        ExcelApp = new Excel.Application();
                    }
                    
                }
                
            }

            try
            {
                WorkBook = Variables[ParseStringForVariable(WorksheetVarName)] as Excel.Workbook;
            }
            catch (KeyNotFoundException)
            {
                WorkBook = null;
                Variables[ParseStringForVariable(WorksheetVarName)] = null;
            }
            //ExcelApp = Excel.ApplicationClass
            
            
            
        }
        protected override void CleanupAction()
        {
            Variables[ParseStringForVariable(WorksheetVarName)] = WorkBook;
        }

        public override void Dispose()
        {
            if (WorkBook != null)
            {
                try
                {
                    WorkBook.Close(SaveChanges: false);
                    Variables[ParseStringForVariable(WorksheetVarName)] = null;
                    WorkBook = null;
                }
                catch
                { }
            }
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
        public Int32 SheetNumber = 1;

        [ActionArgument]
        public String Address = "A1";
        protected override void DoAction()
        {
            Excel.Worksheet sheet = WorkBook.Worksheets.get_Item(SheetNumber) as Excel.Worksheet;

            Variables[ParseStringForVariable(ResultVarName)] = sheet.get_Range(Address).Value2;
        }
        [ActionArgument]
        public String ResultVarName = "CellValue";
    }

    [ActionCategory("Office", iconName="excel", displayName="Set Cell Value")]
    public class SetCellValue : ExcelAction
    {
        public SetCellValue(Dictionary<String, Object> v)
            : base("Get Value From Cell", v)
        { }

        [ActionArgument]
        public Int32 SheetNumber = 1;

        [ActionArgument]
        public String Address = "A1";

        [ActionArgument]
        public String Value = "";

        protected override void DoAction()
        {
            Excel.Worksheet sheet = WorkBook.Worksheets.Item[SheetNumber] as Excel.Worksheet;

            sheet.Range[Address].Value2 = ParseStringForVariable(Value);

            // ReSharper disable RedundantCast - Cast is needed to avoid ambiguity between _Worksheet.Calculate and 
            //                                   DocEvents_Event.Calculate
            ((Excel._Worksheet)sheet).Calculate();
            // ReSharper restore RedundantCast
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

        public ShowWorksheet(Dictionary<String, Object> v)
            : base("Show workbook", v)
        { }

        protected override void DoAction()
        {
            WorkBook.Application.Visible = true;
            (WorkBook as Excel._Workbook).Activate();
        }
    }
}
