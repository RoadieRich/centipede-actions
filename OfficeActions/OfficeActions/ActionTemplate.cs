using System;
using System.Collections.Generic;
using CentipedeInterfaces;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Action = Centipede.Action;


namespace OfficeActions
{
    public abstract class WordAction : Action
    {
        protected WordAction(String name, IDictionary<string, object> variables, ICentipedeCore c)
            : base(name, variables, c)
        { }

        protected Word.Application WordApp;

    }

    public abstract class ExcelAction : Action
    {
        protected ExcelAction(String name, IDictionary<string, object> variables, ICentipedeCore c)
            : base(name, variables, c)
        { }

        protected Excel.Workbook WorkBook;

        [ActionArgument]
        public String WorksheetVarName = "Worksheet";

        protected static volatile Excel.Application ExcelApp;
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
            if (WorkBook == null)
            {
                return;
            }
            try
            {
                WorkBook.Close(false);
                Variables[ParseStringForVariable(WorksheetVarName)] = null;
                WorkBook = null;
            }
            catch
            {
                Console.Out.Write("Error in dispose of {0}", Name);
            }
        }
    }

    [ActionCategory("Office", displayName = "Open Excel Document", iconName = "excel")]
    public class OpenExcelDocument : ExcelAction
    {
        public OpenExcelDocument(IDictionary<string, object> variables, ICentipedeCore c)
            : base("Open Excel Document", variables, c)
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
        public GetCellValue(IDictionary<string, object> v, ICentipedeCore c)
            : base("Get Value From Cell", v, c)
        { }

        [ActionArgument]
        public Int32 SheetNumber = 1;

        [ActionArgument]
        public String Address = "A1";
        protected override void DoAction()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)WorkBook.Worksheets.Item[SheetNumber];

            Variables[ParseStringForVariable(ResultVarName)] = sheet.Range[Address].Value2;
        }
        [ActionArgument]
        public String ResultVarName = "CellValue";
    }

    [ActionCategory("Office", iconName="excel", displayName="Set Cell Value")]
    public class SetCellValue : ExcelAction
    {
        public SetCellValue(IDictionary<string, object> v, ICentipedeCore c)
            : base("Get Value From Cell", v, c)
        { }

        [ActionArgument]
        public Int32 SheetNumber = 1;

        [ActionArgument]
        public String Address = "A1";

        [ActionArgument]
        public String Value = "";

        protected override void DoAction()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)WorkBook.Worksheets.Item[SheetNumber];

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
        public SaveWorkbook(IDictionary<string, object> v, ICentipedeCore c)
            : base("Save Workbook", v, c)
        { }
        
        [ActionArgument]
        public String Filename = "";

        [ActionArgument]
        public Boolean KeepOpen;

        protected override void DoAction()
        {
            WorkBook.SaveAs(Filename);
            if (KeepOpen)
            {
                return;
            }
            WorkBook.Close();
            WorkBook = null;
        }
    }

    [ActionCategory("Office", iconName="excel", displayName="Show Worksheet")]
    public class ShowWorksheet : ExcelAction
    {

        public ShowWorksheet(IDictionary<string, object> v, ICentipedeCore c)
            : base("Show workbook", v, c)
        { }

        protected override void DoAction()
        {
            WorkBook.Application.Visible = true;
            // ReSharper disable RedundantCast - needed to solve ambiguity between Excel._Workbook.Activate() and 
            //                                   Excel.WorkbookEvents_Event.Activate.
            
            (WorkBook as Excel._Workbook).Activate();
            // ReSharper restore RedundantCast
        }
    }
}
