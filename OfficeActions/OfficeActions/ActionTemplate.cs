using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

        [ActionArgument(Literal = true)]
        public String WorksheetVarName = "Worksheet";

        protected static volatile Excel.Application ExcelApp;
        protected Excel.Workbook WorkBook;
        private Object _lockObject = new Object();

        protected sealed override void InitAction()
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
                WorkBook = Variables[WorksheetVarName] as Excel.Workbook;
            }
            catch (Exception)
            {
                WorkBook = null;
                //Variables.Add(ParseStringForVariable(WorksheetVarName), null);
            }
            //ExcelApp = Excel.ApplicationClass
            
            
            
        }
        protected sealed override void CleanupAction()
        {
            Variables[WorksheetVarName] = WorkBook;
        }

        public sealed override void Dispose()
        {
            if (WorkBook == null)
            {
                return;
            }
            try
            {
                WorkBook.Close(false);
                Variables[WorksheetVarName] = null;
                WorkBook = null;
            }
            catch
            {
                Console.Out.Write("Error in dispose of {0}", Name);
            }
        }
    }

    [ActionCategory("Office", DisplayName = "Open Excel Document", iconName = "excel")]
    public class OpenExcelDocument : ExcelAction
    {
        public OpenExcelDocument(IDictionary<string, object> variables, ICentipedeCore c)
            : base("Open Excel Document", variables, c)
        { }

        [ActionArgument]
        public String Filename = "";

        [ActionArgument]
        public bool Visible;

       // [ActionArgument(displayName = "Variable to store document")]
       // public String ExcelDocumentVar = "WordDoc";

        protected override void DoAction()
        {
            WorkBook = ExcelApp.Workbooks.Open(ParseStringForVariable(Filename));
            if (Visible)
            {
                ExcelApp.Visible = Visible;
                WorkBook.Activate();
            }
        }
    }


    [ActionCategory("Office", DisplayName = "Create New Excel Document", iconName = "excel")]
    public class NewExcelDocument : ExcelAction
    {
        public NewExcelDocument(IDictionary<string, object> variables, ICentipedeCore c)
            : base("Create New Excel Document", variables, c)
        { }
        
        [ActionArgument]
        public bool Visible = true;

       // [ActionArgument(displayName = "Variable to store document")]
       // public String ExcelDocumentVar = "WordDoc";

        protected override void DoAction()
        {
            WorkBook = ExcelApp.Workbooks.Add();
            if (Visible)
            {
                ExcelApp.Visible = Visible;
                WorkBook.Activate();
            }
        }
    }

    [ActionCategory("Office", iconName="excel", DisplayName="Get Cell Value")]
    public class GetCellValue : ExcelAction
    {
        public GetCellValue(IDictionary<string, object> v, ICentipedeCore c)
            : base("Get Value From Cell", v, c)
        { }

        #region Arguments

        [ActionArgument(DisplayName = "Sheet index", Usage = "The first sheet is 1, second is 2, etc.")]
        public String SheetNumber = "1";

        [ActionArgument]
        public String Address = "A1";

        [ActionArgument]
        public String ResultVarName = "CellValue";

        #endregion


        protected override void DoAction()
        {
            if (String.IsNullOrWhiteSpace(ResultVarName))
            {
                throw new ActionException("Result Var Name is empty", this);
            }
            int sheetNo;
            int.TryParse(ParseStringForVariable(SheetNumber), out sheetNo);

            Excel.Worksheet sheet;
            try
            {
                sheet = (Excel.Worksheet)WorkBook.Worksheets.Item[sheetNo];
            }
            catch (COMException e)
            {
                throw new ActionException("Invalid sheet index",e, this);
            }

            Excel.Range range;
            try
            {
                range = sheet.Range[ParseStringForVariable(this.Address)];
            }
            catch (COMException e)
            {
                throw new ActionException("Invalid Address", e, this);
            }
            Variables[ParseStringForVariable(ResultVarName)] = range.Value2;
        }
    }

    [ActionCategory("Office", iconName="excel", DisplayName = "Set Cell Value")]
    public class SetCellValue : ExcelAction
    {
        public SetCellValue(IDictionary<string, object> v, ICentipedeCore c)
            : base("Set Cell Value", v, c)
        { }

        [ActionArgument]
        public String SheetNumber = "1";

        [ActionArgument]
        public String Address = "A1";

        [ActionArgument]
        public String Value = "";

        protected override void DoAction()
        {



            int sheetNo;
            int.TryParse(ParseStringForVariable(SheetNumber), out sheetNo);

            Excel.Worksheet sheet = (Excel.Worksheet)WorkBook.Worksheets[sheetNo];

            sheet.Range[ParseStringForVariable(Address)].Value2 = ParseStringForVariable(Value);

            // ReSharper disable RedundantCast - Cast is needed to avoid ambiguity between _Worksheet.Calculate and 
            //                                   DocEvents_Event.Calculate
            ((Excel._Worksheet)sheet).Calculate();
            // ReSharper restore RedundantCast
        }
    }

    [ActionCategory("Office", iconName="excel", DisplayName = "Save Workbook")]
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

    [ActionCategory("Office", iconName="excel", DisplayName="Show Worksheet")]
    public class ShowWorksheet : ExcelAction
    {

        public ShowWorksheet(IDictionary<string, object> v, ICentipedeCore c)
            : base("Show workbook", v, c)
        { }

        protected override void DoAction()
        {
            WorkBook.Application.Visible = true;

            // ReSharper disable RedundantCast - cast is needed to solve ambiguity between Excel._Workbook.Activate() 
            //                                   and Excel.WorkbookEvents_Event.Activate.
            (WorkBook as Excel._Workbook).Activate();
            // ReSharper restore RedundantCast
        }
    }

    [ActionCategory("Office", iconName="excel", DisplayName="Run Macro")]
    public class RunMacro : ExcelAction
    {

        public RunMacro(IDictionary<string, object> v, ICentipedeCore c)
            : base("Show workbook", v, c)
        { }

        
        [ActionArgument(DisplayName = "Macro Name")]
        public String MacroName = "";

        [ActionArgument(DisplayName = "Macro Arguments", Usage = "Arguments, separated by commas")]
        public string MacroArguments = "";

        protected override void DoAction()
        {
            ExcelApp.GetType().InvokeMember("Run",                                          // it'd be nice if there
                                            System.Reflection.BindingFlags.Default |        // was an easier way of
                                            System.Reflection.BindingFlags.InvokeMethod,    // doing this, but we're
                                            null,                                           // bound by c# being a
                                            ExcelApp,                                       // strongly typed language
                                            this.GetArgsArray());


        }

        private object[] GetArgsArray()
        {
            return this.MacroArguments.Split(',')
                                      .Select(s => s.Trim())
                                      .Cast<Object>()
                                      .ToArray();
        }
    }



}
