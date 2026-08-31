using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using System.Threading;
using System.Web.Script.Serialization;
using ExcelDna.Integration;

namespace XmuxAddIn
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class XmuxBridge
    {
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly Dictionary<int, object> handles = new Dictionary<int, object> { { 0, WorkbookHandle.Instance } };
        private readonly object gate = new object();
        private int nextHostHandle = -1;

        private readonly dynamic excelWindow;
        private readonly int excelWindowHandle;

        public XmuxBridge() : this((object)((dynamic)ExcelDnaUtil.Application).ActiveWindow) { }

        public XmuxBridge(object window)
        {
            if (window == null) throw new ArgumentNullException("window");
            excelWindow = window;
            excelWindowHandle = Convert.ToInt32(excelWindow.Hwnd, CultureInfo.InvariantCulture);
        }

        public string handshake() { return Handshake(); }
        public string execute(string opsJson) { return Execute(opsJson); }
        public string readExternalWorkbook(string requestJson) { return ReadExternalWorkbook(requestJson); }
        public string readNativeEditorState() { return ReadNativeEditorState(); }
        public void close()
        {
            lock (gate) ResetBatchHandles();
        }

        private string Handshake()
        {
            try
            {
                var workbookUrl = RunInExcel(delegate
                {
                    dynamic workbook = CurrentWorkbook();
                    return workbook == null ? string.Empty : Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                });
                return json.Serialize(new Dictionary<string, object>
                {
                    { "workbookUrl", workbookUrl },
                    { "capabilities", new object[]
                        {
                            new Dictionary<string, string>
                            {
                                { "name", "ExcelApiOnline" },
                                { "version", "1.1" }
                            }
                        }
                    }
                });
            }
            catch (Exception exception) { return Failure(new Dictionary<string, object>(), "excel", exception.Message); }
        }

        private string Execute(string opsJson)
        {
            lock (gate)
            {
                var values = new Dictionary<string, object>();
                try
                {
                    var ops = json.Deserialize<List<BridgeOperation>>(opsJson);
                    if (ops == null) throw new DispatchException("protocol", "The operation batch is null.");
                    RunInExcel(delegate { foreach (var operation in ops) Dispatch(operation, values); return true; });
                    return json.Serialize(new Dictionary<string, object> { { "values", values } });
                }
                catch (DispatchException exception) { return Failure(values, exception.Code, exception.Message); }
                catch (COMException exception) { return Failure(values, IsEditMode(exception) ? "cellEditMode" : "excel", exception.Message); }
                catch (Exception exception) { return Failure(values, IsEditMode(exception) ? "cellEditMode" : "excel", exception.Message); }
            }
        }

        /// <summary>
        /// Opens an external workbook in its own invisible Excel instance.  It deliberately
        /// never borrows the in-process application's window: previewing a file must not
        /// change the user's active workbook, calculation state, links, or macro security.
        /// </summary>
        private string ReadExternalWorkbook(string requestJson)
        {
            dynamic application = null;
            dynamic workbooks = null;
            dynamic workbook = null;
            dynamic worksheet = null;
            dynamic worksheets = null;
            dynamic cells = null;
            dynamic firstCell = null;
            dynamic range = null;
            object automationSecurity = null;
            Exception primaryFailure = null;
            Exception cleanupFailure = null;
            object values = null;
            try
            {
                var request = json.Deserialize<ExternalWorkbookRequest>(requestJson);
                ValidateExternalRequest(request);
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) throw new InvalidOperationException("Excel COM automation is unavailable.");
                application = Activator.CreateInstance(excelType);
                application.Visible = false;
                application.DisplayAlerts = false;
                application.AskToUpdateLinks = false;
                application.EnableEvents = false;
                application.Calculation = -4135;
                // msoAutomationSecurityForceDisable.  Keep the prior value only for a
                // successful restore before Quit; this instance is otherwise disposable.
                automationSecurity = application.AutomationSecurity;
                application.AutomationSecurity = 3;
                workbooks = application.Workbooks;
                var unavailablePassword = Guid.NewGuid().ToString("N");
                workbook = workbooks.Open(
                    request.Path,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Password: unavailablePassword,
                    WriteResPassword: unavailablePassword,
                    IgnoreReadOnlyRecommended: true,
                    Notify: false,
                    AddToMru: false);
                worksheets = workbook.Worksheets;
                worksheet = worksheets[request.Sheet];
                cells = worksheet.Cells;
                firstCell = cells[request.Area.Top, request.Area.Left];
                range = firstCell.Resize[request.Area.Height, request.Area.Width];
                values = DisplayTextMatrix(range, request.Area.Height, request.Area.Width);
            }
            catch (Exception exception)
            {
                primaryFailure = exception;
            }
            finally
            {
                try { if (workbook != null) workbook.Close(false); }
                catch (Exception exception) { cleanupFailure = CombineCleanupFailures(cleanupFailure, exception); }
                try { if (application != null && automationSecurity != null) application.AutomationSecurity = automationSecurity; }
                catch (Exception exception) { cleanupFailure = CombineCleanupFailures(cleanupFailure, exception); }
                try { if (application != null) application.Quit(); }
                catch (Exception exception) { cleanupFailure = CombineCleanupFailures(cleanupFailure, exception); }
                cleanupFailure = CaptureRelease(range, cleanupFailure);
                cleanupFailure = CaptureRelease(firstCell, cleanupFailure);
                cleanupFailure = CaptureRelease(cells, cleanupFailure);
                cleanupFailure = CaptureRelease(worksheet, cleanupFailure);
                cleanupFailure = CaptureRelease(worksheets, cleanupFailure);
                cleanupFailure = CaptureRelease(workbook, cleanupFailure);
                cleanupFailure = CaptureRelease(workbooks, cleanupFailure);
                cleanupFailure = CaptureRelease(application, cleanupFailure);
            }
            if (primaryFailure != null)
            {
                var message = primaryFailure.Message;
                if (cleanupFailure != null)
                    message += " External workbook cleanup also failed: " + CleanupFailureMessage(cleanupFailure);
                return json.Serialize(new Dictionary<string, object> { { "error", message } });
            }
            if (cleanupFailure != null)
                return json.Serialize(new Dictionary<string, object> { { "error", "External workbook cleanup failed: " + CleanupFailureMessage(cleanupFailure) } });
            return json.Serialize(new Dictionary<string, object> { { "values", values } });
        }

        /// <summary>
        /// Reads Excel's focused formula editor through its process-local native editor window.
        /// The same observer drives the Tab hook, so the highlight is Excel's actual selection.
        /// </summary>
        private string ReadNativeEditorState()
        {
            return NativeEditorKeyboardHook.ReadState(excelWindowHandle);
        }

        private static T RunInExcel<T>(Func<T> work)
        {
            if (ExcelDnaUtil.MainManagedThreadId != 0 &&
                Thread.CurrentThread.ManagedThreadId == ExcelDnaUtil.MainManagedThreadId)
                return work();
            T result = default(T);
            Exception failure = null;
            var completed = new ManualResetEvent(false);
            var state = 0; // queued, running, cancelled, completed
            ExcelAsyncUtil.QueueAsMacro(delegate
            {
                if (Interlocked.CompareExchange(ref state, 1, 0) != 0)
                {
                    completed.Close();
                    return;
                }
                try { result = work(); }
                catch (Exception exception) { failure = exception; }
                finally
                {
                    Interlocked.Exchange(ref state, 3);
                    completed.Set();
                }
            });
            if (!completed.WaitOne(TimeSpan.FromSeconds(2)))
            {
                if (Interlocked.CompareExchange(ref state, 2, 0) == 0)
                    throw new TimeoutException("Excel is in edit mode or remained busy for 2 seconds.");
                // A running mutation cannot be cancelled. Wait for its real outcome rather than
                // returning a failure followed by a delayed write.
                completed.WaitOne();
            }
            completed.Close();
            if (failure != null) throw failure;
            return result;
        }

        private void Dispatch(BridgeOperation operation, Dictionary<string, object> values)
        {
            if (operation == null) throw new DispatchException("protocol", "The operation is null.");
            object target;
            if (!handles.TryGetValue(operation.On, out target)) throw new DispatchException("protocol", "Unknown handle " + operation.On + ".");
            if (operation.Op == "call")
            {
                if (operation.Id <= 0) throw new DispatchException("protocol", "Call result handles must be positive.");
                handles[operation.Id] = Call(target, operation.Member, operation.Args ?? new object[0]);
                return;
            }
            if (operation.Op == "load")
            {
                var next = Load(target, operation.Properties ?? new string[0]);
                object old;
                if (values.TryGetValue(operation.On.ToString(CultureInfo.InvariantCulture), out old) && old is Dictionary<string, object>)
                    foreach (var entry in next) ((Dictionary<string, object>)old)[entry.Key] = entry.Value;
                else values[operation.On.ToString(CultureInfo.InvariantCulture)] = next;
                return;
            }
            throw new DispatchException("protocol", "Unknown operation " + operation.Op + ".");
        }

        private object Call(object target, string member, object[] args)
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic workbook = CurrentWorkbook();
            if (member == "set") { Set(target, StringArg(args, 0), Arg(args, 1)); return target; }
            if (target is WorkbookHandle)
            {
                if (member == "worksheets") return new WorksheetsHandle(workbook.Worksheets);
                if (member == "getSelectedRange") return new RangeHandle(excelWindow.RangeSelection);
                if (member == "getSelectedRanges") return new SelectedAreasHandle(excelWindow.RangeSelection);
                if (member == "func") return new FunctionHandle(StringArg(args, 0), RangeArgument(Arg(args, 1)));
                if (member == "getNameRange") return NameRange(workbook, StringArg(args, 0));
                if (member == "getTable") return TableByName(workbook, StringArg(args, 0));
                if (member == "names.add") { workbook.Names.Add(Name: StringArg(args, 0), RefersTo: "=" + QualifiedAddress(RangeArgument(Arg(args, 1)))); return target; }
                if (member == "linkedWorkbooks.refreshAll") { RefreshLinks(workbook); return target; }
                if (member == "application.calculate") { Calculate(app, StringArg(args, 0)); return target; }
            }
            if (target is WorksheetsHandle)
            {
                var worksheets = ((WorksheetsHandle)target).Worksheets;
                if (member == "getItem") return new WorksheetHandle(worksheets.Item[StringArg(args, 0)]);
                if (member == "getItemOrNullObject") return WorksheetByName(worksheets, StringArg(args, 0));
                if (member == "getActiveWorksheet") return new WorksheetHandle(excelWindow.ActiveSheet);
                if (member == "add") { dynamic added = worksheets.Add(After: worksheets.Item[worksheets.Count], Type: Type.Missing); added.Name = StringArg(args, 0); return new WorksheetHandle(added); }
                if (member == "onSelectionChanged.add" || member == "onSingleClicked.add") return target;
            }
            if (target is WorksheetHandle)
            {
                dynamic sheet = ((WorksheetHandle)target).Worksheet;
                if (sheet == null) return NullForWorksheetCall(member);
                if (member == "getRange") return new RangeHandle(sheet.Range[StringArg(args, 0)]);
                if (member == "getCell") return new RangeHandle(sheet.Cells[IntArg(args, 0) + 1, IntArg(args, 1) + 1]);
                if (member == "getRangeByIndexes") return new RangeHandle(sheet.Cells[IntArg(args, 0) + 1, IntArg(args, 1) + 1].Resize[IntArg(args, 2), IntArg(args, 3)]);
                if (member == "getUsedRange") return UsedRange(sheet.Cells, BoolArg(args, 0));
                if (member == "activate") { sheet.Activate(); return target; }
                if (member == "delete") { DeleteSheet(app, sheet); return target; }
                if (member == "copy") return CopySheet(sheet, args);
                if (member == "freezePanes.freezeRows" || member == "freezePanes.freezeColumns") { Freeze(sheet, member, IntArg(args, 0)); return target; }
                if (member == "charts.add") return AddChart(sheet, StringArg(args, 0), RangeArgument(Arg(args, 1)), Optional(args, 2));
                if (member == "tables.add") return AddTable(sheet, StringArg(args, 0), BoolArg(args, 1));
                if (member == "pivotTables.add") return AddPivot(sheet, StringArg(args, 0), RangeArgument(Arg(args, 1)), RangeArgument(Arg(args, 2)));
                if (member == "autoFilter.apply") { ApplyFilter(sheet, RangeArgument(Arg(args, 0)), Optional(args, 1), Optional(args, 2)); return target; }
                if (member == "autoFilter.clearCriteria") { if (Convert.ToBoolean(sheet.FilterMode)) sheet.ShowAllData(); return target; }
                if (member == "autoFilter.remove") { sheet.AutoFilterMode = false; return target; }
                if (member == "protection.protect") { sheet.Protect(); return target; }
                if (member == "protection.unprotect") { sheet.Unprotect(Guid.NewGuid().ToString("N")); return target; }
                if (member == "pageLayout.setPrintTitleRows") { sheet.PageSetup.PrintTitleRows = StringArg(args, 0); return target; }
            }
            if (target is RangeHandle)
            {
                dynamic range = ((RangeHandle)target).Range;
                if (range == null) throw new DispatchException("excel", "The range is null.");
                if (member == "worksheet") return new WorksheetHandle(range.Worksheet);
                if (member == "getColumn") return new RangeHandle(range.Columns[IntArg(args, 0) + 1]);
                if (member == "getRow") return new RangeHandle(range.Rows[IntArg(args, 0) + 1]);
                if (member == "getCell") return new RangeHandle(range.Cells[IntArg(args, 0) + 1, IntArg(args, 1) + 1]);
                if (member == "getResizedRange") return new RangeHandle(range.Resize[Convert.ToInt32(range.Rows.Count) + IntArg(args, 0), Convert.ToInt32(range.Columns.Count) + IntArg(args, 1)]);
                if (member == "getUsedRange") return UsedRange(range, BoolArg(args, 0));
                if (member == "insert") { range.Insert(Shift: InsertShift(StringArg(args, 0))); return target; }
                if (member == "delete") { range.Delete(Shift: DeleteShift(StringArg(args, 0))); return target; }
                if (member == "clear") { Clear(range, Optional(args, 0)); return target; }
                if (member == "select") { range.Select(); return target; }
                if (member == "sort") { Sort(range, Arg(args, 0), BoolArg(args, 1), BoolArg(args, 2)); return target; }
                if (member == "merge") { range.Merge(Optional(args, 0)); return target; }
                if (member == "unmerge") { range.UnMerge(); return target; }
                if (member == "autoFill") { range.AutoFill(RangeArgument(Arg(args, 0)), AutoFillType(StringArg(args, 1))); return target; }
                if (member == "copyFrom") { CopyFrom(range, RangeArgument(Arg(args, 0)), Optional(args, 1), Optional(args, 2), Optional(args, 3)); return target; }
                if (member == "moveTo") { range.Cut(RangeArgument(Arg(args, 0))); return target; }
                if (member == "removeDuplicates") return RemoveDuplicates(range, Arg(args, 0), BoolArg(args, 1));
                if (member == "dataValidation.clear") { range.Validation.Delete(); return target; }
                if (member == "format.autofitColumns") { range.Columns.AutoFit(); return target; }
                if (member == "format.autofitRows") { range.Rows.AutoFit(); return target; }
                if (member == "conditionalFormats.add") return AddConditionalFormat(range, StringArg(args, 0));
                if (member == "replaceAll") return ReplaceAll(range, StringArg(args, 0), StringArg(args, 1), Optional(args, 2));
            }
            if (target is TableHandle)
            {
                dynamic table = ((TableHandle)target).Table;
                if (member == "getRange") return new RangeHandle(table == null ? null : table.Range);
                if (member == "getDataBodyRange") return new RangeHandle(table == null ? null : table.DataBodyRange);
                if (member == "columns.add") return AddTableColumn(table, args);
            }
            if (target is TableColumnHandle && member == "getDataBodyRange") return new RangeHandle(((TableColumnHandle)target).Column.DataBodyRange);
            if (target is PivotHandle)
            {
                dynamic pivot = ((PivotHandle)target).Pivot;
                if (member == "hierarchies.getItem") return new OpaqueHandle(pivot.PivotFields(StringArg(args, 0)), "pivot hierarchy");
                if (member == "rowHierarchies.add") { pivot.AddFields(RowFields: PivotFieldName(Arg(args, 0)), AddToTable: true); return target; }
                if (member == "columnHierarchies.add") { pivot.AddFields(ColumnFields: PivotFieldName(Arg(args, 0)), AddToTable: true); return target; }
                if (member == "dataHierarchies.add") return new PivotDataHandle(pivot.AddDataField(pivot.PivotFields(PivotFieldName(Arg(args, 0)))));
            }
            throw NoDispatch(member);
        }

        private Dictionary<string, object> Load(object target, string[] properties)
        {
            if (target is WorksheetsHandle) return LoadWorksheets((WorksheetsHandle)target, properties);
            if (target is WorksheetHandle) return LoadWorksheet((WorksheetHandle)target, properties);
            if (target is SelectedAreasHandle) return LoadSelectedAreas((SelectedAreasHandle)target, properties);
            if (target is FunctionHandle) return LoadFunction((FunctionHandle)target, properties);
            if (target is TableHandle) return LoadTable((TableHandle)target, properties);
            if (target is RemoveDuplicatesHandle) return LoadDuplicates((RemoveDuplicatesHandle)target, properties);
            if (target is ReplaceHandle) return LoadReplace((ReplaceHandle)target, properties);
            if (target is WorkbookHandle) return LoadWorkbook(properties);
            if (!(target is RangeHandle)) throw NoDispatch(properties.Length == 0 ? "" : properties[0]);
            return LoadRange((RangeHandle)target, properties);
        }

        private Dictionary<string, object> LoadWorkbook(string[] properties)
        {
            dynamic app = ExcelDnaUtil.Application;
            var loaded = new Dictionary<string, object>();
            var nameProperties = new List<string>();
            var linkProperties = new List<string>();
            foreach (var property in properties)
            {
                if (property == "application/calculationMode") loaded[property] = CalculationMode(app.Calculation);
                else if (property.StartsWith("names/", StringComparison.Ordinal)) nameProperties.Add(property);
                else if (property.StartsWith("linkedWorkbooks/", StringComparison.Ordinal)) linkProperties.Add(property);
                else throw NoDispatch(property);
            }
            dynamic workbook = CurrentWorkbook();
            if (nameProperties.Count != 0) loaded["names/items"] = LoadNames(workbook, nameProperties);
            if (linkProperties.Count != 0) loaded["linkedWorkbooks/items"] = LoadLinks(workbook, linkProperties);
            return loaded;
        }

        private Dictionary<string, object> LoadRange(RangeHandle target, string[] properties)
        {
            dynamic range = target.Range;
            var loaded = new Dictionary<string, object>();
            foreach (var property in properties)
            {
                if (range == null) { loaded[property] = NullRangeValue(property); continue; }
                if (property == "address") loaded[property] = QualifiedAddress(range);
                else if (property == "isNullObject") loaded[property] = false;
                else if (property == "text") loaded[property] = TextMatrix(range);
                else if (property == "formulas") loaded[property] = FormulaMatrix(range);
                else if (property == "values") loaded[property] = ValueMatrix(range);
                else if (property == "valueTypes") loaded[property] = ValueTypes(range);
                else if (property == "numberFormat") loaded[property] = NumberFormatMatrix(range);
                else if (property == "cellCount") loaded[property] = Convert.ToDouble(range.CountLarge);
                else if (property == "format/columnWidth") loaded[property] = ColumnPointWidth(range);
                else if (property == "format/rowHeight") loaded[property] = Convert.ToDouble(range.RowHeight, CultureInfo.InvariantCulture);
                else if (property == "rowCount") loaded[property] = Convert.ToInt32(range.Rows.Count);
                else if (property == "columnCount") loaded[property] = Convert.ToInt32(range.Columns.Count);
                else if (property == "rowIndex") loaded[property] = Convert.ToInt32(range.Row) - 1;
                else if (property == "columnIndex") loaded[property] = Convert.ToInt32(range.Column) - 1;
                else if (property == "rowHidden") loaded[property] = Convert.ToBoolean(range.EntireRow.Hidden);
                else if (property == "columnHidden") loaded[property] = Convert.ToBoolean(range.EntireColumn.Hidden);
                else if (property == "worksheet/name") loaded[property] = Convert.ToString(range.Worksheet.Name);
                else throw NoDispatch(property);
            }
            return loaded;
        }

        private Dictionary<string, object> LoadWorksheet(WorksheetHandle target, string[] properties)
        {
            dynamic sheet = target.Worksheet;
            var loaded = new Dictionary<string, object>();
            var tableProperties = new List<string>();
            foreach (var property in properties)
            {
                if (sheet == null) { loaded[property] = property == "isNullObject" ? (object)true : property == "visibility" ? "Visible" : string.Empty; continue; }
                if (property == "name") loaded[property] = Convert.ToString(sheet.Name);
                else if (property == "id") loaded[property] = Convert.ToString(sheet.CodeName);
                else if (property == "visibility") loaded[property] = Visibility(sheet.Visible);
                else if (property == "isNullObject") loaded[property] = false;
                else if (property.StartsWith("tables/", StringComparison.Ordinal)) tableProperties.Add(property);
                else throw NoDispatch(property);
            }
            if (sheet != null && tableProperties.Count != 0)
                loaded["tables/items"] = LoadSheetTables(sheet, tableProperties);
            return loaded;
        }

        private Dictionary<string, object> LoadWorksheets(WorksheetsHandle target, string[] properties)
        {
            var items = new List<object>();
            foreach (dynamic sheet in target.Worksheets)
            {
                var child = nextHostHandle--; handles[child] = new WorksheetHandle(sheet);
                var item = new Dictionary<string, object> { { "id", child } };
                foreach (var property in properties) if (property.StartsWith("items/", StringComparison.Ordinal)) item[property.Substring(6)] = WorksheetProperty(sheet, property.Substring(6)); else throw NoDispatch(property);
                items.Add(item);
            }
            return new Dictionary<string, object> { { "items", items } };
        }

        private Dictionary<string, object> LoadSelectedAreas(SelectedAreasHandle target, string[] properties)
        {
            var loaded = new Dictionary<string, object>(); dynamic range = target.Range;
            foreach (var property in properties)
            {
                if (property == "address") loaded[property] = range == null ? "" : QualifiedAddress(range);
                else if (property == "worksheet/name") loaded[property] = range == null ? "" : Convert.ToString(range.Worksheet.Name);
                else if (property == "areas/items/cellCount")
                {
                    var items = new List<object>(); if (range != null) foreach (dynamic area in range.Areas) { var id = nextHostHandle--; handles[id] = new RangeHandle(area); items.Add(new Dictionary<string, object> { { "id", id }, { "cellCount", Convert.ToDouble(area.CountLarge) } }); }
                    loaded["areas/items"] = items;
                }
                else throw NoDispatch(property);
            }
            return loaded;
        }

        private static Dictionary<string, object> LoadFunction(FunctionHandle target, string[] properties)
        {
            var loaded = new Dictionary<string, object>();
            foreach (var property in properties)
            {
                if (property != "value") throw NoDispatch(property);
                loaded[property] = FunctionValue(target.Name, target.Range);
            }
            return loaded;
        }

        private static object FunctionValue(string name, dynamic range)
        {
            var operation = name.ToUpperInvariant();
            if (operation != "COUNTA" && operation != "SUM" && operation != "AVERAGE" &&
                operation != "MIN" && operation != "MAX" && operation != "COUNT" &&
                operation != "COUNTBLANK") throw NoDispatch("func " + name);
            dynamic app = ExcelDnaUtil.Application;
            var source = Convert.ToString(range.Address[true, true, 1, true], CultureInfo.InvariantCulture);
            object value = app.Evaluate(operation + "(" + source + ")");
            var error = value as ErrorWrapper;
            return error == null ? JsonValue(value) : ExcelError(error.ErrorCode);
        }

        private static string ExcelError(int code)
        {
            if (code == 2000) return "#NULL!";
            if (code == 2007) return "#DIV/0!";
            if (code == 2015) return "#VALUE!";
            if (code == 2023) return "#REF!";
            if (code == 2029) return "#NAME?";
            if (code == 2036) return "#NUM!";
            if (code == 2042) return "#N/A";
            if (code == 2043) return "#GETTING_DATA";
            if (code == 2045) return "#SPILL!";
            if (code == 2046) return "#CONNECT!";
            if (code == 2047) return "#BLOCKED!";
            if (code == 2048) return "#UNKNOWN!";
            if (code == 2049) return "#FIELD!";
            if (code == 2050) return "#CALC!";
            if (code == 2051) return "#BUSY!";
            if (code == 2052) return "#PYTHON!";
            if (code == 2053) return "#DATA!";
            throw new DispatchException("excel", "Unknown Excel error code " + code.ToString(CultureInfo.InvariantCulture) + ".");
        }

        private static Dictionary<string, object> LoadTable(TableHandle target, string[] properties)
        {
            dynamic table = target.Table; var loaded = new Dictionary<string, object>();
            foreach (var property in properties)
            {
                if (property == "isNullObject") loaded[property] = table == null;
                else if (property == "name") loaded[property] = table == null ? "" : Convert.ToString(table.Name);
                else throw NoDispatch(property);
            }
            return loaded;
        }
        private static Dictionary<string, object> LoadDuplicates(RemoveDuplicatesHandle target, string[] properties) { return Values(properties, "removed", target.Removed, "uniqueRemaining", target.Remaining); }
        private static Dictionary<string, object> LoadReplace(ReplaceHandle target, string[] properties) { return Values(properties, "value", target.Value, null, null); }

        private void Set(object target, string path, object value)
        {
            dynamic app = ExcelDnaUtil.Application;
            if (target is WorkbookHandle && path == "application.calculationMode") { app.Calculation = Calculation(String(value)); return; }
            if (target is WorksheetHandle)
            {
                dynamic sheet = ((WorksheetHandle)target).Worksheet; if (sheet == null) throw new DispatchException("excel", "The worksheet is null.");
                if (path == "name") { sheet.Name = String(value); return; }
                if (path.StartsWith("pageLayout.", StringComparison.Ordinal)) { SetPageLayout(sheet, path, value); return; }
            }
            if (target is TableHandle)
            {
                dynamic table = ((TableHandle)target).Table; if (table == null) throw new DispatchException("excel", "The table is null.");
                if (path == "name") { table.Name = String(value); return; } if (path == "style") { table.TableStyle = String(value); return; }
            }
            if (target is PivotDataHandle)
            {
                dynamic field = ((PivotDataHandle)target).Field;
                if (path == "summarizeBy") { field.Function = Summary(String(value)); return; }
                if (path == "showAs") { SetShowAs(field, value); return; }
            }
            if (target is OpaqueHandle) { SetOpaque((OpaqueHandle)target, path, value); return; }
            if (!(target is RangeHandle)) throw NoDispatch(path);
            dynamic range = ((RangeHandle)target).Range; if (range == null) throw new DispatchException("excel", "The range is null.");
            if (path == "formulas") { SetFormulaMatrix(range, ComMatrix(value)); return; }
            if (path == "numberFormat") { SetNumberFormat(range, value); return; }
            if (path == "rowHidden") { range.EntireRow.Hidden = Bool(value); return; }
            if (path == "columnHidden") { range.EntireColumn.Hidden = Bool(value); return; }
            if (path == "dataValidation.rule") { SetValidation(range, value); return; }
            if (path == "format.fill.color") { range.Interior.Color = Color(value); return; }
            if (path == "format.font.bold") { range.Font.Bold = Bool(value); return; }
            if (path == "format.font.italic") { range.Font.Italic = Bool(value); return; }
            if (path == "format.font.color") { range.Font.Color = Color(value); return; }
            if (path == "format.horizontalAlignment") { range.HorizontalAlignment = Horizontal(String(value)); return; }
            if (path == "format.columnWidth") { SetColumnPointWidth(range, Convert.ToDouble(value, CultureInfo.InvariantCulture)); return; }
            if (path == "format.rowHeight") { range.RowHeight = Convert.ToDouble(value, CultureInfo.InvariantCulture); return; }
            if (path == "format.wrapText") { range.WrapText = Bool(value); return; }
            if (path.StartsWith("format.borders.", StringComparison.Ordinal)) { SetBorder(range, path, value); return; }
            throw NoDispatch(path);
        }

        // COM helpers and enum mappings deliberately stay explicit: Office.js words are not COM values.
        private static void SetPageLayout(dynamic sheet, string path, object value) { dynamic p = sheet.PageSetup; if (path == "pageLayout.orientation") p.Orientation = String(value) == "Landscape" ? 2 : 1; else if (path == "pageLayout.paperSize") p.PaperSize = Paper(String(value)); else if (path == "pageLayout.printGridlines") p.PrintGridlines = Bool(value); else if (path == "pageLayout.centerHorizontally") p.CenterHorizontally = Bool(value); else if (path == "pageLayout.zoom.horizontalFitToPages") { p.Zoom = false; p.FitToPagesWide = Convert.ToInt32(value); } else if (path == "pageLayout.zoom.verticalFitToPages") { p.Zoom = false; p.FitToPagesTall = Convert.ToInt32(value); } else throw NoDispatch(path); }
        private static void SetOpaque(OpaqueHandle opaque, string path, object value)
        {
            dynamic o = opaque.Value;
            if (opaque.Label == "chart" && path == "title.text") { o.HasTitle = true; o.ChartTitle.Text = String(value); return; }
            if (opaque.Label == "conditional format")
            {
                if (path == "cellValue.format.fill.color") { o.Interior.Color = Color(value); return; }
                if (path == "cellValue.format.font.color") { o.Font.Color = Color(value); return; }
                if (path == "cellValue.rule") { SetCellValueRule(o, value); return; }
                if (path == "colorScale.criteria") { SetColorScaleCriteria(opaque, value); return; }
            }
            throw NoDispatch(path);
        }
        private static void SetCellValueRule(dynamic condition, object value)
        {
            var rule = value as Dictionary<string, object>;
            if (rule == null || !rule.ContainsKey("operator") || !rule.ContainsKey("formula1"))
                throw new DispatchException("protocol", "Conditional cell-value rule is invalid.");
            condition.Modify(1, ConditionOperator(String(rule["operator"])), String(rule["formula1"]),
                rule.ContainsKey("formula2") ? String(rule["formula2"]) : Type.Missing);
        }
        private static void SetColorScaleCriteria(OpaqueHandle opaque, object value)
        {
            var criteria = value as Dictionary<string, object>;
            if (criteria == null) throw new DispatchException("protocol", "Color-scale criteria are invalid.");
            var hasMinimum = criteria.ContainsKey("minimum");
            var hasMiddle = criteria.ContainsKey("middle");
            var hasMaximum = criteria.ContainsKey("maximum");
            if (!hasMinimum || !hasMaximum || (hasMiddle && criteria.Count != 3) ||
                (!hasMiddle && criteria.Count != 2))
                throw new DispatchException("protocol", "Color-scale criteria must contain minimum and maximum, with an optional middle.");
            var cardinality = hasMiddle ? 3 : 2;
            dynamic scale = opaque.Value;
            if (opaque.ColorScaleCardinality != cardinality)
            {
                scale.Delete();
                scale = opaque.Range.FormatConditions.AddColorScale(cardinality);
                opaque.Value = scale;
                opaque.ColorScaleCardinality = cardinality;
            }
            ApplyScaleCriterion(scale.ColorScaleCriteria[1], criteria, "minimum");
            if (cardinality == 3) ApplyScaleCriterion(scale.ColorScaleCriteria[2], criteria, "middle");
            ApplyScaleCriterion(scale.ColorScaleCriteria[cardinality], criteria, "maximum");
        }
        private static void ApplyScaleCriterion(dynamic criterion, Dictionary<string, object> all, string key)
        {
            object raw;
            if (!all.TryGetValue(key, out raw)) return;
            var value = raw as Dictionary<string, object>;
            if (value == null) throw new DispatchException("protocol", "Color-scale " + key + " is invalid.");
            object type;
            if (value.TryGetValue("type", out type)) criterion.Type = ColorScaleType(String(type));
            object color;
            if (value.TryGetValue("color", out color)) criterion.FormatColor.Color = Color(color);
            object formula;
            if (value.TryGetValue("formula", out formula)) criterion.Value = String(formula);
        }
        private static void SetShowAs(dynamic field, object value) { var map = value as Dictionary<string, object>; if (map == null || !map.ContainsKey("calculation")) throw new DispatchException("protocol", "showAs needs a calculation."); field.Calculation = ShowAs(String(map["calculation"])); }
        private static void SetValidation(dynamic range, object value)
        {
            var rule = value as Dictionary<string, object>;
            if (rule == null) throw new DispatchException("protocol", "data validation rule must be an object.");
            dynamic validation = range.Validation;
            validation.Delete();
            object listValue;
            if (rule.TryGetValue("list", out listValue))
            {
                var list = listValue as Dictionary<string, object>;
                object source;
                if (list == null || !list.TryGetValue("source", out source))
                    throw new DispatchException("protocol", "List validation needs a source.");
                validation.Add(Type: 3, AlertStyle: 1, Operator: 1, Formula1: String(source));
                object dropdown;
                if (list.TryGetValue("inCellDropDown", out dropdown)) validation.InCellDropdown = Bool(dropdown);
                return;
            }
            validation.Add(Type: ValidationType(rule), AlertStyle: 1, Operator: ValidationOperator(rule), Formula1: RuleValue(rule, "formula1"), Formula2: RuleValue(rule, "formula2"));
        }
        private static void SetBorder(dynamic range, string path, object value) { var split = path.Split('.'); if (split.Length != 4) throw NoDispatch(path); dynamic border = range.Borders[Border(split[2])]; if (split[3] == "color") border.Color = Color(value); else if (split[3] == "style") border.LineStyle = BorderStyle(String(value)); else throw NoDispatch(path); }

        private static string QualifiedAddress(dynamic range)
        {
            var sheet = Convert.ToString(range.Worksheet.Name, CultureInfo.InvariantCulture);
            var escaped = sheet.Replace("'", "''");
            return "'" + escaped + "'!" + Convert.ToString(range.Address, CultureInfo.InvariantCulture);
        }

        private static object NullRangeValue(string property) { if (property == "isNullObject") return true; if (property == "address" || property == "worksheet/name") return ""; if (property == "rowCount" || property == "columnCount" || property == "cellCount" || property == "rowIndex" || property == "columnIndex") return 0; if (property == "rowHidden" || property == "columnHidden") return false; return new object[0]; }
        private static object DisplayTextMatrix(dynamic range, int height, int width)
        {
            dynamic application = null;
            dynamic worksheetFunction = null;
            object result = null;
            Exception primaryFailure = null;
            try
            {
                application = range.Application;
                worksheetFunction = application.WorksheetFunction;
                var rows = new List<object>();
                for (var row = 1; row <= height; row++)
                {
                    var cells = new List<object>();
                    for (var column = 1; column <= width; column++)
                    {
                        dynamic cell = range.Cells[row, column];
                        cells.Add(DisplayTextAndRelease(worksheetFunction, cell));
                    }
                    rows.Add(cells);
                }
                result = rows;
            }
            catch (Exception exception) { primaryFailure = exception; }
            Exception releaseFailure = null;
            try { ReleaseCom(worksheetFunction); }
            catch (Exception exception) { releaseFailure = CombineCleanupFailures(releaseFailure, exception); }
            try { ReleaseCom(application); }
            catch (Exception exception) { releaseFailure = CombineCleanupFailures(releaseFailure, exception); }
            if (primaryFailure != null && releaseFailure != null)
                throw new AggregateException(primaryFailure, releaseFailure);
            if (primaryFailure != null) throw primaryFailure;
            if (releaseFailure != null) throw releaseFailure;
            return result;
        }

        private static string DisplayTextAndRelease(dynamic worksheetFunction, dynamic cell)
        {
            string result = null;
            Exception primaryFailure = null;
            try { result = DisplayText(worksheetFunction, cell); }
            catch (Exception exception) { primaryFailure = exception; }
            Exception releaseFailure = null;
            try { ReleaseCom(cell); }
            catch (Exception exception) { releaseFailure = exception; }
            if (primaryFailure != null && releaseFailure != null)
                throw new AggregateException(primaryFailure, releaseFailure);
            if (primaryFailure != null) throw primaryFailure;
            if (releaseFailure != null) throw releaseFailure;
            return result;
        }

        private static string DisplayText(dynamic worksheetFunction, dynamic cell)
        {
            object value = cell.Value2;
            if (value == null || value is DBNull) return string.Empty;
            var error = value as ErrorWrapper;
            if (error != null) return ExcelError(error.ErrorCode);
            if (value is string) return (string)value;
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            var format = Convert.ToString(cell.NumberFormatLocal, CultureInfo.CurrentCulture);
            if (string.IsNullOrEmpty(format)) format = "General";
            return Convert.ToString(
                worksheetFunction.Text(value, format),
                CultureInfo.CurrentCulture) ?? string.Empty;
        }
        private static void ValidateExternalRequest(ExternalWorkbookRequest request)
        {
            if (request == null || request.Area == null ||
                string.IsNullOrWhiteSpace(request.Path) || string.IsNullOrWhiteSpace(request.Sheet) ||
                request.Area.Top < 1 || request.Area.Left < 1 ||
                request.Area.Height < 1 || request.Area.Width < 1)
                throw new DispatchException("protocol", "External workbook request has an invalid path, sheet, or area.");
        }
        private static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        private static Exception CaptureRelease(object value, Exception failure)
        {
            try
            {
                if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
                return failure;
            }
            catch (Exception exception) { return CombineCleanupFailures(failure, exception); }
        }

        private static Exception CombineCleanupFailures(Exception existing, Exception next)
        {
            if (existing == null) return next;
            var aggregate = existing as AggregateException;
            if (aggregate == null) return new AggregateException(existing, next);
            var failures = new List<Exception>(aggregate.Flatten().InnerExceptions) { next };
            return new AggregateException(failures);
        }

        private static string CleanupFailureMessage(Exception failure)
        {
            var aggregate = failure as AggregateException;
            if (aggregate == null) return failure.Message;
            var messages = new List<string>();
            foreach (var exception in aggregate.Flatten().InnerExceptions) messages.Add(exception.Message);
            return string.Join(" | ", messages);
        }
        private static object TextMatrix(dynamic range)
        {
            return DisplayTextMatrix(
                range,
                Convert.ToInt32(range.Rows.Count),
                Convert.ToInt32(range.Columns.Count));
        }

        private static object NumberFormatMatrix(dynamic range)
        {
            var rows = new List<object>();
            for (var row = 1; row <= Convert.ToInt32(range.Rows.Count); row++)
            {
                var cells = new List<object>();
                for (var column = 1; column <= Convert.ToInt32(range.Columns.Count); column++)
                {
                    dynamic cell = range.Cells[row, column];
                    try
                    {
                        cells.Add(Convert.ToString(
                            cell.NumberFormat,
                            CultureInfo.InvariantCulture) ?? string.Empty);
                    }
                    finally { ReleaseCom(cell); }
                }
                rows.Add(cells);
            }
            return rows;
        }

        private static object FormulaMatrix(dynamic range)
        {
            object values = Formula2(range);
            int height = Convert.ToInt32(range.Rows.Count);
            int width = Convert.ToInt32(range.Columns.Count);
            return NormalizeMatrix(values, height, width, FormulaOrValue);
        }

        private static object Formula2(dynamic range)
        {
            try { return range.Formula2; }
            catch (RuntimeBinderException) { return range.Formula; }
            catch (COMException exception) when (IsMissingFormula2Member(exception)) { return range.Formula; }
        }

        private static void SetFormulaMatrix(dynamic range, object values)
        {
            try { range.Formula2 = values; }
            catch (RuntimeBinderException) { range.Formula = values; }
            catch (COMException exception) when (IsMissingFormula2Member(exception)) { range.Formula = values; }
        }

        private static bool IsMissingFormula2Member(COMException exception)
        {
            return exception.ErrorCode == unchecked((int)0x80020003) ||
                exception.ErrorCode == unchecked((int)0x80020006);
        }

        private static void SetNumberFormat(dynamic range, object value)
        {
            var matrix = (Array)ComMatrix(value);
            var rows = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            var columns = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            if (matrix.GetLength(0) == 1 && matrix.GetLength(1) == 1)
            {
                range.NumberFormat = matrix.GetValue(0, 0);
                return;
            }
            if (matrix.GetLength(0) != rows || matrix.GetLength(1) != columns)
                throw new DispatchException("protocol", "Number-format matrix must be one cell or match the target range.");
            range.NumberFormat = matrix;
        }

        private static object ColumnPointWidth(dynamic range)
        {
            dynamic columns = range.Columns;
            try
            {
                var count = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
                if (count == 0) return null;
                double? width = null;
                for (var index = 1; index <= count; index++)
                {
                    dynamic column = columns[index];
                    try
                    {
                        var current = Convert.ToDouble(column.Width, CultureInfo.InvariantCulture);
                        if (width == null) width = current;
                        else if (Math.Abs(width.Value - current) > 0.01) return null;
                    }
                    finally { ReleaseCom(column); }
                }
                return width;
            }
            finally { ReleaseCom(columns); }
        }

        private static void SetColumnPointWidth(dynamic range, double points)
        {
            if (double.IsNaN(points) || double.IsInfinity(points) || points <= 0)
                throw new DispatchException("protocol", "Column width must be a positive point value.");
            dynamic columns = range.Columns;
            try
            {
                var count = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    dynamic column = columns[index];
                    try
                    {
                        for (var attempt = 0; attempt < 4; attempt++)
                        {
                            var actual = Convert.ToDouble(column.Width, CultureInfo.InvariantCulture);
                            if (Math.Abs(actual - points) <= 0.01) break;
                            if (actual <= 0)
                                throw new DispatchException("excel", "Excel cannot set a point width on a zero-width column.");
                            var comWidth = Convert.ToDouble(column.ColumnWidth, CultureInfo.InvariantCulture);
                            column.ColumnWidth = comWidth * points / actual;
                        }
                    }
                    finally { ReleaseCom(column); }
                }
            }
            finally { ReleaseCom(columns); }
        }

        private static object ValueMatrix(dynamic range)
        {
            int height = Convert.ToInt32(range.Rows.Count);
            int width = Convert.ToInt32(range.Columns.Count);
            object values = range.Value2;
            return NormalizeMatrix(values, height, width, FormulaOrValue);
        }

        private static object ValueTypes(dynamic range)
        {
            int height = Convert.ToInt32(range.Rows.Count);
            int width = Convert.ToInt32(range.Columns.Count);
            object values = range.Value2;
            return NormalizeMatrix(values, height, width, ValueType);
        }

        private static object NormalizeMatrix(
            object source,
            int height,
            int width,
            Func<object, object> normalize)
        {
            var array = source as Array;
            if (array == null)
            {
                if (height != 1 || width != 1)
                    throw new DispatchException("excel", "Excel returned a scalar for a multi-cell range.");
                return new object[] { new object[] { normalize(source) } };
            }
            if (array.Rank != 2 || array.GetLength(0) != height || array.GetLength(1) != width)
                throw new DispatchException("excel", "Excel returned an unexpected range matrix shape.");
            var rows = new List<object>();
            var rowBase = array.GetLowerBound(0);
            var columnBase = array.GetLowerBound(1);
            for (var row = 0; row < height; row++)
            {
                var cells = new List<object>();
                for (var column = 0; column < width; column++)
                    cells.Add(normalize(array.GetValue(rowBase + row, columnBase + column)));
                rows.Add(cells);
            }
            return rows;
        }

        private static object FormulaOrValue(object value)
        {
            if (value == null || value is DBNull) return string.Empty;
            var error = value as ErrorWrapper;
            return error == null ? JsonValue(value) : ExcelError(error.ErrorCode);
        }

        private static object ValueType(object value)
        {
            if (value == null || value is DBNull) return "Empty";
            if (value is ErrorWrapper) return "Error";
            if (value is string) return "String";
            if (value is bool) return "Boolean";
            return "Double";
        }
        private static object JsonValue(object value) { var array = value as Array; if (array == null) return value is DBNull ? null : value; var rows = new List<object>(); if (array.Rank == 1) { for (var i = array.GetLowerBound(0); i <= array.GetUpperBound(0); i++) rows.Add(JsonValue(array.GetValue(i))); return rows; } for (var r = array.GetLowerBound(0); r <= array.GetUpperBound(0); r++) { var row = new List<object>(); for (var c = array.GetLowerBound(1); c <= array.GetUpperBound(1); c++) row.Add(JsonValue(array.GetValue(r,c))); rows.Add(row); } return rows; }
        private static object ComMatrix(object value) { var rows = value as IEnumerable; if (rows == null || value is string) throw new DispatchException("protocol", "A matrix is required."); var data = new List<List<object>>(); foreach (var sourceRow in rows) { var row = sourceRow as IEnumerable; if (row == null || sourceRow is string) throw new DispatchException("protocol", "Matrix rows are required."); var output = new List<object>(); foreach (var cell in row) output.Add(cell); data.Add(output); } if (data.Count == 0) return new object[0,0]; var result = Array.CreateInstance(typeof(object), data.Count, data[0].Count); for (var r=0;r<data.Count;r++) { if (data[r].Count != data[0].Count) throw new DispatchException("protocol", "Matrix rows have different widths."); for(var c=0;c<data[r].Count;c++) result.SetValue(data[r][c],r,c); } return result; }

        private dynamic CurrentWorkbook()
        {
            dynamic sheet = excelWindow.ActiveSheet;
            if (sheet == null) throw new DispatchException("excel", "The Excel window has no active worksheet.");
            return sheet.Parent;
        }

        private static void RefreshLinks(dynamic workbook)
        {
            object links = workbook.LinkSources(1);
            if (links != null) workbook.UpdateLink(links, 1);
        }

        private dynamic RangeArgument(object value) { var reference = value as Dictionary<string, object>; if (reference == null || !reference.ContainsKey("handle")) throw new DispatchException("protocol", "A range handle argument is required."); var id = Convert.ToInt32(reference["handle"]); object handle; if (!handles.TryGetValue(id, out handle) || !(handle is RangeHandle)) throw new DispatchException("protocol", "Unknown range handle " + id + "."); return ((RangeHandle)handle).Range; }
        private void ResetBatchHandles()
        {
            var released = new HashSet<object>(ReferenceComparer.Instance);
            Exception releaseFailure = null;
            try
            {
                foreach (var handle in handles.Values)
                {
                    try { ReleaseHandle(handle, released); }
                    catch (Exception exception) { releaseFailure = CombineCleanupFailures(releaseFailure, exception); }
                }
            }
            finally
            {
                handles.Clear();
                handles[0] = WorkbookHandle.Instance;
                nextHostHandle = -1;
            }
            if (releaseFailure != null) throw releaseFailure;
        }
        private static void ReleaseHandle(object handle, HashSet<object> released)
        {
            if (handle is WorksheetsHandle) ReleaseTrackedCom(((WorksheetsHandle)handle).Worksheets, released);
            else if (handle is WorksheetHandle) ReleaseTrackedCom(((WorksheetHandle)handle).Worksheet, released);
            else if (handle is RangeHandle) ReleaseTrackedCom(((RangeHandle)handle).Range, released);
            else if (handle is SelectedAreasHandle) ReleaseTrackedCom(((SelectedAreasHandle)handle).Range, released);
            else if (handle is FunctionHandle) ReleaseTrackedCom(((FunctionHandle)handle).Range, released);
            else if (handle is NameHandle) ReleaseTrackedCom(((NameHandle)handle).Name, released);
            else if (handle is TableHandle) ReleaseTrackedCom(((TableHandle)handle).Table, released);
            else if (handle is TableColumnHandle) ReleaseTrackedCom(((TableColumnHandle)handle).Column, released);
            else if (handle is OpaqueHandle)
            {
                var opaque = (OpaqueHandle)handle;
                ReleaseTrackedCom(opaque.Value, released);
                ReleaseTrackedCom(opaque.Range, released);
            }
            else if (handle is PivotHandle) ReleaseTrackedCom(((PivotHandle)handle).Pivot, released);
            else if (handle is PivotDataHandle) ReleaseTrackedCom(((PivotDataHandle)handle).Field, released);
        }
        private static void ReleaseTrackedCom(object value, HashSet<object> released)
        {
            if (value != null && Marshal.IsComObject(value) && released.Add(value))
                Marshal.ReleaseComObject(value);
        }
        private static object Arg(object[] args, int index) { if (args == null || args.Length <= index) throw new DispatchException("protocol", "Missing argument " + index + "."); return args[index]; }
        private static object Optional(object[] args, int index) { return args != null && args.Length > index ? args[index] : null; }
        private static string StringArg(object[] args, int index) { return String(Arg(args, index)); }
        private static string String(object value) { return value == null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture); }
        private static int IntArg(object[] args, int index) { return Convert.ToInt32(Arg(args,index), CultureInfo.InvariantCulture); }
        private static bool BoolArg(object[] args, int index) { return Bool(Arg(args,index)); }
        private static bool Bool(object value) { return value is bool && (bool)value; }
        private static bool IsEditMode(Exception exception)
        {
            if (exception is TimeoutException) return true;
            var message = exception.Message ?? "";
            return message.IndexOf("edit mode", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        private string Failure(Dictionary<string, object> values, string code, string message) { return json.Serialize(new Dictionary<string, object> { { "values", values }, { "failure", new Dictionary<string,string> { { "code", code }, { "message", message } } } }); }
        private static DispatchException NoDispatch(string member) { return new DispatchException("dispatch", "no dispatch for \"" + member + "\""); }

        private sealed class WorkbookHandle { internal static readonly WorkbookHandle Instance = new WorkbookHandle(); }
        private sealed class WorksheetsHandle { internal dynamic Worksheets; internal WorksheetsHandle(dynamic value) { Worksheets = value; } }
        private sealed class WorksheetHandle { internal dynamic Worksheet; internal WorksheetHandle(dynamic value) { Worksheet = value; } }
        private sealed class RangeHandle { internal dynamic Range; internal RangeHandle(dynamic value) { Range = value; } }
        private sealed class SelectedAreasHandle { internal dynamic Range; internal SelectedAreasHandle(dynamic value) { Range = value; } }
        private sealed class FunctionHandle { internal string Name; internal dynamic Range; internal FunctionHandle(string name, dynamic range) { Name = name; Range = range; } }
        private sealed class NameHandle { internal dynamic Name; internal NameHandle(dynamic value) { Name = value; } }
        private sealed class TableHandle { internal dynamic Table; internal TableHandle(dynamic value) { Table = value; } }
        private sealed class TableColumnHandle { internal dynamic Column; internal TableColumnHandle(dynamic value) { Column = value; } }
        private sealed class RemoveDuplicatesHandle { internal int Removed; internal int Remaining; internal RemoveDuplicatesHandle(int removed, int remaining) { Removed=removed; Remaining=remaining; } }
        private sealed class ReplaceHandle { internal int Value; internal ReplaceHandle(int value) { Value=value; } }
        private sealed class OpaqueHandle
        {
            internal dynamic Value;
            internal string Label;
            internal dynamic Range;
            internal int ColorScaleCardinality;
            internal OpaqueHandle(dynamic value, string label) { Value = value; Label = label; }
            internal OpaqueHandle(dynamic value, string label, dynamic range, int colorScaleCardinality)
            {
                Value = value; Label = label; Range = range; ColorScaleCardinality = colorScaleCardinality;
            }
        }
        private sealed class PivotHandle { internal dynamic Pivot; internal PivotHandle(dynamic value) { Pivot=value; } }
        private sealed class PivotDataHandle { internal dynamic Field; internal PivotDataHandle(dynamic value) { Field=value; } }
        private sealed class ReferenceComparer : IEqualityComparer<object>
        {
            internal static readonly ReferenceComparer Instance = new ReferenceComparer();
            public bool Equals(object left, object right) { return ReferenceEquals(left, right); }
            public int GetHashCode(object value) { return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value); }
        }
        private sealed class DispatchException : Exception { internal string Code; internal DispatchException(string code, string message) : base(message) { Code=code; } }
        public sealed class ExternalWorkbookRequest
        {
            public string Path { get; set; }
            public string Sheet { get; set; }
            public ExternalWorkbookArea Area { get; set; }
        }
        public sealed class ExternalWorkbookArea
        {
            public int Top { get; set; }
            public int Left { get; set; }
            public int Height { get; set; }
            public int Width { get; set; }
        }
        public sealed class BridgeOperation { public string Op { get; set; } public int Id { get; set; } public int On { get; set; } public string Member { get; set; } public object[] Args { get; set; } public string[] Properties { get; set; } }

        // Dynamic COM implementation details.
        private static void DeleteSheet(dynamic app, dynamic sheet)
        {
            var alerts = Convert.ToBoolean(app.DisplayAlerts);
            try
            {
                app.DisplayAlerts = false;
                sheet.Delete();
            }
            finally
            {
                app.DisplayAlerts = alerts;
            }
        }

        private static WorksheetHandle WorksheetByName(dynamic worksheets, string name) { foreach (dynamic sheet in worksheets) if (string.Equals(Convert.ToString(sheet.Name), name, StringComparison.OrdinalIgnoreCase)) return new WorksheetHandle(sheet); return new WorksheetHandle(null); }
        private static object NullForWorksheetCall(string member) { if (member == "getRange" || member == "getCell" || member == "getRangeByIndexes" || member == "getUsedRange") return new RangeHandle(null); throw new DispatchException("excel", "The worksheet is null."); }
        private static RangeHandle UsedRange(dynamic source, bool valuesOnly)
        {
            if (!valuesOnly)
            {
                dynamic worksheet = source.Worksheet;
                dynamic worksheetUsed = worksheet.UsedRange;
                var pristineA1 = Convert.ToDouble(worksheetUsed.CountLarge) == 1 &&
                    Convert.ToInt32(worksheetUsed.Row) == 1 &&
                    Convert.ToInt32(worksheetUsed.Column) == 1 &&
                    IsPristineA1(worksheetUsed.Cells[1, 1], worksheet);
                if (pristineA1) return new RangeHandle(null);
                dynamic app = ExcelDnaUtil.Application;
                dynamic intersection = app.Intersect(source, worksheetUsed);
                return new RangeHandle(intersection);
            }
            var lookIn = valuesOnly ? -4163 : -4123;
            dynamic lastCell = source.Cells[source.Rows.Count, source.Columns.Count];
            dynamic firstCell = source.Cells[1, 1];
            dynamic firstByRows = source.Find("*", lastCell, lookIn, 2, 1, 1, false, false, false);
            if (firstByRows == null) return new RangeHandle(null);
            dynamic firstByColumns = source.Find("*", lastCell, lookIn, 2, 2, 1, false, false, false);
            dynamic lastByRows = source.Find("*", firstCell, lookIn, 2, 1, 2, false, false, false);
            dynamic lastByColumns = source.Find("*", firstCell, lookIn, 2, 2, 2, false, false, false);
            dynamic sheet = source.Worksheet;
            dynamic used = sheet.Range[
                sheet.Cells[firstByRows.Row, firstByColumns.Column],
                sheet.Cells[lastByRows.Row, lastByColumns.Column]];
            return new RangeHandle(used);
        }

        private static bool IsPristineA1(dynamic cell, dynamic sheet)
        {
            try
            {
                if (cell.Value2 != null || Convert.ToBoolean(cell.HasFormula)) return false;
                if (!string.Equals(Convert.ToString(cell.NumberFormat), "General", StringComparison.Ordinal))
                    return false;
                dynamic application = ExcelDnaUtil.Application;
                if (Convert.ToInt32(cell.Interior.Pattern) != -4142 ||
                    Convert.ToInt32(cell.Interior.ColorIndex) != -4142 ||
                    Convert.ToInt32(cell.Interior.PatternColorIndex) != -4105 ||
                    Convert.ToBoolean(cell.Font.Bold) || Convert.ToBoolean(cell.Font.Italic) ||
                    Convert.ToInt32(cell.Font.ColorIndex) != -4105 ||
                    !string.Equals(Convert.ToString(cell.Font.Name), Convert.ToString(application.StandardFont), StringComparison.CurrentCulture) ||
                    Math.Abs(Convert.ToDouble(cell.Font.Size) - Convert.ToDouble(application.StandardFontSize)) > 0.001 ||
                    Convert.ToInt32(cell.Font.Underline) != -4142 ||
                    Convert.ToBoolean(cell.Font.Strikethrough) ||
                    Convert.ToBoolean(cell.Font.Superscript) || Convert.ToBoolean(cell.Font.Subscript) ||
                    Convert.ToBoolean(cell.Font.Shadow) || Convert.ToBoolean(cell.Font.OutlineFont) ||
                    Convert.ToBoolean(cell.WrapText) || Convert.ToBoolean(cell.ShrinkToFit) ||
                    Convert.ToInt32(cell.HorizontalAlignment) != -4105 ||
                    Convert.ToInt32(cell.VerticalAlignment) != -4107 ||
                    Convert.ToInt32(cell.IndentLevel) != 0 || Convert.ToInt32(cell.Orientation) != 0 ||
                    Convert.ToBoolean(cell.AddIndent) || Convert.ToInt32(cell.ReadingOrder) != -5002 ||
                    Convert.ToInt32(cell.FormatConditions.Count) != 0 ||
                    Convert.ToInt32(cell.Validation.Type) != -4142 ||
                    Convert.ToBoolean(cell.MergeCells) ||
                    !Convert.ToBoolean(cell.Locked) || Convert.ToBoolean(cell.FormulaHidden) ||
                    Convert.ToBoolean(cell.EntireRow.Hidden) ||
                    Convert.ToBoolean(cell.EntireColumn.Hidden))
                    return false;
                if (Math.Abs(Convert.ToDouble(cell.ColumnWidth) - Convert.ToDouble(sheet.StandardWidth)) > 0.001 ||
                    Math.Abs(Convert.ToDouble(cell.RowHeight) - Convert.ToDouble(sheet.StandardHeight)) > 0.001)
                    return false;
                foreach (var edge in new[] { 5, 6, 7, 8, 9, 10, 11, 12 })
                    if (Convert.ToInt32(cell.Borders[edge].LineStyle) != -4142) return false;
                return true;
            }
            catch { return false; }
        }
        private static TableHandle TableByName(dynamic workbook, string name) { foreach (dynamic sheet in workbook.Worksheets) foreach (dynamic table in sheet.ListObjects) if (string.Equals(Convert.ToString(table.Name), name, StringComparison.OrdinalIgnoreCase)) return new TableHandle(table); return new TableHandle(null); }
        private static RangeHandle NameRange(dynamic workbook, string name) { try { dynamic named = workbook.Names.Item(name); return new RangeHandle(named.RefersToRange); } catch { return new RangeHandle(null); } }
        private WorksheetHandle CopySheet(dynamic sheet, object[] args)
        {
            dynamic relative = null;
            var reference = Optional(args, 1) as Dictionary<string, object>;
            if (reference != null && reference.ContainsKey("handle"))
            {
                object target;
                var id = Convert.ToInt32(reference["handle"]);
                if (!handles.TryGetValue(id, out target) || !(target is WorksheetHandle))
                    throw new DispatchException("protocol", "Unknown worksheet handle " + id + ".");
                relative = ((WorksheetHandle)target).Worksheet;
                if (relative == null) throw new DispatchException("excel", "The relative worksheet is null.");
            }
            var position = StringArg(args, 0);
            if (position == "Before") sheet.Copy(Before: relative ?? sheet);
            else if (position == "After") sheet.Copy(After: relative ?? sheet);
            else throw new DispatchException("protocol", "Unknown sheet copy position " + position + ".");
            return new WorksheetHandle(excelWindow.ActiveSheet);
        }
        private void Freeze(dynamic sheet, string member, int count)
        {
            if (count < 0) throw new DispatchException("protocol", "Freeze count cannot be negative.");
            sheet.Activate();
            dynamic window = excelWindow;
            window.FreezePanes = false;
            if (member.EndsWith("Rows", StringComparison.Ordinal)) window.SplitRow = count;
            else window.SplitColumn = count;
            window.FreezePanes = Convert.ToInt32(window.SplitRow) > 0 ||
                Convert.ToInt32(window.SplitColumn) > 0;
        }
        private static OpaqueHandle AddChart(dynamic sheet, string type, dynamic source, object seriesBy)
        {
            dynamic chart = sheet.ChartObjects().Add(100, 100, 375, 225).Chart;
            var direction = String(seriesBy);
            if (string.IsNullOrEmpty(direction) || direction == "Auto") chart.SetSourceData(source);
            else chart.SetSourceData(source, direction == "Columns" ? 2 : 1);
            chart.ChartType = ChartType(type);
            return new OpaqueHandle(chart, "chart");
        }
        private static TableHandle AddTable(dynamic sheet, string address, bool headers) { return new TableHandle(sheet.ListObjects.Add(1, sheet.Range[address], Type.Missing, headers ? 1 : 2)); }
        private PivotHandle AddPivot(dynamic sheet,string name,dynamic source,dynamic destination) { dynamic cache=CurrentWorkbook().PivotCaches().Create(1,Convert.ToString(source.Address[true, true, 1, true])); return new PivotHandle(cache.CreatePivotTable(destination,name)); }
        private static TableColumnHandle AddTableColumn(dynamic table, object[] args) { if (table == null) throw new DispatchException("excel","The table is null."); dynamic column=table.ListColumns.Add(Optional(args,0) == null ? Type.Missing : Convert.ToInt32(Optional(args,0))+1); if(Optional(args,2)!=null) column.Name=String(Optional(args,2)); if(Optional(args,1)!=null) column.DataBodyRange.Value2=ComMatrix(Optional(args,1)); return new TableColumnHandle(column); }
        private static void ApplyFilter(dynamic sheet,dynamic range,object column,object criteria)
        {
            if (column == null) { range.AutoFilter(); return; }
            var field = Convert.ToInt32(column) + 1;
            var rule = criteria as Dictionary<string, object>;
            if (rule == null) throw new DispatchException("protocol", "AutoFilter criteria must be an object.");
            object filterOn;
            if (!rule.TryGetValue("filterOn", out filterOn)) throw new DispatchException("protocol", "AutoFilter criteria need filterOn.");
            var kind = String(filterOn);
            object criterion;
            if (kind == "Values")
            {
                if (!rule.TryGetValue("values", out criterion)) throw new DispatchException("protocol", "Values filter needs values.");
                range.AutoFilter(Field: field, Criteria1: EnumerableArray(criterion), Operator: 7);
                return;
            }
            if (!rule.TryGetValue("criterion1", out criterion)) throw new DispatchException("protocol", "AutoFilter criteria need criterion1.");
            if (kind == "TopItems") { range.AutoFilter(Field: field, Criteria1: String(criterion), Operator: 3); return; }
            if (kind == "Custom") { range.AutoFilter(Field: field, Criteria1: String(criterion)); return; }
            throw new DispatchException("protocol", "Unknown AutoFilter type " + kind + ".");
        }
        private static object[] EnumerableArray(object value)
        {
            var values = value as IEnumerable;
            if (values == null || value is string) throw new DispatchException("protocol", "Filter values must be an array.");
            var result = new List<object>();
            foreach (var item in values) result.Add(item);
            return result.ToArray();
        }
        private static void Calculate(dynamic app,string kind) { if(kind=="Full") app.CalculateFull(); else if(kind=="FullRebuild") app.CalculateFullRebuild(); else app.Calculate(); }
        private static object InsertShift(string value)
        {
            if (value == "Down") return -4121;
            if (value == "Right") return -4161;
            throw new DispatchException("protocol", "Unknown insert shift " + value + ".");
        }
        private static object DeleteShift(string value)
        {
            if (value == "Up") return -4162;
            if (value == "Left") return -4159;
            throw new DispatchException("protocol", "Unknown delete shift " + value + ".");
        }
        private static void Clear(dynamic range, object type) { var value=String(type); if(value=="Contents") range.ClearContents(); else if(value=="Formats") range.ClearFormats(); else range.Clear(); }
        private static void Sort(dynamic range, object fields, bool matchCase, bool headers) { var list=fields as IEnumerable; object first=null; if(list!=null) foreach(var x in list){first=x;break;} var map=first as Dictionary<string,object>; var key=map != null && map.ContainsKey("key") ? Convert.ToInt32(map["key"])+1 : 1; var ascending=map==null || !map.ContainsKey("ascending") || Bool(map["ascending"]); range.Sort(Key1:range.Columns[key],Order1:ascending?1:2,Header:headers?1:2,MatchCase:matchCase); }
        private static object AutoFillType(string value) { return value=="FillCopy" ? 2 : 0; }
        private static void CopyFrom(dynamic destination,dynamic source,object copyType,object skipBlanks,object transpose) { source.Copy(); destination.PasteSpecial(CopyType(copyType), -4142, Bool(skipBlanks), Bool(transpose)); ((dynamic)ExcelDnaUtil.Application).CutCopyMode=false; }
        private static object CopyType(object value) { var word=String(value); if(word=="Values") return -4163; if(word=="Formats") return -4122; if(word=="Formulas") return -4123; return -4104; }
        private static RemoveDuplicatesHandle RemoveDuplicates(dynamic range, object columns, bool header)
        {
            var source = columns as IEnumerable;
            if (source == null || columns is string)
                throw new DispatchException("protocol", "Duplicate columns must be an array.");
            var selected = new List<int>();
            foreach (var column in source)
            {
                var index = Convert.ToInt32(column, CultureInfo.InvariantCulture);
                if (index < 0 || index >= Convert.ToInt32(range.Columns.Count))
                    throw new DispatchException("protocol", "A duplicate column is outside the range.");
                selected.Add(index + 1);
            }
            if (selected.Count == 0)
                throw new DispatchException("protocol", "At least one duplicate column is required.");

            var rows = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            var seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            var removed = 0;
            var firstDataRow = header ? 2 : 1;
            const int chunkLimit = 4096;
            for (var chunkStart = firstDataRow; chunkStart <= rows; chunkStart += chunkLimit)
            {
                var chunkRows = Math.Min(chunkLimit, rows - chunkStart + 1);
                var columnValues = new List<object>();
                foreach (var column in selected)
                    columnValues.Add(ReadColumnChunk(range, chunkStart, chunkRows, column));
                for (var row = 0; row < chunkRows; row++)
                {
                    var keyParts = new List<string>();
                    foreach (var values in columnValues)
                    {
                        var key = DuplicateKey(BulkValue(values, chunkRows, 1, row, 0));
                        keyParts.Add(key.Length.ToString(CultureInfo.InvariantCulture) + ":" + key);
                    }
                    if (!seen.Add(string.Join("|", keyParts))) removed++;
                }
            }
            var comColumns = new object[selected.Count];
            for (var index = 0; index < selected.Count; index++) comColumns[index] = selected[index];
            range.RemoveDuplicates(comColumns, header ? 1 : 2);
            return new RemoveDuplicatesHandle(removed, seen.Count);
        }

        private static object ReadColumnChunk(dynamic range, int firstRow, int rows, int column)
        {
            dynamic cells = null;
            dynamic firstCell = null;
            dynamic chunk = null;
            try
            {
                cells = range.Cells;
                firstCell = cells[firstRow, column];
                chunk = firstCell.Resize[rows, 1];
                return chunk.Value2;
            }
            finally
            {
                ReleaseCom(chunk);
                ReleaseCom(firstCell);
                ReleaseCom(cells);
            }
        }

        private static object BulkValue(object values, int rows, int columns, int row, int column)
        {
            var matrix = values as Array;
            if (matrix == null)
            {
                if (rows != 1 || columns != 1)
                    throw new DispatchException("excel", "Excel returned a scalar for a multi-cell range.");
                return values;
            }
            if (matrix.Rank != 2 || matrix.GetLength(0) != rows || matrix.GetLength(1) != columns)
                throw new DispatchException("excel", "Excel returned an unexpected range matrix shape.");
            return matrix.GetValue(matrix.GetLowerBound(0) + row, matrix.GetLowerBound(1) + column);
        }

        private static string DuplicateKey(object value)
        {
            if (value == null || value is DBNull) return "empty";
            var error = value as ErrorWrapper;
            if (error != null) return "error:" + error.ErrorCode.ToString(CultureInfo.InvariantCulture);
            var text = value as string;
            if (text != null) return "string:" + text;
            if (value is bool) return "boolean:" + ((bool)value ? "1" : "0");
            if (value is double) return "number:" + ((double)value).ToString("R", CultureInfo.InvariantCulture);
            if (value is float) return "number:" + ((float)value).ToString("R", CultureInfo.InvariantCulture);
            if (value is decimal) return "number:" + ((decimal)value).ToString(CultureInfo.InvariantCulture);
            return value.GetType().FullName + ":" + Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static ReplaceHandle ReplaceAll(
            dynamic range,
            string find,
            string replacement,
            object criteria)
        {
            var map = criteria as Dictionary<string, object>;
            var completeMatch = map != null && map.ContainsKey("completeMatch") && Bool(map["completeMatch"]);
            var matchCase = map != null && map.ContainsKey("matchCase") && Bool(map["matchCase"]);
            var count = 0;
            dynamic current = range.Find(find, Type.Missing, -4123, completeMatch ? 1 : 2, 1, 1, matchCase, false, false);
            if (current == null) return new ReplaceHandle(0);
            var firstAddress = Convert.ToString(current.Address, CultureInfo.InvariantCulture);
            var limit = Convert.ToDouble(range.CountLarge, CultureInfo.InvariantCulture);
            try
            {
                while (count < limit)
                {
                    count++;
                    dynamic next = range.FindNext(current);
                    var nextAddress = next == null ? null : Convert.ToString(next.Address, CultureInfo.InvariantCulture);
                    ReleaseCom(current);
                    current = null;
                    if (next == null)
                        break;
                    if (string.Equals(nextAddress, firstAddress, StringComparison.Ordinal))
                    {
                        ReleaseCom(next);
                        break;
                    }
                    current = next;
                }
            }
            finally { ReleaseCom(current); }
            range.Replace(find, replacement, completeMatch ? 1 : 2, 1, matchCase, false, false, false);
            return new ReplaceHandle(count);
        }
        private static OpaqueHandle AddConditionalFormat(dynamic range,string type) { if(type=="CellValue") return new OpaqueHandle(range.FormatConditions.Add(1, 5, "0"),"conditional format"); if(type=="ColorScale") return new OpaqueHandle(range.FormatConditions.AddColorScale(3),"conditional format", range, 3); if(type=="DataBar") return new OpaqueHandle(range.FormatConditions.AddDatabar(),"conditional format"); throw new DispatchException("dispatch","Unsupported conditional format type "+type+"."); }
        private static object WorksheetProperty(dynamic sheet,string property) { if(property=="name") return Convert.ToString(sheet.Name); if(property=="visibility") return Visibility(sheet.Visible); if(property=="id") return Convert.ToString(sheet.CodeName); throw NoDispatch("items/"+property); }
        private static string Visibility(object value) { var n=Convert.ToInt32(value); return n==-1 ? "Visible" : n==2 ? "VeryHidden" : "Hidden"; }
        private List<object> LoadSheetTables(dynamic sheet, IEnumerable<string> properties)
        {
            var fields = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in properties)
            {
                if (!property.StartsWith("tables/items/", StringComparison.Ordinal)) throw NoDispatch(property);
                var field = property.Substring("tables/items/".Length);
                if (field != "name" && field != "showHeaders") throw NoDispatch(property);
                fields.Add(field);
            }
            var items = new List<object>();
            foreach (dynamic table in sheet.ListObjects)
            {
                var id = nextHostHandle--;
                handles[id] = new TableHandle(table);
                var item = new Dictionary<string, object> { { "id", id } };
                if (fields.Contains("name")) item["name"] = Convert.ToString(table.Name);
                if (fields.Contains("showHeaders")) item["showHeaders"] = Convert.ToBoolean(table.ShowHeaders);
                items.Add(item);
            }
            return items;
        }
        private List<object> LoadNames(dynamic workbook, IEnumerable<string> properties)
        {
            var fields = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in properties)
            {
                if (!property.StartsWith("names/items/", StringComparison.Ordinal)) throw NoDispatch(property);
                var field = property.Substring("names/items/".Length);
                if (field != "name" && field != "formula" && field != "scope") throw NoDispatch(property);
                fields.Add(field);
            }
            var items = new List<object>();
            foreach (dynamic named in workbook.Names)
            {
                if (NameScope(named) != "Workbook") continue;
                var id = nextHostHandle--;
                handles[id] = new NameHandle(named);
                var item = new Dictionary<string, object> { { "id", id } };
                if (fields.Contains("name")) item["name"] = Convert.ToString(named.Name);
                if (fields.Contains("formula")) item["formula"] = Convert.ToString(named.RefersTo);
                if (fields.Contains("scope")) item["scope"] = NameScope(named);
                items.Add(item);
            }
            return items;
        }
        private static List<object> LoadLinks(dynamic workbook, IEnumerable<string> properties)
        {
            foreach (var property in properties)
                if (property != "linkedWorkbooks/items/id") throw NoDispatch(property);
            var items = new List<object>();
            object links = workbook.LinkSources(1);
            var enumerable = links as IEnumerable;
            if (enumerable != null)
                foreach (var link in enumerable)
                    items.Add(new Dictionary<string, object> { { "id", String(link) } });
            return items;
        }
        private static string NameScope(dynamic named)
        {
            var name = Convert.ToString(named.Name, CultureInfo.InvariantCulture) ?? string.Empty;
            return name.IndexOf("!", StringComparison.Ordinal) >= 0 ? "Worksheet" : "Workbook";
        }
        private static Dictionary<string,object> Values(string[] properties,string key1,object value1,string key2,object value2) { var result=new Dictionary<string,object>(); foreach(var p in properties) { if(p==key1) result[p]=value1; else if(p==key2) result[p]=value2; else throw NoDispatch(p); } return result; }
        private static object Calculation(string word) { return word=="Manual" ? -4135 : word=="AutomaticExceptTables" ? 2 : -4105; }
        private static string CalculationMode(object value) { var n=Convert.ToInt32(value); return n==-4135 ? "Manual" : n==2 ? "AutomaticExceptTables" : "Automatic"; }
        private static object Horizontal(string word) { if(word=="Center") return -4108; if(word=="Right") return -4152; if(word=="Justify") return -4130; return -4131; }
        private static object Border(string word) { if(word=="EdgeTop") return 8; if(word=="EdgeBottom") return 9; if(word=="EdgeLeft") return 7; if(word=="EdgeRight") return 10; if(word=="InsideHorizontal") return 12; if(word=="InsideVertical") return 11; return 5; }
        private static object BorderStyle(string word) { if(word=="Continuous") return 1; if(word=="Dash") return -4115; if(word=="DashDot") return 4; if(word=="DashDotDot") return 5; if(word=="Dot") return -4118; if(word=="Double") return -4119; if(word=="None") return -4142; throw new DispatchException("protocol","Unknown border style "+word+"."); }
        private static int Color(object value) { var text=String(value); if(!text.StartsWith("#",StringComparison.Ordinal) || text.Length!=7) throw new DispatchException("protocol","Colors must be #RRGGBB."); return int.Parse(text.Substring(5,2),NumberStyles.HexNumber) * 65536 + int.Parse(text.Substring(3,2),NumberStyles.HexNumber) * 256 + int.Parse(text.Substring(1,2),NumberStyles.HexNumber); }
        private static object Paper(string word) { if(word=="A4") return 9; if(word=="Letter") return 1; if(word=="A3") return 8; if(word=="Legal") return 5; throw new DispatchException("protocol","Unknown paper size "+word+"."); }
        private static object Summary(string word) { if(word=="Average") return -4106; if(word=="Count") return -4112; if(word=="Max") return -4136; if(word=="Min") return -4139; if(word=="Product") return -4149; return -4157; }
        private static object ShowAs(string word) { if(word=="PercentOfGrandTotal") return 8; if(word=="PercentOfRowTotal") return 6; if(word=="PercentOfColumnTotal") return 7; if(word=="RunningTotal") return 5; return -4143; }
        private static object ValidationType(Dictionary<string,object> rule) { var type=String(RuleValue(rule,"type")); if(type=="List") return 3; if(type=="WholeNumber") return 1; if(type=="Decimal") return 2; if(type=="Date") return 4; if(type=="Time") return 5; if(type=="TextLength") return 6; if(type=="Custom") return 7; throw new DispatchException("protocol","Unknown validation type "+type+"."); }
        private static object ValidationOperator(Dictionary<string,object> rule) { var op=String(RuleValue(rule,"operator")); if(op=="Between") return 1; if(op=="NotBetween") return 2; if(op=="EqualTo") return 3; if(op=="NotEqualTo") return 4; if(op=="GreaterThan") return 5; if(op=="LessThan") return 6; if(op=="GreaterThanOrEqual") return 7; if(op=="LessThanOrEqual") return 8; return 1; }
        private static object RuleValue(Dictionary<string,object> rule,string key) { object value; return rule.TryGetValue(key,out value) ? value : Type.Missing; }
        private static object ChartType(string word) { if(word=="Bar" || word=="BarClustered") return 57; if(word=="Column" || word=="ColumnClustered") return 51; if(word=="Line") return 4; if(word=="Pie") return 5; if(word=="Area") return 1; if(word=="Scatter" || word=="XYScatter") return -4169; throw new DispatchException("protocol","Unknown chart type "+word+"."); }
        private static object ConditionOperator(string word) { if (word == "Between") return 1; if (word == "EqualTo") return 3; if (word == "GreaterThan") return 5; if (word == "GreaterThanOrEqual") return 7; if (word == "LessThan") return 6; if (word == "LessThanOrEqual") return 8; throw new DispatchException("protocol", "Unknown conditional-format operator " + word + "."); }
        private static object ColorScaleType(string word) { if (word == "LowestValue") return 1; if (word == "HighestValue") return 2; if (word == "Number") return 0; if (word == "Percent") return 3; if (word == "Formula") return 4; if (word == "Percentile") return 5; throw new DispatchException("protocol", "Unknown color-scale criterion type " + word + "."); }
        private string PivotFieldName(object value)
        {
            var reference = value as Dictionary<string, object>;
            if (reference == null || !reference.ContainsKey("handle")) return String(value);
            object target;
            var id = Convert.ToInt32(reference["handle"]);
            if (!handles.TryGetValue(id, out target) || !(target is OpaqueHandle))
                throw new DispatchException("protocol", "Unknown pivot hierarchy handle " + id + ".");
            dynamic field = ((OpaqueHandle)target).Value;
            return Convert.ToString(field.Name);
        }
    }
}
