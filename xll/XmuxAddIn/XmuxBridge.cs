using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
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

        /*
         * The pane calls `handshake` and `execute` in lower camel case, because that is what
         * every other object it touches looks like. Whether WebView2's host-object proxy
         * resolves a COM member case-insensitively is not something this machine can find
         * out, and being wrong about it costs a whole session on the target PC: every call
         * rejects, `startBridgeHost` reports no host, and the pane says it only works in
         * Excel while sitting inside Excel.
         *
         * So both spellings exist. `AutoDual` publishes each public method under its own
         * name, the forwarders cost nothing, and the coin flip is gone. Delete them once a
         * real machine has shown which casing arrives.
         */
        public string handshake()
        {
            return Handshake();
        }

        public string execute(string opsJson)
        {
            return Execute(opsJson);
        }

        public string Handshake()
        {
            try
            {
                var workbookUrl = RunInExcel(delegate
                {
                    dynamic application = ExcelDnaUtil.Application;
                    dynamic workbook = application.ActiveWorkbook;
                    return workbook == null ? string.Empty : Convert.ToString(workbook.FullName);
                });
                return json.Serialize(new Dictionary<string, object>
                {
                    { "workbookUrl", workbookUrl },
                    { "capabilities", new object[0] }
                });
            }
            catch (Exception exception)
            {
                return Failure(new Dictionary<string, object>(), "excel", exception.Message);
            }
        }

        public string Execute(string opsJson)
        {
            lock (gate)
            {
                var values = new Dictionary<string, object>();
                try
                {
                    var ops = json.Deserialize<List<BridgeOperation>>(opsJson);
                    if (ops == null)
                    {
                        throw new ArgumentException("The operation batch is null.");
                    }

                    RunInExcel(delegate
                    {
                        foreach (var operation in ops)
                        {
                            Dispatch(operation, values);
                        }
                        return true;
                    });
                    return json.Serialize(new Dictionary<string, object> { { "values", values } });
                }
                catch (DispatchException exception)
                {
                    return Failure(values, exception.Code, exception.Message);
                }
                catch (Exception exception)
                {
                    return Failure(values, "excel", exception.Message);
                }
            }
        }

        private void Dispatch(BridgeOperation operation, Dictionary<string, object> values)
        {
            if (operation == null)
            {
                throw new DispatchException("protocol", "The operation is null.");
            }

            object target;
            if (!handles.TryGetValue(operation.On, out target))
            {
                throw new DispatchException("protocol", "Unknown handle " + operation.On + ".");
            }

            if (operation.Op == "call")
            {
                var result = Call(target, operation.Member, operation.Args);
                handles[operation.Id] = result;
                return;
            }

            if (operation.Op == "load")
            {
                values[operation.On.ToString()] = Load(target, operation.Properties);
                return;
            }

            throw new DispatchException("protocol", "Unknown operation " + operation.Op + ".");
        }

        private static object Call(object target, string member, object[] args)
        {
            dynamic application = ExcelDnaUtil.Application;
            if (target is WorkbookHandle && member == "worksheets")
            {
                return new WorksheetsHandle(application.ActiveWorkbook.Worksheets);
            }
            if (target is WorksheetsHandle && member == "getItem")
            {
                return new WorksheetHandle(((dynamic)target).Worksheets.Item[Convert.ToString(Argument(args, 0))]);
            }
            if (target is WorksheetHandle && member == "getRange")
            {
                return new RangeHandle(((dynamic)target).Worksheet.Range[Convert.ToString(Argument(args, 0))]);
            }

            throw NoDispatch(member);
        }

        private static Dictionary<string, object> Load(object target, string[] properties)
        {
            if (!(target is RangeHandle))
            {
                throw NoDispatch(properties == null || properties.Length == 0 ? "" : properties[0]);
            }

            var loaded = new Dictionary<string, object>();
            foreach (var property in properties ?? new string[0])
            {
                dynamic range = ((RangeHandle)target).Range;
                if (property == "address")
                {
                    loaded[property] = Convert.ToString(range.Address);
                }
                else if (property == "text")
                {
                    loaded[property] = JsonValue(range.Text);
                }
                else if (property == "formulas")
                {
                    loaded[property] = JsonValue(range.Formula);
                }
                else
                {
                    throw NoDispatch(property);
                }
            }
            return loaded;
        }

        private static object Argument(object[] args, int index)
        {
            if (args == null || args.Length <= index)
            {
                throw new DispatchException("protocol", "Missing argument " + index + ".");
            }
            return args[index];
        }

        private static object JsonValue(object value)
        {
            var array = value as Array;
            if (array == null)
            {
                return value;
            }

            var rows = new List<object>();
            for (var row = array.GetLowerBound(0); row <= array.GetUpperBound(0); row++)
            {
                var columns = new List<object>();
                for (var column = array.GetLowerBound(1); column <= array.GetUpperBound(1); column++)
                {
                    columns.Add(JsonValue(array.GetValue(row, column)));
                }
                rows.Add(columns);
            }
            return rows;
        }

        private T RunInExcel<T>(Func<T> action)
        {
            T result = default(T);
            Exception failure = null;
            using (var completed = new ManualResetEventSlim(false))
            {
                ExcelAsyncUtil.QueueAsMacro(delegate
                {
                    try
                    {
                        result = action();
                    }
                    catch (Exception exception)
                    {
                        failure = exception;
                    }
                    finally
                    {
                        completed.Set();
                    }
                });

                if (!completed.Wait(TimeSpan.FromSeconds(30)))
                {
                    throw new TimeoutException("Excel did not enter macro context within 30 seconds.");
                }
            }

            if (failure != null)
            {
                throw failure;
            }
            return result;
        }

        private string Failure(Dictionary<string, object> values, string code, string message)
        {
            return json.Serialize(new Dictionary<string, object>
            {
                { "values", values },
                { "failure", new Dictionary<string, string> { { "code", code }, { "message", message } } }
            });
        }

        private static DispatchException NoDispatch(string member)
        {
            return new DispatchException("dispatch", "no dispatch for \"" + member + "\"");
        }

        private sealed class WorkbookHandle
        {
            internal static readonly WorkbookHandle Instance = new WorkbookHandle();
        }

        private sealed class WorksheetsHandle
        {
            internal dynamic Worksheets { get; private set; }
            internal WorksheetsHandle(dynamic worksheets) { Worksheets = worksheets; }
        }

        private sealed class WorksheetHandle
        {
            internal dynamic Worksheet { get; private set; }
            internal WorksheetHandle(dynamic worksheet) { Worksheet = worksheet; }
        }

        private sealed class RangeHandle
        {
            internal dynamic Range { get; private set; }
            internal RangeHandle(dynamic range) { Range = range; }
        }

        private sealed class DispatchException : Exception
        {
            internal string Code { get; private set; }
            internal DispatchException(string code, string message) : base(message) { Code = code; }
        }

        public sealed class BridgeOperation
        {
            public string Op { get; set; }
            public int Id { get; set; }
            public int On { get; set; }
            public string Member { get; set; }
            public object[] Args { get; set; }
            public string[] Properties { get; set; }
        }
    }
}
