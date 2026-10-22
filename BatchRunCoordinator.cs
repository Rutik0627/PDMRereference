using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace PDMReReferenceAddin
{
    public sealed class BatchRowResult
    {
        public int ExcelRowNumber { get; set; }
        public string Status { get; set; }
        public string SourcePart { get; set; }
        public string ReplacementPart { get; set; }
        public string Message { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime FinishedAt { get; set; }
    }

    public sealed class BatchProgressEventArgs : EventArgs
    {
        public int Completed { get; private set; }
        public int Total { get; private set; }
        public BatchRowResult Result { get; private set; }
        public string Message { get; private set; }

        public BatchProgressEventArgs(int completed, int total,
            BatchRowResult result, string message)
        {
            Completed = completed;
            Total = total;
            Result = result;
            Message = message;
        }
    }

    // Implement this interface later with the PDM/ SOLIDWORKS operations.
    // Keep all mutating operations out of the coordinator itself.
    public interface IReReferenceRowProcessor
    {
        BatchRowResult Process(BatchPlanRow row, CancellationToken cancellationToken);
    }

    // Safe processor for development without access to a test vault.
    // It validates the row and records that no PDM operation was attempted.
    public sealed class DryRunRowProcessor : IReReferenceRowProcessor
    {
        public BatchRowResult Process(BatchPlanRow row, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            BatchRowResult result = new BatchRowResult
            {
                ExcelRowNumber = row.ExcelRowNumber,
                SourcePart = row.SourcePart,
                ReplacementPart = row.ReplacementPart,
                StartedAt = DateTime.Now,
                FinishedAt = DateTime.Now
            };

            if (!String.Equals(row.ValidationStatus, "VALID", StringComparison.OrdinalIgnoreCase))
            {
                result.Status = "REVIEW";
                result.Message = String.IsNullOrWhiteSpace(row.ValidationMessage)
                    ? "Row failed workbook validation; skipped."
                    : row.ValidationMessage;
            }
            else
            {
                result.Status = "DRY_RUN_ONLY";
                result.Message = "Validated only. No PDM files or references were changed.";
            }
            return result;
        }
    }

    public sealed class BatchRunCoordinator
    {
        public event EventHandler<BatchProgressEventArgs> ProgressChanged;

        public IList<BatchRowResult> Run(IList<BatchPlanRow> rows,
            IReReferenceRowProcessor processor, string logPath,
            CancellationToken cancellationToken)
        {
            if (rows == null) throw new ArgumentNullException("rows");
            if (processor == null) throw new ArgumentNullException("processor");
            if (String.IsNullOrWhiteSpace(logPath)) throw new ArgumentException("A log path is required.", "logPath");

            List<BatchRowResult> results = new List<BatchRowResult>();
            string directory = Path.GetDirectoryName(Path.GetFullPath(logPath));
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);

            using (StreamWriter log = new StreamWriter(logPath, false, new UTF8Encoding(true)))
            {
                log.WriteLine("Timestamp,ExcelRow,Status,SourcePart,ReplacementPart,StartedAt,FinishedAt,Message");
                log.Flush();

                for (int i = 0; i < rows.Count; i++)
                {
                    BatchRowResult result;
                    if (cancellationToken.IsCancellationRequested)
                    {
                        result = new BatchRowResult
                        {
                            ExcelRowNumber = rows[i].ExcelRowNumber,
                            SourcePart = rows[i].SourcePart,
                            ReplacementPart = rows[i].ReplacementPart,
                            Status = "CANCELLED_NOT_STARTED",
                            Message = "Cancellation requested; this row was not processed.",
                            StartedAt = DateTime.MinValue,
                            FinishedAt = DateTime.Now
                        };
                    }
                    else
                    {
                        try
                        {
                            result = processor.Process(rows[i], cancellationToken);
                            if (result == null)
                                throw new InvalidOperationException("Processor returned no result.");
                        }
                        catch (OperationCanceledException)
                        {
                            result = new BatchRowResult
                            {
                                ExcelRowNumber = rows[i].ExcelRowNumber,
                                SourcePart = rows[i].SourcePart,
                                ReplacementPart = rows[i].ReplacementPart,
                                Status = "CANCELLED",
                                Message = "Processing was cancelled.",
                                StartedAt = DateTime.Now,
                                FinishedAt = DateTime.Now
                            };
                        }
                        catch (Exception ex)
                        {
                            result = new BatchRowResult
                            {
                                ExcelRowNumber = rows[i].ExcelRowNumber,
                                SourcePart = rows[i].SourcePart,
                                ReplacementPart = rows[i].ReplacementPart,
                                Status = "ERROR",
                                Message = ex.GetType().Name + ": " + ex.Message,
                                StartedAt = DateTime.Now,
                                FinishedAt = DateTime.Now
                            };
                        }
                    }

                    results.Add(result);
                    log.WriteLine(String.Join(",", new string[] {
                        Csv(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")),
                        Csv(result.ExcelRowNumber.ToString()), Csv(result.Status),
                        Csv(result.SourcePart), Csv(result.ReplacementPart),
                        Csv(result.StartedAt == DateTime.MinValue ? "" : result.StartedAt.ToString("yyyy-MM-dd HH:mm:ss")),
                        Csv(result.FinishedAt == DateTime.MinValue ? "" : result.FinishedAt.ToString("yyyy-MM-dd HH:mm:ss")),
                        Csv(result.Message)
                    }));
                    log.Flush();
                    EventHandler<BatchProgressEventArgs> handler = ProgressChanged;
                    if (handler != null)
                        handler(this, new BatchProgressEventArgs(i + 1, rows.Count, result, result.Message));
                }
            }
            return results.AsReadOnly();
        }

        private static string Csv(string value)
        {
            value = value ?? "";
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
