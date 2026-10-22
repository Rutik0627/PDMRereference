PDM Re-Reference Add-in - Integrated Read-Only Dry Run

Included source files:
- ReReferenceAddin.cs: replace the current ReReferenceAddin.cs contents with this file.
- PdmDryRunRowProcessor.cs: add this new file to the same Visual Studio project.
- BatchPlanReader.cs and BatchRunCoordinator.cs: current workbook reader and coordinator versions for reference; if these already exist in the project, do NOT add duplicate copies.

Project requirements:
- .NET Framework 4.7.2 (or compatible version already used by the project)
- Reference EPDM.Interop.epdm.dll from the installed SOLIDWORKS PDM client
- System.Windows.Forms and System.Drawing references
- Microsoft.CSharp reference if Visual Studio does not add it automatically (the processor uses dynamic COM binding)
- Ensure the project compiles with x64/Any CPU settings compatible with the installed PDM client, as used by your current project.

Run only from the pdmdev vault for initial testing. The implementation uses the vault session that invoked the add-in; it does not log in to a named vault.

What the Excel dry run does:
1. Reads and validates the workbook rows.
2. Resolves source and replacement names in the invoking PDM vault. Names without extensions are searched as .SLDPRT and .SLDASM.
3. Flags missing or ambiguous names for review.
4. For uniquely resolved source files, enumerates parent references, ignores drawings, deduplicates parent assemblies, and records each parent's state and local path where available.
5. Writes a CSV report beside the workbook.

Safety boundary:
This is read-only inspection only. It does not check out files, change workflow states, open/save SOLIDWORKS documents, replace references, check files in, or rename/obsolete a source file. It does not yet compare reference versions/configurations or execute the VBA macro's state transitions. Those are later milestones after the report is verified in pdmdev.

Important:
PDM API late-bound member names can vary by client/interoperability version. If a row reports a PDM inspection error, retain the CSV and exact error text so the API call can be adjusted. Do not enable mutating operations as a workaround.
