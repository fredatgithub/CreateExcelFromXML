namespace XmlToExcel
{
  internal class Issue
  {
    //<Issue TypeId="InconsistentNaming" File="file1.cs" Offset="265-275" Line="12" Message="Name 'XMLService' does not match rule 'Types and namespaces'. Suggested name is 'XmlService'." />
    public string TypeId { get; set; }
    public string FileName { get; set; }
    public string Offset { get; set; }
    public int LineNumber { get; set; }
    public string Message { get; set; }

    public Issue(string typeId, string fileName, string offset, int lineNumber, string message)
    {
      TypeId = typeId;
      FileName = fileName;
      Offset = offset;
      LineNumber = lineNumber;
      Message = message;
    }
  }
}