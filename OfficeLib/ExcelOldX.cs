// Decompiled with JetBrains decompiler
// Type: OfficeLib.ExcelOldX
// Assembly: OfficeLib, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4D1E5114-8013-4B62-83A4-AF6178892448
// Assembly location: D:\Diploma\ImportStudentV2\bin\Debug\netcoreapp3.1\OfficeLib.dll

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.IO;
using System.Threading.Tasks;

namespace OfficeLib
{
  public class ExcelOldX
  {
    public HSSFWorkbook Document { get; private set; }

    protected DataFormatter Formatter { get; set; }

    public ExcelOldX(Stream stream = null)
    {
      if (stream != null)
        this.Document = new HSSFWorkbook(stream);
      this.Formatter = new DataFormatter();
    }

    public async Task Open(string path)
    {
      try
      {
                //byte[] bytes = await File.ReadAllBytesAsync(path);
                byte[] bytes = File.ReadAllBytes(path);
                this.Document = new HSSFWorkbook((Stream) new MemoryStream(bytes));
                bytes = (byte[]) null;
      }
      catch (Exception ex)
      {
        throw ex;
      }
    }

    public void Save(string path)
    {
      using (FileStream out1 = new FileStream(path, FileMode.OpenOrCreate))
        this.Document.Write((Stream) out1);
    }

    public void Save() => this.Save("test.xls");

    public void Merge(int startRow, int startCell, int endRow, int endCell, int numPage = 1)
    {
      try
      {
        this.Document.GetSheetAt(--numPage).AddMergedRegion(new CellRangeAddress(--startRow, --endRow, --startCell, --endCell));
      }
      catch
      {
      }
    }

    private ICell GetCell(int row, int cell, int numPage = 1)
    {
      try
      {
        return this.Document[--numPage].GetRow(--row).GetCell(--cell, MissingCellPolicy.CREATE_NULL_AS_BLANK);
      }
      catch
      {
        return (ICell) null;
      }
    }

    public DateTime? ReadDate(int row, int cell, int numPage = 1)
    {
      try
      {
        return new DateTime?(this.GetCell(row, cell, numPage).DateCellValue);
      }
      catch
      {
        return new DateTime?();
      }
    }

    public string Read(int row, int cell, int numPage = 1) => this.Formatter.FormatCellValue(this.GetCell(row, cell, numPage));

    public void Write(int row, int cell, int data, int numPage = 1) => this.Write(row, cell, data.ToString(), numPage);

    public void Write(int row, int cell, string data, int numPage = 1)
    {
      try
      {
        this.GetCell(row, cell, numPage).SetCellValue(data);
      }
      catch
      {
      }
    }

    public void Replace(string start, int end) => this.Replace(start, end.ToString());

    public void Replace(string start, int end, int numPage) => this.Replace(start, end.ToString(), numPage);

    public void Replace(string start, string end)
    {
      for (int index = 0; index < this.Document.Count; ++index)
        this.Replace(start, end, index + 1);
    }

    public void Find(string word)
    {
      for (int index1 = 0; index1 < this.Document.Count; ++index1)
      {
        int lastRowNum = this.Document[index1].LastRowNum;
        for (int index2 = 0; index2 < lastRowNum; ++index2)
        {
          int lastCellNum = (int) this.Document[index1].GetRow(index2).LastCellNum;
          for (int cell = 0; cell < lastCellNum; ++cell)
          {
            if (this.Read(index2, cell, index1) == word)
              Console.WriteLine("true");
          }
        }
      }
    }

    public void Replace(string start, string end, int numPage)
    {
      --numPage;
      for (int index = 0; index < this.Document[numPage].LastRowNum; ++index)
      {
        for (int cell = 0; cell < (int) this.Document[numPage].GetRow(index).LastCellNum; ++cell)
        {
          if (this.Read(index, cell, numPage) == start)
            this.Write(index, cell, end, numPage);
        }
      }
    }
  }
}
