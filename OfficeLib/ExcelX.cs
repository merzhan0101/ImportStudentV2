// Decompiled with JetBrains decompiler
// Type: OfficeLib.ExcelX
// Assembly: OfficeLib, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 4D1E5114-8013-4B62-83A4-AF6178892448
// Assembly location: D:\Diploma\ImportStudentV2\bin\Debug\netcoreapp3.1\OfficeLib.dll

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Threading.Tasks;

namespace OfficeLib
{
  public class ExcelX
  {
    public XSSFWorkbook Document { get; private set; }

    protected DataFormatter Formatter { get; set; }

    public ExcelX(Stream stream = null)
    {
      if (stream != null)
        this.Document = new XSSFWorkbook(stream);
      this.Formatter = new DataFormatter();
    }

    public async Task Open(string path)
    {
      try
      {
        //byte[] bytes = await File.ReadAllBytesAsync(path);
        byte[] bytes = File.ReadAllBytes(path);
        this.Document = new XSSFWorkbook((Stream) new MemoryStream(bytes));
        bytes = (byte[]) null;
      }
      catch (Exception ex)
      {
        throw ex;
      }
    }

    public void Save(string path = "test.xls")
    {
      using (FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate))
        this.Document.Write((Stream) fileStream);
    }

    private ICell GetCell(int row, string cell, int numPage = 1)
    {
      string str = "abcdefghijklmnopqrstuvwxyz";
      int cell1 = 0;
      int length = cell.Length;
      if (length > 1)
        cell1 = 26 * (length - 1);
      char ch1 = cell[length - 1];
      foreach (char ch2 in str)
      {
        if ((int) ch1 != (int) ch2)
          ++cell1;
        else
          break;
      }
      return this.GetCell(row, cell1, numPage);
    }

    private ICell GetCell(int row, int cell, int numPage = 1)
    {
      try
      {
        return this.Document[--numPage].GetRow(--row).GetCell(--cell);
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

    public string Read(int row, string cell, int numPage = 1) => this.Formatter.FormatCellValue(this.GetCell(row, cell, numPage));

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
