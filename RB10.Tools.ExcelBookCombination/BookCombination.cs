using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RB10.Tools.ExcelBookCombination
{
    class BookCombination
    {
        private string _targetFolder;
        private string _outputParentFolder;

        public BookCombination(string targetFolder, string outputPrentFolder)
        {
            _targetFolder = targetFolder;
            _outputParentFolder = outputPrentFolder;
        }

        public void Run()
        {
            // 出力フォルダー作成
            var outputFolder = System.IO.Path.Combine(_outputParentFolder, System.IO.Path.GetFileName(_targetFolder));
            if(!System.IO.Directory.Exists(outputFolder))
            {
                System.IO.Directory.CreateDirectory(outputFolder);
            }
            else
            {
                foreach (var file in System.IO.Directory.GetFiles(outputFolder))
                {
                    try
                    {
                        var fileInfo = new System.IO.FileInfo(file);
                        fileInfo.Delete();
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            // ファイルを読み込む
            List<(string ColA, string ColB)> allLines = new List<(string ColA, string ColB)>();
            foreach (var file in System.IO.Directory.GetFiles(_targetFolder, "*.xls*", System.IO.SearchOption.TopDirectoryOnly))
            {
                allLines.AddRange(ReadFile(file));
            }

            // A列の重複を削除
            var targetLines = new List<(string ColA, string ColB)>();
            foreach (var group in allLines.GroupBy(x => x.ColA))
            {
                targetLines.Add(group.First());
            }

            // ①のファイル用に加工する
            DataSet1.DataTable1DataTable dt1 = EditFor1(targetLines);

            // ②のファイル用に加工する
            DataSet1.DataTable2DataTable dt2 = EditFor2(targetLines);

            // ファイル保存
            SaveFile(Path.Combine(outputFolder, targetLines.Count.ToString() + ".xlsx"), targetLines);

            // ①のファイルを保存する
            SaveTsvFile(Path.Combine(outputFolder, "①.tsv"), dt1);

            // ②のファイルを保存する
            SaveTsvFile(Path.Combine(outputFolder, "②.tsv"), dt2);
        }

        private List<(string ColA, string ColB)> ReadFile(string file)
        {
            IWorkbook book;
            using (FileStream infile = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // エクセルファイルを開いて内容を取得（ワークブックオブジェクトを作成）
                book = WorkbookFactory.Create(infile, ImportOption.All);
            }

            var ret = new List<(string ColA, string ColB)>();
            for (int i = 0; i < book.NumberOfSheets; i++)
            {
                // シートインデックスを指定してシートオブジェクトを取得
                ISheet sheet = book.GetSheetAt(i);

                // シートで使用されている行の最終行インデックスを取得
                for (int j = 0; j <= sheet.LastRowNum; j++)
                {
                    // 行オブジェクトを取得
                    IRow row = sheet.GetRow(j);
                    if (row == null) continue;

                    string colA = row.GetCell(0).GetCellValue();
                    string colB = row.GetCell(1).GetCellValue();
                    
                    if (string.IsNullOrEmpty(colA) && string.IsNullOrEmpty(colB)) continue;

                    var values = (colA, colB);
                    ret.Add(values);
                }
            }

            return ret;
        }

        private string GetCellValue(ICell cell)
        {
            string cellStr = string.Empty;
            switch (cell.CellType)
            {
                // 文字列型
                case CellType.String:
                    cellStr = cell.StringCellValue;
                    break;
                // 数値型（日付の場合もここに入る）
                case CellType.Numeric:
                    // セルが日付情報が単なる数値かを判定
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        // 日付型
                        // 本来はスタイルに合わせてフォーマットすべきだが、
                        // うまく表示できないケースが若干見られたので固定のフォーマットとして取得
                        cellStr = cell.DateCellValue.ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    else
                    {
                        // 数値型
                        cellStr = cell.NumericCellValue.ToString();
                    }
                    break;

                // bool型(文字列でTrueとか入れておけばbool型として扱われた)
                case CellType.Boolean:
                    cellStr = cell.BooleanCellValue.ToString();
                    break;

                // 入力なし
                case CellType.Blank:
                    cellStr = cell.ToString();
                    break;

                // 数式
                case CellType.Formula:
                    // 下記で数式の文字列が取得される
                    //cellStr = cell.CellFormula.ToString();

                    // 数式の元となったセルの型を取得して同様の処理を行う
                    // コメントは省略
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            cellStr = cell.StringCellValue;
                            break;
                        case CellType.Numeric:

                            if (DateUtil.IsCellDateFormatted(cell))
                            {
                                cellStr = cell.DateCellValue.ToString("yyyy/MM/dd HH:mm:ss");
                            }
                            else
                            {
                                cellStr = cell.NumericCellValue.ToString();
                            }
                            break;
                        case CellType.Boolean:
                            cellStr = cell.BooleanCellValue.ToString();
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Error:
                            cellStr = cell.ErrorCellValue.ToString();
                            break;
                        case CellType.Unknown:
                            break;
                        default:
                            break;
                    }
                    break;

                // エラー
                case CellType.Error:
                    cellStr = cell.ErrorCellValue.ToString();
                    break;

                // 型不明なセル
                case CellType.Unknown:
                    break;
                // もっと不明なセル（あぶない刑事をなぜか思い出しました）
                default:
                    break;
            }

            return cellStr;
        }

        private DataSet1.DataTable1DataTable EditFor1(List<(string ColA, string ColB)> targetLines)
        {
            var dt = new DataSet1.DataTable1DataTable();

            foreach (var line in targetLines)
            {
                var row = dt.NewDataTable1Row();

                row.DataColumnA = line.ColA;
                row.DataColumnB = line.ColB;
                row.DataColumnD = "3";
                row.DataColumnE = line.ColA;
                row.DataColumnF = "AAAAA";
                row.DataColumnG = "New";
                row.DataColumnH = "XXXXX";
                row.DataColumnP = "5";

                dt.AddDataTable1Row(row);
            }

            return dt;
        }

        private DataSet1.DataTable2DataTable EditFor2(List<(string ColA, string ColB)> targetLines)
        {
            var dt = new DataSet1.DataTable2DataTable();

            foreach (var line in targetLines)
            {
                var row = dt.NewDataTable2Row();

                row.DataColumnA = line.ColA;
                row.DataColumnG = "PartialUpdate";
                if(decimal.TryParse(line.ColB, out var value))
                {
                    row.DataColumnN = (value - 5000).ToString();
                    row.DataColumnO = (System.Math.Round(value, 1)).ToString();
                }
                else
                {
                    row.DataColumnN = "N/A";
                    row.DataColumnO = "N/A";
                }

                dt.AddDataTable2Row(row);
            }

            return dt;
        }

        private void SaveFile(string fileName, List<(string ColA, string ColB)> targetLines)
        {
            var book = new XSSFWorkbook();
            var sheet = book.CreateSheet("Sheet1");

            int rowIndex = 0;
            foreach (var item in targetLines)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                var cell1 = row.GetCell(0) ?? row.CreateCell(0);
                var cell2 = row.GetCell(1) ?? row.CreateCell(1);

                cell1.SetCellValue(item.ColA);
                cell2.SetCellValue(item.ColB);

                rowIndex++;
            }

            using (var fs = new FileStream(fileName, FileMode.Create))
            {
                book.Write(fs);
            }
        }

        private void SaveTsvFile(string fileName, DataTable dt)
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataRow row in dt.Rows)
            {
                var values = new List<string>();
                foreach (DataColumn col in dt.Columns)
                {
                    values.Add(row.IsNull(col.ColumnName) ? "" : row.Field<string>(col.ColumnName));
                }

                sb.AppendLine(string.Join("\t", values));
            }

            System.IO.File.WriteAllText(fileName, sb.ToString());
        }
    }
}
