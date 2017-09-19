using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RB10.Tools.ExcelBookCombination
{
    public static class CellExtension
    {
        public static string GetCellValue(this ICell cell)
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
    }
}
