using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

public static class MyFunctions
{
    [ExcelFunction(Description = "Counts the number of cells with the specified background color.")]
    public static int COUNT_IF_COLOR(
        [ExcelArgument(Description = "The range of cells to evaluate.")] ExcelReference rng,
        [ExcelArgument(Description = "The background color to search for.")] int color)
    {
        int count = 0;
        object[,] values = (object[,])rng.GetValue();

        for (int row = 1; row <= values.GetLength(0); row++)
        {
            for (int col = 1; col <= values.GetLength(1); col++)
            {
                ExcelReference cellRef = new ExcelReference(row, col, row, col, rng.SheetId);
                if (CellColorEquals(cellRef, color))
                {
                    count++;
                }
            }
        }

        return count;
    }

    // Helper function to compare cell color with the given color
    private static bool CellColorEquals(ExcelReference cellRef, int colorToMatch)
    {
        Application excelApp = (Application)ExcelDnaUtil.Application;
        Range cell = excelApp.ActiveSheet.Range[cellRef.ToA1String()];
        int cellColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(cell.Interior.Color));
        return cellColor == colorToMatch;
    }
}
