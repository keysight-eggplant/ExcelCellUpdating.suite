Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\durgaprasad.surath\OneDrive - Ideabytes\Documents\Eggplant\EggplantSuites\Personal\RandD\ExcelCellUpdating.suite\Resources\SampleSheetExcelWrite.xlsm")
objExcel.Run "RecalculateCells"
objWorkbook.Close True
objExcel.Quit
