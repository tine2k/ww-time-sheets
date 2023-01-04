import org.apache.commons.lang3.time.DateFormatUtils
import org.apache.poi.ss.usermodel.*
import org.supercsv.cellprocessor.Optional
import org.supercsv.cellprocessor.ParseDate
import org.supercsv.cellprocessor.ParseDouble
import org.supercsv.cellprocessor.constraint.NotNull
import org.supercsv.cellprocessor.ift.CellProcessor
import org.supercsv.io.CsvMapReader
import org.supercsv.prefs.CsvPreference
import java.awt.Desktop
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.FileReader
import java.math.RoundingMode
import java.net.URI
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.temporal.TemporalAdjusters.firstDayOfMonth
import java.time.temporal.TemporalAdjusters.lastDayOfMonth
import java.util.*

fun main() {
    Locale.setDefault(Locale.GERMAN)

    val (dataFile, templateFile, outputFile, firstDayOfLastMonth) = createFileReferences(
        File("/Users/tine2k/Downloads"),
        File("/Users/tine2k/Downloads/out")
    )

    val groupedRecords = readRecords(dataFile)

    writeToXls(groupedRecords, templateFile, outputFile)

    val desktop = Desktop.getDesktop()
    desktop.open(outputFile)

    desktop.open(outputFile.parentFile)

    val year = firstDayOfLastMonth.format(DateTimeFormatter.ofPattern("yyyy"))
    val month = firstDayOfLastMonth.format(DateTimeFormatter.ofPattern("MMMM"))
    desktop.mail(URI.create("mailto:Invoicing@convista.com?subject=Zeiterfassung%20${month}%20${year}&body=Liebes%20Invoicing-Team%2C%0D%0A%0D%0Aanbei%20die%20Rechnung%20f%C3%BCr%20den%20Zeitraum%20${month}%20${year}.%20Vielen%20Dank.%0D%0A%0D%0AMit%20freundlichen%20Gr%C3%BC%C3%9Fen%0D%0AMartin%20Maier-Moessner"))
}

private fun createFileReferences(inputFolder: File, outputFolder: File): Config {
    val currentDate = LocalDate.now()
    val firstDayOfLastMonth = currentDate.with(firstDayOfMonth()).minusMonths(1)
    val lastDayOfLastMonth = currentDate.with(lastDayOfMonth()).minusMonths(1)
    val monthYear = firstDayOfLastMonth.format(DateTimeFormatter.ofPattern("MMyyyy"))
    val dayFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd")
    val start = firstDayOfLastMonth.format(dayFormatter)
    val end = lastDayOfLastMonth.format(dayFormatter)

    val dataFile = File(inputFolder, "User_ Maier-Moessner Martin_${start}_${end}.csv")
    val templateFileName = "ConVista ZE_8902_Martin Maier-Moessner_${monthYear}.xlsx"
    val templateFile = File(inputFolder, templateFileName)
    val outputFile = File(outputFolder, templateFileName)

    println("Running for month: ${firstDayOfLastMonth.format(DateTimeFormatter.ofPattern("MMMM yyyy"))}")

    assert(dataFile.exists()) { "data file does not exist!" }
    assert(templateFile.exists()) { "template file does not exist!" }
    assert(!outputFile.exists()) { "output file can not exist!" }

    return Config(dataFile, templateFile, outputFile, firstDayOfLastMonth)
}

fun readRecords(dataFile: File): List<Record> {
    val beanReader = CsvMapReader(FileReader(dataFile), CsvPreference.STANDARD_PREFERENCE)

    val header = beanReader.getHeader(true)
    val processors = arrayOf<CellProcessor>(
        Optional(),
        Optional(),
        ParseDouble(),
        ParseDate("yyyy-MM-dd"),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        Optional(),
        NotNull()
    );

    val records = mutableListOf<Record>()
    while (true) {
        val line = beanReader.read(header, processors) ?: break
        records.add(Record(line["Hours"] as Double, line["Work date"] as Date, line["Work Description"] as String))
    }

    return records.groupingBy { it.date }
        .reduce { _, accumulator, element -> accumulator + element }.values.sortedBy { it.date }
}

fun writeToXls(records: Collection<Record>, templateFile: File, outputFile: File) {
    FileInputStream(templateFile).use { inp ->
        val wb: Workbook = WorkbookFactory.create(inp)
        val sheet: Sheet = wb.getSheetAt(0)

        // get cell styles if first column is predefined
        val dateCellStyle = sheet.getRow(7).getCell(1).cellStyle
        val timeCellStyle = sheet.getRow(7).getCell(2).cellStyle

        // otherwise set these
//        val dateCellStyle = wb.createCellStyle()
//        dateCellStyle.locked = false
//        dateCellStyle.dataFormat = 16
//        val timeCellStyle = wb.createCellStyle()
//        timeCellStyle.locked = false
//        timeCellStyle.dataFormat = 20

        // set records
        records.forEachIndexed { index, record ->
            val row: Row = sheet.getRow(index + 7)
            fillCell(row, 1, record.date, dateCellStyle)
            fillCell(row, 2, formatDate(hoursAsDate(8.0)), timeCellStyle)
            fillCell(row, 3, formatDate(hoursAsDate(8.0 + record.hours)), timeCellStyle)
            fillCell(row, 4, null, timeCellStyle)
            fillCell(row, 7, 672002.0)
            fillCell(row, 9, record.description)
            fillCell(row, 10, "Standard")
            fillCell(row, 11, "02")
        }

        // recalc all formulas
        val createHelper = wb.creationHelper
        val evaluator = createHelper.createFormulaEvaluator()
        evaluator.clearAllCachedResultValues()
        evaluator.evaluateAll()

        // hour assertion
        val cellValue = round(sheet.getRow(59).getCell(5).numericCellValue)
        val recordValue = round(records.sumOf { it.hours })
        assert(cellValue == recordValue) { "Hour count must be equal! ($cellValue != $recordValue)" }

        FileOutputStream(outputFile).use { fileOut -> wb.write(fileOut) }
    }
}

fun round(d: Double): Double = d.toBigDecimal().setScale(2, RoundingMode.HALF_UP).toDouble()

fun hoursAsDate(number: Double): Date {
    return Date((number * 60 * 60 * 1000).toLong())
}

fun fillCell(row: Row, i: Int, value: String?) {
    val cell: Cell = row.getCell(i) ?: row.createCell(i)
    cell.setCellValue(value)
}

fun fillCell(row: Row, i: Int, value: Double) {
    val cell: Cell = row.getCell(i) ?: row.createCell(i)
    cell.setCellValue(value)
}

fun fillCell(row: Row, i: Int, value: Date?, cellStyle: CellStyle) {
    val cell: Cell = row.getCell(i) ?: row.createCell(i)
    cell.cellStyle = cellStyle
    cell.setCellValue(value)
}

fun fillCell(row: Row, i: Int, value: Double, cellStyle: CellStyle) {
    val cell: Cell = row.getCell(i) ?: row.createCell(i)
    cell.setCellValue(value)
    cell.cellStyle = cellStyle
}

fun formatDate(number: Date): Double {
    return DateUtil.convertTime(DateFormatUtils.format(number, "HH:mm:ss"))
}

data class Record(
    val hours: Double,
    val date: Date,
    val description: String
) {
    operator fun plus(increment: Record): Record {
        return Record(hours + increment.hours, date, description + "; " + increment.description)
    }
}

data class Config(
    val dataFile: File,
    val templateFile: File,
    val outputFile: File,
    val firstDayOfPeriod: LocalDate
)
