import org.apache.commons.lang3.time.DateFormatUtils
import org.supercsv.cellprocessor.Optional
import org.supercsv.cellprocessor.ParseDate
import org.supercsv.cellprocessor.ParseDouble
import org.supercsv.cellprocessor.constraint.NotNull
import org.supercsv.cellprocessor.ift.CellProcessor
import org.supercsv.io.CsvMapReader
import org.supercsv.prefs.CsvPreference
import java.awt.Desktop
import java.io.File
import java.io.FileReader
import java.net.URI
import java.time.LocalDate
import java.time.format.DateTimeFormatter.ofPattern
import java.time.temporal.TemporalAdjusters.firstDayOfMonth
import java.time.temporal.TemporalAdjusters.lastDayOfMonth
import java.util.*
import java.util.Locale.getDefault

class X

fun main() {
    Locale.setDefault(Locale.GERMAN)

    val file = X::class.java.classLoader.getResourceAsStream("settings.properties")
    assert(file != null) { "settings.properties must exist on the class path!" }
    val properties = Properties()
    properties.load(file)
    assert(properties.getProperty("startHour").toDoubleOrNull() != null) { "startHour must be numeric" }

    val (dataFile, firstDayOfLastMonth) = createFileReference(properties.getProperty("dataFileLocation"))

    val groupedRecords = readRecords(dataFile)
    groupedRecords.forEach { group ->
        println("${formatDate(group.key)} ---------- ${group.value.sumOf { it.hours }}")
        group.value.forEach { issue -> println("${issue.issue}: ${issue.description}") }
        println()
    }

    val desktop = Desktop.getDesktop()
    val mailUri = properties.getProperty("mailURI")
    if (mailUri != null) {
        val year = firstDayOfLastMonth.format(ofPattern("yyyy"))
        val month = firstDayOfLastMonth.format(ofPattern("MMMM"))
        desktop.mail(URI.create(mailUri.replace("\$month", month).replace("\$year", year)))
    }
}

private fun createFileReference(dataFileName: String): Config {
    val currentDate = LocalDate.now()
    val firstDayOfLastMonth = currentDate.plusDays(12).with(firstDayOfMonth()).minusMonths(1)
    val lastDayOfLastMonth = firstDayOfLastMonth.with(lastDayOfMonth())
    val dayFormatter = ofPattern("yyyy-MM-dd")
    val start = firstDayOfLastMonth.format(dayFormatter)
    val end = lastDayOfLastMonth.format(dayFormatter)

    val dataFile = File(dataFileName.replace("\$start", start).replace("\$end", end))

    println("Running for month: ${firstDayOfLastMonth.format(ofPattern("MMMM yyyy"))}\n")

    assert(dataFile.exists()) { "data file does not exist!" }

    return Config(dataFile, firstDayOfLastMonth)
}

fun readRecords(dataFile: File): Map<Date, List<Record>> {
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
        Optional(),
        NotNull()
    )

    val records = mutableListOf<Record>()
    while (true) {
        val line = beanReader.read(header, processors) ?: break
        records.add(
            Record(
                line["Hours"] as Double,
                line["Work date"] as Date,
                line["Issue Key"] as String,
                line["Work Description"] as String
            )
        )
    }

    return records.groupBy { it.date }
}

private fun formatDate(date: Date): String = DateFormatUtils.format(date, "EEEE, dd.MM.yyyy").uppercase(getDefault())

data class Record(
    val hours: Double,
    val date: Date,
    val issue: String,
    val description: String
)

data class Config(
    val dataFile: File,
    val firstDayOfPeriod: LocalDate
)
