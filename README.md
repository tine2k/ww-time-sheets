# Time Sheet Generator

- reads a CSV file exported from Jira Tempo
- aggregates all work items by day and concatenates the notes to a single string
- writes the aggregated items to proprietary xls template file
- opens the generated file in excel and its folder
- (optional) opens the platform's email client with a predefined email message

# How to run
- export the time sheet of last month to `dataFileLocation` (e.g. your browser's downloads folder)
- open projekt in your favourite IDE (= IntelliJ IDEA)
- copy `settings.properties.sample` to `settings.properties` and configure your setup
- run `Main.kt`
