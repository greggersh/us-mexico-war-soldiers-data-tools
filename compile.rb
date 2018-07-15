require 'csv'
require 'roo'
require 'roo-xls'

# Go through every spreadsheet in the data directory and concatenate the contents into a new CSV file.

all_rows = []
all_rows << ["Last Name", "First Name", "Rank", "Age", "Height", "Place of birth", "Occupation", "Enlistment Date", "Re-enlistment Date", "Enlistment Place or State of Birth", "Regiment", "Company", "Discharge Date", "Discharge Cause", "Date of Muster Out", "Muster Remarks", "General Remarks", "Unit", "Indexer", "Source of Data", "Title of Source", "Source File"]

Dir.glob('data/*.xls').each do |file|
  sheet = Roo::Spreadsheet.open(file, extension: :xls)
  sheet_info = [sheet.cell('B', 2), sheet.cell('B', 3), sheet.cell('B', 4), sheet.cell('B', 5), file.split("data/").last]
  first_data_row = 8
  last_data_row = sheet.last_row - 1
  first_data_row.upto(last_data_row) do |index|
    all_rows << sheet.row(index) + sheet_info
  end
end

CSV.open('./output.csv', 'wb') do |csv|
  all_rows.each do |row|
    csv << row
  end
end