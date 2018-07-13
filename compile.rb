require 'csv'
require 'roo'
require 'roo-xls'

# Go through every spreadsheet in the data directory and concatenate the contents into a new CSV file.

all_rows = []

Dir.glob('data/*.xls').each do |file|
  sheet = Roo::Spreadsheet.open(file, extension: :xls)
  first_data_row = 8
  last_data_row = sheet.last_row - 1
  first_data_row.upto(last_data_row) do |index|
    all_rows << sheet.row(index)
  end
end

CSV.open('./output.csv', 'wb') do |csv|
  all_rows.each do |row|
    csv << row
  end
end
