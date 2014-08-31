require 'time'

require 'calco/engines/csv_engine'

#
# Next example presents uses the CSV engine that outputs formulas and save time 
# values as formulas. LibreOffice opens CSV files containing formulas and is 
# able to apply them.
#
# The example shows how to define an aggregation function (sum) using a Range
# object marked with the as_grouping method (extension added by Calco to the
# Ruby Range class).
#
# The example uses the Sheet#replace_and_clear method that changes the 
# definition set for some column.
#
# The spreadsheet is a kind ot timesheet.
#

doc = spreadsheet(Calco::CSVEngine.new) do

  definitions do

    set start_time: ''
    set end_time: ''

    function duration: end_time - start_time

    function total: sum(duration[(1..-1).as_grouping])

  end

  sheet do

    column value_of(:start_time), :title => "Start"
    column value_of(:end_time), :title => "End"

    column :duration, :title => "Duration", :id => :duration

  end

end

data = [
  {:start_time => Time.parse("12:10"), :end_time => Time.parse("15:30")},
  {:start_time => Time.parse("11:00"), :end_time => Time.parse("16:30")},
  {:start_time => Time.parse("10:01"), :end_time => Time.parse("12:05")},
]

doc.save($stdout) do |spreadsheet|

  sheet = doc.current

  sheet.write_row 0

  data.each_with_index do |entry, i|

    sheet.write_row i + 1, entry

  end

  sheet.empty_row
  
  sheet.replace_and_clear :duration => :total

  sheet.write_row data.length + 1

end
