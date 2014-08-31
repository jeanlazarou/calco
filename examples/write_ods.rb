require 'csv'
require 'date'
require 'tmpdir'

#
# Next example uses the use of the office engine and reads data from a CSV file.
#
# It saves an office file in the temporary directory named "res.ods" and uses
# a template file named "ages.ods"
#
# It shows the definition of a cell with a conditional style. Depending on the
# current value of the cell the style name "adult" or "default" is applied.
# The "adult" style should be defined in the template file.
#

require 'calco/engines/office_engine'

output_file = File.join(Dir.tmpdir, "res.ods")

def relative_path file
  File.join(File.dirname(__FILE__), file)
end

engine = Calco::OfficeEngine.new(relative_path('ages.ods'))

doc = spreadsheet(engine) do

  definitions do

    set name: ''
    set birth_date: ''
    
    function age: year(today) - year(birth_date)
    
  end
  
  sheet('Main') do

    has_titles true

    column value_of(:name)
    column value_of(:birth_date)
    
    column :age, style: _if(current > 19, '"adult"', '"default"')

  end
  
end

count = 0

doc.save(output_file) do |spreadsheet|
  
  sheet = doc.current
  
  CSV.foreach(relative_path('data.csv'), :headers => true) do |row|
    
    count += 1

    sheet[:name] = row[0]
    sheet[:birth_date] = Date.parse(row[1])
    
    sheet.write_row count
    
  end
  
end

puts "Wrote #{output_file} (#{count} rows)"
