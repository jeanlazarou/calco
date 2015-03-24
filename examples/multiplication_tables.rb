require 'date'
require 'tmpdir'

#
# Next example shows the use of the office engine, no headers and empty row.
#
# It saves an office file in the temporary directory named "res.ods" and uses
# a template file named "multiplication_tables.ods"
#

require 'calco/engines/office_engine'

output_file = File.join(Dir.tmpdir, "res.ods")

def relative_path file
  File.join(File.dirname(__FILE__), file)
end

engine = Calco::OfficeEngine.new(relative_path('multiplication_tables.ods'))

doc = spreadsheet(engine) do

  definitions do

    set multiplicand: 7
    set multiplier: 0
    
    function result: multiplicand * multiplier
    
  end
  
  sheet('x') do

    column value_of(:multiplier)
    column 'x'
    column value_of(:multiplicand)
    column '='
    
    column :result

  end
  
end

count = 0

doc.save(output_file) do |spreadsheet|
  
  sheet = doc.current

  7.step(10) do |multiplicand|
  
    sheet[:multiplicand] = multiplicand
    
    1.step(10) do |multiplier|
      
      sheet[:multiplier] = multiplier
      
      sheet.write_row count
      
      count += 1

    end
    
    count += 1
      
    sheet.empty_row
    
  end
  
end

puts "Wrote #{output_file} (#{count} rows)"
