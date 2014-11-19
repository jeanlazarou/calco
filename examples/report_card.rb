require 'date'
require 'tmpdir'

#
# Next example uses the office engine.
#
# It saves an office file, in the temporary directory, named "res.ods" and uses
# a template file named "report_cards.ods"
#
# Shows applying styles
#
# The example fills an academic report card, such a report contains students'
# grades for different disciplines. Each sheet cover a different period.
#
# Here are the grades:
#    A -> Excellent
#    B -> Very Good
#    C -> Good
#    D -> Acceptable
#    F -> Fail
#

require 'calco/engines/office_engine'

output_file = File.join(Dir.tmpdir, "res.ods")

def relative_path file
  File.join(File.dirname(__FILE__), file)
end

engine = Calco::OfficeEngine.new(relative_path('report_cards.ods'))

doc = spreadsheet(engine) do

  definitions do

    set name: ''
    
    set art: ''
    set english: ''
    set geography: ''
    set history: ''
    set math: ''
    set music: ''
    set physics: ''
    
  end
  
  sheet('Period 1') do

    has_titles true
    
    column value_of(:name)
    
    highlight_f = _if(current == '"F"', '"fail"', '"default"')
    
    column value_of(:art), style: highlight_f
    column value_of(:english), style: highlight_f
    column value_of(:geography), style: highlight_f
    column value_of(:history), style: highlight_f
    column value_of(:history), style: highlight_f
    column value_of(:math), style: highlight_f
    column value_of(:music), style: highlight_f
    column value_of(:physics), style: highlight_f
    
  end
  
end

Strudent = Struct.new(:name, :art, :english, :geography, :history, :math, :music, :physics)

period_1 = [
  Strudent.new('Alan', 'A', 'B', 'C', 'D', 'B', 'B', 'B'),
  Strudent.new('Greg', 'A', 'B', 'C', 'C', 'F', 'A', 'F'),
  Strudent.new('Iñes', 'A', 'A', 'A', 'B', 'A', 'A', 'B'),
  Strudent.new('Jack', 'C', 'C', 'F', 'C', 'C', 'C', 'C'),
  Strudent.new( 'Jim', 'C', 'D', 'D', 'F', 'C', 'B', 'D'),
  Strudent.new('Luis', 'A', 'A', 'D', 'A', 'A', 'B', 'B'),
  Strudent.new('Phil', 'B', 'A', 'F', 'D', 'D', 'B', 'F'),
  Strudent.new( 'Tom', 'C', 'F', 'C', 'B', 'B', 'D', 'C'),
]

doc.save(output_file) do |spreadsheet|

  sheet = doc.sheet["Period 1"]
  
  period_1.each_with_index do |student, i|
  
    sheet.record_assign student.to_h
    
    sheet.write_row i + 1
    
  end
  
end

puts "Wrote #{output_file} (1st period)"
