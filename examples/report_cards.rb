require 'date'
require 'tmpdir'

#
# Next example uses the office engine.
#
# It saves an office file, in the temporary directory, named "res.ods" and uses
# a template file named "report_cards.ods"
#
# Shows writing several sheets + applying styles + appending a new sheet (named
# "Info")
#
# The example fills academic report cards, such reports contain students'
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
    
    set author: 'Author'
    set updated_at: 'Last update'
    
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
  
  sheet('Period 2') do

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

  sheet('Info') do
  
    column value_of(:author)
    column value_of(:updated_at)
    
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
period_2 = [
  Strudent.new('Alan', 'B', 'B', 'B', 'C', 'A', 'B', 'B'),
  Strudent.new('Greg', 'B', 'B', 'C', 'C', 'D', 'A', 'D'),
  Strudent.new('Iñes', 'A', 'A', 'A', 'B', 'A', 'A', 'B'),
  Strudent.new('Jack', 'D', 'B', 'B', 'C', 'B', 'C', 'D'),
  Strudent.new( 'Jim', 'B', 'F', 'D', 'D', 'B', 'B', 'C'),
  Strudent.new('Luis', 'A', 'A', 'B', 'B', 'A', 'B', 'C'),
  Strudent.new('Phil', 'A', 'A', 'D', 'B', 'D', 'B', 'F'),
  Strudent.new( 'Tom', 'C', 'D', 'D', 'D', 'D', 'F', 'B'),
]

doc.save(output_file) do |spreadsheet|

  sheet = doc.sheet["Period 1"]
  
  period_1.each_with_index do |student, i|
  
    sheet.record_assign student.to_h
    
    sheet.write_row i + 1
    
  end
  
  sheet = doc.sheet["Period 2"]
  
  period_2.each_with_index do |student, i|
  
    sheet.record_assign student.to_h
    
    sheet.write_row i + 1

  end
  
  sheet = doc.sheet["Info"]

  sheet.write_row 0
  
  sheet[:author] = 'Jean Lazarou'
  sheet[:updated_at] = Date.today.to_s
  sheet.write_row 1
  
end

puts "Wrote #{output_file} (3 sheets)"
