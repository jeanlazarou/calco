require 'date'

require 'calco'

#
# Next example shows how to extend the DSL with new built-in functions.
# It outputs as simple text the resulting spreadsheet (two rows: the headers
# and one row of data)
#

# Next functions does not exist in a spreadsheet software.
# we only show how to declare functions that are missing...
Calco::BuiltinFunction.declare :now, 0, Integer
Calco::BuiltinFunction.declare :age, 1, Integer

doc = spreadsheet do

  definitions do

    set some_date: Date.today

    function now: now
    function age: age(some_date)
    
  end
  
  sheet do

    column :now, :title => "Today"

    column value_of(:some_date)
    column :age,  :title => "Age"

  end

end

hash_printer = proc {|k, v| puts "#{k}: #{v}"}

doc.row_to_hash(0).each &hash_printer

doc.current[:some_date] = Date.new(1934, 10, 3)
puts
doc.row_to_hash(1).each &hash_printer
