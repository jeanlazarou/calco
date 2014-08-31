require 'date'

require 'calco'

#
# Next example uses date functions and outputs to the console the result as 
# simple text.
#

doc = spreadsheet do

  definitions do

    set some_date: Date.today

    function some_year: year(some_date)
    function age: year(today) - year(some_date)
    
  end
  
  sheet do

    column value_of(:some_date)

    column :some_year
    column :age

  end

end

doc.save($stdout) do

  sheet = doc.current
  
  sheet[:some_date] = Date.new(1934, 10, 3)
  sheet.write_row 3
  
  sheet[:some_date] = Date.new(2004, 6, 19)
  sheet.write_row 5
  
end
